"""Service for NW Reports stored procedures and data retrieval."""

from __future__ import annotations

import html
from typing import Dict, List, Optional, Tuple

from app.config.db import get_connection_scope
from app.models.nw_reports_models import (
    NewlyCreatedName,
    OpenNote,
    ProjectAnalytics,
    RegionSpecificAnalytics,
    RetainedName,
    RootConcept,
    SummaryType,
    VoteByGroup,
    VotedParticipant,
    WordReportReplacement,
    WordReportResult,
)
from app.utils.logging_utils import get_logger
# OPTIMIZATION: Use centralized db_utils to reduce code duplication
from app.utils.db_utils import execute_sp_single_result, execute_sp_multiple_results, rows_to_dicts

logger = get_logger(__name__)


def _safe_fetch_sp_results(cursor, sp_name: str, presentation_id: int):
    """Safely fetch results from a stored procedure, handling None descriptions.

    Args:
        cursor: Database cursor after SP execution
        sp_name: Name of the stored procedure (for logging)
        presentation_id: Presentation ID (for logging)

    Returns:
        Tuple of (columns, rows) or (None, None) if no results
    """
    if cursor.description is None:
        logger.warning(
            "No result set returned from %s for presentation_id=%d",
            sp_name,
            presentation_id
        )
        return None, None

    columns = [column[0] for column in cursor.description]
    rows = cursor.fetchall()

    if not rows:
        logger.info(
            "No data returned from %s for presentation_id=%d",
            sp_name,
            presentation_id
        )
        return columns, []

    return columns, rows


class NWReportsService:
    """Service for handling NW Reports stored procedures."""

    def check_has_participants(self, presentation_id: int) -> int:
        """Check if presentation has participant voting enabled.

        OPTIMIZED: Now uses execute_sp_single_result from db_utils.

        Args:
            presentation_id: The presentation ID.

        Returns:
            1 if participant voting is enabled, 0 otherwise.
        """
        logger.debug(
            "Checking if presentation %d has participant voting enabled",
            presentation_id,
        )

        # OPTIMIZATION: Single line replaces 20+ lines of boilerplate code
        result = execute_sp_single_result(
            "[BI_GUIDELINES].[dbo].[nw_IsParticipantVoted]",
            (presentation_id,),
            timeout=30
        )

        if result:
            value = int(result.get("HasVoted", 0)) if "HasVoted" in result else int(list(result.values())[0] if result.values() else 0)
            logger.info("Presentation %d has participant voting: %s", presentation_id, value)
            return value

        return 0

    def is_participant_voted(
        self, presentation_id: int, participant_id: int
    ) -> bool:
        """Check if participant has voted using nw_IsParticipantVoted SP.

        Args:
            presentation_id: The presentation ID.
            participant_id: The participant ID.

        Returns:
            True if participant has voted, False otherwise.
        """
        logger.debug(
            "Checking if participant %d voted for presentation_id=%d",
            participant_id,
            presentation_id,
        )

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[nw_IsParticipantVoted](?, ?)}",
                (presentation_id, participant_id)
            )

            row = cursor.fetchone()
            if row:
                columns = [column[0] for column in cursor.description]
                row_dict = dict(zip(columns, row))
                has_voted = row_dict.get("HasVoted", False)
                logger.info(
                    "Participant %d voted status for presentation %d: %s",
                    participant_id,
                    presentation_id,
                    has_voted,
                )
                return bool(has_voted)

            return False

    # Word Report SPs

    def get_word_report_replacements(
        self, presentation_id: int
    ) -> List[WordReportReplacement]:
        """Get values to replace in Word template from nw_wdValuesToReplace SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of WordReportReplacement objects.
        """
        logger.debug(
            "Fetching Word report replacements for presentation_id=%d",
            presentation_id,
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_wdValuesToReplace](?)}",
                    (presentation_id,)
                )

                # This SP might return multiple result sets, iterate through all
                columns = None
                rows = None
                result_set_num = 0

                while True:
                    if cursor.description is not None:
                        result_set_num += 1
                        temp_columns = [column[0] for column in cursor.description]
                        temp_rows = cursor.fetchall()

                        logger.debug(
                            "nw_wdValuesToReplace result set %d: columns=%s, rows=%d",
                            result_set_num,
                            temp_columns,
                            len(temp_rows)
                        )

                        # Use the first result set that has data with 'key' and 'value' columns
                        if temp_rows and 'key' in temp_columns and 'value' in temp_columns:
                            columns = temp_columns
                            rows = temp_rows
                            logger.info("Found replacement data in result set #%d", result_set_num)
                            break

                    # Try to move to next result set
                    if not cursor.nextset():
                        break

                if columns is None or rows is None or not rows:
                    logger.warning(
                        "SP nw_wdValuesToReplace returned no data for presentation_id=%d "
                        "(processed %d result sets)",
                        presentation_id,
                        result_set_num
                    )
                    return []

                replacements = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    # Use 'key' and 'value' as column names
                    placeholder = row_dict.get("key", "")
                    value = row_dict.get("value", "")

                    logger.debug("Placeholder: '%s' -> Value: '%s'", placeholder, value)

                    replacements.append(
                        WordReportReplacement(
                            placeholder=placeholder,
                            value=value,
                        )
                    )

                logger.info(
                    "Retrieved %d Word report replacements for presentation_id=%d",
                    len(replacements),
                    presentation_id,
                )
                return replacements

        except Exception as e:
            logger.error(
                "Error fetching Word report replacements for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    def get_word_report_results_phonetics(
        self, presentation_id: int, summary_type: SummaryType
    ) -> List[WordReportResult]:
        """Get results for Word phonetics report from nw_wdGetResults_Phonetics SP.

        This SP returns MULTIPLE result sets. We need to iterate through all of them
        to find the one that matches our SummaryType.

        Args:
            presentation_id: The presentation ID.
            summary_type: Type of summary to retrieve.

        Returns:
            List of WordReportResult objects.
        """
        logger.debug(
            "Fetching Word phonetics results for presentation_id=%d, summary_type=%s",
            presentation_id,
            summary_type,
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_wdGetResults_Phonetics](?, ?)}",
                    (presentation_id, summary_type.value)
                )

                # This SP returns MULTIPLE result sets, we need to iterate through all
                columns = None
                rows = None
                result_set_num = 0
                target_columns_found = False

                while True:
                    if cursor.description is not None:
                        result_set_num += 1
                        temp_columns = [column[0] for column in cursor.description]
                        temp_rows = cursor.fetchall()

                        logger.debug(
                            "nw_wdGetResults_Phonetics result set %d: columns=%s, rows=%d",
                            result_set_num,
                            temp_columns,
                            len(temp_rows)
                        )

                        # For ByTheNumbers, look for result set with Positive, Neutral, Reconsider columns
                        if summary_type == SummaryType.BY_THE_NUMBERS:
                            if any(col in temp_columns for col in ['Positive', 'Neutral', 'Reconsider', 'TotNewNames']):
                                columns = temp_columns
                                rows = temp_rows
                                target_columns_found = True
                                logger.info("Found ByTheNumbers result set #%d", result_set_num)
                                break
                        # For other summary types, look for Name, NameRationale, NameCategory
                        elif 'Name' in temp_columns and ('NameRationale' in temp_columns or 'NameCategory' in temp_columns):
                            # Only use result sets that have actual data
                            if temp_rows:
                                columns = temp_columns
                                rows = temp_rows
                                target_columns_found = True
                                logger.info("Found data result set #%d with %d rows", result_set_num, len(temp_rows))
                                break  # Stop after finding first result set with data

                    # Try to move to next result set
                    if not cursor.nextset():
                        break

                if not target_columns_found or columns is None or rows is None:
                    logger.warning(
                        "No matching result set found from nw_wdGetResults_Phonetics for "
                        "presentation_id=%d, summary_type=%s (processed %d result sets)",
                        presentation_id,
                        summary_type.value,
                        result_set_num
                    )
                    return []

                results = []
                for row in rows:
                    row_dict = dict(zip(columns, row))

                    # For ByTheNumbers, structure is different
                    if summary_type == SummaryType.BY_THE_NUMBERS:
                        # Return a single result with the counts
                        results.append(
                            WordReportResult(
                                name="Summary",
                                positive_count=row_dict.get("Positive", 0),
                                neutral_count=row_dict.get("Neutral", 0),
                                reconsider_count=row_dict.get("Reconsider", 0),
                                new_names_count=row_dict.get("TotNewNames", 0),
                                total_count=row_dict.get("TotalNames", 0),
                            )
                        )
                    elif summary_type == SummaryType.EXPLORE:
                        # Refined Creative Direction table packs multiple values using $$ or ##
                        explore_value = row_dict.get("Explore") or row_dict.get("NameCategory") or ""
                        name_value = row_dict.get("Name", "") or ""
                        rationale_value = row_dict.get("NameRationale", "") or ""

                        # Detect delimiter used to pack multiple entries
                        delimiter = None
                        for candidate in (name_value, rationale_value):
                            if isinstance(candidate, str) and "##" in candidate:
                                delimiter = "##"
                                break
                        if delimiter is None:
                            for candidate in (name_value, rationale_value):
                                if isinstance(candidate, str) and "$$" in candidate:
                                    delimiter = "$$"
                                    break

                        names = []
                        rationales = []
                        if delimiter:
                            names = [n.strip() for n in str(name_value).split(delimiter) if n and n.strip()]
                            rationales = [r.strip() for r in str(rationale_value).split(delimiter) if r is not None]
                            # Keep list lengths aligned
                            while len(rationales) < len(names):
                                rationales.append(rationales[-1] if rationales else "")
                        else:
                            names = [name_value]
                            rationales = [rationale_value]

                        for idx, name_item in enumerate(names):
                            rationale_item = rationales[idx] if idx < len(rationales) else ""
                            results.append(
                                WordReportResult(
                                    name=name_item,
                                    direction=rationale_item,
                                    rationale=rationale_item,
                                    category=explore_value,
                                    original_rationale=explore_value,
                                )
                            )
                    else:
                        # Regular name results
                        # Columns from SP: Name, Pronunciation (phonetics), NameRationale, NameCategory
                        results.append(
                            WordReportResult(
                                name=row_dict.get("Name", ""),
                                pronunciation=row_dict.get("Pronunciation") or row_dict.get("NamePronunciation") or row_dict.get("Positive Pronounciation"),
                                category=row_dict.get("NameCategory", ""),
                                rationale=row_dict.get("NameRationale", ""),
                                vote=row_dict.get("Vote"),
                                name_rationale_part1=row_dict.get("NameRationalePart1"),
                                name_rationale_part2=row_dict.get("NameRationalePart2"),
                            )
                        )

                logger.info(
                    "Retrieved %d Word phonetics results for presentation_id=%d, summary_type=%s",
                    len(results),
                    presentation_id,
                    summary_type.value
                )
                return results

        except Exception as e:
            logger.error(
                "Error fetching Word phonetics results for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    def get_word_report_results(
        self, presentation_id: int, summary_type: SummaryType = SummaryType.BY_THE_NUMBERS
    ) -> List[WordReportResult]:
        """Get general results for Word report from nw_wdGetResults SP.

        Args:
            presentation_id: The presentation ID.
            summary_type: Type of summary to retrieve (default: BY_THE_NUMBERS).

        Returns:
            List of WordReportResult objects.
        """
        logger.debug(
            "Fetching Word report results for presentation_id=%d, summary_type=%s",
            presentation_id,
            summary_type,
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_wdGetResults](?, ?)}",
                    (presentation_id, summary_type.value)
                )

                columns, rows = _safe_fetch_sp_results(
                    cursor, "nw_wdGetResults", presentation_id
                )

                if columns is None or rows is None:
                    return []

                results = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    results.append(
                        WordReportResult(
                            name=row_dict.get("Name", ""),
                            category=row_dict.get("Category"),
                            rationale=row_dict.get("Rationale"),
                            vote=row_dict.get("Vote"),
                        )
                    )

                logger.info(
                    "Retrieved %d Word report results for presentation_id=%d",
                    len(results),
                    presentation_id,
                )
                return results

        except Exception as e:
            logger.error(
                "Error fetching Word report results for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    def get_newly_created_names_for_word(
        self, presentation_id: int
    ) -> List[WordReportResult]:
        """Get newly created names for Word report from nw_CombineNewNames SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of WordReportResult objects with newly created names.
        """
        logger.debug(
            "Fetching newly created names for Word report for presentation_id=%d",
            presentation_id,
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_CombineNewNames](?)}",
                    (presentation_id,)
                )

                # This SP may return multiple result sets, iterate through them
                columns = None
                rows = None
                result_set_num = 0

                while True:
                    if cursor.description is not None:
                        result_set_num += 1
                        temp_columns = [column[0] for column in cursor.description]
                        temp_rows = cursor.fetchall()

                        # Use result sets that have data OR if we have no data yet
                        if temp_rows or columns is None:
                            columns = temp_columns
                            rows = temp_rows
                            logger.debug("nw_CombineNewNames result set %d: %d columns, %d rows",
                                       result_set_num, len(columns), len(rows))

                    # Try to move to next result set
                    if not cursor.nextset():
                        break

                if columns is None or rows is None:
                    logger.warning(
                        "No result set from nw_CombineNewNames for presentation_id=%d",
                        presentation_id
                    )
                    return []

                # Find the name column index
                name_col_idx = None
                for idx, col in enumerate(columns):
                    col_lower = col.lower() if col else ''
                    if col_lower in ['name', 'newname']:
                        name_col_idx = idx
                        break

                results = []
                for row in rows:
                    row_dict = dict(zip(columns, row))

                    def clean_value(value):
                        if isinstance(value, str):
                            return html.unescape(value).strip()
                        return value

                    # Extract and clean fields
                    name_value = clean_value(row_dict.get("Name") or row_dict.get("NewName", ""))
                    original_name_value = clean_value(row_dict.get("NewName") or row_dict.get("OriginalName", ""))
                    category_value = clean_value(row_dict.get("NameCategory") or row_dict.get("Category"))
                    rationale_value = clean_value(row_dict.get("NameRationale") or row_dict.get("Rationale"))

                    # Detect delimiter used to pack multiple values
                    delimiter = None
                    for candidate in (name_value, original_name_value, category_value, rationale_value):
                        if isinstance(candidate, str) and "##" in candidate:
                            delimiter = "##"
                            break
                    if delimiter is None:
                        for candidate in (name_value, original_name_value, category_value, rationale_value):
                            if isinstance(candidate, str) and "$$" in candidate:
                                delimiter = "$$"
                                break

                    # Check if name contains group delimiters (## or $$)
                    if delimiter and name_value and isinstance(name_value, str) and delimiter in name_value:
                        # Split all fields by the delimiter
                        names = [n.strip() for n in str(name_value).split(delimiter) if n and n.strip()]

                        # Split category and rationale if they also contain the delimiter
                        if category_value and delimiter in str(category_value):
                            categories = [c.strip() for c in str(category_value).split(delimiter)]
                        else:
                            categories = [category_value] * len(names)

                        if rationale_value and delimiter in str(rationale_value):
                            rationales = [r.strip() for r in str(rationale_value).split(delimiter)]
                        else:
                            rationales = [rationale_value] * len(names)

                        if original_name_value and isinstance(original_name_value, str) and delimiter in original_name_value:
                            original_names = [n.strip() for n in str(original_name_value).split(delimiter)]
                        else:
                            original_names = [original_name_value] * len(names)

                        # Pad arrays to match length
                        while len(categories) < len(names):
                            categories.append(category_value)
                        while len(rationales) < len(names):
                            rationales.append(rationale_value)
                        while len(original_names) < len(names):
                            original_names.append(original_name_value)

                        # Create a result for each expanded name
                        for i, expanded_name in enumerate(names):
                            if expanded_name:  # Skip empty names
                                results.append(
                                    WordReportResult(
                                        name=expanded_name,
                                        original_name=original_names[i] if i < len(original_names) else original_name_value,
                                        category=categories[i] if i < len(categories) else category_value,
                                        rationale=rationales[i] if i < len(rationales) else rationale_value,
                                        vote=row_dict.get("Vote"),
                                    )
                                )
                    else:
                        # No grouping, add normally
                        results.append(
                            WordReportResult(
                                name=name_value,
                                original_name=original_name_value,
                                category=category_value,
                                rationale=rationale_value,
                                vote=row_dict.get("Vote"),
                            )
                        )

                # Split rationale into two parts when using "+" convention (keeps original category/rationale)
                processed_results = []
                for result in results:
                    if result.rationale and isinstance(result.rationale, str) and '+' in result.rationale:
                        parts = result.rationale.split('+', 1)
                        result.rationale = parts[0].strip()
                        result.name_rationale_part2 = parts[1].strip() if len(parts) > 1 else ""
                        if not result.category:
                            result.category = result.name_rationale_part2
                    processed_results.append(result)
                results = processed_results

                logger.info(
                    "Retrieved %d newly created names for Word report for presentation_id=%d",
                    len(results),
                    presentation_id,
                )
                return results

        except Exception as e:
            logger.error(
                "Error fetching newly created names for Word report for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    # Analytics SPs

    def get_project_analytics(self, presentation_id: int) -> List[ProjectAnalytics]:
        """Get project analytics from NW_ProjectAnalytics SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of ProjectAnalytics objects.
        """
        logger.debug(
            "Fetching project analytics for presentation_id=%d", presentation_id
        )

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[NW_ProjectAnalytics](?)}",
                (presentation_id,)
            )

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            analytics = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                analytics.append(
                    ProjectAnalytics(
                        metric_name=row_dict.get("MetricName", ""),
                        metric_value=str(row_dict.get("MetricValue", "")),
                    )
                )

            logger.info(
                "Retrieved %d project analytics for presentation_id=%d",
                len(analytics),
                presentation_id,
            )
            return analytics

    def get_presentation_id_by_name(self, display_name: str) -> Optional[int]:
        """Get presentation ID from display name using nw_GetPresentationId SP.

        Args:
            display_name: The display name of the presentation.

        Returns:
            Presentation ID if found, None otherwise.
        """
        logger.debug("Fetching presentation ID for display_name=%s", display_name)

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[nw_GetPresentationId](?)}",
                (display_name,)
            )

            row = cursor.fetchone()
            if row:
                columns = [column[0] for column in cursor.description]
                row_dict = dict(zip(columns, row))
                presentation_id = row_dict.get("PresentationId")
                logger.info(
                    "Found presentation_id=%s for display_name=%s",
                    presentation_id,
                    display_name,
                )
                return presentation_id

            logger.warning("No presentation found for display_name=%s", display_name)
            return None

    def get_region_specific_analytics(
        self, presentation_id: int
    ) -> List[RegionSpecificAnalytics]:
        """Get region-specific analytics from NW_RegionSpecificAnalytics SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of RegionSpecificAnalytics objects.
        """
        logger.debug(
            "Fetching region-specific analytics for presentation_id=%d",
            presentation_id,
        )

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[NW_RegionSpecificAnalytics](?)}",
                (presentation_id,)
            )

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            analytics = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                analytics.append(
                    RegionSpecificAnalytics(
                        region=row_dict.get("Region", ""),
                        metric_name=row_dict.get("MetricName", ""),
                        metric_value=str(row_dict.get("MetricValue", "")),
                    )
                )

            logger.info(
                "Retrieved %d region-specific analytics for presentation_id=%d",
                len(analytics),
                presentation_id,
            )
            return analytics

    # BSR SPs

    def get_bsr_excel_report(self, project_id: int) -> Dict:
        """Get BSR Excel report data from bsr_GetExcelReport SP.

        Args:
            project_id: The project ID.

        Returns:
            Dictionary containing BSR report data.
        """
        logger.debug("Fetching BSR Excel report for project_id=%d", project_id)

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[bsr_GetExcelReport](?)}",
                (project_id,)
            )

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            data = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                data.append(row_dict)

            logger.info(
                "Retrieved %d BSR Excel report rows for project_id=%d",
                len(data),
                project_id,
            )
            return {"data": data}

    def get_bsr_project_concepts(self, project_id: int) -> List[Dict]:
        """Get BSR project concepts from direct table query.

        Args:
            project_id: The project ID.

        Returns:
            List of concept dictionaries.
        """
        logger.debug("Fetching BSR project concepts for project_id=%d", project_id)

        with get_connection_scope(timeout=30) as cursor:
            query = """
                SELECT ConceptId, ConceptName, Description
                FROM [BI_GUIDELINES].[dbo].[bsr_ProjectConcepts]
                WHERE ProjectId = ?
            """
            cursor.execute(query, (project_id,))

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            concepts = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                concepts.append(row_dict)

            logger.info(
                "Retrieved %d BSR project concepts for project_id=%d",
                len(concepts),
                project_id,
            )
            return concepts

    def get_bsr_page_comments(self, project_id: int) -> List[Dict]:
        """Get BSR page comments from direct table query.

        Args:
            project_id: The project ID.

        Returns:
            List of comment dictionaries.
        """
        logger.debug("Fetching BSR page comments for project_id=%d", project_id)

        with get_connection_scope(timeout=30) as cursor:
            query = """
                SELECT CommentId, PageNumber, CommentText, CreatedBy
                FROM [BI_GUIDELINES].[dbo].[BSR_PageComments]
                WHERE ProjectId = ?
            """
            cursor.execute(query, (project_id,))

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            comments = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                comments.append(row_dict)

            logger.info(
                "Retrieved %d BSR page comments for project_id=%d",
                len(comments),
                project_id,
            )
            return comments


# Global service instance
nw_reports_service = NWReportsService()
