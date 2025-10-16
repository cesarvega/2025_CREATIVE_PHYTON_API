"""Service for NW Reports stored procedures and data retrieval."""

from __future__ import annotations

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

    def get_retained_names(
        self, presentation_id: int, include_recraft: bool = True
    ) -> List[RetainedName]:
        """Get retained names from nw_dlRetainedNames_withRecraft SP.

        Args:
            presentation_id: The presentation ID.
            include_recraft: Whether to include recraft option.

        Returns:
            List of RetainedName objects.
        """
        logger.debug(
            "Fetching retained names for presentation_id=%d", presentation_id
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_dlRetainedNames_withRecraft](?)}",
                    (presentation_id,)
                )

                columns, rows = _safe_fetch_sp_results(
                    cursor, "nw_dlRetainedNames_withRecraft", presentation_id
                )

                if columns is None or rows is None:
                    return []

                retained_names = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    retained_names.append(
                        RetainedName(
                            name=row_dict.get("Name", ""),
                            category=row_dict.get("Category"),
                            rationale=row_dict.get("Rationale"),
                            vote=row_dict.get("Vote"),
                            recraft_flag=row_dict.get("RecraftFlag", False),
                        )
                    )

                logger.info(
                    "Retrieved %d retained names for presentation_id=%d",
                    len(retained_names),
                    presentation_id,
                )
                return retained_names

        except Exception as e:
            logger.error(
                "Error fetching retained names for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            # Return empty list instead of raising to allow report generation to continue
            return []

    def get_newly_created_names(self, presentation_id: int) -> List[NewlyCreatedName]:
        """Get newly created names from nw_dlNewlyCreatedNames SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of NewlyCreatedName objects.
        """
        logger.debug(
            "Fetching newly created names for presentation_id=%d", presentation_id
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_dlNewlyCreatedNames](?)}",
                    (presentation_id,)
                )

                columns, rows = _safe_fetch_sp_results(
                    cursor, "nw_dlNewlyCreatedNames", presentation_id
                )

                if columns is None or rows is None:
                    return []

                new_names = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    new_names.append(
                        NewlyCreatedName(
                            name=row_dict.get("Name", ""),
                            category=row_dict.get("Category"),
                            rationale=row_dict.get("Rationale"),
                        )
                    )

                logger.info(
                    "Retrieved %d newly created names for presentation_id=%d",
                    len(new_names),
                    presentation_id,
                )
                return new_names

        except Exception as e:
            logger.error(
                "Error fetching newly created names for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    def get_roots_to_explore(self, presentation_id: int) -> List[RootConcept]:
        """Get roots/concepts to explore from nw_dlRootsOrConceptsToExplore SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of RootConcept objects.
        """
        logger.debug(
            "Fetching roots to explore for presentation_id=%d", presentation_id
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_dlRootsOrConceptsToExplore](?)}",
                    (presentation_id,)
                )

                columns, rows = _safe_fetch_sp_results(
                    cursor, "nw_dlRootsOrConceptsToExplore", presentation_id
                )

                if columns is None or rows is None:
                    return []

                concepts = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    concepts.append(
                        RootConcept(
                            concept=row_dict.get("Concept", ""),
                            description=row_dict.get("Description"),
                        )
                    )

                logger.info(
                    "Retrieved %d roots to explore for presentation_id=%d",
                    len(concepts),
                    presentation_id,
                )
                return concepts

        except Exception as e:
            logger.error(
                "Error fetching roots to explore for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    def get_roots_to_avoid(self, presentation_id: int) -> List[RootConcept]:
        """Get roots/concepts to avoid from nw_dlRootsOrConceptsToAvoid SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of RootConcept objects.
        """
        logger.debug("Fetching roots to avoid for presentation_id=%d", presentation_id)

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_dlRootsOrConceptsToAvoid](?)}",
                    (presentation_id,)
                )

                columns, rows = _safe_fetch_sp_results(
                    cursor, "nw_dlRootsOrConceptsToAvoid", presentation_id
                )

                if columns is None or rows is None:
                    return []

                concepts = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    concepts.append(
                        RootConcept(
                            concept=row_dict.get("Concept", ""),
                            description=row_dict.get("Description"),
                        )
                    )

                logger.info(
                    "Retrieved %d roots to avoid for presentation_id=%d",
                    len(concepts),
                    presentation_id,
                )
                return concepts

        except Exception as e:
            logger.error(
                "Error fetching roots to avoid for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    def get_open_notes(self, presentation_id: int) -> List[OpenNote]:
        """Get open notes from nw_dlOpenNotes SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of OpenNote objects.
        """
        logger.debug("Fetching open notes for presentation_id=%d", presentation_id)

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_dlOpenNotes](?)}",
                    (presentation_id,)
                )

                columns, rows = _safe_fetch_sp_results(
                    cursor, "nw_dlOpenNotes", presentation_id
                )

                if columns is None or rows is None:
                    return []

                notes = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    notes.append(
                        OpenNote(
                            note_id=row_dict.get("NoteId", 0),
                            note_text=row_dict.get("NoteText", ""),
                            created_by=row_dict.get("CreatedBy"),
                            created_date=row_dict.get("CreatedDate"),
                        )
                    )

                logger.info(
                    "Retrieved %d open notes for presentation_id=%d",
                    len(notes),
                    presentation_id,
                )
                return notes

        except Exception as e:
            logger.error(
                "Error fetching open notes for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    def get_votes_by_groups(self, presentation_id: int) -> List[VoteByGroup]:
        """Get votes by groups from nw_Votesbygroups SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of VoteByGroup objects.
        """
        logger.debug(
            "Fetching votes by groups for presentation_id=%d", presentation_id
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_Votesbygroups](?)}",
                    (presentation_id,)
                )

                columns, rows = _safe_fetch_sp_results(
                    cursor, "nw_Votesbygroups", presentation_id
                )

                if columns is None or rows is None:
                    return []

                votes = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    votes.append(
                        VoteByGroup(
                            group_name=row_dict.get("GroupName", ""),
                            positive_votes=row_dict.get("PositiveVotes", 0),
                            neutral_votes=row_dict.get("NeutralVotes", 0),
                            negative_votes=row_dict.get("NegativeVotes", 0),
                            total_votes=row_dict.get("TotalVotes", 0),
                        )
                    )

                logger.info(
                    "Retrieved %d vote groups for presentation_id=%d",
                    len(votes),
                    presentation_id,
                )
                return votes

        except Exception as e:
            logger.error(
                "Error fetching votes by groups for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    def get_voted_participants(self, presentation_id: int) -> List[VotedParticipant]:
        """Get voted participants from nw_VotedParticipants SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of VotedParticipant objects.
        """
        logger.debug(
            "Fetching voted participants for presentation_id=%d", presentation_id
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_VotedParticipants](?)}",
                    (presentation_id,)
                )

                columns, rows = _safe_fetch_sp_results(
                    cursor, "nw_VotedParticipants", presentation_id
                )

                if columns is None or rows is None:
                    return []

                participants = []
                for row in rows:
                    row_dict = dict(zip(columns, row))
                    participants.append(
                        VotedParticipant(
                            participant_id=row_dict.get("ParticipantId", 0),
                            participant_name=row_dict.get("ParticipantName", ""),
                            email=row_dict.get("Email"),
                            voted_date=row_dict.get("VotedDate"),
                        )
                    )

                logger.info(
                    "Retrieved %d voted participants for presentation_id=%d",
                    len(participants),
                    presentation_id,
                )
                return participants

        except Exception as e:
            logger.error(
                "Error fetching voted participants for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return []

    def check_has_participants(self, presentation_id: int) -> int:
        """Check if presentation has participant voting enabled.

        Args:
            presentation_id: The presentation ID.

        Returns:
            1 if participant voting is enabled, 0 otherwise.
        """
        logger.debug(
            "Checking if presentation %d has participant voting enabled",
            presentation_id,
        )

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_IsParticipantVoted](?)}",
                    (presentation_id,)
                )

                row = cursor.fetchone()
                if row:
                    # SP returns 1 or 0
                    result = int(row[0]) if row[0] is not None else 0
                    logger.info(
                        "Presentation %d has participant voting: %s",
                        presentation_id,
                        result,
                    )
                    return result

                return 0

        except Exception as e:
            logger.error(
                "Error checking participant voting for presentation_id=%d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
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

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[nw_wdValuesToReplace](?)}",
                (presentation_id,)
            )

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            replacements = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                replacements.append(
                    WordReportReplacement(
                        placeholder=row_dict.get("Placeholder", ""),
                        value=row_dict.get("Value", ""),
                    )
                )

            logger.info(
                "Retrieved %d Word report replacements for presentation_id=%d",
                len(replacements),
                presentation_id,
            )
            return replacements

    def get_word_report_results_phonetics(
        self, presentation_id: int, summary_type: SummaryType
    ) -> List[WordReportResult]:
        """Get results for Word phonetics report from nw_wdGetResults_Phonetics SP.

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

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[nw_wdGetResults_Phonetics](?, ?)}",
                (presentation_id, summary_type.value)
            )

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            results = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                results.append(
                    WordReportResult(
                        name=row_dict.get("Name", ""),
                        category=row_dict.get("Category"),
                        rationale=row_dict.get("Rationale"),
                        vote=row_dict.get("Vote"),
                        name_rationale_part1=row_dict.get("NameRationalePart1"),
                        name_rationale_part2=row_dict.get("NameRationalePart2"),
                    )
                )

            logger.info(
                "Retrieved %d Word phonetics results for presentation_id=%d",
                len(results),
                presentation_id,
            )
            return results

    def get_word_report_results(self, presentation_id: int) -> List[WordReportResult]:
        """Get general results for Word report from nw_wdGetResults SP.

        Args:
            presentation_id: The presentation ID.

        Returns:
            List of WordReportResult objects.
        """
        logger.debug(
            "Fetching Word report results for presentation_id=%d", presentation_id
        )

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[nw_wdGetResults](?)}",
                (presentation_id,)
            )

            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

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
