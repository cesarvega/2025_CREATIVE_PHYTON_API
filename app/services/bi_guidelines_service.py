"""Service for BI_GUIDELINES database operations."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from app.config.db import get_connection_scope
from app.models.response_models import ActivePresentation, TemplateGroup
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class BIGuidelinesService:
    """Service for handling BI_GUIDELINES database queries."""

    def get_active_presentations(
        self,
        search: Optional[str] = None,
        page: int = 1,
        limit: int = 50,
    ) -> Tuple[List[ActivePresentation], int]:
        """Retrieve paginated and filtered active presentations from BI_GUIDELINES.

        Args:
            search: Optional search term to filter by project or display name.
            page: Page number (1-indexed).
            limit: Number of results per page.

        Returns:
            Tuple containing (list of ActivePresentation objects, total count).

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If query execution fails.
        """
        offset = (page - 1) * limit

        # Build query with optional search filter
        base_query = """
            SELECT [PresentationId], [Project], [DisplayName], [UploadedBy],
                   [UploadedDate], [PresentationStatus], [LastUpdateDate]
            FROM [BI_GUIDELINES].[dbo].[nw_Master]
            WHERE [PresentationStatus] = 'OPEN'
        """

        if search:
            search_pattern = f"%{search}%"
            query = base_query + """
                AND ([Project] LIKE ? OR [DisplayName] LIKE ?)
                ORDER BY [LastUpdateDate] DESC
                OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
            """
            count_query = """
                SELECT COUNT(*) as Total
                FROM [BI_GUIDELINES].[dbo].[nw_Master]
                WHERE [PresentationStatus] = 'OPEN'
                AND ([Project] LIKE ? OR [DisplayName] LIKE ?)
            """
            query_params = (search_pattern, search_pattern, offset, limit)
            count_params = (search_pattern, search_pattern)
        else:
            query = base_query + """
                ORDER BY [LastUpdateDate] DESC
                OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
            """
            count_query = """
                SELECT COUNT(*) as Total
                FROM [BI_GUIDELINES].[dbo].[nw_Master]
                WHERE [PresentationStatus] = 'OPEN'
            """
            query_params = (offset, limit)
            count_params = ()

        logger.debug(
            "BI_GUIDELINES - Fetching active presentations: search=%s, page=%d, limit=%d, offset=%d",
            search or "(none)",
            page,
            limit,
            offset,
        )

        with get_connection_scope(timeout=30) as cursor:
            # Get total count
            cursor.execute(count_query, count_params)
            total_row = cursor.fetchone()
            total = total_row.Total if total_row else 0

            # Get paginated results
            cursor.execute(query, query_params)
            rows = cursor.fetchall()

            presentations = []
            for row in rows:
                # Build link from display name
                link = f"https://tools.brandinstitute.com/nw/#/main/{row.DisplayName}"

                presentations.append(
                    ActivePresentation(
                        presentation_id=row.PresentationId,
                        project=row.Project,
                        display_name=row.DisplayName,
                        uploaded_by=row.UploadedBy,
                        uploaded_date=row.UploadedDate,
                        link=link,
                        presentation_status=row.PresentationStatus,
                        last_update_date=row.LastUpdateDate,
                    )
                )

            logger.info(
                "BI_GUIDELINES - Retrieved %d active presentations (total: %d) for search='%s', page=%d",
                len(presentations),
                total,
                search or "",
                page,
            )

            return presentations, total

    def get_bsr_display_names(
        self,
        search: Optional[str] = None,
        page: int = 1,
        limit: int = 50,
    ) -> Tuple[List[str], int]:
        """Retrieve paginated and filtered BSR display names from BI_GUIDELINES.

        Args:
            search: Optional search term to filter display names.
            page: Page number (1-indexed).
            limit: Number of results per page.

        Returns:
            Tuple containing (list of display names, total count).

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If query execution fails.
        """
        offset = (page - 1) * limit

        # Build query with optional search filter
        # Note: Use bsr_master table to match the getBSRDisplayNames SP
        if search:
            search_pattern = f"%{search}%"
            query = """
                SELECT displayname
                FROM [BI_GUIDELINES].[dbo].[bsr_master]
                WHERE displayname LIKE ?
                ORDER BY displayname
                OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
            """
            count_query = """
                SELECT COUNT(*) as Total
                FROM [BI_GUIDELINES].[dbo].[bsr_master]
                WHERE displayname LIKE ?
            """
            query_params = (search_pattern, offset, limit)
            count_params = (search_pattern,)
        else:
            query = """
                SELECT displayname
                FROM [BI_GUIDELINES].[dbo].[bsr_master]
                ORDER BY displayname
                OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
            """
            count_query = """
                SELECT COUNT(*) as Total
                FROM [BI_GUIDELINES].[dbo].[bsr_master]
            """
            query_params = (offset, limit)
            count_params = ()

        logger.debug(
            "BI_GUIDELINES - Fetching BSR display names: search=%s, page=%d, limit=%d, offset=%d",
            search or "(none)",
            page,
            limit,
            offset,
        )

        with get_connection_scope(timeout=30) as cursor:
            # Get total count
            cursor.execute(count_query, count_params)
            total_row = cursor.fetchone()
            total = total_row.Total if total_row else 0

            # Get paginated results
            cursor.execute(query, query_params)
            rows = cursor.fetchall()
            display_names = [row.displayname for row in rows if row.displayname]

            logger.info(
                "BI_GUIDELINES - Retrieved %d BSR display names (total: %d) for search='%s', page=%d",
                len(display_names),
                total,
                search or "",
                page,
            )

            return display_names, total

    def get_project_info(self, project_id: int) -> dict | None:
        """Retrieve project information using the stored procedure.

        Args:
            project_id: The project ID to retrieve information for.

        Returns:
            Dictionary containing project details or None if not found.

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If query execution fails.
        """
        logger.debug("BI_GUIDELINES - Fetching project info for project_id=%d", project_id)

        with get_connection_scope(timeout=30) as cursor:
            # Execute stored procedure
            cursor.execute("{CALL [BI_GUIDELINES].[dbo].[nw_PresentationInfo_Nw2_apr2020](?)}", (project_id,))

            # Get column names and row
            columns = [column[0] for column in cursor.description]
            row = cursor.fetchone()

            if not row:
                logger.warning("BI_GUIDELINES - Project %d not found", project_id)
                return None

            # Convert row to dictionary
            project_details = dict(zip(columns, row))

            logger.info("BI_GUIDELINES - Retrieved project info for project_id=%d", project_id)
            return project_details

    def get_template_groups(self) -> Tuple[List[TemplateGroup], int]:
        """Retrieve template groups from BI_GUIDELINES database.

        Executes the getNW_TemplateGroups stored procedure to get
        available background templates grouped by category.

        Returns:
            Tuple containing (list of TemplateGroup objects, total count).

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If query execution fails.
        """
        logger.debug("BI_GUIDELINES - Fetching template groups")

        with get_connection_scope(timeout=30) as cursor:
            # Execute stored procedure
            cursor.execute("{CALL [BI_GUIDELINES].[dbo].[getNW_TemplateGroups]}")

            # Get column names and rows
            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            # Log columns for debugging
            logger.info("BI_GUIDELINES - Template groups columns: %s", columns)
            if rows:
                logger.info("BI_GUIDELINES - First row sample: %s", dict(zip(columns, rows[0])))

            template_groups = []
            for idx, row in enumerate(rows):
                # Convert row to dictionary for easier access
                row_dict = dict(zip(columns, row))

                # The SP might return multiple columns or a single concatenated column
                # Try to get individual columns first
                temp_group_value = row_dict.get("'TempGroup'") or row_dict.get('TempGroup')
                template_file_name = row_dict.get("'TemplateFileName'") or row_dict.get('TemplateFileName')

                # If we have a concatenated value, parse it
                if temp_group_value and not template_file_name:
                    # Parse the value if it contains delimited data (e.g., "Category~TemplateName~Path")
                    parts = str(temp_group_value).split('~')
                    category = parts[0] if len(parts) > 0 else None
                    template_name = parts[1] if len(parts) > 1 else str(temp_group_value)
                    template_file_name = parts[2] if len(parts) > 2 else None
                elif not temp_group_value and row_dict:
                    # Fallback: use first value if column name is unexpected
                    temp_group_value = list(row_dict.values())[0] if row_dict else None
                    parts = str(temp_group_value).split('~') if temp_group_value else []
                    category = parts[0] if len(parts) > 0 else None
                    template_name = parts[1] if len(parts) > 1 else str(temp_group_value) if temp_group_value else ''
                    template_file_name = parts[2] if len(parts) > 2 else None
                else:
                    # We have separate columns
                    if temp_group_value:
                        parts = str(temp_group_value).split('~')
                        category = parts[0] if len(parts) > 0 else None
                        template_name = parts[1] if len(parts) > 1 else str(temp_group_value)
                    else:
                        category = None
                        template_name = ''

                template_groups.append(
                    TemplateGroup(
                        template_group_id=idx + 1,  # Use index as ID since SP doesn't return one
                        template_name=template_name,
                        category=category,
                        template_file_name=template_file_name,
                    )
                )

            total = len(template_groups)

            logger.info(
                "BI_GUIDELINES - Retrieved %d template groups",
                total,
            )

            return template_groups, total

    def get_background_templates_by_names(self, template_names: List[str]) -> Dict[str, Dict[str, any]]:
        """Query nw_Templates table to get background info for given template names.

        Args:
            template_names: List of template names (e.g., ['BMW_1', 'BrandDNA', 'Kitchen2'])

        Returns:
            Dictionary mapping template_name to dict with template_id and template_file_name
            Example: {
                'BMW_1': {
                    'template_id': 123,
                    'template_file_name': 'images/BackGrounds/Backgrounds2019/BMW_1.jpg'
                }
            }

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If query execution fails.
        """
        if not template_names:
            logger.warning("BI_GUIDELINES - No template names provided")
            return {}

        logger.debug(
            "BI_GUIDELINES - Fetching background templates for names: %s",
            template_names
        )

        # Build parameterized query with IN clause
        placeholders = ','.join(['?'] * len(template_names))
        query = f"""
            SELECT templateid, TemplateName, TemplateFileName
            FROM [BI_GUIDELINES].[dbo].[nw_Templates]
            WHERE TemplateName IN ({placeholders})
        """

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(query, template_names)
            rows = cursor.fetchall()

            # Build dictionary mapping
            template_map = {}
            for row in rows:
                template_map[row.TemplateName] = {
                    'template_id': row.templateid,
                    'template_file_name': row.TemplateFileName
                }

            logger.info(
                "BI_GUIDELINES - Retrieved %d background templates out of %d requested",
                len(template_map),
                len(template_names),
            )

            # Log any missing templates
            missing = set(template_names) - set(template_map.keys())
            if missing:
                logger.warning(
                    "BI_GUIDELINES - Templates not found in database: %s",
                    list(missing)
                )

            return template_map


# Global service instance
bi_guidelines_service = BIGuidelinesService()
