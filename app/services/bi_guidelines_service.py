"""Service for BI_GUIDELINES database operations."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from app.config.db import get_connection_scope
from app.models.response_models import ActivePresentation, TemplateGroup
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class BIGuidelinesService:
    """Service for handling BI_GUIDELINES database queries."""

    def _get_ci(self, row_dict: dict, *candidates: str):
        """Case-insensitive getter for row dictionaries from pyodbc.

        Tries multiple candidate column names and returns the first match.
        """
        lowered = {str(k).lower(): v for k, v in row_dict.items()}
        for key in candidates:
            if key is None:
                continue
            v = lowered.get(str(key).lower())
            if v is not None:
                return v
        return None

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

    def get_bsr_active_presentations(
        self,
        search: Optional[str] = None,
        page: int = 1,
        limit: int = 50,
    ) -> Tuple[List[ActivePresentation], int]:
        """Retrieve paginated and filtered active BSR presentations via stored procedure.

        Executes [BI_GUIDELINES].[dbo].[BSR_ActivePresentations] and maps results to
        the same schema as ActivePresentation. The stored procedure returns all BSR
        presentations ordered by PresentationId DESC with a pre-generated BSR link.

        Since the SP takes no parameters, filtering by status='OPEN', search, and
        pagination are applied in Python to match the behavior of get_active_presentations.
        """
        logger.debug(
            "BI_GUIDELINES - Fetching BSR active presentations: search=%s, page=%d, limit=%d",
            search or "(none)",
            page,
            limit,
        )

        with get_connection_scope(timeout=30) as cursor:
            # Execute stored procedure to get BSR presentations
            # SP returns: PresentationId, Project, DisplayName, UploadedBy, UploadedDate,
            #             Link (pre-generated BSR link), PresentationStatus, LastUpdateDate
            cursor.execute("{CALL [BI_GUIDELINES].[dbo].[BSR_ActivePresentations]}")

            # Fetch all rows and column names
            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

        # Convert to dictionaries for easier processing and case-insensitive access
        row_dicts: List[dict] = [dict(zip(columns, row)) for row in rows]

        # Filter by PresentationStatus = 'OPEN' (similar to get_active_presentations)
        def is_open(d: dict) -> bool:
            status = self._get_ci(d, "PresentationStatus", "Status", "status")
            return str(status).upper() == "OPEN" if status is not None else False

        filtered = [d for d in row_dicts if is_open(d)]

        # Optional search on project or display name (case-insensitive, contains)
        if search:
            term = str(search).lower()
            def matches(d: dict) -> bool:
                project_val = self._get_ci(d, "Project")
                display_val = self._get_ci(d, "DisplayName", "displayname")
                project = str(project_val).lower() if project_val is not None else ""
                display = str(display_val).lower() if display_val is not None else ""
                return term in project or term in display
            filtered = [d for d in filtered if matches(d)]

        total = len(filtered)

        # Stable sort by last update date descending when available
        def last_update_key(d: dict):
            val = self._get_ci(d, "LastUpdateDate", "lastupdatedate", "Last_Update_Date")
            return (val is None, val)  # None sorts last

        filtered.sort(key=last_update_key, reverse=True)

        # Pagination
        start = max(0, (page - 1) * limit)
        end = start + limit
        page_rows = filtered[start:end]

        # Map to ActivePresentation
        presentations: List[ActivePresentation] = []
        for d in page_rows:
            presentation_id = self._get_ci(d, "PresentationId", "PresentationID", "presentationid")
            project = self._get_ci(d, "Project") or ""
            display_name = self._get_ci(d, "DisplayName", "displayname") or ""
            uploaded_by = self._get_ci(d, "UploadedBy", "uploadedby")
            uploaded_date = self._get_ci(d, "UploadedDate", "uploadeddate")
            presentation_status = self._get_ci(d, "PresentationStatus", "Status", "status") or "OPEN"
            last_update_date = self._get_ci(d, "LastUpdateDate", "lastupdatedate", "Last_Update_Date")

            # Use the Link column from the SP which already contains the BSR link
            # Format: https://tools.brandinstitute.com/bsr/#/main/{DisplayName}
            link = self._get_ci(d, "Link", "link") or f"https://tools.brandinstitute.com/bsr/#/main/{display_name}"

            presentations.append(
                ActivePresentation(
                    presentation_id=int(presentation_id) if presentation_id is not None else 0,
                    project=str(project),
                    display_name=str(display_name),
                    uploaded_by=str(uploaded_by) if uploaded_by is not None else None,
                    uploaded_date=uploaded_date,  # pyodbc returns datetime already
                    link=link,
                    presentation_status=str(presentation_status),
                    last_update_date=last_update_date,
                )
            )

        logger.info(
            "BI_GUIDELINES - Retrieved %d BSR active presentations (total: %d) for search='%s', page=%d",
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

    def get_bsr_project_info(self, project_id: int) -> dict | None:
        """Retrieve BSR project information using BSR_PresentationInfo stored procedure.

        This method executes the BSR_PresentationInfo stored procedure and optionally
        retrieves project categories if available.

        Args:
            project_id: The BSR project ID to retrieve information for.

        Returns:
            Dictionary containing BSR project details with categories or None if not found.
            Structure:
            {
                "project": str,
                "displayname": str,
                "uploadedby": str,
                "uploadeddate": datetime,
                "presentationid": int,
                "presentationtype": str,
                "presentationstatus": str,
                "link": str,
                "lastupdatedate": datetime,
                "slidenumber": int,
                "iswideppt": int,
                "categories": [
                    {
                        "category": str,
                        "elements": str
                    }
                ]
            }

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If query execution fails.
        """
        logger.debug("BI_GUIDELINES - Fetching BSR project info for project_id=%d", project_id)

        with get_connection_scope(timeout=30) as cursor:
            # Execute BSR_PresentationInfo stored procedure
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[BSR_PresentationInfo](?)}",
                (project_id,)
            )

            # Get column names and row
            columns = [column[0] for column in cursor.description]
            row = cursor.fetchone()

            if not row:
                logger.warning("BI_GUIDELINES - BSR Project %d not found", project_id)
                return None

            # Convert row to dictionary
            project_details = dict(zip(columns, row))

            # Get project ID for categories lookup
            presentation_id = project_details.get("presentationid") or project_id

            # Fetch categories for this project
            categories = self._get_bsr_project_categories(presentation_id, cursor)
            project_details["categories"] = categories

            logger.info(
                "BI_GUIDELINES - Retrieved BSR project info for project_id=%d with %d categories",
                project_id,
                len(categories)
            )
            return project_details

    def _get_bsr_project_categories(
        self, project_id: int, cursor=None
    ) -> List[dict]:
        """Retrieve BSR project categories using bsr_GetCategoryValues stored procedure.

        Args:
            project_id: The BSR project ID to retrieve categories for.
            cursor: Optional existing cursor to reuse connection.

        Returns:
            List of category dictionaries with structure:
            [
                {
                    "category": "Category Name",
                    "elements": "Element1, Element2, ..."
                }
            ]

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If query execution fails.
        """
        logger.debug("BI_GUIDELINES - Fetching BSR categories for project_id=%d", project_id)

        def _fetch_categories(cursor):
            # Execute bsr_GetCategoryValues stored procedure
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[bsr_GetCategoryValues](?)}",
                (project_id,)
            )

            # Get column names and rows
            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            categories = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                # Map database column names to response format
                category_item = {
                    "category": row_dict.get("category") or "",
                    "elements": row_dict.get("CategoriesElements") or "",
                }
                categories.append(category_item)

            logger.debug(
                "BI_GUIDELINES - Retrieved %d categories for project_id=%d",
                len(categories),
                project_id
            )
            return categories

        # If cursor provided, use it; otherwise create new connection
        if cursor:
            return _fetch_categories(cursor)
        else:
            with get_connection_scope(timeout=30) as new_cursor:
                return _fetch_categories(new_cursor)

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

    def reload_project_sounds(self, display_name: str) -> None:
        """Reload MP3 file paths for a NW project by executing the stored procedure.

        Executes [BI_GUIDELINES].[dbo].[NW_UpdateMP3FilePath] to resynchronize
        audio file paths for a project based on existing files in the cloud.

        Args:
            display_name: The display name of the project to reload sounds for.

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If stored procedure execution fails.
        """
        if not display_name or not display_name.strip():
            raise ValueError("Display name cannot be empty")

        display_name = display_name.strip()

        logger.debug(
            "BI_GUIDELINES - Reloading project sounds for display_name=%s",
            display_name
        )

        with get_connection_scope(timeout=30) as cursor:
            # Execute stored procedure to update MP3 file paths
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[NW_UpdateMP3FilePath](?)}",
                (display_name,)
            )
            
            # Commit the transaction
            cursor.connection.commit()

            logger.info(
                "BI_GUIDELINES - Successfully reloaded sounds for project: %s",
                display_name
            )

    def update_project_details(
        self,
        presentation_id: int,
        display_name: str,
        presentation_status: str,
        bsr_display_name: Optional[str] = None,
    ) -> None:
        """Update NW project presentation master details by executing the stored procedure.

        Executes [BI_GUIDELINES].[dbo].[nw_UpdatePresentationMaster] to modify
        the master details of an existing NW presentation. The LastUpdateDate is
        automatically updated in the database.

        Args:
            presentation_id: ID of the presentation to update.
            display_name: The new display name for the NW project.
            presentation_status: The new presentation status (e.g., 'OPEN', 'CLOSED').
            bsr_display_name: The BSR display name associated with this presentation (optional).

        Raises:
            ValueError: If display name is already in use by another presentation.
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If stored procedure execution fails.
        """
        logger.debug(
            "BI_GUIDELINES - Updating project presentation_id=%d, display_name=%s, status=%s",
            presentation_id,
            display_name,
            presentation_status
        )

        with get_connection_scope(timeout=30) as cursor:
            # First, verify the presentation exists and get current display name
            cursor.execute(
                "SELECT DisplayName, PresentationId FROM [BI_GUIDELINES].[dbo].[nw_Master] WHERE PresentationId = ?",
                (presentation_id,)
            )
            current_record = cursor.fetchone()

            if not current_record:
                raise ValueError(f"Presentation with ID {presentation_id} not found")

            current_display_name = current_record.DisplayName

            # If display name is changing, check if the new name is already in use
            if display_name != current_display_name:
                logger.debug(
                    "BI_GUIDELINES - Display name changing from '%s' to '%s', checking availability",
                    current_display_name,
                    display_name
                )

                # Check if the new display name is already in use by another presentation
                cursor.execute(
                    "SELECT PresentationId, Project FROM [BI_GUIDELINES].[dbo].[nw_Master] WHERE DisplayName = ?",
                    (display_name,)
                )
                existing_record = cursor.fetchone()

                if existing_record and existing_record.PresentationId != presentation_id:
                    raise ValueError(
                        f"Display name '{display_name}' is already in use by project '{existing_record.Project}' "
                        f"(PresentationId: {existing_record.PresentationId})"
                    )

            # Execute stored procedure to update presentation master details
            # SP signature: nw_UpdatePresentationMaster @PresentationId, @DisplayName,
            #               @PresentationStatus, @BsrDisplayName
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[nw_UpdatePresentationMaster](?, ?, ?, ?)}",
                (
                    presentation_id,
                    display_name,
                    presentation_status,
                    bsr_display_name,
                )
            )

            logger.info(
                "BI_GUIDELINES - Successfully updated project presentation_id=%d, display_name=%s, status=%s",
                presentation_id,
                display_name,
                presentation_status
            )


# Global service instance
bi_guidelines_service = BIGuidelinesService()
