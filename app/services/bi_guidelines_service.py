"""Service for BI_GUIDELINES database operations."""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from app.config.db import get_connection_scope
from app.config.settings import settings
from app.models.response_models import ActivePresentation, TemplateGroup
from app.utils.logging_utils import get_logger
from app.utils.path_utils import sanitize_folder_name

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
                # FIXED: Build correct link - NW presentations use nw.bipresents.com
                link = f"https://nw.bipresents.com/{row.DisplayName}"

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

            # FIXED: Always generate correct link - ignore old Link from SP
            # This endpoint is specifically for BSR/NSR presentations
            # Check for PresentationType field from SP (BSR, NSR, BSR-Japan, NSR-Japan, NW, DW)
            presentation_type = self._get_ci(d, "PresentationType", "presentationtype", "Type")

            # Generate link with first 2 letters of project + presentation_id
            project_prefix = project[:2].lower() if project and len(project) >= 2 else ""
            link_id = f"{project_prefix}{presentation_id}" if project_prefix else str(presentation_id)
            
            # Generate appropriate link based on type
            if presentation_type:
                ptype = str(presentation_type).upper()
                
                if "NSR" in ptype:
                    # NSR or NSR-Japan -> https://bipresents.com/{project_prefix}{presentation_id}
                    link = f"https://bipresents.com/{link_id}"
                elif "BSR" in ptype:
                    # BSR or BSR-Japan -> https://bipresents.com/{project_prefix}{presentation_id}
                    link = f"https://bipresents.com/{link_id}"
                elif ptype in ["NW", "DM"]:
                    # NW or DayMaster -> https://nw.bipresents.com/{display_name}
                    link = f"https://nw.bipresents.com/{display_name}"
                else:
                    # Unknown type: default to bipresents
                    link = f"https://bipresents.com/{link_id}"
            else:
                # CRITICAL FIX: SP doesn't provide PresentationType, but this endpoint is for BSR/NSR
                # Always use project_prefix + presentation_id format
                link = f"https://bipresents.com/{link_id}"

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

            # DEBUG: Log what columns the stored procedure actually returns
            logger.info(
                "BSR_PresentationInfo SP returned columns for project_id=%d: %s",
                project_id,
                list(project_details.keys())
            )

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

    def update_bsr_presentation(
        self,
        *,
        presentation_id: int,
        display_name: str,
        status: str
    ) -> None:
        """Update BSR presentation master details.

        Executes the BSR_UpdatePresentationMaster stored procedure to update
        an existing BSR presentation's display name and status.

        Args:
            presentation_id: The presentation ID to update
            display_name: New display name
            status: New status (OPEN, CLOSED, etc.)

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If update fails.
        """
        logger.debug(
            "BI_GUIDELINES - Updating BSR presentation %d with display_name='%s', status='%s'",
            presentation_id,
            display_name,
            status
        )

        with get_connection_scope(timeout=30) as cursor:
            # First, get the current display name to check if it's changing
            cursor.execute(
                "SELECT DisplayName FROM [BI_GUIDELINES].[dbo].[bsr_Master] WHERE PresentationId = ?",
                (presentation_id,)
            )
            current_record = cursor.fetchone()

            if not current_record:
                raise ValueError(f"BSR Presentation with ID {presentation_id} not found")

            current_display_name = current_record.DisplayName

            # If display name is changing, rename the physical folder
            if display_name != current_display_name:
                logger.debug(
                    "BI_GUIDELINES - BSR display name changing from '%s' to '%s'",
                    current_display_name,
                    display_name
                )

                # Rename the physical folder before updating the database
                folder_renamed = self._rename_project_folder(
                    old_display_name=current_display_name,
                    new_display_name=display_name,
                    project_type="bsr"
                )

                if not folder_renamed:
                    logger.warning(
                        "Failed to rename BSR folder from '%s' to '%s', but continuing with database update",
                        current_display_name,
                        display_name
                    )

                # Update slide image paths in the database
                self._update_slide_image_paths(
                    cursor=cursor,
                    presentation_id=presentation_id,
                    old_display_name=current_display_name,
                    new_display_name=display_name,
                    project_type="bsr"
                )

            # Execute BSR_UpdatePresentationMaster stored procedure
            cursor.execute(
                "{CALL [BI_GUIDELINES].[dbo].[BSR_UpdatePresentationMaster](?, ?, ?)}",
                (presentation_id, display_name, status)
            )

            # Commit the transaction
            cursor.connection.commit()

            logger.info(
                "BI_GUIDELINES - Successfully updated BSR presentation %d",
                presentation_id
            )

    def get_template_groups(self) -> Tuple[List[TemplateGroup], int, int, int]:
        """Retrieve template groups from BI_GUIDELINES database.

        Executes the getNW_TemplateGroups stored procedure to get
        available background templates grouped by category.

        Returns:
            Tuple containing (list of TemplateGroup objects, total count, custom_count, system_count).
            Note: custom_count is always 0 (legacy field for compatibility).

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If query execution fails.
        """
        logger.debug("BI_GUIDELINES - Fetching template groups")

        with get_connection_scope(timeout=30) as cursor:
            # Execute stored procedure for system templates
            # SP now returns: TemplateId, TemplateName, TemplateGroup, TemplateFileName
            cursor.execute("{CALL [BI_GUIDELINES].[dbo].[getNW_TemplateGroups]}")

            rows = cursor.fetchall()

            template_groups = []
            for row in rows:
                # Construct preview URL and thumbnail URL from template file name
                # DB path: "images/BackGrounds/BMW_blue.jpg"
                # Full URL: "https://tools.brandinstitute.com/nw2/assets/images/BackGrounds/BMW_blue.jpg"
                # Thumbnail: "https://tools.brandinstitute.com/nw2/assets/images/BackGrounds/thumbnails/BMW_blue.jpg"
                preview_url = None
                thumbnail_url = None
                
                if row.TemplateFileName:
                    # Check if running in production
                    from app.config.settings import settings
                    from pathlib import Path
                    from app.utils.image_processor import get_thumbnail_filename
                    
                    if settings.environment == "production":
                        base_url = "https://tools.brandinstitute.com/nw2/assets/"
                    else:
                        # For development, point to production server since images are stored there
                        base_url = "https://tools.brandinstitute.com/nw2/assets/"
                    
                    # Full resolution image
                    # Normalize path separators to forward slashes for URLs
                    preview_url = f"{base_url}{row.TemplateFileName.replace(chr(92), '/')}"
                    
                    # Thumbnail (300x169px for fast loading)
                    # Maintain subdirectory structure in thumbnails
                    # Example: images/BackGrounds/Backgrounds2019/BlueCrag.jpg
                    #       -> images/BackGrounds/Backgrounds2019/thumbnails/BlueCrag.jpg
                    path_parts = Path(row.TemplateFileName)
                    thumbnail_filename = get_thumbnail_filename(path_parts.name)
                    
                    # Build thumbnail path maintaining subdirectory structure
                    # If path is: images/BackGrounds/Backgrounds2019/file.jpg
                    # Parent is: images/BackGrounds/Backgrounds2019
                    # Thumbnail: images/BackGrounds/Backgrounds2019/thumbnails/file.jpg
                    if path_parts.parent and str(path_parts.parent) != '.':
                        thumbnail_path = f"{path_parts.parent}/thumbnails/{thumbnail_filename}"
                    else:
                        thumbnail_path = f"thumbnails/{thumbnail_filename}"
                    
                    # Normalize to forward slashes for URL
                    thumbnail_url = f"{base_url}{thumbnail_path.replace(chr(92), '/')}"
                
                template_groups.append(
                    TemplateGroup(
                        template_group_id=row.TemplateId,
                        template_name=row.TemplateName,
                        category=row.TemplateGroup,
                        template_file_name=row.TemplateFileName,
                        thumbnail_url=thumbnail_url,
                        preview_url=preview_url
                    )
                )

            system_count = len(template_groups)
            total = system_count
            custom_count = 0

            logger.info(
                "BI_GUIDELINES - Retrieved %d template groups",
                total
            )

            return template_groups, total, custom_count, system_count

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

    def _rename_project_folder(
        self,
        old_display_name: str,
        new_display_name: str,
        project_type: str = "nw",
    ) -> bool:
        """Rename physical project folder when display name changes.

        Args:
            old_display_name: Previous display name
            new_display_name: New display name
            project_type: 'nw' for NW_Files/nw2, 'bsr' for NW_Files/bipresents

        Returns:
            True if folder was renamed successfully or didn't exist, False if error occurred

        Note:
            This method logs errors but doesn't raise exceptions to avoid blocking
            database updates if folder operations fail.
        """
        try:
            # Sanitize folder names
            old_folder_name = sanitize_folder_name(old_display_name)
            new_folder_name = sanitize_folder_name(new_display_name)

            # If folder names are the same after sanitization, no need to rename
            if old_folder_name == new_folder_name:
                logger.debug(
                    "Folder names are identical after sanitization ('%s'), skipping rename",
                    old_folder_name
                )
                return True

            # Determine base directory based on project type
            if project_type.lower() in ("bsr", "bipresents"):
                base_dir = settings.get_base_dir_for_project_type("bipresents")
            else:
                base_dir = settings.get_base_dir_for_project_type("nw")

            old_path = base_dir / old_folder_name
            new_path = base_dir / new_folder_name

            # Check if old folder exists
            if not old_path.exists():
                logger.warning(
                    "Old folder does not exist, cannot rename: %s",
                    old_path
                )
                # Return True because this is not a critical error - folder might not have been created yet
                return True

            # Check if new folder already exists
            if new_path.exists():
                logger.error(
                    "Cannot rename folder: destination already exists. Old: %s, New: %s",
                    old_path,
                    new_path
                )
                return False

            # Rename the folder
            shutil.move(str(old_path), str(new_path))

            logger.info(
                "Successfully renamed project folder from '%s' to '%s' (type: %s)",
                old_folder_name,
                new_folder_name,
                project_type
            )
            return True

        except Exception as e:
            logger.error(
                "Error renaming project folder from '%s' to '%s': %s",
                old_display_name,
                new_display_name,
                str(e),
                exc_info=True
            )
            return False

    def _update_slide_image_paths(
        self,
        cursor,
        presentation_id: int,
        old_display_name: str,
        new_display_name: str,
        project_type: str = "nw",
    ) -> int:
        """Update SlideBGFileName paths in detail tables when display name changes.

        Args:
            cursor: Database cursor
            presentation_id: ID of the presentation
            old_display_name: Previous display name
            new_display_name: New display name
            project_type: 'nw' or 'bsr'

        Returns:
            Number of rows updated
        """
        try:
            old_folder = sanitize_folder_name(old_display_name)
            new_folder = sanitize_folder_name(new_display_name)

            if old_folder == new_folder:
                logger.debug("Folder names identical after sanitization, no path updates needed")
                return 0

            # Determine table and slide path pattern based on project type
            if project_type.lower() in ("bsr", "bipresents"):
                table_name = "bsr_Details"
            else:
                table_name = "nw_Details"

            # First, check what URLs actually exist in the database
            check_sql = f"""
                SELECT TOP 5 SlideNumber, SlideBGFileName
                FROM [BI_GUIDELINES].[dbo].[{table_name}]
                WHERE PresentationId = ?
                ORDER BY SlideNumber
            """
            cursor.execute(check_sql, (presentation_id,))
            sample_rows = cursor.fetchall()
            
            if sample_rows:
                logger.info(
                    "Sample URLs for presentation %d (first 5 slides):",
                    presentation_id
                )
                for row in sample_rows:
                    logger.info("  Slide %d: %s", row.SlideNumber, row.SlideBGFileName)
            else:
                logger.warning("No slides found for presentation_id=%d", presentation_id)
                return 0

            # Extract the actual folder name from the first URL in the database
            # Format in DB: nw_slides/FOLDER_NAME/001.jpg or bsr_slides/FOLDER_NAME/001.jpg
            first_url = sample_rows[0].SlideBGFileName if sample_rows else ""
            
            if not first_url:
                logger.warning("No valid URL found for presentation_id=%d", presentation_id)
                return 0

            # Determine the slide root and extract actual folder from DB
            if project_type.lower() in ("bsr", "bipresents"):
                slide_root = "bsr_slides"
            else:
                slide_root = "nw_slides"

            # Extract the actual folder name from the URL
            # Example: "nw_slides/TEST_NW_11_07_2025_TEST1/001.jpg" -> "TEST_NW_11_07_2025_TEST1"
            if slide_root in first_url:
                parts = first_url.split('/')
                try:
                    root_index = parts.index(slide_root)
                    actual_old_folder = parts[root_index + 1] if len(parts) > root_index + 1 else None
                except (ValueError, IndexError):
                    actual_old_folder = None
            else:
                actual_old_folder = None

            if not actual_old_folder:
                logger.warning(
                    "Could not extract folder name from URL: %s",
                    first_url
                )
                return 0

            logger.info(
                "Detected actual folder in database: '%s', updating to: '%s'",
                actual_old_folder,
                new_folder
            )

            # Build patterns WITHOUT leading slash (matching the DB format)
            old_pattern = f"{slide_root}/{actual_old_folder}/"
            new_pattern = f"{slide_root}/{new_folder}/"

            # Update all slides for this presentation that contain the old folder name
            update_sql = f"""
                UPDATE [BI_GUIDELINES].[dbo].[{table_name}]
                SET SlideBGFileName = REPLACE(SlideBGFileName, ?, ?)
                WHERE PresentationId = ?
                AND SlideBGFileName LIKE ?
            """

            cursor.execute(
                update_sql,
                (old_pattern, new_pattern, presentation_id, f"%{old_pattern}%")
            )

            rows_affected = cursor.rowcount

            logger.info(
                "Updated %d slide image paths from '%s' to '%s' (presentation_id=%d, type=%s)",
                rows_affected,
                old_pattern,
                new_pattern,
                presentation_id,
                project_type
            )

            return rows_affected

        except Exception as e:
            logger.error(
                "Error updating slide image paths for presentation %d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            return 0

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

                # Rename the physical folder before updating the database
                folder_renamed = self._rename_project_folder(
                    old_display_name=current_display_name,
                    new_display_name=display_name,
                    project_type="nw"
                )

                if not folder_renamed:
                    logger.warning(
                        "Failed to rename folder from '%s' to '%s', but continuing with database update",
                        current_display_name,
                        display_name
                    )

                # Update slide image paths in the database
                self._update_slide_image_paths(
                    cursor=cursor,
                    presentation_id=presentation_id,
                    old_display_name=current_display_name,
                    new_display_name=display_name,
                    project_type="nw"
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

    # --- Custom Themes Methods ---

    def create_background_template(
        self,
        *,
        template_name: str,
        template_group: str,
        file_name: str
    ) -> dict:
        """Create a new background template in nw_Templates table.
        
        Args:
            template_name: Name of the template (unique)
            template_group: Group/category for organization
            file_name: Filename of the uploaded image (will be saved to BackGrounds folder)
            
        Returns:
            Dictionary with created template data including template_id
            
        Raises:
            ValueError: If template name already exists
            DatabaseConnectionError: If connection fails
            DatabaseTransactionError: If insert fails
        """
        logger.debug(
            "BI_GUIDELINES - Creating background template '%s' in group '%s'",
            template_name,
            template_group
        )

        with get_connection_scope(timeout=30) as cursor:
            # Check if template name already exists
            cursor.execute(
                """
                SELECT TemplateId FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                WHERE TemplateName = ?
                """,
                (template_name,)
            )
            existing = cursor.fetchone()
            
            if existing:
                raise ValueError(
                    f"Template '{template_name}' already exists"
                )

            # Insert into nw_Templates using stored procedure
            cursor.execute(
                """
                EXEC [BI_GUIDELINES].[dbo].[nw_InsertTemplate]
                    @TemplateName = ?,
                    @TemplateGroup = ?,
                    @TemplateFileName = ?
                """,
                (template_name, template_group, file_name)
            )
            
            # Get the inserted template ID
            cursor.execute(
                """
                SELECT TemplateId, TemplateName, TemplateGroup, TemplateFileName
                FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                WHERE TemplateName = ?
                """,
                (template_name,)
            )
            result = cursor.fetchone()
            cursor.connection.commit()

            logger.info(
                "BI_GUIDELINES - Created background template '%s' with ID %d",
                template_name,
                result.TemplateId
            )

            return {
                "template_id": result.TemplateId,
                "template_name": result.TemplateName,
                "template_group": result.TemplateGroup,
                "template_path": result.TemplateFileName
            }

    def get_background_templates(
        self,
        *,
        template_group: Optional[str] = None,
        page: int = 1,
        limit: int = 50
    ) -> Tuple[List[dict], int]:
        """Retrieve background templates with optional pagination and filtering.
        
        Args:
            template_group: Filter by group/category (optional)
            page: Page number (1-indexed)
            limit: Results per page
            
        Returns:
            Tuple of (list of template dicts, total count)
        """
        offset = (page - 1) * limit

        logger.debug(
            "BI_GUIDELINES - Fetching background templates: group=%s, page=%d",
            template_group or "(all)",
            page
        )

        with get_connection_scope(timeout=30) as cursor:
            # Build query based on filter
            if template_group:
                count_query = """
                    SELECT COUNT(*) as total
                    FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                    WHERE TemplateGroup = ?
                """
                query = """
                    SELECT TemplateId, TemplateName, TemplateGroup, TemplateFileName
                    FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                    WHERE TemplateGroup = ?
                    ORDER BY TemplateName
                    OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
                """
                # Get total count
                cursor.execute(count_query, (template_group,))
                total = cursor.fetchone().total
                
                # Get paginated results
                cursor.execute(query, (template_group, offset, limit))
            else:
                count_query = """
                    SELECT COUNT(*) as total
                    FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                """
                query = """
                    SELECT TemplateId, TemplateName, TemplateGroup, TemplateFileName
                    FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                    ORDER BY TemplateGroup, TemplateName
                    OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
                """
                # Get total count
                cursor.execute(count_query)
                total = cursor.fetchone().total
                
                # Get paginated results
                cursor.execute(query, (offset, limit))
            
            rows = cursor.fetchall()

            templates = []
            for row in rows:
                templates.append({
                    "template_id": row.TemplateId,
                    "template_name": row.TemplateName,
                    "template_group": row.TemplateGroup,
                    "template_path": row.TemplateFileName
                })

            logger.info(
                "BI_GUIDELINES - Retrieved %d background templates (total: %d)",
                len(templates),
                total
            )

            return templates, total

    def get_background_template(self, template_id: int) -> Optional[dict]:
        """Get a single background template by ID.
        
        Args:
            template_id: Template ID
            
        Returns:
            Template dict or None if not found
        """
        logger.debug("BI_GUIDELINES - Fetching background template ID: %d", template_id)

        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                """
                SELECT TemplateId, TemplateName, TemplateGroup, TemplateFileName
                FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                WHERE TemplateId = ?
                """,
                (template_id,)
            )
            row = cursor.fetchone()

            if not row:
                logger.warning("Background template %d not found", template_id)
                return None

            return {
                "template_id": row.TemplateId,
                "template_name": row.TemplateName,
                "template_group": row.TemplateGroup,
                "template_path": row.TemplateFileName
            }

    def update_background_template(
        self,
        *,
        template_id: int,
        template_name: Optional[str] = None,
        template_group: Optional[str] = None
    ) -> dict:
        """Update a background template's name or group.
        
        Note: TemplateFileName cannot be updated. To change the image, delete and recreate.
        
        Args:
            template_id: Template ID to update
            template_name: New template name (optional)
            template_group: New group/category (optional)
            
        Returns:
            Updated template dict
            
        Raises:
            ValueError: If template not found or no fields to update
        """
        logger.debug(
            "BI_GUIDELINES - Updating background template %d",
            template_id
        )

        with get_connection_scope(timeout=30) as cursor:
            # Verify template exists
            cursor.execute(
                """
                SELECT TemplateId, TemplateName, TemplateGroup, TemplateFileName
                FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                WHERE TemplateId = ?
                """,
                (template_id,)
            )
            row = cursor.fetchone()
            
            if not row:
                raise ValueError(f"Background template {template_id} not found")

            # Build UPDATE statement dynamically
            update_fields = []
            params = []

            if template_name is not None:
                # Check for duplicate name
                cursor.execute(
                    """
                    SELECT TemplateId FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                    WHERE TemplateName = ? AND TemplateId != ?
                    """,
                    (template_name, template_id)
                )
                if cursor.fetchone():
                    raise ValueError(f"Template name '{template_name}' already exists")
                
                update_fields.append("TemplateName = ?")
                params.append(template_name)
            
            if template_group is not None:
                update_fields.append("TemplateGroup = ?")
                params.append(template_group)

            if not update_fields:
                raise ValueError("No fields to update")

            # Execute update using stored procedure or direct SQL
            cursor.execute(
                """
                EXEC [BI_GUIDELINES].[dbo].[nw_UpdateTemplate]
                    @TemplateId = ?,
                    @TemplateName = ?,
                    @TemplateGroup = ?
                """,
                (template_id, template_name or row.TemplateName, template_group or row.TemplateGroup)
            )
            
            cursor.connection.commit()

            logger.info("BI_GUIDELINES - Updated background template %d", template_id)

            # Return updated template
            return self.get_background_template(template_id)

    def delete_background_template(
        self,
        *,
        template_id: int
    ) -> Tuple[bool, str]:
        """Delete a background template.
        
        Args:
            template_id: Template ID to delete
            
        Returns:
            Tuple of (success bool, template_path for file cleanup)
            
        Raises:
            ValueError: If template not found
        """
        logger.debug(
            "BI_GUIDELINES - Deleting background template %d",
            template_id
        )

        with get_connection_scope(timeout=30) as cursor:
            # Get template path before deleting
            cursor.execute(
                """
                SELECT TemplateFileName FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                WHERE TemplateId = ?
                """,
                (template_id,)
            )
            row = cursor.fetchone()
            
            if not row:
                raise ValueError(f"Background template {template_id} not found")
            
            template_path = row.TemplateFileName

            # Delete from database
            cursor.execute(
                """
                DELETE FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                WHERE TemplateId = ?
                """,
                (template_id,)
            )

            cursor.connection.commit()

            logger.info(
                "BI_GUIDELINES - Deleted background template %d",
                template_id
            )

            return True, template_path


# Global service instance
bi_guidelines_service = BIGuidelinesService()

