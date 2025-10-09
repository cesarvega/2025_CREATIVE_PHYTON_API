"""Service for BI_GUIDELINES database operations."""

from __future__ import annotations

from typing import List, Optional, Tuple

from app.config.db import get_connection_scope
from app.models.response_models import ActivePresentation
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


# Global service instance
bi_guidelines_service = BIGuidelinesService()
