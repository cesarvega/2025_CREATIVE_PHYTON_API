"""Service for DayMaster database operations."""

from __future__ import annotations

from typing import Optional

from app.config.db import get_connection_scope
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class DayMasterService:
    """Service for handling DayMaster database queries (projects, etc.)."""

    def get_projects(
        self,
        search: Optional[str] = None,
        page: int = 1,
        limit: int = 50,
    ) -> tuple[list[str], int]:
        """Retrieve paginated and filtered projects from DayMaster database.

        Args:
            search: Optional search term to filter project names.
            page: Page number (1-indexed).
            limit: Number of results per page.

        Returns:
            Tuple containing (list of project names, total count of matching projects).

        Raises:
            DatabaseConnectionError: If connection to database fails.
            DatabaseTransactionError: If query execution fails.
        """
        offset = (page - 1) * limit

        # Build query with optional search filter
        # Note: Filter by creative=1 to match the getSalesBoardProjects_Creative SP
        if search:
            # Use LIKE with wildcards for partial matching
            search_pattern = f"%{search}%"
            query = """
                SELECT ProjectName
                FROM dbo.Projects
                WHERE creative = 1
                AND ProjectName LIKE ?
                ORDER BY ProjectName
                OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
            """
            count_query = """
                SELECT COUNT(*) as Total
                FROM dbo.Projects
                WHERE creative = 1
                AND ProjectName LIKE ?
            """
            query_params = (search_pattern, offset, limit)
            count_params = (search_pattern,)
        else:
            # No search filter - get all creative projects
            query = """
                SELECT ProjectName
                FROM dbo.Projects
                WHERE creative = 1
                ORDER BY ProjectName
                OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
            """
            count_query = """
                SELECT COUNT(*) as Total
                FROM dbo.Projects
                WHERE creative = 1
            """
            query_params = (offset, limit)
            count_params = ()

        logger.debug(
            "DayMaster - Fetching projects: search=%s, page=%d, limit=%d, offset=%d",
            search or "(none)",
            page,
            limit,
            offset,
        )

        with get_connection_scope(timeout=30, use_daymaster=True) as cursor:
            # Get total count
            cursor.execute(count_query, count_params)
            total_row = cursor.fetchone()
            total = total_row.Total if total_row else 0

            # Get paginated results
            cursor.execute(query, query_params)
            rows = cursor.fetchall()
            projects = [row.ProjectName for row in rows if row.ProjectName]

            logger.info(
                "DayMaster - Retrieved %d projects (total: %d) for search='%s', page=%d",
                len(projects),
                total,
                search or "",
                page,
            )

            return projects, total


# Global service instance
daymaster_service = DayMasterService()
