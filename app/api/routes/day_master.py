"""Routes for DayMaster database operations (projects, etc.)."""

from typing import Optional

from fastapi import APIRouter, HTTPException, Query

from app.models.response_models import ProjectsResponse
from app.services.daymaster_service import daymaster_service
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/daymaster", tags=["DayMaster"])
logger = get_logger(__name__)


@router.get(
    "/projects",
    response_model=ProjectsResponse,
    summary="Get paginated and filterable list of projects",
    description=(
        "Retrieve projects from DayMaster database with optional search filtering and pagination. "
        "This endpoint is optimized for large datasets (8000+ projects) and supports:\n\n"
        "- **Search filtering**: Filter projects by partial name match\n"
        "- **Pagination**: Control page number and results per page\n"
        "- **Sorted results**: Projects are always sorted alphabetically\n\n"
        "Ideal for autocomplete inputs, dropdowns with search, and infinite scroll implementations."
    ),
    response_description=(
        "Paginated list of project names with metadata including total count for pagination controls."
    ),
    responses={
        200: {
            "description": "Successfully retrieved projects",
            "content": {
                "application/json": {
                    "example": {
                        "projects": ["_brin2", "007", "0367", "0599_CTID"],
                        "page": 2,
                        "limit": 50,
                        "total": 8123,
                    }
                }
            },
        },
        400: {
            "description": "Invalid query parameters",
            "content": {
                "application/json": {
                    "example": {"detail": "Page must be greater than 0"}
                }
            },
        },
        500: {
            "description": "Database connection or query error",
            "content": {
                "application/json": {
                    "example": {"detail": "Failed to retrieve projects from database"}
                }
            },
        },
    },
)
async def get_projects(
    search: Optional[str] = Query(
        None,
        description="Optional search term to filter project names (partial match)",
        min_length=1,
        max_length=100,
    ),
    page: int = Query(
        1,
        ge=1,
        description="Page number (1-indexed)",
    ),
    limit: int = Query(
        50,
        ge=1,
        le=500,
        description="Number of results per page (max 500)",
    ),
) -> ProjectsResponse:
    """Retrieve paginated and optionally filtered projects from SQL Server.

    This endpoint queries the `daymaster.dbo.SalesBoardProjects` table and returns
    project names in alphabetical order. Use the `search` parameter to filter by
    partial name match, and `page`/`limit` parameters to control pagination.

    Example usage:
    - Get first 50 projects: `GET /api/daymaster/projects`
    - Search for "brin": `GET /api/daymaster/projects?search=brin`
    - Get page 2 with 100 results: `GET /api/daymaster/projects?page=2&limit=100`
    - Search with pagination: `GET /api/daymaster/projects?search=brin&page=2&limit=50`
    """
    try:
        logger.info(
            "DayMaster projects request: search=%s, page=%d, limit=%d",
            search or "(none)",
            page,
            limit,
        )

        # Call service to retrieve projects
        projects, total = daymaster_service.get_projects(
            search=search,
            page=page,
            limit=limit,
        )

        return ProjectsResponse(
            projects=projects,
            page=page,
            limit=limit,
            total=total,
        )

    except Exception as e:
        logger.error("Error retrieving projects: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve projects from database",
        ) from e
