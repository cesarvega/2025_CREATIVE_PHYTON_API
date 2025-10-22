"""Routes for BI Guidelines API.

This module provides endpoints to query and create BI Guideline presentations
from the BI_GUIDELINES database.
"""

import pyodbc
from typing import Optional
from fastapi import APIRouter, HTTPException, Query

from app.config.db import (
    DatabaseConnectionError,
    DatabaseTransactionError,
    get_connection_scope,
    get_db_connection,
)
from app.models.bi_guidelines_models import NWMasterRequest, ReloadProjectSoundsRequest, ProjectUpdateRequest
from app.models.presentation_models import PresentationData
from app.models.response_models import (
    ActivePresentationsResponse,
    DisplayNamesResponse,
    TemplateGroupsResponse,
)
from app.services.bi_guidelines_service import bi_guidelines_service
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/bi_guidelines", tags=["BI Guidelines"])
logger = get_logger(__name__)


class PresentationNotFoundError(Exception):
    """Custom exception for when a presentation is not found."""


@router.get(
    "/nw-active-presentations",
    response_model=ActivePresentationsResponse,
    summary="Get paginated and filterable list of active NW presentations",
    description=(
        "Retrieve active NW presentations from BI_GUIDELINES database with optional search filtering and pagination. "
        "This endpoint is optimized for large datasets (5000+ presentations) and supports:\n\n"
        "- **Search filtering**: Filter by project name or display name (partial match)\n"
        "- **Pagination**: Control page number and results per page\n"
        "- **Sorted results**: Presentations are sorted by last update date (newest first)\n\n"
        "Ideal for autocomplete inputs, dropdowns with search, and infinite scroll implementations."
    ),
)
async def get_nw_active_presentations(
    search: Optional[str] = Query(
        None,
        description="Optional search term to filter by project or display name (partial match)",
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
) -> ActivePresentationsResponse:
    """Get paginated active NW presentations from BI_GUIDELINES database.

    This endpoint queries the nw_Master table for presentations with status 'OPEN'
    and returns them sorted by last update date in descending order.

    Example usage:
    - Get first 50 presentations: `GET /api/bi_guidelines/nw-active-presentations`
    - Search for "SOLE": `GET /api/bi_guidelines/nw-active-presentations?search=SOLE`
    - Get page 2 with 100 results: `GET /api/bi_guidelines/nw-active-presentations?page=2&limit=100`
    """
    try:
        presentations, total = bi_guidelines_service.get_active_presentations(
            search=search,
            page=page,
            limit=limit,
        )
        return ActivePresentationsResponse(
            presentations=presentations,
            page=page,
            limit=limit,
            total=total,
        )
    except Exception as e:
        logger.error("Error retrieving active presentations: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve active presentations from database",
        ) from e


@router.get(
    "/bsr-active-presentations",
    response_model=ActivePresentationsResponse,
    summary="Get paginated and filterable list of active BSR presentations",
    description=(
        "Retrieve active BSR presentations from BI_GUIDELINES via stored procedure with optional search and pagination. "
        "This endpoint mirrors /active-presentations but sources data from the BSR pipeline using "
        "[dbo].[BSR_ActivePresentations]."
    ),
)
async def get_bsr_active_presentations(
    search: Optional[str] = Query(
        None,
        description="Optional search term to filter by project or display name (partial match)",
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
) -> ActivePresentationsResponse:
    """Get paginated active BSR presentations from BI_GUIDELINES database.

    Uses the [dbo].[BSR_ActivePresentations] stored procedure as data source
    and applies search + pagination in the API layer.
    """
    try:
        presentations, total = bi_guidelines_service.get_bsr_active_presentations(
            search=search,
            page=page,
            limit=limit,
        )
        return ActivePresentationsResponse(
            presentations=presentations,
            page=page,
            limit=limit,
            total=total,
        )
    except Exception as e:
        logger.error("Error retrieving BSR active presentations: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve BSR active presentations from database",
        ) from e


@router.get(
    "/display-names",
    response_model=DisplayNamesResponse,
    summary="Get paginated and filterable list of BSR display names",
    description=(
        "Retrieve BSR display names from BI_GUIDELINES database with optional search filtering and pagination. "
        "This endpoint is optimized for large datasets (4000+ names) and supports:\n\n"
        "- **Search filtering**: Filter display names by partial match\n"
        "- **Pagination**: Control page number and results per page\n"
        "- **Sorted results**: Names are sorted alphabetically\n\n"
        "Ideal for autocomplete inputs, dropdowns with search, and infinite scroll implementations."
    ),
)
async def get_bsr_display_names(
    search: Optional[str] = Query(
        None,
        description="Optional search term to filter display names (partial match)",
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
) -> DisplayNamesResponse:
    """Get paginated BSR display names from BI_GUIDELINES database.

    This endpoint queries the nw_Master table for distinct BSRDisplayName values
    and returns them sorted alphabetically.

    Example usage:
    - Get first 50 names: `GET /api/bi_guidelines/display-names`
    - Search for "BRIN": `GET /api/bi_guidelines/display-names?search=BRIN`
    - Get page 2 with 100 results: `GET /api/bi_guidelines/display-names?page=2&limit=100`
    """
    try:
        display_names, total = bi_guidelines_service.get_bsr_display_names(
            search=search,
            page=page,
            limit=limit,
        )
        return DisplayNamesResponse(
            display_names=display_names,
            page=page,
            limit=limit,
            total=total,
        )
    except Exception as e:
        logger.error("Error retrieving BSR display names: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve BSR display names from database",
        ) from e


@router.get("/project-info/{project_id}")
async def get_project_info(project_id: int):
    """Retrieve project information by executing the nw_PresentationInfo_Nw2_apr2020 stored procedure.

    Args:
        project_id: The project ID to retrieve information for.

    Returns:
        Dictionary containing all project details returned by the stored procedure.

    Raises:
        HTTPException: 404 if project not found, 500 if database error occurs.
    """
    try:
        project_details = bi_guidelines_service.get_project_info(project_id)

        if project_details is None:
            raise HTTPException(
                status_code=404,
                detail=f"Project with ID {project_id} not found"
            )

        logger.info("Retrieved project info for project_id=%d", project_id)
        return project_details

    except HTTPException:
        raise

    except DatabaseConnectionError as exc:
        logger.error("Database connection failed: %s", exc)
        raise HTTPException(
            status_code=503,
            detail="Database service temporarily unavailable."
        ) from exc

    except DatabaseTransactionError as exc:
        logger.error("Database query failed for project %d: %s", project_id, exc)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve project information from database."
        ) from exc

    except Exception as e:
        logger.error("Unexpected error retrieving project info for %d: %s", project_id, str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="An unexpected error occurred while retrieving project information."
        ) from e


@router.get(
    "/template-groups",
    response_model=TemplateGroupsResponse,
    summary="Get list of template groups",
    description=(
        "Retrieve available background templates grouped by category from BI_GUIDELINES database. "
        "This endpoint executes the getNW_TemplateGroups stored procedure and returns:\n\n"
        "- **Template information**: ID, name, and category for each template\n"
        "- **Complete list**: All available templates without pagination\n\n"
        "Useful for populating template selection dropdowns and displaying available background options."
    ),
)
async def get_template_groups() -> TemplateGroupsResponse:
    """Get template groups from BI_GUIDELINES database.

    This endpoint queries the getNW_TemplateGroups stored procedure to retrieve
    all available background templates grouped by category.

    Example usage:
    - Get all templates: `GET /api/bi_guidelines/template-groups`
    """
    try:
        template_groups, total = bi_guidelines_service.get_template_groups()
        return TemplateGroupsResponse(
            template_groups=template_groups,
            total=total,
        )
    except Exception as e:
        logger.error("Error retrieving template groups: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve template groups from database",
        ) from e


@router.post(
    "/reload-project-sounds",
    summary="Reload MP3 file paths for a NW project",
    description=(
        "Executes the NW_UpdateMP3FilePath stored procedure to resynchronize "
        "audio file paths for a project based on existing files in the cloud. "
        "This is useful when MP3 files have been updated or moved and need to be "
        "reindexed in the database.\n\n"
        "**Use case**: After uploading or modifying audio files for a project, "
        "call this endpoint to ensure the database references are up to date."
    ),
)
async def reload_project_sounds(request: ReloadProjectSoundsRequest):
    """Reload MP3 file paths for a NW project.

    Executes the stored procedure [dbo].[NW_UpdateMP3FilePath] to update
    the audio file paths associated with a project.

    Args:
        request: Request body containing the display_name of the project.

    Returns:
        Success message confirming the reload operation.

    Raises:
        HTTPException: 400 if display_name is invalid, 500 if database error occurs.

    Example usage:
        POST /api/bi_guidelines/reload-project-sounds
        {
            "display_name": "SOLE_19Nov2024"
        }
    """
    try:
        bi_guidelines_service.reload_project_sounds(request.display_name)

        return {
            "message": f"Successfully reloaded sounds for project '{request.display_name}'",
            "display_name": request.display_name
        }

    except ValueError as e:
        logger.error("Invalid input for reload_project_sounds: %s", str(e))
        raise HTTPException(
            status_code=400,
            detail=str(e)
        ) from e

    except DatabaseConnectionError as exc:
        logger.error("Database connection failed: %s", exc)
        raise HTTPException(
            status_code=503,
            detail="Database service temporarily unavailable."
        ) from exc

    except DatabaseTransactionError as exc:
        logger.error("Database query failed for display_name %s: %s", request.display_name, exc)
        raise HTTPException(
            status_code=500,
            detail="Failed to reload project sounds. Please check the display name and try again."
        ) from exc

    except Exception as e:
        logger.error("Unexpected error reloading sounds for %s: %s", request.display_name, str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="An unexpected error occurred while reloading project sounds."
        ) from e


@router.put(
    "/update-project-details",
    summary="Update NW project presentation master details",
    description=(
        "Executes the nw_UpdatePresentationMaster stored procedure to update "
        "the master details of an existing NW presentation. This endpoint allows you to modify:\n\n"
        "- Display name (NW)\n"
        "- Presentation status (OPEN/CLOSED)\n"
        "- BSR display name (optional association)\n\n"
        "The LastUpdateDate is automatically updated in the database.\n\n"
        "**Use case**: When project master details change (e.g., renaming a project, "
        "changing status, linking to BSR), use this endpoint to keep the database synchronized."
    ),
)
async def update_project_details(request: ProjectUpdateRequest):
    """Update NW project presentation master details.

    Executes the stored procedure [dbo].[nw_UpdatePresentationMaster] to modify
    an existing NW presentation's master details.

    Args:
        request: Request body containing the fields to update.

    Returns:
        Success message confirming the update operation.

    Raises:
        HTTPException: 400 if input is invalid, 500 if database error occurs.

    Example usage:
        PUT /api/bi_guidelines/update-project-details
        {
            "presentation_id": 8286,
            "display_name": "SOLE_19Nov2024",
            "presentation_status": "OPEN",
            "bsr_display_name": "SOLE_BSR_2024"
        }
    """
    try:
        bi_guidelines_service.update_project_details(
            presentation_id=request.presentation_id,
            display_name=request.display_name,
            presentation_status=request.presentation_status,
            bsr_display_name=request.bsr_display_name,
        )

        return {
            "message": f"Successfully updated project with ID {request.presentation_id}",
            "presentation_id": request.presentation_id,
            "display_name": request.display_name,
            "presentation_status": request.presentation_status,
            "bsr_display_name": request.bsr_display_name
        }

    except ValueError as e:
        logger.error("Invalid input for update_project_details: %s", str(e))
        raise HTTPException(
            status_code=400,
            detail=str(e)
        ) from e

    except DatabaseConnectionError as exc:
        logger.error("Database connection failed: %s", exc)
        raise HTTPException(
            status_code=503,
            detail="Database service temporarily unavailable."
        ) from exc

    except DatabaseTransactionError as exc:
        logger.error(
            "Database query failed for presentation_id %d: %s",
            request.presentation_id,
            exc
        )
        raise HTTPException(
            status_code=500,
            detail="Failed to update project details. Please verify the presentation ID and try again."
        ) from exc

    except Exception as e:
        logger.error(
            "Unexpected error updating presentation_id %d: %s",
            request.presentation_id,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail="An unexpected error occurred while updating project details."
        ) from e


# --- Data Models (Pydantic) ---
# Define the structure of the JSON that the API expects to receive.
