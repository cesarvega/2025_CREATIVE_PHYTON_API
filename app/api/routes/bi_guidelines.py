"""Routes for BI Guidelines API.

This module provides endpoints to query and create BI Guideline presentations
from the BI_GUIDELINES database.
"""

import pyodbc
from pathlib import Path
from typing import Literal, Optional
from fastapi import APIRouter, HTTPException, Query, UploadFile, File, Form, Depends

from app.config.db import (
    DatabaseConnectionError,
    DatabaseTransactionError,
    get_connection_scope,
    get_db_connection,
)
from app.config.settings import settings
from app.models.bi_guidelines_models import (
    NWMasterRequest,
    ReloadProjectSoundsRequest,
    ProjectUpdateRequest,
    BSRProjectUpdateRequest,
    BackgroundTemplateCreate,
    BackgroundTemplateUpdate,
)
from app.utils.image_processor import create_thumbnail, get_thumbnail_filename
from app.models.nsr_models import (
    NSRProjectConfigResponse,
    NSRUpdateRuleRequest,
    NSRUpdateRuleResponse,
    NSRProjectsListResponse,
    NSRRuleExistsResponse,
    NSRInitializeResponse,
)
from app.models.presentation_models import PresentationData
from app.models.response_models import (
    ActivePresentationsResponse,
    DisplayNamesResponse,
    TemplateGroup,
    TemplateGroupsResponse,
    CustomThemesListResponse,
    CustomThemeCreateResponse,
    CustomThemeDeleteResponse,
)
from app.services.bi_guidelines_service import bi_guidelines_service
from app.services.nsr_service import nsr_service
from app.utils.file_upload_utils import (
    process_file,
    FileValidationError,
    FileProcessingError,
    cleanup_theme_file,
)
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
    status: Literal["OPEN", "CLOSED", "ALL"] = Query(
        "OPEN",
        description="Filter by presentation status. Use 'ALL' to return both OPEN and CLOSED.",
    ),
    project_type: Optional[Literal["NW", "DW"]] = Query(
        None,
        description="Optional project type filter. 'NW' for standard naming, 'DW' for Design presentations. "
        "When omitted, returns both NW and DW (original mixed behavior).",
    ),
) -> ActivePresentationsResponse:
    """Get paginated active NW presentations from BI_GUIDELINES database.

    This endpoint queries the nw_Master table and returns results sorted by last
    update date in descending order. By default only presentations with status
    'OPEN' are returned, matching the original behavior of this endpoint.

    Example usage:
    - Get first 50 OPEN presentations: `GET /api/bi_guidelines/nw-active-presentations`
    - Get CLOSED presentations: `GET /api/bi_guidelines/nw-active-presentations?status=CLOSED`
    - Get both: `GET /api/bi_guidelines/nw-active-presentations?status=ALL`
    - Search for "SOLE": `GET /api/bi_guidelines/nw-active-presentations?search=SOLE`
    - Get page 2 with 100 results: `GET /api/bi_guidelines/nw-active-presentations?page=2&limit=100`
    - Filter by NW only: `GET /api/bi_guidelines/nw-active-presentations?project_type=NW`
    - Filter by DW only: `GET /api/bi_guidelines/nw-active-presentations?project_type=DW`
    """
    try:
        presentations, total = bi_guidelines_service.get_active_presentations(
            search=search,
            page=page,
            limit=limit,
            status=status,
            project_type=project_type,
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
    status: Literal["OPEN", "CLOSED", "ALL"] = Query(
        "OPEN",
        description="Filter by presentation status. Use 'ALL' to return both OPEN and CLOSED.",
    ),
    project_type: Optional[Literal["BSR", "NSR"]] = Query(
        None,
        description="Optional project type filter. 'BSR' for BSR presentations, 'NSR' for NSR presentations. "
        "When omitted, returns both BSR and NSR (original mixed behavior).",
    ),
) -> ActivePresentationsResponse:
    """Get paginated active BSR presentations from BI_GUIDELINES database.

    Uses the [dbo].[BSR_ActivePresentations] stored procedure as data source
    and applies status filter, search, project_type filter, and pagination in the API layer.
    """
    try:
        presentations, total = bi_guidelines_service.get_bsr_active_presentations(
            search=search,
            page=page,
            limit=limit,
            status=status,
            project_type=project_type,
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
    """Retrieve project information automatically detecting if it's NW, BSR, or NSR type.

    This endpoint intelligently searches for the project in both NW and BSR tables:
    1. First tries to find the project in NW table (nw_PresentationInfo_Nw2_apr2020)
    2. If not found in NW, tries BSR table (BSR_PresentationInfo) - includes BSR and NSR
    3. Returns project details if found in either table

    Args:
        project_id: The project ID to retrieve information for.

    Returns:
        Dictionary containing all project details returned by the stored procedure.
        Includes additional fields:
        - "project_type": "NW" or "BSR" (BSR includes NSR projects)
        - "page_number": The starting slide/page number for the project

    Raises:
        HTTPException: 404 if project not found in either table, 500 if database error occurs.
    """
    try:
        # First, try to get project info from NW table
        project_details = bi_guidelines_service.get_project_info(project_id)

        if project_details is not None:
            # Found in NW table
            project_details["project_type"] = "NW"
            # Add page_number: prefer field from SP, else fetch from nw_Master
            page_number = None
            for key, val in project_details.items():
                if str(key).lower() == "namecandidatestartingslide":
                    page_number = val
                    break
            if page_number is None:
                # Fallback lookup from nw_Master
                try:
                    with get_connection_scope(timeout=30) as cursor:
                        cursor.execute(
                            "SELECT NameCandidateStartingSlide FROM [BI_GUIDELINES].[dbo].[nw_Master] WHERE PresentationId = ?",
                            (project_id,),
                        )
                        row = cursor.fetchone()
                        if row:
                            page_number = row[0]
                except Exception:
                    page_number = page_number  # leave unchanged
            if page_number is not None:
                project_details["page_number"] = page_number
            logger.info("Retrieved NW project info for project_id=%d", project_id)
            return project_details

        # Not found in NW, try BSR table
        logger.debug("Project %d not found in NW table, trying BSR table", project_id)
        project_details = bi_guidelines_service.get_bsr_project_info(project_id)

        if project_details is not None:
            # Found in BSR table (includes BSR and NSR projects)
            project_details["project_type"] = "BSR"
            # Add page_number: prefer field from SP, else default to 1
            page_number = None

            # Try to get from stored procedure result first
            for key, val in project_details.items():
                key_lower = str(key).lower()
                if key_lower in ("slidenumber", "page_number", "pagenumber", "startingslide"):
                    page_number = val
                    logger.debug("Found page_number from SP field '%s': %s", key, val)
                    break

            # If not found in SP, try fallback query to bsr_Master
            if page_number is None:
                try:
                    with get_connection_scope(timeout=30) as cursor:
                        # Query bsr_Master table for NameCandidateStartingSlide (same field as NW)
                        cursor.execute(
                            "SELECT NameCandidateStartingSlide FROM [BI_GUIDELINES].[dbo].[bsr_Master] WHERE PresentationId = ?",
                            (project_id,),
                        )
                        row = cursor.fetchone()
                        if row and row[0] is not None:
                            page_number = row[0]
                            logger.debug("Found page_number from bsr_Master.NameCandidateStartingSlide: %s", page_number)
                except Exception as e:
                    logger.warning("Failed to query page_number from bsr_Master for project %d: %s", project_id, str(e))

            # Default to 1 if still not found
            if page_number is None:
                page_number = 1
                logger.debug("Using default page_number=1 for BSR project %d", project_id)

            project_details["page_number"] = page_number
            logger.info("Retrieved BSR project info for project_id=%d with page_number=%s", project_id, page_number)
            return project_details

        # Not found in either table
        raise HTTPException(
            status_code=404,
            detail=f"Project with ID {project_id} not found in NW or BSR tables"
        )

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
    "/bsr-project-info/{project_id}",
    summary="Get BSR project information by ID",
    description=(
        "Retrieve detailed information for a specific BSR (Board Sales Request) presentation.\n\n"
        "This endpoint executes the BSR_PresentationInfo stored procedure and retrieves associated categories "
        "to provide complete project details including:\n"
        "- Project name and display name\n"
        "- PowerPoint file path\n"
        "- Slide number configuration\n"
        "- Presentation type (BSR or BSR-Japan)\n"
        "- Presentation status (OPEN, CLOSED, etc.)\n"
        "- User information (uploaded by, upload date)\n"
        "- Wide screen setting\n"
        "- Web link (if available)\n"
        "- Creation and update timestamps\n"
        "- Project categories with their elements\n\n"
        "Useful for loading existing BSR presentations for editing or viewing."
    ),
)
async def get_bsr_project_info(project_id: int):
    """Retrieve BSR project information using BSR_PresentationInfo stored procedure.

    This endpoint provides complete BSR presentation information including project categories.
    It executes two stored procedures:
    1. BSR_PresentationInfo - for main presentation details
    2. bsr_GetCategoryValues - for project categories

    Args:
        project_id: The BSR presentation ID to retrieve information for.

    Returns:
        Dictionary containing all BSR project details with categories.

    Raises:
        HTTPException: 404 if project not found, 500 if database error occurs.

    Example response:
        {
            "project": "BSR_Project_2025",
            "displayname": "Test_Presentation",
            "uploadedby": "analyst",
            "uploadeddate": "2025-01-15T10:30:00",
            "presentationid": 12345,
            "presentationtype": "BSR",
            "presentationstatus": "OPEN",
            "link": "https://example.com/presentation",
            "lastupdatedate": "2025-01-15T10:30:00",
            "slidenumber": 3,
            "iswideppt": 0,
            "page_number": 3,
            "categories": [
                {
                    "category": "Technology",
                    "elements": "AI, Machine Learning, Cloud"
                },
                {
                    "category": "Industry",
                    "elements": "Healthcare, Finance"
                }
            ]
        }
    """
    try:
        project_details = bi_guidelines_service.get_bsr_project_info(project_id)

        if project_details is None:
            raise HTTPException(
                status_code=404,
                detail=f"BSR Project with ID {project_id} not found"
            )

        # Add page_number: prefer field from SP, else fetch from bsr_PresentationMaster, else default to 1
        page_number = None

        # Try to get from stored procedure result first
        for key, val in project_details.items():
            key_lower = str(key).lower()
            if key_lower in ("slidenumber", "page_number", "pagenumber", "startingslide"):
                page_number = val
                logger.debug("Found page_number from SP field '%s': %s", key, val)
                break

        # If not found in SP, try fallback query to bsr_Master
        if page_number is None:
            try:
                with get_connection_scope(timeout=30) as cursor:
                    # Query bsr_Master table for NameCandidateStartingSlide (same field as NW)
                    cursor.execute(
                        "SELECT NameCandidateStartingSlide FROM [BI_GUIDELINES].[dbo].[bsr_Master] WHERE PresentationId = ?",
                        (project_id,),
                    )
                    row = cursor.fetchone()
                    if row and row[0] is not None:
                        page_number = row[0]
                        logger.debug("Found page_number from bsr_Master.NameCandidateStartingSlide: %s", page_number)
            except Exception as e:
                logger.warning("Failed to query page_number from bsr_Master for project %d: %s", project_id, str(e))

        # Default to 1 if still not found
        if page_number is None:
            page_number = 1
            logger.debug("Using default page_number=1 for BSR project %d", project_id)

        project_details["page_number"] = page_number
        logger.info("Retrieved BSR project info for project_id=%d with page_number=%s", project_id, page_number)
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
        logger.error("Database query failed for BSR project %d: %s", project_id, exc)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve BSR project information from database."
        ) from exc

    except Exception as e:
        logger.error("Unexpected error retrieving BSR project info for %d: %s", project_id, str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="An unexpected error occurred while retrieving BSR project information."
        ) from e


@router.get(
    "/template-groups",
    response_model=TemplateGroupsResponse,
    summary="Get list of background templates",
    description=(
        "Retrieve available background templates grouped by category from BI_GUIDELINES database. "
        "This endpoint executes the getNW_TemplateGroups stored procedure.\n\n"
        "- **Template information**: ID, name, category, and file path for each template\n"
        "- **Preview URLs**: Full URLs to access the background images\n\n"
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
        template_groups, total, custom_count, system_count = bi_guidelines_service.get_template_groups()
        return TemplateGroupsResponse(
            template_groups=template_groups,
            total=total,
            custom_count=custom_count,
            system_count=system_count,
        )
    except ValueError as e:
        logger.error("Validation error in get_template_groups: %s", str(e))
        raise HTTPException(status_code=400, detail=str(e)) from e
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


@router.put(
    "/update-bsr-project",
    summary="Update BSR project presentation master details",
    description=(
        "Executes the BSR_UpdatePresentationMaster stored procedure to update "
        "the master details of an existing BSR presentation. This endpoint allows you to modify:\n\n"
        "- Display name\n"
        "- Presentation status (OPEN/CLOSED)\n\n"
        "The LastUpdateDate is automatically updated in the database.\n\n"
        "**Use case**: When BSR project details change (e.g., renaming a presentation, "
        "changing status to CLOSED), use this endpoint to keep the database synchronized.\n\n"
        "**Note**: When status changes to CLOSED, the application may trigger additional "
        "processing such as email notifications."
    ),
)
async def update_bsr_project(request: BSRProjectUpdateRequest):
    """Update BSR project presentation master details.

    Executes the stored procedure [dbo].[BSR_UpdatePresentationMaster] to modify
    an existing BSR presentation's master details.

    Args:
        request: Request body containing the fields to update.

    Returns:
        Success message confirming the update operation.

    Raises:
        HTTPException: 400 if input is invalid, 500 if database error occurs.

    Example usage:
        PUT /api/bi_guidelines/update-bsr-project
        {
            "presentation_id": 12345,
            "display_name": "Test_Presentation_Updated",
            "status": "CLOSED"
        }
    """
    try:
        bi_guidelines_service.update_bsr_presentation(
            presentation_id=request.presentation_id,
            display_name=request.display_name,
            status=request.status,
        )

        return {
            "message": f"Successfully updated BSR project with ID {request.presentation_id}",
            "presentation_id": request.presentation_id,
            "display_name": request.display_name,
            "status": request.status
        }

    except ValueError as e:
        logger.error("Invalid input for update_bsr_project: %s", str(e))
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
            "Database query failed for BSR presentation_id %d: %s",
            request.presentation_id,
            exc
        )
        raise HTTPException(
            status_code=500,
            detail="Failed to update BSR project details. Please verify the presentation ID and try again."
        ) from exc

    except Exception as e:
        logger.error(
            "Unexpected error updating BSR presentation_id %d: %s",
            request.presentation_id,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail="An unexpected error occurred while updating BSR project details."
        ) from e


# --- NSR Configuration Endpoints ---


@router.get(
    "/nsr-config/{project_name}",
    response_model=NSRProjectConfigResponse,
    summary="Get NSR project rule configuration",
    description=(
        "Retrieve all NSR validation rules for a specific project.\n\n"
        "**NSR Rules:**\n"
        "- **101-105**: Contains rules (Y, H, W, J, K)\n"
        "- **106, 108, 116, 117**: USAN validation checks\n"
        "- **109-115**: Prefix rules (AR, DEX, ES, STR, LEV, X, RAC)\n\n"
        "Each rule has a state:\n"
        "- **0**: OFF (rule not active)\n"
        "- **1**: ON (rule active)\n\n"
        "If no configuration exists, returns empty rules list (all OFF by default)."
    ),
)
async def get_nsr_project_config(project_name: str) -> NSRProjectConfigResponse:
    """Get NSR rule configuration for a project."""
    try:
        rules = nsr_service.get_project_config(project_name)

        return NSRProjectConfigResponse(
            project_name=project_name,
            rules=rules
        )

    except Exception as e:
        logger.error(
            "Error fetching NSR config for project '%s': %s",
            project_name,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Failed to fetch NSR configuration: {str(e)}"
        ) from e


@router.put(
    "/nsr-config",
    response_model=NSRUpdateRuleResponse,
    summary="Update NSR rule configuration",
    description=(
        "Update (or insert) an NSR validation rule for a project.\n\n"
        "This endpoint automatically determines whether to INSERT or UPDATE:\n"
        "- If the rule doesn't exist for the project → **INSERT**\n"
        "- If the rule already exists → **UPDATE**\n\n"
        "**Request body:**\n"
        "```json\n"
        "{\n"
        '  "project_name": "ATEST01",\n'
        '  "rule_id": 101,\n'
        '  "is_on": 1\n'
        "}\n"
        "```\n\n"
        "**Rule IDs:**\n"
        "- **101**: Contains Y\n"
        "- **102**: Contains H\n"
        "- **103**: Contains W\n"
        "- **104**: Contains J\n"
        "- **105**: Contains K\n"
        "- **106**: Check USAN Violation\n"
        "- **108**: Check USAN Nomenclature\n"
        "- **109**: Prefix AR\n"
        "- **110**: Prefix DEX\n"
        "- **111**: Prefix ES\n"
        "- **112**: Prefix STR\n"
        "- **113**: Prefix LEV\n"
        "- **114**: Prefix X\n"
        "- **115**: Prefix RAC\n"
        "- **116**: Check USAN Search\n"
        "- **117**: Check USAN MedNet"
    ),
)
async def update_nsr_rule(request: NSRUpdateRuleRequest) -> NSRUpdateRuleResponse:
    """Update or insert an NSR rule configuration."""
    try:
        success, operation = nsr_service.upsert_rule(
            project_name=request.project_name,
            rule_id=request.rule_id,
            is_on=request.is_on
        )

        if not success:
            raise HTTPException(
                status_code=500,
                detail=f"Failed to {operation} NSR rule {request.rule_id}"
            )

        return NSRUpdateRuleResponse(
            message=f"NSR rule {request.rule_id} successfully {operation}d",
            project_name=request.project_name,
            rule_id=request.rule_id,
            is_on=request.is_on,
            operation=operation
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(
            "Error updating NSR rule (project='%s', rule=%d, is_on=%d): %s",
            request.project_name,
            request.rule_id,
            request.is_on,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Failed to update NSR rule: {str(e)}"
        ) from e


@router.get(
    "/nsr-projects",
    response_model=NSRProjectsListResponse,
    summary="Get list of active NSR projects",
    description=(
        "Retrieve list of active NSR project display names.\n\n"
        "This endpoint returns all active NSR projects (both 'NSR' and 'NSR-Japan' types) "
        "to populate dropdowns or project selection lists.\n\n"
        "**Uses stored procedure:** `nsr_ActivePresentations`\n\n"
        "**Fallback query:** If SP doesn't exist, queries `bsr_Master` table directly."
    ),
)
async def get_nsr_projects() -> NSRProjectsListResponse:
    """Get list of active NSR projects."""
    try:
        projects = nsr_service.get_active_nsr_projects()

        return NSRProjectsListResponse(
            success=True,
            data=projects
        )

    except Exception as e:
        logger.error(
            "Error fetching NSR projects: %s",
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Failed to fetch NSR projects: {str(e)}"
        ) from e


@router.get(
    "/nsr-config/{project_name}/{rule_id}/exists",
    response_model=NSRRuleExistsResponse,
    summary="Check if NSR rule exists for project",
    description=(
        "Verify if a specific NSR rule already exists for a project.\n\n"
        "This endpoint is useful for determining whether to perform an INSERT or UPDATE operation.\n\n"
        "**Returns:**\n"
        "- `exists: true` → Rule exists (use UPDATE)\n"
        "- `exists: false` → Rule doesn't exist (use INSERT)\n\n"
        "**Uses stored procedure:** `nsr_CheckTableEntry`"
    ),
)
async def check_nsr_rule_exists(project_name: str, rule_id: int) -> NSRRuleExistsResponse:
    """Check if an NSR rule exists for a project."""
    try:
        # Validate rule_id
        if rule_id < 101 or rule_id > 120 or rule_id == 107:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid rule_id: {rule_id}. Must be 101-120 (except 107)"
            )

        exists = nsr_service.check_rule_exists(project_name, rule_id)

        return NSRRuleExistsResponse(
            success=True,
            exists=exists
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(
            "Error checking NSR rule existence (project='%s', rule=%d): %s",
            project_name,
            rule_id,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Failed to check NSR rule existence: {str(e)}"
        ) from e


@router.post(
    "/nsr-config/{project_name}/initialize",
    response_model=NSRInitializeResponse,
    summary="Initialize all NSR rules for a project",
    description=(
        "Create all NSR validation rules (101-120, except 107) for a new project.\n\n"
        "All rules are initialized in **ON** state (is_on=1) by default.\n\n"
        "**Rules created:**\n"
        "- **Contains:** 101, 102, 103, 104, 105\n"
        "- **USAN:** 106, 108, 116, 117\n"
        "- **Prefix:** 109, 110, 111, 112, 113, 114, 115, 120\n"
        "- **Double:** 118, 119\n\n"
        "**Total:** 19 rules (107 is skipped)\n\n"
        "**Use case:** When a new NSR project is created and needs default rule configuration."
    ),
)
async def initialize_nsr_rules(project_name: str) -> NSRInitializeResponse:
    """Initialize all NSR rules for a project."""
    try:
        success, rules_created = nsr_service.initialize_all_rules(project_name)

        if not success:
            raise HTTPException(
                status_code=500,
                detail=f"Failed to initialize NSR rules for project '{project_name}'"
            )

        return NSRInitializeResponse(
            success=True,
            message="All rules initialized successfully",
            displayName=project_name,
            rulesCreated=rules_created
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(
            "Error initializing NSR rules for project '%s': %s",
            project_name,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Failed to initialize NSR rules: {str(e)}"
        ) from e


# --- Custom Themes Endpoints ---


@router.post(
    "/background-templates",
    response_model=CustomThemeCreateResponse,
    status_code=201,
    summary="Upload a new background template",
    description=(
        "Upload a custom background (JPG, PNG, or PPTX).\n\n"
        "**File Requirements:**\n"
        "- Format: JPG, PNG, or PPTX\n"
        "- Max size: 10MB\n"
        "- PPTX files: First slide extracted as JPG (1920x1080)\n"
        "- Images saved directly to BackGrounds folder\n\n"
        "**Storage:**\n"
        "- Production: C:\\inetpub\\wwwroot\\nw2\\assets\\images\\BackGrounds\n"
        "- Development: NW_Files/backgrounds\n"
        "- Files stored with sanitized filenames"
    ),
)
async def create_background_template(
    template_name: str = Form(..., min_length=3, max_length=100),
    template_group: str = Form(..., min_length=3, max_length=100),
    background_file: UploadFile = File(...),
):
    """Upload a background template (image or PPTX).
    
    This endpoint:
    1. Validates the uploaded file (JPG/PNG/PPTX, max 10MB)
    2. For PPTX: Extracts first slide as JPG
    3. Saves to BackGrounds directory
    4. Creates record in nw_Templates table
    5. Returns template information
    """
    try:
        # Read file content
        file_content = await background_file.read()
        
        logger.info(
            "Creating background template '%s' in group '%s' (file: %s, size: %d bytes)",
            template_name,
            template_group,
            background_file.filename,
            len(file_content)
        )
        
        # Determine output directory based on environment
        backgrounds_dir = settings.backgrounds_dir
        
        # Process and save file
        try:
            saved_filename = process_file(
                filename=background_file.filename,
                file_content=file_content,
                output_dir=backgrounds_dir
            )
        except FileValidationError as e:
            logger.warning("File validation failed for '%s': %s", background_file.filename, str(e))
            raise HTTPException(status_code=400, detail=str(e)) from e
        except FileProcessingError as e:
            logger.error("File processing failed for '%s': %s", background_file.filename, str(e))
            raise HTTPException(status_code=500, detail=f"Failed to process file: {str(e)}") from e
        
        # Generate thumbnail for faster loading in frontend
        try:
            thumbnails_dir = backgrounds_dir / "thumbnails"
            thumbnails_dir.mkdir(exist_ok=True)
            
            source_path = backgrounds_dir / saved_filename
            thumbnail_filename = get_thumbnail_filename(saved_filename)
            thumbnail_path = thumbnails_dir / thumbnail_filename
            
            create_thumbnail(source_path, thumbnail_path)
            logger.info("Thumbnail created: %s", thumbnail_filename)
        except Exception as e:
            # Don't fail the entire request if thumbnail creation fails
            logger.warning("Failed to create thumbnail for '%s': %s", saved_filename, str(e))
        
        # Build relative path for database storage
        # Physical path: C:/inetpub/wwwroot/nw2/assets/images/BackGrounds/slide_7.png
        # DB path: images/BackGrounds/slide_7.png
        relative_path = f"images/BackGrounds/{saved_filename}"
        
        # Create database record
        template_data = bi_guidelines_service.create_background_template(
            template_name=template_name,
            template_group=template_group,
            file_name=relative_path
        )
        
        logger.info(
            "Successfully created background template '%s' (ID: %d)",
            template_name,
            template_data['template_id']
        )
        
        # Build response
        theme = TemplateGroup(
            template_group_id=template_data['template_id'],
            template_name=template_data['template_name'],
            category=template_data['template_group'],
            template_file_name=template_data['template_path']
        )
        
        return CustomThemeCreateResponse(
            success=True,
            message="Background template created successfully",
            template=theme
        )
        
    except ValueError as e:
        logger.error("Validation error creating background template: %s", str(e))
        raise HTTPException(status_code=400, detail=str(e)) from e
    except HTTPException:
        raise
    except Exception as e:
        logger.error("Unexpected error creating background template: %s", str(e), exc_info=True)
        # Cleanup on error
        if 'saved_filename' in locals():
            file_path = backgrounds_dir / saved_filename
            if file_path.exists():
                file_path.unlink()
        raise HTTPException(
            status_code=500,
            detail="Failed to create background template"
        ) from e


@router.get(
    "/background-templates",
    response_model=CustomThemesListResponse,
    summary="Get list of background templates",
    description=(
        "Retrieve background templates with optional filtering by group.\n\n"
        "**Filters:**\n"
        "- `template_group`: Filter by group/category\n"
        "- `page`: Page number (default: 1)\n"
        "- `limit`: Results per page (default: 50, max: 100)"
    ),
)
async def get_background_templates(
    template_group: Optional[str] = Query(None, description="Filter by template group"),
    page: int = Query(1, ge=1, description="Page number"),
    limit: int = Query(50, ge=1, le=100, description="Results per page"),
):
    """Get list of background templates with optional pagination and filtering."""
    try:
        templates, total = bi_guidelines_service.get_background_templates(
            template_group=template_group,
            page=page,
            limit=limit
        )
        
        # Convert to TemplateGroup format with preview URLs
        theme_groups = []
        for t in templates:
            # Construct preview URL and thumbnail URL
            preview_url = None
            thumbnail_url = None
            
            if t['template_path']:
                from app.config.settings import settings
                if settings.environment == "production":
                    base_url = "https://tools.brandinstitute.com/nw2/assets/"
                else:
                    base_url = "https://tools.brandinstitute.com/nw2/assets/"
                
                # Full resolution image - normalize path separators
                preview_url = f"{base_url}{t['template_path'].replace(chr(92), '/')}"
                
                # Thumbnail (300x169px for fast loading)
                # Maintain subdirectory structure
                # Example: images/BackGrounds/Backgrounds2019/file.jpg
                #       -> images/BackGrounds/Backgrounds2019/thumbnails/file.jpg
                path_parts = Path(t['template_path'])
                thumbnail_filename = get_thumbnail_filename(path_parts.name)
                
                # Build thumbnail path maintaining subdirectory structure
                if path_parts.parent and str(path_parts.parent) != '.':
                    thumbnail_path = f"{path_parts.parent}/thumbnails/{thumbnail_filename}"
                else:
                    thumbnail_path = f"thumbnails/{thumbnail_filename}"
                
                # Normalize to forward slashes for URL
                thumbnail_url = f"{base_url}{thumbnail_path.replace(chr(92), '/')}"
            
            theme_groups.append(
                TemplateGroup(
                    template_group_id=t['template_id'],
                    template_name=t['template_name'],
                    category=t['template_group'],
                    template_file_name=t['template_path'],
                    thumbnail_url=thumbnail_url,
                    preview_url=preview_url
                )
            )
        
        return CustomThemesListResponse(
            success=True,
            templates=theme_groups,
            total=total,
            page=page,
            limit=limit
        )
        
    except Exception as e:
        logger.error("Error retrieving background templates: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve background templates"
        ) from e


@router.get(
    "/background-templates/{template_id}",
    response_model=TemplateGroup,
    summary="Get a single background template by ID",
    description="Retrieve detailed information about a specific background template.",
)
async def get_background_template(template_id: int):
    """Get a single background template by template ID."""
    try:
        template = bi_guidelines_service.get_background_template(template_id)
        
        if not template:
            raise HTTPException(
                status_code=404,
                detail=f"Background template with ID {template_id} not found"
            )
        
        return TemplateGroup(
            template_group_id=template['template_id'],
            template_name=template['template_name'],
            category=template['template_group'],
            template_file_name=template['template_path']
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error retrieving background template %d: %s", template_id, str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve background template"
        ) from e



@router.patch(
    "/background-templates/{template_id}",
    response_model=TemplateGroup,
    summary="Update a background template",
    description=(
        "Update template name or group.\n\n"
        "**Updatable fields:**\n"
        "- template_name\n"
        "- template_group\n\n"
        "**Note:** To change the image, delete and recreate the template."
    ),
)
async def update_background_template(
    template_id: int,
    template_name: Optional[str] = Form(None, min_length=3, max_length=100),
    template_group: Optional[str] = Form(None, min_length=3, max_length=100),
):
    """Update a background template's metadata (name/group only)."""
    try:
        # Update database
        updated_template = bi_guidelines_service.update_background_template(
            template_id=template_id,
            template_name=template_name,
            template_group=template_group
        )
        
        logger.info("Successfully updated background template %d", template_id)
        
        return TemplateGroup(
            template_group_id=updated_template['template_id'],
            template_name=updated_template['template_name'],
            category=updated_template['template_group'],
            template_file_name=updated_template['template_path']
        )
        
    except ValueError as e:
        logger.error("Validation error updating background template: %s", str(e))
        raise HTTPException(status_code=400, detail=str(e)) from e
    except HTTPException:
        raise
    except Exception as e:
        logger.error("Unexpected error updating background template %d: %s", template_id, str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to update background template"
        ) from e


@router.delete(
    "/background-templates/{template_id}",
    response_model=CustomThemeDeleteResponse,
    summary="Delete a background template",
    description=(
        "Delete a background template.\n\n"
        "This will:\n"
        "- Remove the record from nw_Templates table\n"
        "- Delete the physical image file from BackGrounds folder"
    ),
)
async def delete_background_template(
    template_id: int,
):
    """Delete a background template."""
    try:
        # Delete from database and get file path
        success, template_path = bi_guidelines_service.delete_background_template(
            template_id=template_id
        )
        
        # Delete physical file
        if success and template_path:
            file_path = settings.backgrounds_dir / template_path
            if file_path.exists():
                file_path.unlink()
                logger.info("Deleted physical file: %s", file_path)
        
        logger.info("Successfully deleted background template %d", template_id)
        
        return CustomThemeDeleteResponse(
            success=True,
            message="Background template deleted successfully",
            deleted_template_id=template_id
        )
        
    except ValueError as e:
        logger.error("Validation error deleting background template: %s", str(e))
        raise HTTPException(status_code=400, detail=str(e)) from e
    except HTTPException:
        raise
    except Exception as e:
        logger.error("Unexpected error deleting background template %d: %s", template_id, str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to delete background template"
        ) from e


# --- Data Models (Pydantic) ---
# Define the structure of the JSON that the API expects to receive.
