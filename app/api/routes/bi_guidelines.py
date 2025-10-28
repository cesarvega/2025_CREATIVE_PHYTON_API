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
from app.models.bi_guidelines_models import (
    NWMasterRequest,
    ReloadProjectSoundsRequest,
    ProjectUpdateRequest,
    BSRProjectUpdateRequest,
)
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
    TemplateGroupsResponse,
)
from app.services.bi_guidelines_service import bi_guidelines_service
from app.services.nsr_service import nsr_service
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

        logger.info("Retrieved BSR project info for project_id=%d", project_id)
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
        if rule_id < 101 or rule_id > 117 or rule_id == 107:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid rule_id: {rule_id}. Must be 101-117 (except 107)"
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
        "Create all NSR validation rules (101-117, except 107) for a new project.\n\n"
        "All rules are initialized in **ON** state (is_on=1) by default.\n\n"
        "**Rules created:**\n"
        "- **Contains:** 101, 102, 103, 104, 105\n"
        "- **USAN:** 106, 108, 116, 117\n"
        "- **Prefix:** 109, 110, 111, 112, 113, 114, 115\n\n"
        "**Total:** 16 rules (107 is skipped)\n\n"
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


# --- Data Models (Pydantic) ---
# Define the structure of the JSON that the API expects to receive.
