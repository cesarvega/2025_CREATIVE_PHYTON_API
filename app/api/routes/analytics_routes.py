"""API routes for Global Analytics Statistics.

This module provides endpoints for the Statistics tab functionality,
which generates analytics reports for ALL projects (not individual projects).

IMPORTANT: These endpoints are DIFFERENT from individual project report endpoints.
They do NOT require a presentation_id parameter - they fetch ALL projects.

Endpoints:
- GET /api/analytics/nw/projects - All NW projects analytics
- GET /api/analytics/nw/regions - NW regions aggregated analytics
- GET /api/analytics/bsr/projects - All BSR projects analytics
- GET /api/analytics/bsr/regions - BSR regions aggregated analytics
- POST /api/analytics/download - Generate Excel file
"""

from __future__ import annotations

from typing import Optional
from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from app.models.analytics_models import (
    AnalyticsDownloadRequest,
    BSRProjectAnalyticsResponse,
    BSRRegionAnalyticsResponse,
    NWProjectAnalyticsResponse,
    NWRegionAnalyticsResponse,
)
from app.services.analytics_excel_generator import analytics_excel_generator
from app.services.analytics_service import analytics_service
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/analytics", tags=["Analytics Statistics"])
logger = get_logger(__name__)


@router.get(
    "/nw/projects",
    response_model=NWProjectAnalyticsResponse,
    summary="Get analytics for ALL NW projects",
    description=(
        "Retrieve analytics data for ALL NW projects in the system.\n\n"
        "This endpoint is used by the Statistics tab to generate global reports.\n\n"
        "**IMPORTANT:** This is DIFFERENT from individual project reports:\n"
        "- Does NOT require presentation_id\n"
        "- Returns data for ALL projects\n"
        "- Calculates vote percentages and statistics\n"
        "- Used for Statistics tab Excel generation\n\n"
        "**Data Source:**\n"
        "- Queries BI_GUIDELINES.nw_master and nw_Details tables\n"
        "- Calls nw_GetPresentationId and nw_wdGetResults for additional data\n\n"
        "**Response includes:**\n"
        "- Project name, display name, active lead\n"
        "- Vote percentages (retained, positive, neutral, negative)\n"
        "- Number of newly created names\n"
        "- Active status\n\n"
        "**Use case:** Statistics tab in the NW section"
    ),
)
async def get_nw_projects_analytics() -> NWProjectAnalyticsResponse:
    """Get analytics for ALL NW projects.

    This endpoint retrieves analytics data for all NW projects in the system
    and calculates their statistics.

    Returns:
        NWProjectAnalyticsResponse with list of project analytics
    """
    try:
        logger.info("API: Fetching NW projects analytics")

        projects = await analytics_service.get_nw_projects_analytics()

        logger.info("API: Retrieved %d NW project analytics", len(projects))

        return NWProjectAnalyticsResponse(
            success=True,
            message=f"Successfully retrieved analytics for {len(projects)} NW projects",
            data=projects,
            total=len(projects)
        )

    except Exception as e:
        logger.error("API: Error fetching NW projects analytics: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to retrieve NW projects analytics: {str(e)}"
        ) from e


@router.get(
    "/nw/regions",
    response_model=NWRegionAnalyticsResponse,
    summary="Get aggregated analytics by region for NW projects",
    description=(
        "Retrieve aggregated analytics data by region for NW projects.\n\n"
        "This endpoint aggregates data across all NW projects and groups by region/lead.\n\n"
        "**Response includes:**\n"
        "- Region name\n"
        "- Active lead\n"
        "- Number of projects in region\n"
        "- Average vote percentages\n"
        "- Average names per project\n\n"
        "**Note:** Since external database (DayMaster/BNRS) is not available, "
        "the system aggregates by 'Active Lead' (UploadedBy) instead of true region.\n\n"
        "**Use case:** Statistics tab - Region-Specific sheet"
    ),
)
async def get_nw_regions_analytics() -> NWRegionAnalyticsResponse:
    """Get aggregated analytics by region for NW projects.

    Returns:
        NWRegionAnalyticsResponse with list of region analytics
    """
    try:
        logger.info("API: Fetching NW regions analytics")

        regions = await analytics_service.get_nw_regions_analytics()

        logger.info("API: Retrieved %d NW region analytics", len(regions))

        return NWRegionAnalyticsResponse(
            success=True,
            message=f"Successfully retrieved analytics for {len(regions)} NW regions",
            data=regions,
            total=len(regions)
        )

    except Exception as e:
        logger.error("API: Error fetching NW regions analytics: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to retrieve NW regions analytics: {str(e)}"
        ) from e


@router.get(
    "/bsr/projects",
    response_model=BSRProjectAnalyticsResponse,
    summary="Get analytics for ALL BSR projects",
    description=(
        "Retrieve analytics data for ALL BSR projects in the system.\n\n"
        "This endpoint is used by the Statistics tab to generate global BSR reports.\n\n"
        "**Response includes:**\n"
        "- Project name, display name, active lead\n"
        "- PC access count\n"
        "- Mobile access count\n"
        "- Total access count\n"
        "- Active status\n\n"
        "**Data Source:**\n"
        "- Queries BI_GUIDELINES.bsr_Master, bsr_ProjectConcepts, and BSR_CONCEPTNAMES tables\n\n"
        "**Use case:** Statistics tab in the BSR section"
    ),
)
async def get_bsr_projects_analytics() -> BSRProjectAnalyticsResponse:
    """Get analytics for ALL BSR projects.

    Returns:
        BSRProjectAnalyticsResponse with list of project analytics
    """
    try:
        logger.info("API: Fetching BSR projects analytics")

        projects = await analytics_service.get_bsr_projects_analytics()

        logger.info("API: Retrieved %d BSR project analytics", len(projects))

        return BSRProjectAnalyticsResponse(
            success=True,
            message=f"Successfully retrieved analytics for {len(projects)} BSR projects",
            data=projects,
            total=len(projects)
        )

    except Exception as e:
        logger.error("API: Error fetching BSR projects analytics: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to retrieve BSR projects analytics: {str(e)}"
        ) from e


@router.get(
    "/bsr/regions",
    response_model=BSRRegionAnalyticsResponse,
    summary="Get aggregated analytics by region for BSR projects",
    description=(
        "Retrieve aggregated analytics data by region for BSR projects.\n\n"
        "This endpoint aggregates data across all BSR projects and groups by region/lead.\n\n"
        "**Response includes:**\n"
        "- Region name\n"
        "- Active lead\n"
        "- Number of projects in region\n"
        "- Total PC access count\n"
        "- Total mobile access count\n"
        "- Total access count\n\n"
        "**Note:** Since external database is not available, "
        "the system aggregates by 'Active Lead' instead of true region.\n\n"
        "**Use case:** Statistics tab - Region-Specific sheet for BSR"
    ),
)
async def get_bsr_regions_analytics() -> BSRRegionAnalyticsResponse:
    """Get aggregated analytics by region for BSR projects.

    Returns:
        BSRRegionAnalyticsResponse with list of region analytics
    """
    try:
        logger.info("API: Fetching BSR regions analytics")

        regions = await analytics_service.get_bsr_regions_analytics()

        logger.info("API: Retrieved %d BSR region analytics", len(regions))

        return BSRRegionAnalyticsResponse(
            success=True,
            message=f"Successfully retrieved analytics for {len(regions)} BSR regions",
            data=regions,
            total=len(regions)
        )

    except Exception as e:
        logger.error("API: Error fetching BSR regions analytics: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to retrieve BSR regions analytics: {str(e)}"
        ) from e


@router.get(
    "/generate/{project_type}",
    summary="Generate and download Analytics Excel report (SIMPLE - ONE CALL)",
    description=(
        "Generate and download complete Analytics Excel report with a single API call.\n\n"
        "This endpoint does everything automatically:\n"
        "1. Fetches all projects analytics from database\n"
        "2. Calculates aggregated region statistics\n"
        "3. Generates formatted Excel file with 2 sheets\n"
        "4. Returns Excel file for download\n\n"
        "**Parameters:**\n"
        "- `project_type`: Either \"NW\" or \"BSR\"\n"
        "- `start_date` (optional): Filter projects from this date (format: YYYY-MM-DD)\n"
        "- `end_date` (optional): Filter projects up to this date (format: YYYY-MM-DD)\n\n"
        "**Excel file structure:**\n"
        "- **Sheet 1: Project-Specific** - All projects data (one row per project)\n"
        "- **Sheet 2: Region-Specific** - Aggregated region data (one row per region)\n\n"
        "**Filename format:** `{ProjectType}AnalyticsReport_YYYYMMDD.xlsx`\n\n"
        "**Example:**\n"
        "```\n"
        "GET /api/analytics/generate/NW\n"
        "→ Downloads NWAnalyticsReport_20251030.xlsx\n\n"
        "GET /api/analytics/generate/NW?start_date=2025-01-01&end_date=2025-12-31\n"
        "→ Downloads NWAnalyticsReport_20251030.xlsx (filtered by date range)\n"
        "```\n\n"
        "**Use case:** Statistics tab download button - just one click!"
    ),
)
async def generate_analytics_report(
    project_type: str,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None
):
    """Generate and download complete analytics report with one call.

    Args:
        project_type: "NW" or "BSR"
        start_date: Optional start date for filtering (YYYY-MM-DD)
        end_date: Optional end date for filtering (YYYY-MM-DD)

    Returns:
        StreamingResponse with Excel file

    Raises:
        HTTPException: If project_type is invalid or generation fails
    """
    # Validate project type
    project_type = project_type.upper()
    if project_type not in ["NW", "BSR"]:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid project_type: {project_type}. Must be 'NW' or 'BSR'"
        )

    try:
        logger.info(
            "API: Generating complete %s analytics report (start_date=%s, end_date=%s)",
            project_type,
            start_date,
            end_date
        )

        # Fetch data from database based on project type
        if project_type == "NW":
            projects = await analytics_service.get_nw_projects_analytics(start_date, end_date)
            regions = await analytics_service.get_nw_regions_analytics(start_date, end_date)
        else:  # BSR
            projects = await analytics_service.get_bsr_projects_analytics(start_date, end_date)
            regions = await analytics_service.get_bsr_regions_analytics(start_date, end_date)

        logger.info(
            "API: Fetched %d projects and %d regions for %s",
            len(projects),
            len(regions),
            project_type
        )

        # Convert Pydantic models to dicts for Excel generator
        projects_data = [p.model_dump() for p in projects]
        regions_data = [r.model_dump() for r in regions]

        # Generate Excel file
        excel_buffer, filename = analytics_excel_generator.generate_analytics_excel(
            project_type=project_type,
            projects_data=projects_data,
            regions_data=regions_data
        )

        logger.info("API: Excel file generated successfully: %s", filename)

        # Return as downloadable file
        return StreamingResponse(
            excel_buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename={filename}",
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )

    except ValueError as e:
        logger.error("API: Invalid request for analytics report: %s", str(e))
        raise HTTPException(
            status_code=400,
            detail=str(e)
        ) from e

    except Exception as e:
        logger.error("API: Error generating analytics report: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to generate {project_type} analytics report: {str(e)}"
        ) from e


@router.post(
    "/download",
    summary="Generate and download Analytics Excel report (LEGACY - 3 steps)",
    description=(
        "⚠️ **LEGACY ENDPOINT** - Use GET /api/analytics/generate/{project_type} instead.\n\n"
        "Generate Excel file with analytics statistics for all projects.\n\n"
        "**Workflow:**\n"
        "1. Client calls GET /api/analytics/{type}/projects to get projects data\n"
        "2. Client calls GET /api/analytics/{type}/regions to get regions data\n"
        "3. Client calls this endpoint with the data to generate Excel file\n\n"
        "**Request body:**\n"
        "```json\n"
        "{\n"
        '  "project_type": "NW",\n'
        '  "projects": [...],  // Data from GET /api/analytics/nw/projects\n'
        '  "regions": [...]    // Data from GET /api/analytics/nw/regions\n'
        "}\n"
        "```\n\n"
        "**Excel file structure:**\n"
        "- **Sheet 1: Project-Specific** - All projects data (one row per project)\n"
        "- **Sheet 2: Region-Specific** - Aggregated region data (one row per region)\n\n"
        "**Filename format:** `{ProjectType}AnalyticsReport_YYYYMMDD.xlsx`\n\n"
        "**Response:** Excel file download with proper headers"
    ),
)
async def download_analytics_excel(request: AnalyticsDownloadRequest):
    """Generate and download analytics Excel file.

    Args:
        request: AnalyticsDownloadRequest with project_type and data

    Returns:
        StreamingResponse with Excel file

    Raises:
        HTTPException: If Excel generation fails
    """
    try:
        logger.info(
            "API: Generating %s analytics Excel with %d projects and %d regions",
            request.project_type,
            len(request.projects),
            len(request.regions)
        )

        # Generate Excel file
        excel_buffer, filename = analytics_excel_generator.generate_analytics_excel(
            project_type=request.project_type,
            projects_data=request.projects,
            regions_data=request.regions
        )

        logger.info("API: Analytics Excel file generated successfully: %s", filename)

        # Return as downloadable file
        return StreamingResponse(
            excel_buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename={filename}",
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )

    except ValueError as e:
        logger.error("API: Invalid request for analytics Excel: %s", str(e))
        raise HTTPException(
            status_code=400,
            detail=str(e)
        ) from e

    except Exception as e:
        logger.error("API: Error generating analytics Excel: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to generate analytics Excel file: {str(e)}"
        ) from e
