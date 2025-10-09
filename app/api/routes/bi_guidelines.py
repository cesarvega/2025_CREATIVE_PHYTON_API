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
from app.models.bi_guidelines_models import NWMasterRequest
from app.models.presentation_models import PresentationData
from app.models.response_models import (
    ActivePresentationsResponse,
    DisplayNamesResponse,
)
from app.services.bi_guidelines_service import bi_guidelines_service
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/bi_guidelines", tags=["BI Guidelines"])
logger = get_logger(__name__)


class PresentationNotFoundError(Exception):
    """Custom exception for when a presentation is not found."""


@router.get(
    "/active-presentations",
    response_model=ActivePresentationsResponse,
    summary="Get paginated and filterable list of active presentations",
    description=(
        "Retrieve active presentations from BI_GUIDELINES database with optional search filtering and pagination. "
        "This endpoint is optimized for large datasets (5000+ presentations) and supports:\n\n"
        "- **Search filtering**: Filter by project name or display name (partial match)\n"
        "- **Pagination**: Control page number and results per page\n"
        "- **Sorted results**: Presentations are sorted by last update date (newest first)\n\n"
        "Ideal for autocomplete inputs, dropdowns with search, and infinite scroll implementations."
    ),
)
async def get_active_presentations(
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
    """Get paginated active presentations from BI_GUIDELINES database.

    This endpoint queries the nw_Master table for presentations with status 'OPEN'
    and returns them sorted by last update date in descending order.

    Example usage:
    - Get first 50 presentations: `GET /api/bi_guidelines/active-presentations`
    - Search for "SOLE": `GET /api/bi_guidelines/active-presentations?search=SOLE`
    - Get page 2 with 100 results: `GET /api/bi_guidelines/active-presentations?page=2&limit=100`
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


@router.get("/nw-master")
async def get_nw_master():
    """Return the top 1000 rows from the nw_Master table."""
    try:
        with get_connection_scope() as cursor:
            query = """
            SELECT TOP (1000) [PresentationId], [Project], [DisplayName], [MainPptFileName],
                [NameCandidateFileName], [NameCandidateBGType], [NameCandidateBGName],
                [NameCandidateStartingSlide], [UploadedBy], [UploadedDate], [PresentationStatus],
                [PresentationOpenDate], [NotesExplore], [NotesAvoid], [LastUpdateDate],
                [PresentationType], [BSRDisplayName], [isParticipantsVote],
                [show_EngKat_in_groups], [isWideScreenPPT], [isAWSLinkReq], [isWide]
            FROM [BI_GUIDELINES].[dbo].[nw_Master]
            """
            cursor.execute(query)
            columns = [column[0] for column in cursor.description]
            rows = cursor.fetchall()
            result = [dict(zip(columns, row)) for row in rows]
            
        logger.info("Retrieved %d records from nw_Master", len(result))
        return result

    except DatabaseConnectionError as exc:
        logger.error("Database connection failed: %s", exc)
        raise HTTPException(
            status_code=503,
            detail="Database service temporarily unavailable."
        ) from exc

    except DatabaseTransactionError as exc:
        logger.error("Database query failed: %s", exc)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve data from database."
        ) from exc


@router.post("/nw-master/validate")
async def validate_nw_master(data: NWMasterRequest):
    """Validate nw_master data using the Pydantic model."""

    return {"validated_data": data.dict()}


@router.post("/presentations/", status_code=201)
async def create_presentation(data: PresentationData):
    """Create a new presentation master record and associated detail records."""
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()

            # Log database connection info
            cursor.execute("SELECT DB_NAME(), @@SERVERNAME;")
            db_row = cursor.fetchone()
            if db_row:
                logger.debug("Connected to DB: %s on Server: %s", db_row[0], db_row[1])

            cursor.execute("SET NOCOUNT ON;")

            # Insert master record
            master_sql = """
                EXEC [dbo].[nw_InsertPresentationMaster_sep2025]
                    @Project=?, @DisplayName=?, @MainPptFileName=?, @NameCandidateFileName=?,
                    @NameCandidateBGType=?, @NameCandidateBGName=?, @NameCandidateStartingSlide=?,
                    @PresentationType=?, @UploadedBy=?, @BSRDisplayName=?,
                    @isParticipantsVote=?, @isWideScreenPPT=?, @isAWSLinkReq=?;
            """
            master_params = (
                data.project,
                data.display_name,
                data.powerpoint_file,
                data.excel_file,
                data.background_type,
                data.background_name,
                data.page_number,
                data.presentation_type,
                data.user_name,
                data.bsr_display_name,
                data.participant_vote,
                data.is_wide_ppt,
                data.is_aws_email,
            )

            logger.info("Executing master stored procedure for project: %s", data.project)
            cursor.execute(master_sql, master_params)

            # Retrieve the generated PresentationId
            presentation_id = None
            while True:
                try:
                    row = cursor.fetchone()
                    if row:
                        presentation_id = row[0]
                        logger.info("PresentationId created: %s", presentation_id)
                        break
                except pyodbc.Error:  # pylint: disable=c-extension-no-member
                    pass
                if not cursor.nextset():
                    break

            if not presentation_id:
                raise PresentationNotFoundError(
                    "Could not get PresentationId from the master record."
                )

            # Insert detail records
            detail_sql = """
                EXEC [dbo].[nw_InsertPresentationDetail_copy]
                    @PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?,
                    @SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?,
                    @NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?,
                    @TemplateId=?, @NameSubGroup=?;
            """
            for item in data.details:
                detail_params = (
                    presentation_id,
                    item.slide_number,
                    item.slide_type,
                    item.slide_bg_file_name,
                    item.slide_description,
                    item.group_name,
                    item.category,
                    item.name,
                    item.rationale,
                    item.notation,
                    item.kana,
                    item.logo_filename,
                    item.template_id,
                    item.name_sub_group,
                )
                cursor.execute(detail_sql, detail_params)

            cursor.close()
            # Connection will auto-commit on successful exit from context manager
            logger.info("Presentation %s committed to database successfully", presentation_id)

        return {
            "message": "Presentation created successfully.",
            "presentation_id": presentation_id,
        }

    except DatabaseConnectionError as exc:
        logger.error("Database connection failed: %s", exc)
        raise HTTPException(
            status_code=503,
            detail="Database service temporarily unavailable."
        ) from exc

    except PresentationNotFoundError as exc:
        logger.error("Presentation creation failed: %s", exc)
        raise HTTPException(
            status_code=500,
            detail="Failed to create presentation record."
        ) from exc

    except DatabaseTransactionError as exc:
        logger.error("Database transaction failed: %s", exc)
        raise HTTPException(
            status_code=500,
            detail="Failed to save presentation to database."
        ) from exc


@router.get("/presentations/{presentation_id}")
async def get_presentation(presentation_id: int):
    """Retrieve a presentation by its ID, including master and detail records."""
    try:
        with get_connection_scope() as cursor:
            # Get master record
            master_query = """
            SELECT [PresentationId], [Project], [DisplayName], [MainPptFileName],
                   [NameCandidateFileName], [NameCandidateBGType], [NameCandidateBGName],
                   [NameCandidateStartingSlide], [UploadedBy], [UploadedDate],
                   [PresentationStatus], [PresentationOpenDate], [NotesExplore],
                   [NotesAvoid], [LastUpdateDate], [PresentationType], [BSRDisplayName],
                   [isParticipantsVote], [show_EngKat_in_groups], [isWideScreenPPT],
                   [isAWSLinkReq], [isWide]
            FROM [BI_GUIDELINES].[dbo].[nw_Master]
            WHERE PresentationId = ?
            """
            cursor.execute(master_query, (presentation_id,))
            master_columns = [column[0] for column in cursor.description]
            master_row = cursor.fetchone()
            
            if not master_row:
                raise HTTPException(
                    status_code=404,
                    detail=f"Presentation with ID {presentation_id} not found."
                )
            
            master_data = dict(zip(master_columns, master_row))

            # Get detail records
            detail_query = """
            SELECT [PresentationId], [SlideNumber], [SlideType], [SlideBGFileName],
                   [SlideDescription], [NameGroup], [NameCategory], [Name],
                   [NameRationale], [NameNotation], [KanaNames], [NameLogo],
                   [TemplateId], [NameSubGroup]
            FROM [BI_GUIDELINES].[dbo].[nw_Details]
            WHERE PresentationId = ?
            ORDER BY SlideNumber
            """
            cursor.execute(detail_query, (presentation_id,))
            detail_columns = [column[0] for column in cursor.description]
            detail_rows = cursor.fetchall()
            detail_data = [dict(zip(detail_columns, row)) for row in detail_rows]

        logger.info("Retrieved presentation %s with %d detail records", presentation_id, len(detail_data))
        return {"master": master_data, "details": detail_data}

    except HTTPException:
        raise

    except DatabaseConnectionError as exc:
        logger.error("Database connection failed: %s", exc)
        raise HTTPException(
            status_code=503,
            detail="Database service temporarily unavailable."
        ) from exc

    except DatabaseTransactionError as exc:
        logger.error("Database query failed for presentation %s: %s", presentation_id, exc)
        raise HTTPException(
            status_code=500,
            detail="Failed to retrieve presentation from database."
        ) from exc


# --- Data Models (Pydantic) ---
# Define the structure of the JSON that the API expects to receive.
