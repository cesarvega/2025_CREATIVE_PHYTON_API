"""Routes for presentation creation and assembly orchestration."""

import base64
import time
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile, Form
from fastapi.responses import FileResponse
from urllib.parse import quote

from app.api.dependencies import (
    parse_build_metadata,
    parse_presentation_metadata,
    parse_simple_dw_metadata,
    parse_bsr_metadata,
)
from app.config.settings import settings
from app.models.presentation_models import (
    CreatePresentationMetadata,
    CreatePresentationRequest,
    CreatePresentationResponse,
    SimpleDWMetadata,
    PresentationBuildMetadata,
    PresentationBuildResponse,
    GenerateBackupRequest,
    GenerateBackupResponse,
    BSRCreatePresentationMetadata,
    BSRCreatePresentationResponse,
)
from app.models.nw_reports_models import (
    CreateFeedbackTemplateRequest,
    CreateFeedbackTemplateResponse,
    DownloadResultsRequest,
    DownloadResultsResponse,
)
from app.models.bsr_reports_models import (
    BSRGenerateReportRequest,
    BSRGenerateReportResponse,
)
from app.services.bsr_report_orchestrator_service import bsr_report_orchestrator_service as bsr_report_orchestrator
from app.models.response_models import ReplaceProjectImagesResponse, CreateSimpleDWResponse
from app.services.presentation_service import presentation_service
from app.services.report_orchestrator_service import report_orchestrator_service
from app.services.feedback_template_generator import feedback_template_generator
from app.services.bi_guidelines_service import bi_guidelines_service
from app.services.pptx_service import pptx_service
from app.config.db import get_connection_scope
from app.utils.download_utils import (
    build_api_download_url,
    create_download_token,
    decode_download_token,
    guess_media_type,
)
from pathlib import Path
from app.utils.logging_utils import get_logger
from app.api.dependencies import (
    validate_project_type,
    validate_pptx_file,
    validate_file_size,
)

router = APIRouter(prefix="/presentations", tags=["Presentation Creation"])
logger = get_logger(__name__)


@router.post(
    "/create",
    response_model=CreatePresentationResponse,
    summary="Create a complete presentation from Excel + PPTX inputs with physical file generation",
    description=(
        "Upload the Excel candidate workbook, a base PPTX template, and a metadata JSON payload "
        "to generate a fully populated presentation. This endpoint:\n\n"
        "1. **Processes Excel data** - Extracts names, categories, groups, and rationales\n"
        "2. **Converts PPTX to images** - Generates JPG images from each slide of the original PPTX\n"
        "3. **Generates slide metadata** - Creates detailed slide information for database storage\n"
        "4. **Applies template rotation** (optional) - Rotates through specified templates for variety\n"
        "5. **Generates physical PowerPoint file** - Combines original PPTX slides with template-generated slides:\n"
        "   - Original PPTX slides (before `page_number`)\n"
        "   - Template-generated slides from Excel data (groups, individuals, multi-name slides)\n"
        "   - Original PPTX slides (after generated slides)\n"
        "6. **Persists to database** - Saves presentation metadata and slide details to nw_Master and nw_Details\n"
        "7. **Returns complete response** - Includes presentation ID, generated file paths, and processing stats\n\n"
        "The generated PowerPoint file is a complete, ready-to-present deck that seamlessly integrates "
        "your original slides with dynamically generated content from the Excel data."
    ),
    response_description=(
        "Creation status with presentation ID, total slide count, processing time, "
        "and paths to generated PowerPoint files (.pptx and optional .pptm)."
    ),
    responses={
        400: {
            "description": "Invalid or missing files provided in the multipart request.",
            "content": {
                "application/json": {
                    "example": {"detail": "Excel file must be .xlsx or .xls"}
                }
            },
        },
        422: {
            "description": "Metadata payload failed validation.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": [
                            {
                                "loc": ["metadata", "project"],
                                "msg": "Field required",
                                "type": "missing",
                            }
                        ]
                    }
                }
            },
        },
        500: {
            "description": "Unexpected error during orchestration.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to create presentation: PowerPoint automation unavailable"
                    }
                }
            },
        },
    },
)
async def create_presentation(
    metadata: CreatePresentationMetadata = Depends(parse_presentation_metadata),
    excel_file: UploadFile = File(..., description="Excel file (.xlsx or .xls)"),
    pptx_file: UploadFile = File(..., description="PowerPoint file (.pptx)"),
) -> CreatePresentationResponse:
    """Create a presentation by combining JSON metadata with uploaded template files.

    The client must submit a `metadata` field containing a JSON object with all
    presentation parameters (project name, display name, background settings, flags,
    etc.) together with two file uploads: the Excel candidate sheet (`excel_file`) and
    the PowerPoint template (`pptx_file`).

    **Optional Physical PowerPoint Generation:**
    If `generate_physical_pptx=true` in metadata, the service will also generate a
    complete physical PowerPoint file combining:
    - Original slides from the uploaded PPTX (before generated content)
    - Dynamically generated slides from Excel data using templates
    - Original slides from the uploaded PPTX (after generated content)
    
    This requires additional template configuration fields:
    - `template_pack`: Template directory name (e.g., "BackgroundTemplates")
    - `base_template`, `multi_template`, `group_template`, `separator_template`, `summary_template`
    - `include_macro_version`: Generate .pptm file (default: false)

    The service then executes the full pipeline:
    Excel processing â†’ PPTX conversion â†’ slide generation â†’ template application â†’
    database persistence â†’ [optional] physical PowerPoint generation.
    """
    try:
        # Validate Excel file
        if not excel_file.filename or not excel_file.filename.lower().endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="Excel file must be .xlsx or .xls")
        
        # Validate PPTX file
        if not pptx_file.filename or not pptx_file.filename.lower().endswith(".pptx"):
            raise HTTPException(status_code=400, detail="PPTX file must be .pptx")
        
        # Read file contents with size validation
        excel_content = await excel_file.read()
        pptx_content = await pptx_file.read()
        
        # Validate sizes
        max_excel_size = 10 * 1024 * 1024  # 10MB
        max_pptx_size = 50 * 1024 * 1024   # 50MB
        
        if len(excel_content) == 0:
            raise HTTPException(status_code=400, detail="Empty Excel file provided")
        if len(excel_content) > max_excel_size:
            raise HTTPException(status_code=413, detail="Excel file too large. Maximum size is 10MB")
            
        if len(pptx_content) == 0:
            raise HTTPException(status_code=400, detail="Empty PPTX file provided")
        if len(pptx_content) > max_pptx_size:
            raise HTTPException(status_code=413, detail="PPTX file too large. Maximum size is 50MB")

        # Create request object from metadata + files
        request: CreatePresentationRequest = metadata.to_service_request(
            excel_content=excel_content,
            excel_filename=excel_file.filename,
            pptx_content=pptx_content,
            pptx_filename=pptx_file.filename,
        )

        logger.info(
            "Starting presentation creation for project: %s, display_name: %s",
            metadata.project,
            metadata.display_name,
        )

        start_time = time.time()

        # Orchestrate complete presentation creation
        result = presentation_service.create_presentation(request)

        processing_time = time.time() - start_time

        logger.info(
            "Presentation created successfully. ID: %s, Slides: %d, Time: %.2fs",
            result.presentation_id,
            result.total_slides,
            processing_time,
        )

        return result

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error creating presentation: %s", str(e))
        raise HTTPException(
            status_code=500, detail=f"Failed to create presentation: {str(e)}"
        ) from e


@router.post(
    "/replace-project-images/",
    summary="Replace slide images for an existing project",
    response_model=ReplaceProjectImagesResponse,
)
async def replace_project_images(
    powerpointFile: UploadFile = File(..., description="PowerPoint file (.pptx)"),
    project_name: str = Form(..., description="Existing project folder name"),
    project_type: str = Depends(validate_project_type),
):
    """Replace images and thumbnails in an existing project folder.

    - project_name: Must match an existing folder under the project type base dir.
    - project_type: 'bipresents' or 'nw' (used to resolve base directory).
    - powerpoint file: Uploaded PPTX used to export new slide images.
    """
    try:
        # Validate file type and size (50MB default for PPTX)
        validated_file = await validate_pptx_file(powerpointFile)
        file_content = await validate_file_size(validated_file, 50)

        # Basic validation for project_name (no path separators or whitespace trim changes)
        if not project_name or project_name.strip() != project_name or ("/" in project_name or "\\" in project_name):
            raise HTTPException(status_code=400, detail="Invalid project name")

        result = pptx_service.replace_project_images(
            file_content=file_content,
            filename=validated_file.filename,
            project_name=project_name,
            project_type=project_type,
        )

        return ReplaceProjectImagesResponse(
            message=result["message"],
            project_name=result["project_name"],
            project_type=result["project_type"],
            total_images=result["total_images"],
            images=result["images"],
            thumbnails=result.get("thumbnails", []),
        )

    except FileNotFoundError as fnf:
        raise HTTPException(status_code=404, detail=str(fnf)) from fnf
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve)) from ve
    except Exception as e:
        logger.error("Error replacing project images: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Replace images error: {str(e)}"
        ) from e


@router.post(
    "/create-simple-dw",
    summary="Create DW slide images from a PPTX using simplified metadata",
    response_model=CreateSimpleDWResponse,
)
async def create_simple_dw_presentation(
    metadata: SimpleDWMetadata = Depends(parse_simple_dw_metadata),
    powerpointFile: UploadFile = File(..., description="PowerPoint file (.pptx)"),
):
    """Simplified create endpoint for DW that only requires a PPTX.

    This endpoint reuses the image export logic from the presentations routes,
    saving all slide images into the existing DW project's folder:
    C:\\inetpub\\wwwroot\\nw2\\nw_slides\\{projectName}

    Notes:
    - projectName must match an existing folder under nw2/nw_slides
    - displayName and widePresentation are accepted for compatibility/logging
    - Files are exported as 001.jpg, 002.jpg, ... and Thumbnails/...
    """
    try:
        validated_file = await validate_pptx_file(powerpointFile)
        file_content = await validate_file_size(validated_file, 50)

        projectName = metadata.projectName
        displayName = metadata.displayName
        userName = metadata.userName
        slideType = (metadata.slideType or "Image").strip() or "Image"
        presentationType = (metadata.presentationType or "Design").strip() or "Design"

        # Basic validation for projectName (disallow path separators and trimmed changes)
        if (
            not projectName
            or projectName.strip() != projectName
            or ("/" in projectName or "\\" in projectName)
        ):
            raise HTTPException(status_code=400, detail="Invalid project name")

        logger.info(
            "DW create-simple request: projectName='%s', displayName='%s', user='%s'",
            projectName,
            displayName,
            userName,
        )

        # Generate images under displayName folder (nw_slides/[displayName]/001.jpg)
        result = pptx_service.convert_pptx_to_images(
            file_content,
            validated_file.filename,
            metadata.displayName,
            settings.PROJECT_TYPE_DW,
        )

        # After generating images, record slide details with SlideType='Design'
        # into nw_InsertPresentationDetail_copy for the matching presentation
        presentation_id: Optional[int] = None
        try:
            with get_connection_scope(timeout=30) as cursor:
                # Resolve PresentationId from nw_Master via Project + DisplayName
                cursor.execute(
                    (
                        "SELECT TOP 1 PresentationId "
                        "FROM [BI_GUIDELINES].[dbo].[nw_Master] "
                        "WHERE Project = ? AND DisplayName = ? "
                        "ORDER BY PresentationId DESC"
                    ),
                    (projectName, displayName),
                )
                row = cursor.fetchone()
                if not row:
                    # If no master exists, create a minimal master record (Design/DW)
                    master_sql = (
                        "EXEC [dbo].[nw_InsertPresentationMaster_sep2025] "
                        "@Project=?, @DisplayName=?, @MainPptFileName=?, @NameCandidateFileName=?, "
                        "@NameCandidateBGType=?, @NameCandidateBGName=?, @NameCandidateStartingSlide=?, "
                        "@PresentationType=?, @UploadedBy=?, @BSRDisplayName=?, "
                        "@isParticipantsVote=?, @isWideScreenPPT=?, @isAWSLinkReq=?;"
                    )
                    master_params = (
                        projectName,
                        displayName,
                        validated_file.filename or "",
                        "",  # NameCandidateFileName not used in DW simple flow
                        "Default",  # BG type
                        "",  # BG name
                        1,  # Starting slide
                        presentationType,  # PresentationType from frontend (e.g., 'Design')
                        userName or "",
                        "",  # BSRDisplayName
                        0,  # isParticipantsVote
                        1 if getattr(metadata, "widePresentation", False) else 0,  # isWideScreenPPT
                        0,  # isAWSLinkReq
                    )

                    cursor.execute(master_sql, master_params)

                    # Try to fetch PresentationId from result set(s)
                    while True:
                        try:
                            created_row = cursor.fetchone()
                            if created_row:
                                presentation_id = int(created_row[0])
                                break
                        except Exception:
                            # Some drivers/SPs return no row; proceed to nextset
                            pass
                        if not cursor.nextset():
                            break

                    # Fallback lookup if SP did not return ID
                    if not presentation_id:
                        cursor.execute(
                            (
                                "SELECT TOP 1 PresentationId "
                                "FROM [BI_GUIDELINES].[dbo].[nw_Master] "
                                "WHERE Project = ? AND DisplayName = ? "
                                "ORDER BY PresentationId DESC"
                            ),
                            (projectName, displayName),
                        )
                        row2 = cursor.fetchone()
                        if not row2:
                            raise HTTPException(
                                status_code=404,
                                detail=(
                                    "Presentation master not found and could not be created"
                                ),
                            )
                        presentation_id = int(row2[0])
                else:
                    presentation_id = int(row[0])

                # Insert a detail row for each generated image with the requested SlideType
                insert_sql = (
                    "EXEC [dbo].[nw_InsertPresentationDetail_copy] "
                    "@PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?, "
                    "@SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?, "
                    "@NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?, "
                    "@TemplateId=?, @NameSubGroup=?;"
                )

                total = int(result.get("total_images", 0))
                for idx in range(1, total + 1):
                    # Use relative image path as background filename reference
                    # and store userName in SlideDescription for traceability
                    rel_path = result.get("images", [])[idx - 1] if result.get("images") else ""
                    cursor.execute(
                        insert_sql,
                        (
                            presentation_id,  # @PresentationId
                            idx,              # @SlideNumber
                            slideType,        # @SlideType
                            rel_path,         # @SlideBGFileName
                            "",              # @SlideDescription (left empty per DW spec)
                            "", "", "", "", "", "", "",  # group/category/name fields
                            0,                # @TemplateId
                            "",              # @NameSubGroup
                        ),
                    )
        except HTTPException:
            raise
        except Exception as db_exc:
            logger.error("Failed to insert DW detail records: %s", str(db_exc), exc_info=True)
            raise HTTPException(
                status_code=500,
                detail=f"Failed to record DW slide details: {str(db_exc)}",
            ) from db_exc

        return CreateSimpleDWResponse(
            message=result["message"],
            project_name=result.get("project_name", projectName),
            display_name=displayName,
            project_type=result.get("project_type", settings.PROJECT_TYPE_DW),
            slide_type=slideType,
            user_name=userName,
            total_images=result["total_images"],
            images=result["images"],
            thumbnails=result.get("thumbnails", []),
            presentation_id=presentation_id,
        )

    except FileNotFoundError as fnf:
        raise HTTPException(status_code=404, detail=str(fnf)) from fnf
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve)) from ve
    except Exception as e:
        logger.error(
            "Error creating DW images from PPTX: %s", str(e), exc_info=True
        )
        raise HTTPException(
            status_code=500, detail=f"DW create-simple error: {str(e)}"
        ) from e


@router.post(
    "/create-bsr",
    summary="Create BSR presentation from PowerPoint file",
    response_model=BSRCreatePresentationResponse,
    description=(
        "Create a new BSR (Board Sales Request) presentation by uploading a PowerPoint file.\n\n"
        "This endpoint:\n"
        "1. Validates presentation doesn't already exist\n"
        "2. Validates display name hasn't been used\n"
        "3. Converts PowerPoint slides to images\n"
        "4. Reads slide titles from the PowerPoint\n"
        "5. Creates presentation record in database\n"
        "6. Inserts slide details including summary slide at specified position\n\n"
        "**Slide Insertion Logic:**\n"
        "If the PowerPoint has 5 slides and slide_number=3:\n"
        "- Slide 1 (Image) -> Slide #1\n"
        "- Slide 2 (Image) -> Slide #2\n"
        "- Summary (NameSummary) -> Slide #3 (inserted at slide_number)\n"
        "- Slide 3 (Image) -> Slide #4\n"
        "- Slide 4 (Image) -> Slide #5\n"
        "- Slide 5 (Image) -> Slide #6\n"
        "Total: 6 slides (5 original + 1 summary)\n\n"
        "**Data Storage:**\n"
        "- Master record: bsr_InsertPresentationMaster\n"
        "- Detail records: bsr_InsertPresentationDetail (one per slide)\n"
        "- Images stored in: BRS_slides/{ProjectName}/\n"
    ),
    responses={
        200: {
            "description": "BSR presentation created successfully",
            "content": {
                "application/json": {
                    "example": {
                        "message": "BSR Presentation created successfully",
                        "presentation_id": 12345,
                        "total_slides": 6,
                        "processing_time_seconds": 8.45,
                    }
                }
            },
        },
        400: {
            "description": "Invalid request or presentation already exists",
            "content": {
                "application/json": {
                    "examples": {
                        "already_exists": {
                            "summary": "Presentation already exists",
                            "value": {
                                "detail": "Presentation already exists for project 'BSR_Project' with display name 'BSR_Presentation'"
                            }
                        },
                        "display_name_used": {
                            "summary": "Display name already used",
                            "value": {
                                "detail": "Display name 'BSR_Presentation' has already been used"
                            }
                        }
                    }
                }
            },
        },
        422: {
            "description": "Metadata validation failed",
            "content": {
                "application/json": {
                    "example": {
                        "detail": [
                            {
                                "loc": ["metadata", "slide_number"],
                                "msg": "Field required",
                                "type": "missing",
                            }
                        ]
                    }
                }
            },
        },
        500: {
            "description": "Unexpected error during creation",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to create BSR presentation: Database connection error"
                    }
                }
            },
        },
    },
)
async def create_bsr_presentation(
    metadata: BSRCreatePresentationMetadata = Depends(parse_bsr_metadata),
    pptx_file: UploadFile = File(..., description="PowerPoint file (.pptx)"),
) -> BSRCreatePresentationResponse:
    """Create a BSR presentation from a PowerPoint file.

    This endpoint handles the complete BSR presentation creation workflow:
    1. Validation checks (existence, display name uniqueness)
    2. PowerPoint conversion to images
    3. Slide title extraction
    4. Database persistence

    Args:
        metadata: BSR presentation metadata (project_name, display_name, etc.)
        pptx_file: PowerPoint file to process

    Returns:
        BSRCreatePresentationResponse with presentation ID and statistics

    Raises:
        HTTPException 400: If validation fails or presentation already exists
        HTTPException 500: If creation process fails
    """
    try:
        # Validate file
        if not pptx_file.filename or not pptx_file.filename.lower().endswith(".pptx"):
            raise HTTPException(status_code=400, detail="PPTX file must be .pptx")

        # Read file content
        pptx_content = await pptx_file.read()

        # Validate size (max 50MB for PowerPoint files)
        max_pptx_size = 50 * 1024 * 1024
        if len(pptx_content) == 0:
            raise HTTPException(status_code=400, detail="Empty PPTX file provided")
        if len(pptx_content) > max_pptx_size:
            raise HTTPException(status_code=413, detail="PPTX file too large. Maximum size is 50MB")

        logger.info(
            "Starting BSR presentation creation: project=%s, display_name=%s",
            metadata.project_name,
            metadata.display_name,
        )

        start_time = time.time()

        # Call service to create BSR presentation
        result = presentation_service.create_bsr_presentation(
            project_name=metadata.project_name,
            display_name=metadata.display_name,
            slide_number=metadata.slide_number,
            presentation_type=metadata.presentation_type,
            user_name=metadata.user_name,
            is_wide_ppt=metadata.is_wide_ppt,
            pptx_content=pptx_content,
            pptx_filename=pptx_file.filename,
        )

        processing_time = time.time() - start_time

        logger.info(
            "BSR presentation created successfully. ID: %s, Slides: %d, Time: %.2fs",
            result.get("presentation_id"),
            result.get("total_slides", 0),
            processing_time,
        )

        return BSRCreatePresentationResponse(
            message="BSR Presentation created successfully",
            presentation_id=result.get("presentation_id"),
            total_slides=result.get("total_slides", 0),
            processing_time_seconds=processing_time,
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error creating BSR presentation: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Failed to create BSR presentation: {str(e)}"
        ) from e


@router.post(
    "/createTemplate",
    response_model=PresentationBuildResponse,
    summary="Assemble presentation files directly from an Excel candidate sheet",
    description=(
        "Generate PPTX artifacts using metadata and an Excel workbook without uploading a template file. "
        "The metadata defines template packs, slide ranges, and output behaviour. Optionally returns base64 "
        "content or download URLs for the generated files."
    ),
    response_description="Generated artifact metadata including slide range and file locations.",
    responses={
        400: {
            "description": "Invalid Excel file or missing templates on disk.",
            "content": {
                "application/json": {
                    "example": {"detail": "Excel file must be .xlsx or .xls"}
                }
            },
        },
        422: {
            "description": "Metadata payload failed validation (slide ranges, template pack, etc.).",
            "content": {
                "application/json": {
                    "example": {
                        "detail": [
                            {
                                "loc": ["metadata", "slide_end"],
                                "msg": "slide_end must be greater than or equal to slide_start",
                                "type": "value_error",
                            }
                        ]
                    }
                }
            },
        },
        500: {
            "description": "Unexpected error while composing the presentation files.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to assemble presentation from provided inputs."
                    }
                }
            },
        },
    },
)
async def build_presentation_files(
    metadata: PresentationBuildMetadata = Depends(parse_build_metadata),
    excel_file: UploadFile = File(..., description="Excel file (.xlsx or .xls)"),
) -> PresentationBuildResponse:
    try:
        # Validate Excel file
        if not excel_file.filename or not excel_file.filename.lower().endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="Excel file must be .xlsx or .xls")
        
        # Read file content with size validation
        excel_content = await excel_file.read()
        
        # Validate size
        max_excel_size = 10 * 1024 * 1024  # 10MB
        if len(excel_content) == 0:
            raise HTTPException(status_code=400, detail="Empty Excel file provided")
        if len(excel_content) > max_excel_size:
            raise HTTPException(status_code=413, detail="Excel file too large. Maximum size is 10MB")

        build_request = metadata.to_creation_request(
            excel_content=excel_content,
            excel_filename=excel_file.filename,
            pptx_filename=metadata.base_template,
        )
        build_options = metadata.to_build_options()

        start_time = time.time()
        artifacts = presentation_service.build_presentation_files(
            build_request=build_request,
            options=build_options,
        )
        processing_time = time.time() - start_time

        printable_path = (
            str(artifacts.printable_path)
            if artifacts.printable_path is not None
            else None
        )
        macro_path = (
            str(artifacts.macro_path) if artifacts.macro_path is not None else None
        )

        printable_base64 = None
        macro_base64 = None

        if metadata.return_bytes:
            if artifacts.printable_path and artifacts.printable_path.exists():
                printable_base64 = base64.b64encode(
                    artifacts.printable_path.read_bytes()
                ).decode("utf-8")
            if artifacts.macro_path and artifacts.macro_path.exists():
                macro_base64 = base64.b64encode(
                    artifacts.macro_path.read_bytes()
                ).decode("utf-8")

        logger.info(
            "Presentation assembly finished for %s/%s in %.2fs (slides=%d)",
            metadata.project,
            metadata.display_name,
            processing_time,
            artifacts.total_slides,
        )

        return PresentationBuildResponse(
            message=f"Presentation assembled in {processing_time:.2f}s",
            total_slides=artifacts.total_slides,
            slide_start=artifacts.slide_start,
            slide_end=artifacts.slide_end,
            printable_pptx=printable_path,
            macro_pptx=macro_path,
            printable_base64=printable_base64,
            macro_base64=macro_base64,
            download_urls=artifacts.download_urls,
            api_downloads={
                key: build_api_download_url(path)
                for key, path in (
                    ("printable", artifacts.printable_path),
                    ("macro", artifacts.macro_path),
                )
                if path is not None
            },
            summary=artifacts.summary,
            warnings=artifacts.warnings,
        )

    except HTTPException:
        raise
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:  # pylint: disable=broad-except
        logger.error("Error assembling presentation: %s", exc)
        raise HTTPException(
            status_code=500,
            detail="Failed to assemble presentation from provided inputs.",
        ) from exc


@router.get(
    "/exists",
    summary="Check if a presentation exists",
    response_description="Returns a boolean indicating if the presentation exists.",
    responses={
        200: {
            "description": "Check completed successfully.",
            "content": {
                "application/json": {
                    "example": {"exists": True}
                }
            },
        },
        500: {
            "description": "Database error during check.",
            "content": {
                "application/json": {
                    "example": {"detail": "Database error while checking for presentation."}
                }
            },
        },
    },
)
async def check_presentation_exists(
    project_name: str,
    display_name: str,
    exclude_id: Optional[int] = None,
) -> dict[str, bool]:
    """
    Check if a presentation with the given project name and display name already exists.
    An optional `exclude_id` can be provided to exclude a specific presentation ID
    from the check, which is useful for update operations.
    """
    try:
        exists = presentation_service.presentation_exists(
            project_name=project_name,
            display_name=display_name,
            exclude_id=exclude_id,
        )
        return {"exists": exists}
    except HTTPException:
        raise

@router.get(
    "/files/{token}",
    summary="Download a generated presentation artifact",
    responses={
        200: {
            "description": "Binary contents of the requested presentation artifact.",
            "content": {"application/octet-stream": {"schema": {"type": "string", "format": "binary"}}},
        },
        400: {"description": "Invalid download token."},
        403: {"description": "Token does not reference an authorized file."},
        404: {"description": "Requested file not found."},
    },
)
async def download_generated_file(token: str) -> FileResponse:
    """Download a generated presentation file using a secure token.
    
    Args:
        token: The secure download token encoding the file path.
        
    Returns:
        FileResponse with the requested file.
    """
    file_path = decode_download_token(token)
    media_type = guess_media_type(file_path)

    # Build response and force attachment filename to avoid browsers/proxies
    # inferring a name based on the URL token or changing extensions.
    response = FileResponse(
        path=file_path,
        media_type=media_type,
        filename=file_path.name,  # best-effort for Starlette versions that support it
    )
    # Force Content-Disposition with RFC 5987 UTF-8 filename
    utf8_name = quote(file_path.name)
    response.headers["Content-Disposition"] = (
        f"attachment; filename=\"{file_path.name}\"; filename*=UTF-8''{utf8_name}"
    )
    # Prevent type sniffing that may alter extensions
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Content-Transfer-Encoding"] = "binary"
    return response


@router.post(
    "/test-slide-generation",
    summary="Test slide generation with multi-candidate layout",
    description=(
        "Test endpoint to verify slide generation logic in isolation. "
        "Allows testing the multi-candidate name layout without going through the full pipeline."
    ),
)
async def test_slide_generation(
    template_name: str = "template_default_withgroups2019.pptx",
    names: str = "John Doe##Jane Smith##Bob Johnson##Alice Williams##Charlie Brown##David Miller##Emma Davis##Frank Wilson##Grace Martinez##Henry Anderson##Ivy Thomas##Jack Taylor##Kelly Moore##Liam Jackson##Mia White##Noah Harris##Olivia Martin##Peter Thompson##Quinn Garcia##Rachel Robinson##Steve Clark##Tina Rodriguez##Uma Lewis##Victor Lee##Wendy Walker##Xavier Hall##Yara Allen##Zack Young",
) -> dict:
    """Test slide generation with specified template and names.

    Args:
        template_name: Name of the template file to use (must exist in templates directory)
        names: Delimited string of names to layout (use ## as delimiter)

    Returns:
        Result of the slide generation test including success status and file path
    """
    try:
        from app.services.pptx_builder_service import PPTXBuilderService, tokenize_delimited_block
        from app.utils.path_utils import resolve_project_output
        import pythoncom

        # Initialize COM for this thread
        pythoncom.CoInitialize()

        try:
            # Parse names
            names_list = tokenize_delimited_block(names)
            if not names_list:
                raise HTTPException(status_code=400, detail="No names provided")

            logger.info("Testing slide generation with %d names", len(names_list))

            # Initialize builder service
            templates_dir = Path(settings.templates_dir) if hasattr(settings, 'templates_dir') else Path("templates")
            subdir = "BackgroundDefaultTemplate"
            templates_dir = templates_dir / subdir
            template_path = templates_dir / template_name
            
            print(template_path)

            if not template_path.exists():
                raise HTTPException(
                    status_code=404,
                    detail=f"Template not found: {template_name}. Available templates are in {templates_dir}"
                )

            builder = PPTXBuilderService(
                templates_dir=templates_dir,
                base_template_path=template_path,
            )

            # Create output directory
            output_dir, _ = resolve_project_output("test_slide_generation", None)
            output_dir.mkdir(parents=True, exist_ok=True)

            # Generate a test presentation with just one slide
            from pptx import Presentation
            import win32com.client

            # Load template
            prs = Presentation(str(template_path))

            # Use first slide or add one
            if prs.slides:
                slide = prs.slides[0]
            else:
                slide = prs.slides.add_slide(prs.slide_layouts[5])

            # Save the presentation first
            test_output_path = output_dir / f"test_slide_{int(time.time())}.pptx"
            prs.save(str(test_output_path))

            # Now open with COM automation to test the layout logic
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1

            presentation = powerpoint.Presentations.Open(str(test_output_path.absolute()))
            com_slide = presentation.Slides(1)

            # Collect placeholder shapes (use Name Candidate for template_default_withgroups2019.pptx)
            placeholder_shapes = builder._collect_placeholder_shapes(com_slide, "Name Candidate")

            # Debug: List all text shapes in the slide
            logger.info("=== DIAGNOSTIC: Analyzing slide shapes ===")
            all_text_shapes = list(builder._iter_text_shapes(com_slide))
            logger.info("Total text shapes found: %d", len(all_text_shapes))

            for idx, shape in enumerate(all_text_shapes[:20]):  # Limit to first 20
                try:
                    text = shape.TextFrame.TextRange.Text
                    logger.info("  Shape %d: '%s'", idx + 1, text[:50] if text else "(empty)")
                except Exception as e:
                    logger.info("  Shape %d: (no text - %s)", idx + 1, str(e))

            if not placeholder_shapes:
                logger.warning("No 'Name Candidate' placeholder shapes found in template")
                warnings_list = ["No 'Name Candidate' placeholder shapes found in template. Check logs for all text shapes found."]
            else:
                logger.info("Found %d 'Name Candidate' placeholder shapes", len(placeholder_shapes))
                warnings_list = []

                # Log the placeholder shapes found
                for idx, pshape in enumerate(placeholder_shapes):
                    try:
                        ptext = pshape.TextFrame.TextRange.Text
                        logger.info("  Placeholder %d: '%s'", idx + 1, ptext[:30])
                    except:
                        pass

            # Test the simplified placeholder assignment
            builder._assign_names_to_placeholders(
                placeholder_shapes=placeholder_shapes or [],
                names=names_list,
                context="test slide",
            )
            success = True

            # Save and close properly
            try:
                presentation.Save()
            except Exception as e:
                logger.warning("Could not save presentation: %s", e)

            try:
                presentation.Close()
            except Exception as e:
                logger.warning("Could not close presentation: %s", e)

            try:
                powerpoint.Quit()
            except Exception as e:
                logger.warning("Could not quit PowerPoint: %s", e)

            logger.info("Test slide generation completed: success=%s", success)

            return {
                "success": success,
                "names_count": len(names_list),
                "names": names_list[:10],  # First 10 names for brevity
                "placeholder_count": len(placeholder_shapes) if placeholder_shapes else 0,
                "output_path": str(test_output_path),
                "warnings": warnings_list,
                "message": "Slide generation test completed. Check the output file to verify the layout."
            }

        finally:
            # Uninitialize COM
            pythoncom.CoUninitialize()

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error in test slide generation: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to test slide generation: {str(e)}"
        ) from e


@router.post(
    "/test-table-layout",
    summary="Test table layout with simple list of names",
    description="Simple endpoint to test dynamic table layout with just a list of names",
)
async def test_table_layout(names: list[str]) -> dict:
    """Test table layout with a simple list of names.

    Args:
        names: List of candidate names

    Returns:
        Result of the table generation including success status and file path
    """
    try:
        from app.services.pptx_builder_service import PPTXBuilderService
        from app.utils.path_utils import resolve_project_output
        import pythoncom

        # Initialize COM for this thread
        pythoncom.CoInitialize()

        try:
            if not names:
                raise HTTPException(status_code=400, detail="No names provided")

            logger.info("Testing table layout with %d names", len(names))

            # Initialize builder service
            templates_dir = Path(settings.templates_dir) if hasattr(settings, 'templates_dir') else Path("templates")
            subdir = "BackgroundDefaultTemplate"
            templates_dir = templates_dir / subdir
            template_name = "template_default_withgroups2019.pptx"
            template_path = templates_dir / template_name

            if not template_path.exists():
                raise HTTPException(
                    status_code=404,
                    detail=f"Template not found: {template_name}"
                )

            builder = PPTXBuilderService(
                templates_dir=templates_dir,
                base_template_path=template_path,
            )

            # Create output directory
            output_dir, _ = resolve_project_output("test_table_layout", None)
            output_dir.mkdir(parents=True, exist_ok=True)

            # Generate test presentation
            from pptx import Presentation
            import win32com.client

            # Load template
            prs = Presentation(str(template_path))

            # Use first slide or add one
            if prs.slides:
                slide = prs.slides[0]
            else:
                slide = prs.slides.add_slide(prs.slide_layouts[5])

            # Save the presentation first
            test_output_path = output_dir / f"table_test_{int(time.time())}.pptx"
            prs.save(str(test_output_path))

            # Open with COM automation
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1

            presentation = powerpoint.Presentations.Open(str(test_output_path.absolute()))
            com_slide = presentation.Slides(1)

            # Collect placeholder shapes
            placeholder_shapes = builder._collect_placeholder_shapes(com_slide, "Name Candidate")

            logger.info("Found %d placeholder shapes", len(placeholder_shapes))

            # Test the simplified placeholder assignment
            cleaned_names = [n for n in names if n]
            builder._assign_names_to_placeholders(
                placeholder_shapes=placeholder_shapes or [],
                names=cleaned_names,
                context="test table",
            )
            success = True

            # Save and close properly
            try:
                presentation.Save()
            except Exception as e:
                logger.warning("Could not save: %s", e)

            try:
                presentation.Close()
            except Exception as e:
                logger.warning("Could not close: %s", e)

            try:
                powerpoint.Quit()
            except Exception as e:
                logger.warning("Could not quit: %s", e)

            logger.info("Table layout test completed: success=%s", success)

            return {
                "success": success,
                "names_count": len(names),
                "output_path": str(test_output_path),
                "message": "Table layout test completed. Check the output file."
            }

        finally:
            pythoncom.CoUninitialize()

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error in test table layout: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to test table layout: {str(e)}"
        ) from e


@router.post(
    "/download-results",
    response_model=DownloadResultsResponse,
    summary="Download NW results (generates Excel and Word in a ZIP)",
    description=(
        "Always generates two files for the specified presentation and packages them in a ZIP file:\n\n"
        "- **Excel**: Results workbook (retained, new names, explore/avoid, notes) and, when applicable, votes and participants sheets (internal logic).\n"
        "- **Word**: NW report based on a template, with tables and chart.\n\n"
        "**Output**: ZIP file that contains both reports (Excel and Word).\n\n"
        "Minimum input: only `presentation_id`. The endpoint resolves which sheets to include and returns a download token for the ZIP file."
    ),
    response_description=(
        "Generation status with path, filename, and a download token for the ZIP. "
        "Use `/presentations/files/{token}` to download it. The ZIP contains both files (Excel and Word)."
    ),
    responses={
        200: {
            "description": "Reports generated in a ZIP file (Excel and Word).",
            "content": {
                "application/json": {
                    "example": {
                        "success": True,
                        "message": "Generated 2 reports (Excel and Word) in ZIP archive",
                        "file_path": "C:/.../NW_Files/downloads/TestPresentation_Reports_20250122_143022.zip",
                        "file_name": "TestPresentation_Reports_20250122_143022.zip",
                        "download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                        "excel_download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                        "word_download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                        "excel_file_name": "TestPresentation_20250122_143022.xlsx",
                        "word_file_name": "TestPresentation_20250122_143022.doc",
                        "report_type": "excel",
                        "presentation_id": 12345,
                        "generated_at": "2025-01-22T14:30:22.123456",
                        "warnings": []
                    }
                }
            },
        },
        400: {
            "description": "Invalid request parameters.",
            "content": {
                "application/json": {
                    "example": {"detail": "Presentation with ID 12345 not found"}
                }
            },
        },
        404: {
            "description": "Required template file not found.",
            "content": {
                "application/json": {
                    "example": {"detail": "Word template not found at: C:/Templates/Nomenclature Workshop Report_Template_new_xxx.doc"}
                }
            },
        },
        500: {
            "description": "Unexpected error during report generation.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to generate report: Database connection error"
                    }
                }
            },
        },
    },
)
async def download_results(request: DownloadResultsRequest) -> DownloadResultsResponse:
    """Generate and download NW presentation results.

    This endpoint orchestrates the generation of various NW report types:

    **Excel Report (report_type: "excel"):**
    - Generates a complete workbook with multiple sheets
    - Includes retained names, newly created names, roots/concepts, and notes
    - Optionally includes voting data and participant information
    - Data is retrieved from stored procedures:
      - nw_dlRetainedNames_withRecraft
      - nw_CombineNewNames
      - nw_dlRootsOrConceptsToExplore
      - nw_dlRootsOrConceptsToAvoid
      - nw_dlOpenNotes
      - nw_Votesbygroups (if include_votes=true)
      - nw_VotedParticipants (if include_participants=true)

    **Analytics Report (report_type: "analytics"):**
    - Generates from NWAnalytics.xlsx template
    - Fills Project-Specific and Region-Specific sheets
    - Data from stored procedures:
      - NW_ProjectAnalytics
      - NW_RegionSpecificAnalytics

    **Word Report (report_type: "word"):**
    - Generates a Word document from a template
    - Replaces placeholders (cover, headers, footers, text boxes)
    - Tablas pobladas por tipo (Positive, Neutral, Reconsider, New Names, Explore)
    - Inserta grÃ¡fico de torta (By The Numbers)
    - Procedimientos utilizados:
      - nw_wdValuesToReplace (placeholders)
      - nw_wdGetResults_Phonetics / nw_CombineNewNames

    **BSR Report (report_type: "bsr"):** [Not Yet Implemented]
    - Would generate BSR-specific Excel and Word documents
    - Would use stored procedures:
      - bsr_GetExcelReport
      - Plus direct queries to bsr_ProjectConcepts and BSR_PageComments tables

    Args:
        request: Download results request containing:
            - presentation_id: The presentation to generate report for
            - report_type: Type of report (excel, word, analytics, bsr)
            - summary_type: Optional, for Word phonetics reports
            - mobile_link: Optional, for BSR downloads
            - include_votes: Whether to include votes sheet
            - include_participants: Whether to include participants sheet

    Returns:
        DownloadResultsResponse with:
            - success: Boolean indicating if generation succeeded
            - message: Descriptive message
            - file_path: Full path to generated file
            - file_name: Name of generated file
            - download_token: Secure token for downloading the file
            - report_type: Type of report generated
            - presentation_id: ID of presentation
            - generated_at: Timestamp of generation
            - warnings: List of any warnings during generation

    Raises:
        HTTPException 400: If presentation ID is invalid or not found
        HTTPException 404: If required template files are missing
        HTTPException 500: If report generation fails
        HTTPException 501: If report type is not yet implemented
    """
    try:
        logger.info(
            "Received download results request: presentation_id=%d",
            request.presentation_id,
        )

        # Generate report using orchestrator service
        response = report_orchestrator_service.generate_report(request)

        logger.info(
            "Reports (Excel+Word) generated for presentation %d",
            request.presentation_id,
        )

        return response

    except ValueError as e:
        logger.error("Invalid request: %s", str(e))
        raise HTTPException(status_code=400, detail=str(e)) from e

    except FileNotFoundError as e:
        logger.error("Template file not found: %s", str(e))
        raise HTTPException(status_code=404, detail=str(e)) from e

    except NotImplementedError as e:
        logger.error("Feature not implemented: %s", str(e))
        raise HTTPException(status_code=501, detail=str(e)) from e

    except HTTPException:
        raise

    except Exception as e:
        logger.error("Error generating report: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to generate report: {str(e)}"
        ) from e


@router.post(
    "/create-feedback-template",
    response_model=CreateFeedbackTemplateResponse,
    summary="Create NW Feedback Template document (InputDocumentRationales)",
    description=(
        "Generate a complete Feedback Template document from the InputDocumentRationales template.\n\n"
        "This endpoint replicates the 'Create Feedback Template' button functionality from the original app. "
        "It generates a fully populated Word document with:\n\n"
        "**Document Generation Process:**\n"
        "1. **Template Selection** - Automatically selects the correct InputDocumentRationales template:\n"
        "   - `InputDocumentRationales.doc` for Normal presentations\n"
        "   - `InputDocumentRationales_Phonetics.doc` for Phonetics presentations\n"
        "   - `InputDocumentRationales_Katakana.doc` for Katakana presentations\n\n"
        "2. **Placeholder Replacement** - Replaces all placeholders throughout the document:\n"
        "   - `<Client>`, `<ProjectName>`, `<DateNormalCase>`, etc.\n"
        "   - Searches in body, headers, footers, text boxes, and shapes\n"
        "   - Data from `nw_wdValuesToReplace` stored procedure\n\n"
        "3. **Table Population** - Fills all feedback tables with project data:\n"
        "   - **Positive Names** table (from `Positive` or `Positive_Phonetics`)\n"
        "   - **Neutral Names** table (from `Neutral` or `Negative_Phonetics`)\n"
        "   - **Reconsider Names** table (from `Reconsider` or `Reconsider_Phonetics`)\n"
        "   - **Newly Created Names** table (from `nw_CombineNewNames`)\n"
        "   - **Roots/Concepts to Explore** table (from `Explore`)\n"
        "   - Uses Open Sans 12pt font, rationales in 10pt\n"
        "   - Handles HTML content in rationale fields\n"
        "   - Processes name groups (## or $$ delimiters)\n\n"
        "4. **Pie Chart Generation** - Creates and inserts a pie chart:\n"
        "   - Uses data from `ByTheNumbers` summary type\n"
        "   - Shows Positive/Neutral/Reconsider distribution\n"
        "   - Corporate colors: Green (92,184,92), Brown (183,122,51), Purple (153,0,76)\n"
        "   - Generated in Excel and pasted into PieChart bookmark\n\n"
        "5. **Cleanup** - Removes all bookmarks from the final document\n\n"
        "**Data Sources:**\n"
        "- `nw_wdValuesToReplace` - Placeholder values (client, project, date, etc.)\n"
        "- `nw_wdGetResults_Phonetics` - Name results by category (Positive, Neutral, Reconsider, Explore)\n"
        "- `nw_CombineNewNames` - Newly created names with categories and rationales\n\n"
        "**Output Format:**\n"
        "- Word 97-2003 Document (.doc)\n"
        "- Saved to NW_Downloads directory\n"
        "- Filename: `Feedback_Template_{DisplayName}_{timestamp}.doc`\n\n"
        "The generated document is ready for client delivery and contains all project feedback data."
    ),
    response_description=(
        "Generation status with file path, filename, and secure download token. "
        "Use the download token with the /presentations/files/{token} endpoint to retrieve the file."
    ),
    responses={
        200: {
            "description": "Feedback template generated successfully.",
            "content": {
                "application/json": {
                    "example": {
                        "success": True,
                        "message": "Feedback template generated successfully",
                        "file_path": "C:/output/NW_Downloads/Feedback_Template_MyProject_20250117_103045.doc",
                        "file_name": "Feedback_Template_MyProject_20250117_103045.doc",
                        "download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                        "presentation_id": 12345,
                        "presentation_type": "Phonetics",
                        "generated_at": "2025-01-17T10:30:45.123456",
                        "warnings": []
                    }
                }
            },
        },
        400: {
            "description": "Invalid presentation ID or presentation not found.",
            "content": {
                "application/json": {
                    "example": {"detail": "Presentation with ID 12345 not found"}
                }
            },
        },
        404: {
            "description": "Required template file not found.",
            "content": {
                "application/json": {
                    "example": {"detail": "Feedback template not found: C:/Templates/InputDocumentRationales_Phonetics.doc"}
                }
            },
        },
        500: {
            "description": "Unexpected error during document generation.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to generate feedback template: Word COM automation error"
                    }
                }
            },
        },
    },
)
async def create_feedback_template(request: CreateFeedbackTemplateRequest) -> CreateFeedbackTemplateResponse:
    """Generate a Feedback Template document for an NW presentation.

    This endpoint creates a complete InputDocumentRationales document populated with all
    feedback data for a specific presentation. The document includes:

    - Client and project metadata in headers/footers
    - Positive, Neutral, and Reconsider name tables
    - Newly created names table
    - Roots/concepts to explore table
    - Pie chart showing name distribution

    The template used is automatically selected based on the presentation type
    (Normal, Phonetics, or Katakana).

    Args:
        request: CreateFeedbackTemplateRequest with presentation_id

    Returns:
        CreateFeedbackTemplateResponse with:
            - success: Boolean indicating if generation succeeded
            - message: Descriptive message
            - file_path: Full path to generated Word file
            - file_name: Name of generated file
            - download_token: Secure token for downloading the file
            - presentation_id: ID of presentation
            - presentation_type: Type of presentation (Normal, Phonetics, Katakana)
            - generated_at: Timestamp of generation
            - warnings: List of any warnings during generation

    Raises:
        HTTPException 400: If presentation ID is invalid or not found
        HTTPException 404: If required template files are missing
        HTTPException 500: If document generation fails
    """
    try:
        logger.info(
            "Received create feedback template request: presentation_id=%d",
            request.presentation_id,
        )

        # Get presentation info
        presentation_info = bi_guidelines_service.get_project_info(
            request.presentation_id
        )

        if not presentation_info:
            raise ValueError(
                f"Presentation with ID {request.presentation_id} not found"
            )

        display_name = presentation_info.get("DisplayName", "Unknown")
        presentation_type = presentation_info.get("PresentationType", "Normal")

        logger.info(
            "Generating feedback template for display_name='%s', type='%s'",
            display_name,
            presentation_type,
        )

        # Generate the feedback template document
        warnings = []
        try:
            file_path = feedback_template_generator.generate_feedback_template(
                presentation_id=request.presentation_id,
                display_name=display_name,
            )

            logger.info("Feedback template generated successfully: %s", file_path)

        except Exception as e:
            logger.error("Error generating feedback template: %s", str(e), exc_info=True)
            raise

        # Build download token
        download_token = None
        if file_path and file_path.exists():
            # Return only the token; client prepends /api/presentations/files/
            download_token = create_download_token(file_path)

        return CreateFeedbackTemplateResponse(
            success=True,
            message="Feedback template generated successfully",
            file_path=str(file_path),
            file_name=file_path.name,
            download_token=download_token,
            presentation_id=request.presentation_id,
            presentation_type=presentation_type,
            warnings=warnings,
        )

    except ValueError as e:
        logger.error("Invalid request: %s", str(e))
        raise HTTPException(status_code=400, detail=str(e)) from e

    except FileNotFoundError as e:
        logger.error("Template file not found: %s", str(e))
        raise HTTPException(status_code=404, detail=str(e)) from e

    except HTTPException:
        raise

    except Exception as e:
        logger.error("Error creating feedback template: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to generate feedback template: {str(e)}"
        ) from e


@router.post(
    "/backup",
    response_model=GenerateBackupResponse,
    summary="Generate backup PowerPoint from saved files",
    description=(
        "Generate a complete PowerPoint backup presentation from previously saved Excel and PPTX files.\n\n"
        "This endpoint allows users to generate the backup PowerPoint file at a later time, "
        "even if they didn't select the 'create_backup' option during the initial presentation creation.\n\n"
        "**How it works:**\n"
        "1. Retrieves presentation metadata from the database using the presentation_id\n"
        "2. Loads the saved original Excel file (saved during initial creation)\n"
        "3. Loads the saved original PowerPoint template file\n"
        "4. Processes the Excel data and generates slides\n"
        "5. Creates a complete PowerPoint backup combining:\n"
        "   - Original PPTX slides (before page_number)\n"
        "   - Dynamically generated slides from Excel data\n"
        "   - Original PPTX slides (after generated content)\n\n"
        "**Requirements:**\n"
        "- The presentation must exist in the database\n"
        "- Original Excel file must have been saved (automatic since this update)\n"
        "- Original PowerPoint template must exist\n\n"
        "**Returns:** Complete PowerPoint backup file with download token"
    ),
    response_description=(
        "Generation status with file path, filename, total slides, and secure download token. "
        "Use the download token with the /presentations/files/{token} endpoint to retrieve the file."
    ),
    responses={
        200: {
            "description": "Backup generation response (success or missing files)",
            "content": {
                "application/json": {
                    "examples": {
                        "success": {
                            "summary": "Backup generated successfully",
                            "value": {
                                "success": True,
                                "message": "Backup presentation generated successfully",
                                "presentation_id": 12345,
                                "printable_path": "C:/inetpub/wwwroot/nw2/nw_slides/TestProject/Presentations/backup_20250123_143022.pptx",
                                "file_name": "backup_20250123_143022.pptx",
                                "download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                                "total_slides": "24",
                                "warnings": "None",
                                "missing_files": []
                            }
                        },
                        "missing_files": {
                            "summary": "Required files missing",
                            "value": {
                                "success": False,
                                "message": "Cannot generate backup: 2 required file(s) missing. Please contact IT to upload the missing files.",
                                "presentation_id": 12345,
                                "printable_path": None,
                                "file_name": None,
                                "download_token": None,
                                "total_slides": None,
                                "warnings": "Missing 2 required file(s)",
                                "missing_files": [
                                    {
                                        "file_type": "Excel",
                                        "expected_path": "C:/inetpub/wwwroot/nw2/nw_slides/TestProject/original_data.xlsx",
                                        "instructions": "Please upload the original Excel file to: C:/inetpub/wwwroot/nw2/nw_slides/TestProject/original_data.xlsx"
                                    },
                                    {
                                        "file_type": "PowerPoint",
                                        "expected_path": "C:/inetpub/wwwroot/nw2/nw_slides/TestProject/template.pptx",
                                        "instructions": "Please upload the original PowerPoint template to: C:/inetpub/wwwroot/nw2/nw_slides/TestProject/template.pptx"
                                    }
                                ]
                            }
                        }
                    }
                }
            },
        },
        404: {
            "description": "Presentation not found in database",
            "content": {
                "application/json": {
                    "example": {"detail": "Presentation with ID 12345 not found"}
                }
            },
        },
        500: {
            "description": "Unexpected error during backup generation",
            "content": {
                "application/json": {
                    "example": {"detail": "Failed to generate backup: PowerPoint automation error"}
                }
            },
        },
    },
)
async def generate_backup(request: GenerateBackupRequest) -> GenerateBackupResponse:
    """Generate a backup PowerPoint presentation from saved files.

    This endpoint enables deferred backup generation. Users who didn't create
    a backup during initial presentation creation can generate it later using
    this endpoint.

    Args:
        request: GenerateBackupRequest with presentation_id

    Returns:
        GenerateBackupResponse with:
            - success: Boolean indicating if generation succeeded
            - message: Descriptive message
            - presentation_id: ID of the presentation
            - printable_path: Full path to generated .pptx file
            - file_name: Name of the generated file
            - download_token: Secure token for downloading the file
            - total_slides: Total number of slides in the presentation
            - warnings: Any warnings encountered during generation

    Raises:
        HTTPException 404: If presentation or required files not found
        HTTPException 500: If backup generation fails
    """
    try:
        logger.info(
            "Received generate backup request: presentation_id=%d",
            request.presentation_id,
        )

        # Generate backup using service
        result = presentation_service.generate_backup_presentation(
            presentation_id=request.presentation_id
        )

        # Check if there are missing files
        missing_files = result.get("missing_files", [])

        if missing_files:
            # Files are missing - return response with missing file details
            logger.warning(
                "Cannot generate backup for presentation %d: %d file(s) missing",
                request.presentation_id,
                len(missing_files)
            )

            return GenerateBackupResponse(
                success=False,
                message=f"Cannot generate backup: {len(missing_files)} required file(s) missing. Please contact IT to upload the missing files.",
                presentation_id=request.presentation_id,
                printable_path=None,
                file_name=None,
                download_token=None,
                total_slides=None,
                warnings=result.get("warnings", "None"),
                missing_files=missing_files,
            )

        # Extract file path from result
        printable_path = result.get("printable_path", "")

        if not printable_path:
            raise HTTPException(
                status_code=500,
                detail="Backup generation succeeded but no file path was returned"
            )

        # Convert path string to Path object
        file_path = Path(printable_path)

        if not file_path.exists():
            raise HTTPException(
                status_code=500,
                detail=f"Backup file was generated but not found at: {printable_path}"
            )

        # Create download token
        download_token = create_download_token(file_path)

        logger.info(
            "Backup generated successfully for presentation %d: %s",
            request.presentation_id,
            file_path.name
        )

        return GenerateBackupResponse(
            success=True,
            message="Backup presentation generated successfully",
            presentation_id=request.presentation_id,
            printable_path=printable_path,
            file_name=file_path.name,
            download_token=download_token,
            total_slides=result.get("total_slides", "0"),
            warnings=result.get("warnings", "None"),
            missing_files=[],
        )

    except HTTPException:
        raise

    except Exception as e:
        logger.error(
            "Error generating backup for presentation %d: %s",
            request.presentation_id,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Failed to generate backup: {str(e)}"
        ) from e


@router.get(
    "/download-presentation/{presentation_id}",
    summary="Download generated PowerPoint presentation file",
    description=(
        "Download the physical PowerPoint file (.pptx) generated for a specific presentation.\n\n"
        "This endpoint retrieves the PowerPoint file that was created when `create_backup=1` "
        "was specified during presentation creation.\n\n"
        "**Requirements:**\n"
        "- The presentation must exist in the database\n"
        "- A PowerPoint file must have been generated (create_backup=1)\n"
        "- The file must exist on the server\n\n"
        "**Returns:** Direct file download of the .pptx file"
    ),
    responses={
        200: {
            "description": "PowerPoint file download",
            "content": {"application/vnd.openxmlformats-officedocument.presentationml.presentation": {}},
        },
        404: {"description": "Presentation not found or file doesn't exist"},
        500: {"description": "Server error"},
    },
)
async def download_presentation_file(presentation_id: int) -> FileResponse:
    """Download the generated PowerPoint file for a presentation.

    Args:
        presentation_id: The ID of the presentation to download

    Returns:
        FileResponse with the PowerPoint file

    Raises:
        HTTPException: 404 if presentation or file not found, 500 on server error
    """
    try:
        logger.info("Downloading presentation file for ID: %d", presentation_id)

        # Get presentation info from database
        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                """
                SELECT
                    DisplayName,
                    MainPptFileName,
                    Project
                FROM [BI_GUIDELINES].[dbo].[nw_Master]
                WHERE PresentationId = ?
                """,
                (presentation_id,)
            )

            row = cursor.fetchone()

            if not row:
                logger.warning("Presentation ID %d not found", presentation_id)
                raise HTTPException(
                    status_code=404,
                    detail=f"Presentation with ID {presentation_id} not found"
                )

            display_name = row.DisplayName
            main_ppt_filename = row.MainPptFileName
            project = row.Project

            if not main_ppt_filename:
                logger.warning(
                    "Presentation ID %d has no PowerPoint file (create_backup was not enabled)",
                    presentation_id
                )
                raise HTTPException(
                    status_code=404,
                    detail=f"No PowerPoint file available for presentation ID {presentation_id}. "
                           f"The file was not generated (create_backup=0)."
                )

        # Construct file path
        # Files are typically stored in: nw_slides/{DisplayName}/Presentations/{filename}
        file_path = settings.nw_files_dir / display_name / "Presentations" / main_ppt_filename

        if not file_path.exists():
            logger.error(
                "PowerPoint file not found on disk: %s (presentation_id=%d)",
                file_path,
                presentation_id
            )
            raise HTTPException(
                status_code=404,
                detail=f"PowerPoint file not found on server. Expected path: {main_ppt_filename}"
            )

        # Get file info
        file_size = file_path.stat().st_size
        media_type = guess_media_type(file_path)

        logger.info(
            "Serving PowerPoint file: %s (size: %d bytes, presentation_id: %d)",
            file_path.name,
            file_size,
            presentation_id
        )

        # Create FileResponse with proper headers
        response = FileResponse(
            path=file_path,
            media_type=media_type,
            filename=file_path.name,
        )

        # Set headers for download
        utf8_name = quote(file_path.name)
        response.headers["Content-Disposition"] = (
            f"attachment; filename=\"{file_path.name}\"; filename*=UTF-8''{utf8_name}"
        )
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Content-Transfer-Encoding"] = "binary"

        return response

    except HTTPException:
        raise

    except Exception as e:
        logger.error(
            "Error downloading presentation file for ID %d: %s",
            presentation_id,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Failed to download presentation file: {str(e)}"
        ) from e

@router.post(
    "/generate-bsr-report",
    response_model=BSRGenerateReportResponse,
    summary="Generate BSR Excel and Word reports",
    description=(
        "Generate downloadable Excel and Word reports for a BSR presentation.\n\n"
        "**Excel Report**: Contains name candidates with categories\n"
        "**Word Report**: Contains slide notes, attributes, and key concepts\n\n"
        "Files are saved to: C:\\inetpub\\wwwroot\\CreativePythonAPI\\NW_Files\\downloads\\{DisplayName}.{ext}"
    ),
)
async def generate_bsr_report(
    request: BSRGenerateReportRequest
) -> BSRGenerateReportResponse:
    """Generate BSR reports (Excel + Word)."""
    try:
        start_time = time.time()

        # Generate reports
        result = bsr_report_orchestrator.generate_bsr_reports(
            presentation_id=request.presentation_id,
            display_name=request.display_name
        )

        processing_time = time.time() - start_time
        # Build download tokens and an optional ZIP for both files
        from app.utils.download_utils import create_download_token
        import zipfile
        from datetime import datetime
        
        excel_token = None
        word_token = None
        zip_path = None
        zip_token = None
        
        try:
            from pathlib import Path as _Path
            excel_path = _Path(result.excel_path)
            word_path = _Path(result.word_path) if result.word_path else None
            if excel_path.exists():
                excel_token = create_download_token(excel_path)
            if word_path and word_path.exists():
                word_token = create_download_token(word_path)
            # Create ZIP only if both files exist
            if excel_path.exists() and word_path and word_path.exists():
                downloads_dir = excel_path.parent
                base_name = (request.display_name or excel_path.stem or f"BSR_{request.presentation_id}")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                zip_path = downloads_dir / f"{base_name}_BSR_Reports_{timestamp}.zip"
                with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    zf.write(excel_path, arcname=excel_path.name)
                    zf.write(word_path, arcname=word_path.name)
                zip_token = create_download_token(zip_path)
        except Exception as zip_exc:
            logger.warning("Could not create tokens/ZIP for BSR reports: %s", str(zip_exc))

        return BSRGenerateReportResponse(
            message="BSR reports generated successfully",
            excel_file=str(result.excel_path),
            word_file=str(result.word_path),
            processing_time_seconds=processing_time,
            warnings=result.warnings,
            excel_download_token=excel_token,
            word_download_token=word_token,
            zip_file=(str(zip_path) if zip_path else None),
            zip_download_token=zip_token,
        )
    except Exception as e:
        logger.error("Error generating BSR reports: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to generate BSR reports: {str(e)}"
        ) from e

