"""Routes for presentation creation and assembly orchestration."""

import base64
import time
from pathlib import Path

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.api.dependencies import (
    parse_build_metadata,
    parse_presentation_metadata,
)
from app.config.settings import settings
from app.models.presentation_models import (
    CreatePresentationMetadata,
    CreatePresentationRequest,
    CreatePresentationResponse,
    PresentationBuildMetadata,
    PresentationBuildResponse,
)
from app.services.presentation_service import presentation_service
from app.utils.download_utils import (
    build_api_download_url,
    decode_download_token,
    guess_media_type,
)
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/presentations", tags=["Presentation Creation"])
logger = get_logger(__name__)


@router.post(
    "/create",
    response_model=CreatePresentationResponse,
    summary="Create a complete presentation from Excel + PPTX inputs",
    description=(
        "Upload the Excel candidate workbook, a base PPTX template, and a metadata JSON payload "
        "to generate a fully populated presentation. The endpoint processes Excel data, applies template "
        "rotation, persists database records, and returns diagnostic information about the generated assets."
    ),
    response_description="Creation status, generated slide counts, and source artifact metadata.",
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

    The service then executes the full pipeline:
    Excel processing → PPTX conversion → slide generation → template application →
    database persistence.
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
    return FileResponse(
        path=file_path,
        filename=file_path.name,
        media_type=guess_media_type(file_path),
    )
