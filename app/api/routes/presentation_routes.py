"""Routes for presentation creation orchestration."""

import json
import time

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from pydantic import ValidationError

from app.models.presentation_models import (
    CreatePresentationMetadata,
    CreatePresentationRequest,
    CreatePresentationResponse,
)
from app.services.presentation_service import presentation_service
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/presentation", tags=["Presentation Creation"])
logger = get_logger(__name__)


async def _parse_metadata(
    metadata: str = Form(
        ..., description="JSON payload containing presentation metadata"
    )
) -> CreatePresentationMetadata:
    try:
        payload = json.loads(metadata)
    except json.JSONDecodeError as exc:
        raise HTTPException(
            status_code=400, detail="Metadata payload must be valid JSON"
        ) from exc

    try:
        return CreatePresentationMetadata(**payload)
    except ValidationError as exc:  # pragma: no cover - FastAPI handles response
        raise HTTPException(status_code=422, detail=exc.errors()) from exc


@router.post("/create", response_model=CreatePresentationResponse)
async def create_presentation(
    metadata: CreatePresentationMetadata = Depends(_parse_metadata),
    excel_file: UploadFile = File(
        ..., description="Excel file with candidate data (.xlsx or .xls)"
    ),
    pptx_file: UploadFile = File(..., description="PowerPoint template file (.pptx)"),
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
        # Read file contents
        excel_content = await excel_file.read()
        pptx_content = await pptx_file.read()

        # Validate file types
        if not excel_file.filename.lower().endswith((".xlsx", ".xls")):
            raise HTTPException(
                status_code=400, detail="Excel file must be .xlsx or .xls"
            )

        if not pptx_file.filename.lower().endswith(".pptx"):
            raise HTTPException(status_code=400, detail="PPTX file must be .pptx")

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
