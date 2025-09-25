"""
Routes for presentation creation orchestration.
"""

import json
import time
from typing import Optional

from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from app.models.excel_models import (
    CreatePresentationRequest,
    CreatePresentationResponse,
)
from app.services.presentation_service import presentation_service
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/presentation", tags=["Presentation Creation"])
logger = get_logger(__name__)


@router.post("/create", response_model=CreatePresentationResponse)
async def create_presentation(
    excel_file: UploadFile = File(
        ..., description="Excel file with candidate data (.xlsx or .xls)"
    ),
    pptx_file: UploadFile = File(..., description="PowerPoint template file (.pptx)"),
    project: str = Form(..., description="Project identifier"),
    display_name: str = Form(..., description="Display name for the presentation"),
    presentation_type: str = Form("Nonproprietary", description="Type of presentation"),
    user_name: str = Form(..., description="User creating the presentation"),
    bsr_display_name: Optional[str] = Form(None, description="BSR display name"),
    mobile_link_bsr: Optional[str] = Form(None, description="Mobile link for BSR"),
    participant_vote: int = Form(1, description="Participant vote setting"),
    is_wide_ppt: int = Form(0, description="Wide PPT format flag"),
    is_aws_email: int = Form(0, description="AWS email flag"),
    background_type: str = Form("Image", description="Background type"),
    background_name: str = Form("", description="Background name"),
    page_number: int = Form(1, description="Starting page number"),
    project_type: str = Form(
        "bipresents", description="Project type: 'bipresents' or 'nw'"
    ),
    is_phonetics: bool = Form(False, description="Use phonetics processing"),
    has_groups: bool = Form(False, description="Process with groups"),
    template_rotation: Optional[str] = Form(
        None, description="JSON string of template rotation list"
    ),
) -> CreatePresentationResponse:
    """
    Create a complete presentation by orchestrating Excel processing, PPTX conversion,
    slide generation, and database insertion.

    This endpoint accepts Excel and PPTX files and processes them through the complete
    pipeline: Excel processing → PPTX conversion → slide generation → template application
    → database insertion.

    Args:
        excel_file: Excel file with candidate data (.xlsx or .xls)
        pptx_file: PowerPoint template file (.pptx)
        project: Project identifier
        display_name: Display name for the presentation
        presentation_type: Type of presentation
        user_name: User creating the presentation
        bsr_display_name: BSR display name (optional)
        mobile_link_bsr: Mobile link for BSR (optional)
        participant_vote: Participant vote setting
        is_wide_ppt: Wide PPT format flag
        is_aws_email: AWS email flag
        background_type: Background type
        background_name: Background name
        page_number: Starting page number
        is_phonetics: Use phonetics processing
        has_groups: Process with groups
        template_rotation: JSON string of template rotation list (optional)

    Returns:
        CreatePresentationResponse: Complete presentation creation result

    Raises:
        HTTPException: If presentation creation fails
    """
    try:
        # Parse template rotation if provided
        template_rotation_list = None
        if template_rotation:
            try:
                template_rotation_list = json.loads(template_rotation)
            except json.JSONDecodeError as exc:
                raise HTTPException(
                    status_code=400, detail="Invalid template_rotation JSON format"
                ) from exc

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

        # Create request object
        request = CreatePresentationRequest(
            excel_file=excel_content,
            pptx_file=pptx_content,
            excel_filename=excel_file.filename,
            pptx_filename=pptx_file.filename,
            is_phonetics=is_phonetics,
            has_groups=has_groups,
            project=project,
            display_name=display_name,
            presentation_type=presentation_type,
            user_name=user_name,
            bsr_display_name=bsr_display_name,
            mobile_link_bsr=mobile_link_bsr,
            participant_vote=participant_vote,
            is_wide_ppt=is_wide_ppt,
            is_aws_email=is_aws_email,
            background_type=background_type,
            background_name=background_name,
            page_number=page_number,
            project_type=project_type,
            template_rotation=template_rotation_list,
        )

        logger.info(
            "Starting presentation creation for project: %s, display_name: %s",
            request.project,
            request.display_name,
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
