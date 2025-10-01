"""
API routes for PowerPoint to images conversion.
"""

from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from app.config.settings import settings
from app.models.response_models import PPTXConversionResponse
from app.services.pptx_service import pptx_service
from app.utils.files_utils import FileUtils
from app.utils.logging_utils import get_logger
from app.utils.path_utils import sanitize_folder_name

logger = get_logger(__name__)

router = APIRouter(prefix="/pptx", tags=["PPTX Conversion"])


@router.post(
    "/convert/",
    summary="Convert PPTX to images and extract titles",
    response_model=PPTXConversionResponse,
)
async def convert_pptx_to_images(
    pptx_file: UploadFile = File(..., description="PPTX file to convert"),
    display_name: str = Form(..., description="Display name for the project folder"),
    project_type: str = Form(..., description="Project type: 'bipresents' or 'nw'"),
):
    """
    Receives a PPTX file, display_name, and project_type, converts each slide to an image,
    and extracts slide titles.

    - **pptx_file**: PPTX file to process
    - **display_name**: Display name to use as folder name for storing images
    - **project_type**: Project type ('bipresents' or 'nw') to determine base directory

    Returns the paths to the generated images and thumbnails, and the titles of each slide.
    """
    try:
        if not pptx_file.filename:
            raise HTTPException(status_code=400, detail="No filename provided")

        # Verify that the file is a PPTX
        if not pptx_file.filename.lower().endswith(".pptx"):
            raise HTTPException(status_code=400, detail="File must be a PPTX")

        # Validate display_name
        if not display_name or not display_name.strip():
            raise HTTPException(status_code=400, detail="Display name is required")

        # Validate project_type
        valid_project_types = [
            settings.PROJECT_TYPE_BIPRESENTS,
            settings.PROJECT_TYPE_NW,
        ]
        if not project_type or project_type not in valid_project_types:
            raise HTTPException(
                status_code=400,
                detail=f"Project type must be one of: {', '.join(valid_project_types)}",
            )

        # Clean display_name to be safe for folder names
        clean_display_name = sanitize_folder_name(display_name.strip())

        if not clean_display_name:
            raise HTTPException(status_code=400, detail="Invalid display name")

        # Read file content
        file_content = await pptx_file.read()

        logger.info(
            "Processing PPTX file: %s (%s) with display name: %s, project type: %s",
            pptx_file.filename,
            FileUtils.format_file_size(len(file_content)),
            clean_display_name,
            project_type,
        )

        # Convert PPTX to images using display_name as folder name and project_type
        result = pptx_service.convert_pptx_to_images(
            file_content, pptx_file.filename, clean_display_name, project_type
        )

        return PPTXConversionResponse(
            message=result["message"],
            conversion_id=result["conversion_id"],
            project_type=result["project_type"],
            total_images=result["total_images"],
            images=result["images"],
            thumbnails=result.get("thumbnails", []),
            titles=result["titles"],
            pptx_file=result.get("pptx_file", ""),
        )

    except ValueError as ve:
        logger.error("Validation error: %s", str(ve))
        raise HTTPException(status_code=400, detail=str(ve)) from ve

    except Exception as e:
        logger.error("Error in PPTX conversion: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Conversion error: {str(e)}"
        ) from e
