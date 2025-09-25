"""
Routes for Excel file processing and slide generation.
"""

import uuid
from typing import Optional

from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from app.models.excel_models import ExcelProcessingResponse
from app.services.excel_service import excel_processing_service
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/excel", tags=["Excel Processing"])
logger = get_logger(__name__)


@router.post("/process-excel", response_model=ExcelProcessingResponse)
async def process_excel_file(
    file: UploadFile = File(..., description="Excel file to process (.xlsx or .xls)"),
    is_phonetics: Optional[bool] = Form(default=False, description="Use phonetics processing"),
    has_groups: Optional[bool] = Form(default=False, description="Process with groups"),
) -> ExcelProcessingResponse:
    """
    Process an Excel file and return processed data arrays.

    This endpoint accepts an Excel file (.xlsx or .xls) and processes the data
    according to the business logic, returning the processed arrays
    that can be used for slide generation in subsequent processes.

    Args:
        file: Excel file upload
        is_phonetics: Whether to use phonetics processing
        has_groups: Whether to process with groups

    Returns:
        ExcelProcessingResponse: Contains processed data arrays and metadata

    Raises:
        HTTPException: If file processing fails or file is invalid
    """
    # Validate file type
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file provided")

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Please upload an Excel file (.xlsx or .xls)",
        )

    # Validate file size (limit to 10MB)
    max_file_size = 10 * 1024 * 1024  # 10MB

    try:
        # Read file content
        file_content = await file.read()

        if len(file_content) == 0:
            raise HTTPException(status_code=400, detail="Empty file provided")

        if len(file_content) > max_file_size:
            raise HTTPException(
                status_code=413,
                detail=(
                    f"File too large. Maximum size allowed is "
                    f"{max_file_size / (1024 * 1024):.1f}MB"
                ),
            )

        logger.info(
            "Processing Excel file: %s (%d bytes)", file.filename, len(file_content)
        )

        # Process Excel file
        processed_data = excel_processing_service.process_excel_file(
            file_content=file_content,
            is_phonetics=is_phonetics or False,
            has_groups=has_groups or False,
        )

        # Generate processing ID for tracking
        processing_id = str(uuid.uuid4())

        response = ExcelProcessingResponse(
            message=f"Successfully processed {file.filename}",
            data=processed_data,
            processing_id=processing_id,
        )

        logger.info(
            "Excel processing completed. Processed %d rows with ID: %s",
            processed_data.total_rows_processed,
            processing_id,
        )

        return response

    except HTTPException:
        # Re-raise HTTP exceptions
        raise
    except Exception as e:
        logger.error(
            "Unexpected error processing Excel file %s: %s", file.filename, str(e)
        )
        raise HTTPException(
            status_code=500, detail=f"Internal error processing Excel file: {str(e)}"
        ) from e
