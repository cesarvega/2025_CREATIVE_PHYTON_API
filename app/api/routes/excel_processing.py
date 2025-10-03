"""
Routes for Excel file processing and slide generation.
"""

import uuid
from typing import Optional

from fastapi import APIRouter, Depends, Form, HTTPException, UploadFile

from app.api.dependencies import validate_excel_file_with_size
from app.models.excel_models import ExcelProcessingResponse
from app.services.excel_service import excel_processing_service
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/excel", tags=["Excel Processing"])
logger = get_logger(__name__)


@router.post("/process-excel", response_model=ExcelProcessingResponse)
async def process_excel_file(
    file_data: tuple[UploadFile, bytes] = Depends(validate_excel_file_with_size),
    is_phonetics: Optional[bool] = Form(default=False, description="Use phonetics processing"),
    has_groups: Optional[bool] = Form(default=False, description="Process with groups"),
) -> ExcelProcessingResponse:
    """
    Process an Excel file and return processed data arrays.

    This endpoint accepts an Excel file (.xlsx or .xls) and processes the data
    according to the business logic, returning the processed arrays
    that can be used for slide generation in subsequent processes.

    Args:
        file_data: Tuple of (file, file_content) validated by dependency (max 10MB)
        is_phonetics: Whether to use phonetics processing
        has_groups: Whether to process with groups

    Returns:
        ExcelProcessingResponse: Contains processed data arrays and metadata

    Raises:
        HTTPException: If file processing fails or file is invalid
    """
    file, file_content = file_data

    try:
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

        response = ExcelProcessingResponse.from_processed(
            processed=processed_data,
            message=f"Successfully processed {file.filename}",
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
            "Unexpected error processing Excel file %s: %s", file.filename, str(e), exc_info=True
        )
        raise HTTPException(
            status_code=500, detail=f"Internal error processing Excel file: {str(e)}"
        ) from e
