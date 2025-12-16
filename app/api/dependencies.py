"""FastAPI dependencies for request validation and parsing.

This module contains reusable dependencies for file validation, metadata parsing,
and other common operations used across API routes.
"""

import json

from fastapi import File, Form, HTTPException, UploadFile
from pydantic import ValidationError

from app.config.settings import settings
from app.models.presentation_models import (
    CreatePresentationMetadata,
    PresentationBuildMetadata,
    SimpleDWMetadata,
    BSRCreatePresentationMetadata,
)
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


# File type validation dependencies
async def validate_excel_file(
    file: UploadFile = File(..., description="Excel file (.xlsx or .xls)")
) -> UploadFile:
    """Validate that the uploaded file is an Excel file.
    
    Args:
        file: The uploaded file to validate.
        
    Returns:
        The validated UploadFile.
        
    Raises:
        HTTPException: If the file is not a valid Excel file.
    """
    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename provided")
    
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400,
            detail="Excel file must be .xlsx or .xls"
        )
    
    return file


async def validate_pptx_file(
    file: UploadFile = File(..., description="PowerPoint file (.pptx)")
) -> UploadFile:
    """Validate that the uploaded file is a PowerPoint file.
    
    Args:
        file: The uploaded file to validate.
        
    Returns:
        The validated UploadFile.
        
    Raises:
        HTTPException: If the file is not a valid PowerPoint file.
    """
    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename provided")
    
    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(
            status_code=400,
            detail="PPTX file must be .pptx"
        )
    
    return file


# Metadata parsing dependencies
async def parse_presentation_metadata(
    metadata: str = Form(
        ..., description="JSON payload containing presentation metadata"
    )
) -> CreatePresentationMetadata:
    """Parse and validate presentation creation metadata from form data.
    
    Args:
        metadata: JSON string containing presentation metadata.
        
    Returns:
        Validated CreatePresentationMetadata object.
        
    Raises:
        HTTPException: If metadata is invalid JSON or fails validation.
    """
    try:
        payload = json.loads(metadata)
    except json.JSONDecodeError as exc:
        logger.warning("Invalid JSON in metadata: %s", exc)
        raise HTTPException(
            status_code=400, 
            detail="Metadata payload must be valid JSON"
        ) from exc

    try:
        return CreatePresentationMetadata(**payload)
    except ValidationError as exc:
        logger.warning("Metadata validation failed: %s", exc.errors())
        raise HTTPException(status_code=422, detail=exc.errors()) from exc


async def parse_build_metadata(
    metadata: str = Form(
        ..., description="JSON payload containing ppt assembly metadata"
    )
) -> PresentationBuildMetadata:
    """Parse and validate presentation build metadata from form data.
    
    Args:
        metadata: JSON string containing build metadata.
        
    Returns:
        Validated PresentationBuildMetadata object.
        
    Raises:
        HTTPException: If metadata is invalid JSON or fails validation.
    """
    try:
        payload = json.loads(metadata)
    except json.JSONDecodeError as exc:
        logger.warning("Invalid JSON in build metadata: %s", exc)
        raise HTTPException(
            status_code=400, 
            detail="Metadata payload must be valid JSON"
        ) from exc

    try:
        return PresentationBuildMetadata(**payload)
    except ValidationError as exc:
        logger.warning("Build metadata validation failed: %s", exc.errors())
        raise HTTPException(status_code=422, detail=exc.errors()) from exc


# Form field validation dependencies
async def validate_display_name(
    display_name: str = Form(..., description="Display name for the project folder")
) -> str:
    """Validate and clean display name.
    
    Args:
        display_name: The display name to validate.
        
    Returns:
        The validated display name.
        
    Raises:
        HTTPException: If display name is empty or invalid.
    """
    if not display_name or not display_name.strip():
        raise HTTPException(status_code=400, detail="Display name is required")
    
    return display_name.strip()


async def validate_project_type(
    project_type: str = Form(..., description="Project type: 'bipresents', 'nw', 'dw', or 'nsr'")
) -> str:
    """Validate project type.

    Args:
        project_type: The project type to validate.

    Returns:
        The validated project type.

    Raises:
        HTTPException: If project type is not valid.
    """
    valid_project_types = [
        settings.PROJECT_TYPE_BIPRESENTS,
        settings.PROJECT_TYPE_NW,
        settings.PROJECT_TYPE_DW,
        "nsr",  # NSR support
    ]

    if not project_type or project_type not in valid_project_types:
        raise HTTPException(
            status_code=400,
            detail=f"Project type must be one of: {', '.join(valid_project_types)}",
        )

    return project_type


# File size validation dependencies
async def validate_file_size(
    file: UploadFile,
    max_size_mb: int = None,
) -> bytes:
    """Validate file size and return file content.
    
    This dependency reads the file content and validates it's within size limits.
    
    Args:
        file: The uploaded file to validate.
        max_size_mb: Maximum file size in MB (defaults to settings.max_file_size_mb).
        
    Returns:
        The file content as bytes.
        
    Raises:
        HTTPException: If file is empty or exceeds size limit.
    """
    if max_size_mb is None:
        max_size_mb = settings.max_file_size_mb
    
    max_size_bytes = max_size_mb * 1024 * 1024
    
    # Read file content
    file_content = await file.read()
    
    # Validate not empty
    if len(file_content) == 0:
        raise HTTPException(
            status_code=400,
            detail="Empty file provided"
        )
    
    # Validate size limit
    if len(file_content) > max_size_bytes:
        raise HTTPException(
            status_code=413,
            detail=f"File too large. Maximum size allowed is {max_size_mb}MB"
        )
    
    logger.info(
        "File validated: %s (%d bytes, limit: %dMB)",
        file.filename,
        len(file_content),
        max_size_mb
    )
    
    return file_content


async def validate_excel_file_with_size(
    file: UploadFile = File(..., description="Excel file (.xlsx or .xls)"),
    max_size_mb: int = 10
) -> tuple[UploadFile, bytes]:
    """Validate Excel file type and size, returning both file object and content.
    
    This is a composite dependency that validates both file type and size.
    
    Args:
        file: The uploaded file to validate.
        max_size_mb: Maximum file size in MB (default: 10MB).
        
    Returns:
        Tuple of (UploadFile, file_content_bytes).
        
    Raises:
        HTTPException: If file is invalid, empty, or too large.
    """
    # Validate file type first
    validated_file = await validate_excel_file(file)
    
    # Validate size and get content
    file_content = await validate_file_size(validated_file, max_size_mb)
    
    return validated_file, file_content


async def parse_simple_dw_metadata(
    metadata: str = Form(
        ..., description="JSON payload containing simplified DW metadata"
    )
) -> SimpleDWMetadata:
    """Parse and validate simplified DW metadata from form data.

    Expects a 'metadata' form field containing JSON with keys:
    projectName, displayName, widePresentation, userName, and optional slideType.
    """
    try:
        payload = json.loads(metadata)
    except json.JSONDecodeError as exc:
        logger.warning("Invalid JSON in DW metadata: %s", exc)
        raise HTTPException(
            status_code=400,
            detail="Metadata payload must be valid JSON",
        ) from exc

    try:
        return SimpleDWMetadata(**payload)
    except Exception as exc:
        logger.warning("DW metadata validation failed: %s", str(exc))
        raise HTTPException(status_code=422, detail=str(exc)) from exc


async def validate_pptx_file_with_size(
    file: UploadFile = File(..., description="PowerPoint file (.pptx)"),
    max_size_mb: int = 50
) -> tuple[UploadFile, bytes]:
    """Validate PPTX file type and size, returning both file object and content.

    This is a composite dependency that validates both file type and size.

    Args:
        file: The uploaded file to validate.
        max_size_mb: Maximum file size in MB (default: 50MB for PPTX templates).

    Returns:
        Tuple of (UploadFile, file_content_bytes).

    Raises:
        HTTPException: If file is invalid, empty, or too large.
    """
    # Validate file type first
    validated_file = await validate_pptx_file(file)

    # Validate size and get content
    file_content = await validate_file_size(validated_file, max_size_mb)

    return validated_file, file_content


async def parse_bsr_metadata(
    metadata: str = Form(
        ..., description="JSON payload containing BSR presentation metadata"
    )
) -> BSRCreatePresentationMetadata:
    """Parse and validate BSR presentation creation metadata from form data.

    Expects a 'metadata' form field containing JSON with keys:
    project_name, display_name, slide_number, presentation_type, user_name, is_wide_ppt
    """
    try:
        payload = json.loads(metadata)
    except json.JSONDecodeError as exc:
        logger.warning("Invalid JSON in BSR metadata: %s", exc)
        raise HTTPException(
            status_code=400,
            detail="Metadata payload must be valid JSON",
        ) from exc

    try:
        return BSRCreatePresentationMetadata(**payload)
    except ValidationError as exc:
        logger.warning("BSR metadata validation failed: %s", exc.errors())
        raise HTTPException(status_code=422, detail=exc.errors()) from exc
