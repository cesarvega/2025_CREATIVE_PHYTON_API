"""
Utilities for file upload validation and processing for Background Templates.

Handles:
- File type validation (JPG/PNG/PPTX)
- File size validation (max 10MB)
- Image storage (saves original to BackGrounds folder)
- PPTX conversion (extracts first slide as JPG using PowerPoint COM)

Supports JPG, PNG, and PPTX files. PPTX files are converted to JPG by extracting
the first slide at 1920x1080 resolution.
"""

import io
from pathlib import Path
from typing import Dict

from PIL import Image

from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

# Constants
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB
ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".pptx"}
ALLOWED_MIME_TYPES = {
    "image/jpeg": [".jpg", ".jpeg"],
    "image/png": [".png"],
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": [".pptx"],
}


class FileValidationError(Exception):
    """Custom exception for file validation errors."""
    pass


class FileProcessingError(Exception):
    """Custom exception for file processing errors."""
    pass


def validate_file_extension(filename: str) -> str:
    """Validate file extension.
    
    Args:
        filename: Name of the file to validate
        
    Returns:
        File extension (lowercase, with dot)
        
    Raises:
        FileValidationError: If extension is not allowed
    """
    ext = Path(filename).suffix.lower()
    
    if ext not in ALLOWED_EXTENSIONS:
        raise FileValidationError(
            f"Invalid file type. Allowed types: JPG, PNG, PPTX"
        )
    
    return ext


def validate_file_size(file_size: int) -> None:
    """Validate file size.
    
    Args:
        file_size: Size of file in bytes
        
    Raises:
        FileValidationError: If file size exceeds limit
    """
    if file_size > MAX_FILE_SIZE:
        size_mb = file_size / (1024 * 1024)
        max_mb = MAX_FILE_SIZE / (1024 * 1024)
        raise FileValidationError(
            f"File size ({size_mb:.2f}MB) exceeds maximum allowed size ({max_mb}MB)"
        )


def validate_mime_type(file_content: bytes, expected_extension: str) -> str:
    """Validate MIME type of file content.
    
    Uses file signature (magic bytes) to verify actual file type.
    
    Args:
        file_content: Binary content of the file
        expected_extension: Expected file extension (from filename)
        
    Returns:
        Detected MIME type
        
    Raises:
        FileValidationError: If MIME type doesn't match extension or is invalid
    """
    # Basic magic bytes detection
    
    # JPEG magic bytes
    if file_content.startswith(b'\xff\xd8\xff'):
        mime_type = "image/jpeg"
    # PNG magic bytes
    elif file_content.startswith(b'\x89PNG\r\n\x1a\n'):
        mime_type = "image/png"
    # PPTX magic bytes (ZIP format with specific structure)
    elif file_content.startswith(b'PK\x03\x04'):
        mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    else:
        raise FileValidationError("Unable to determine file type from content")
    
    # Verify MIME type matches expected extension
    if mime_type not in ALLOWED_MIME_TYPES:
        raise FileValidationError(f"Unsupported MIME type: {mime_type}")
    
    allowed_exts = ALLOWED_MIME_TYPES[mime_type]
    if expected_extension not in allowed_exts:
        raise FileValidationError(
            f"File extension '{expected_extension}' does not match detected type '{mime_type}'"
        )
    
    logger.debug("Validated MIME type: %s for extension: %s", mime_type, expected_extension)
    return mime_type


def get_image_dimensions(image: Image.Image) -> str:
    """Get image dimensions as string.
    
    Args:
        image: PIL Image object
        
    Returns:
        Dimensions as string (e.g., "1920x1080")
    """
    return f"{image.width}x{image.height}"


def process_image_file(
    filename: str,
    file_content: bytes,
    output_dir: Path
) -> str:
    """Process and save an image file (JPG/PNG).
    
    Simply validates and saves the image to the specified directory.
    NO resizing, NO thumbnails - just saves the original with a sanitized filename.
    
    Args:
        filename: Original filename
        file_content: Binary content of the image file
        output_dir: Directory where to save the image
        
    Returns:
        Saved filename (sanitized)
        
    Raises:
        FileProcessingError: If image processing fails
    """
    try:
        # Validate extension
        ext = validate_file_extension(filename)
        
        # Validate size
        validate_file_size(len(file_content))
        
        # Validate MIME type
        validate_mime_type(file_content, ext)
        
        # Load image to validate it's a valid image file
        img = Image.open(io.BytesIO(file_content))
        dimensions = get_image_dimensions(img)
        
        # Generate safe filename (remove special characters, spaces)
        import re
        safe_name = re.sub(r'[^\w\-.]', '_', filename)
        safe_name = re.sub(r'_+', '_', safe_name)  # Remove consecutive underscores
        
        # Ensure output directory exists
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Full output path
        output_path = output_dir / safe_name
        
        # Save original image directly
        with open(output_path, 'wb') as f:
            f.write(file_content)
        
        logger.info(
            "Saved image: %s (dimensions: %s, size: %d bytes)",
            safe_name,
            dimensions,
            len(file_content)
        )
        
        return safe_name
        
    except FileValidationError:
        raise
    except Exception as e:
        logger.error("Error processing image file: %s", str(e), exc_info=True)
        raise FileProcessingError(f"Failed to process image: {str(e)}") from e


def process_pptx_file(
    filename: str,
    file_content: bytes,
    output_dir: Path
) -> str:
    """Process and save a PPTX file.
    
    Extracts the first slide as a JPG image using PowerPoint COM automation.
    
    Args:
        filename: Original filename
        file_content: Binary content of the PPTX file
        output_dir: Directory where to save the extracted image
        
    Returns:
        Saved filename (sanitized, will be .jpg)
        
    Raises:
        FileProcessingError: If PPTX processing fails
    """
    try:
        import win32com.client
        import pythoncom
        import re
        
        # Validate extension
        ext = validate_file_extension(filename)
        
        # Validate size
        validate_file_size(len(file_content))
        
        # Validate MIME type
        validate_mime_type(file_content, ext)
        
        # Ensure output directory exists
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Save PPTX to temporary file
        temp_pptx = output_dir / "temp_presentation.pptx"
        with open(temp_pptx, 'wb') as f:
            f.write(file_content)
        
        # Generate safe output filename (change extension to .jpg)
        safe_name = re.sub(r'[^\w\-.]', '_', filename)
        safe_name = re.sub(r'_+', '_', safe_name)
        safe_name = safe_name.rsplit('.', 1)[0] + '.jpg'  # Change extension to .jpg
        
        output_path = output_dir / safe_name
        
        # Convert first slide to image using COM
        pythoncom.CoInitialize()
        try:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = False
            
            presentation = powerpoint.Presentations.Open(
                str(temp_pptx.absolute()), 
                WithWindow=False
            )
            
            if presentation.Slides.Count == 0:
                raise FileProcessingError("PPTX file has no slides")
            
            # Export first slide as JPG (1920x1080)
            presentation.Slides[1].Export(
                str(output_path.absolute()),
                "JPG",
                1920,  # Width
                1080   # Height
            )
            
            presentation.Close()
            powerpoint.Quit()
            
            logger.info(
                "Converted PPTX to image: %s (size: %d bytes)",
                safe_name,
                len(file_content)
            )
            
        finally:
            pythoncom.CoUninitialize()
            # Clean up temporary PPTX file
            if temp_pptx.exists():
                temp_pptx.unlink()
        
        return safe_name
        
    except FileValidationError:
        raise
    except Exception as e:
        logger.error("Error processing PPTX file: %s", str(e), exc_info=True)
        raise FileProcessingError(f"Failed to process PPTX: {str(e)}") from e


def process_file(
    filename: str,
    file_content: bytes,
    output_dir: Path
) -> str:
    """Process and save a file (JPG/PNG/PPTX).
    
    Routes to appropriate processor based on file type.
    
    Args:
        filename: Original filename
        file_content: Binary content of the file
        output_dir: Directory where to save the file
        
    Returns:
        Saved filename (sanitized)
        
    Raises:
        FileValidationError: If validation fails
        FileProcessingError: If processing fails
    """
    ext = Path(filename).suffix.lower()
    
    if ext in {".jpg", ".jpeg", ".png"}:
        return process_image_file(filename, file_content, output_dir)
    elif ext == ".pptx":
        return process_pptx_file(filename, file_content, output_dir)
    else:
        raise FileValidationError(f"Unsupported file extension: {ext}")


def cleanup_theme_file(file_path: Path) -> None:
    """Clean up a background template file.
    
    Args:
        file_path: Path to file to delete
    """
    try:
        if file_path.exists() and file_path.is_file():
            file_path.unlink()
            logger.info("Cleaned up file: %s", file_path)
    except Exception as e:
        logger.error("Error cleaning up file: %s", str(e), exc_info=True)
