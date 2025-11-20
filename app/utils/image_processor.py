"""Image processing utilities for generating thumbnails and optimizing images."""

import logging
from pathlib import Path
from PIL import Image
from typing import Tuple

logger = logging.getLogger(__name__)

# Thumbnail dimensions (16:9 aspect ratio)
THUMBNAIL_WIDTH = 300
THUMBNAIL_HEIGHT = 169

class ImageProcessingError(Exception):
    """Exception raised for image processing errors."""
    pass


def create_thumbnail(
    source_path: Path,
    output_path: Path,
    size: Tuple[int, int] = (THUMBNAIL_WIDTH, THUMBNAIL_HEIGHT),
    quality: int = 85
) -> Path:
    """
    Create a thumbnail from an image file.
    
    Args:
        source_path: Path to the source image
        output_path: Path where thumbnail will be saved
        size: Tuple of (width, height) for thumbnail. Default is (300, 169) for 16:9 ratio
        quality: JPEG quality (1-100). Default is 85 for good quality with small size
    
    Returns:
        Path: Path to the created thumbnail
        
    Raises:
        ImageProcessingError: If thumbnail creation fails
    """
    try:
        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Open and process image
        with Image.open(source_path) as img:
            # Convert to RGB if necessary (for PNG with transparency)
            if img.mode in ('RGBA', 'LA', 'P'):
                # Create white background
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Create thumbnail maintaining aspect ratio
            img.thumbnail(size, Image.Resampling.LANCZOS)
            
            # Save as JPEG with specified quality
            img.save(output_path, 'JPEG', quality=quality, optimize=True)
            
        logger.info(
            "Thumbnail created: %s -> %s (size: %dx%d, quality: %d)",
            source_path.name,
            output_path.name,
            size[0],
            size[1],
            quality
        )
        
        return output_path
        
    except FileNotFoundError as e:
        logger.error("Source image not found: %s", source_path)
        raise ImageProcessingError(f"Source image not found: {source_path}") from e
    except Exception as e:
        logger.error("Failed to create thumbnail for %s: %s", source_path, str(e))
        raise ImageProcessingError(f"Failed to create thumbnail: {str(e)}") from e


def get_thumbnail_filename(original_filename: str) -> str:
    """
    Get the thumbnail filename for an original filename.
    
    Args:
        original_filename: Original image filename (e.g., 'slide_1.png')
    
    Returns:
        str: Thumbnail filename (e.g., 'slide_1.jpg')
        
    Note:
        Thumbnails are always saved as JPEG for better compression,
        so the extension is changed to .jpg
    """
    stem = Path(original_filename).stem
    return f"{stem}.jpg"


def ensure_thumbnail_exists(
    image_path: Path,
    thumbnails_dir: Path
) -> Path:
    """
    Ensure a thumbnail exists for the given image.
    Creates it if it doesn't exist.
    
    Args:
        image_path: Path to the original image
        thumbnails_dir: Directory where thumbnails are stored
    
    Returns:
        Path: Path to the thumbnail
        
    Raises:
        ImageProcessingError: If thumbnail creation fails
    """
    thumbnail_filename = get_thumbnail_filename(image_path.name)
    thumbnail_path = thumbnails_dir / thumbnail_filename
    
    if not thumbnail_path.exists():
        logger.info("Thumbnail doesn't exist, creating: %s", thumbnail_filename)
        create_thumbnail(image_path, thumbnail_path)
    
    return thumbnail_path
