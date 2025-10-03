"""Utilities for secure file downloads and token management.

This module provides functionality for creating secure download tokens,
validating file paths, and managing file downloads through the API.
"""

import base64
import json
from pathlib import Path
from typing import Tuple

from fastapi import HTTPException

from app.config.settings import settings


# Allowed root directories for downloads
ALLOWED_DOWNLOAD_ROOTS = tuple(
    {
        settings.base_dir.resolve(),
        settings.base_dir_bipresents.resolve(),
        settings.base_dir_nw.resolve(),
        settings.nw_files_dir.resolve(),
    }
)


def is_subpath(path: Path, parent: Path) -> bool:
    """Check if a path is a subpath of a parent directory.
    
    Args:
        path: The path to check.
        parent: The parent directory.
        
    Returns:
        True if path is a subpath of parent, False otherwise.
    """
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def create_download_token(path: Path) -> str:
    """Create a secure base64-encoded token for file download.
    
    Args:
        path: The file path to encode in the token.
        
    Returns:
        A URL-safe base64-encoded token.
    """
    payload = json.dumps({"path": str(path.resolve())}, separators=(",", ":"))
    token_bytes = base64.urlsafe_b64encode(payload.encode("utf-8"))
    return token_bytes.decode("utf-8").rstrip("=")


def decode_download_token(token: str) -> Path:
    """Decode and validate a download token.
    
    Args:
        token: The base64-encoded download token.
        
    Returns:
        The validated Path object.
        
    Raises:
        HTTPException: If the token is invalid, the path is unauthorized,
                      or the file doesn't exist.
    """
    # Add padding if necessary
    padding = "=" * (-len(token) % 4)
    
    try:
        decoded = base64.urlsafe_b64decode((token + padding).encode("utf-8"))
        payload = json.loads(decoded.decode("utf-8"))
    except (ValueError, json.JSONDecodeError) as exc:
        raise HTTPException(status_code=400, detail="Invalid download token.") from exc

    raw_path = payload.get("path")
    if not raw_path:
        raise HTTPException(status_code=400, detail="Invalid download token payload.")

    resolved_path = Path(raw_path).resolve()
    
    # Security check: ensure path is within allowed directories
    if not any(is_subpath(resolved_path, root) for root in ALLOWED_DOWNLOAD_ROOTS):
        raise HTTPException(
            status_code=403,
            detail="Requested file is outside authorized directories.",
        )

    if not resolved_path.exists():
        raise HTTPException(status_code=404, detail="Requested file not found.")

    return resolved_path


def build_api_download_url(path: Path, endpoint_prefix: str = "/api/presentations/files") -> str:
    """Build a complete API download URL for a file path.
    
    Args:
        path: The file path to create a download URL for.
        endpoint_prefix: The API endpoint prefix for downloads.
        
    Returns:
        A complete download URL with encoded token.
    """
    token = create_download_token(path)
    return f"{endpoint_prefix}/{token}"


def guess_media_type(path: Path) -> str:
    """Guess the MIME media type based on file extension.
    
    Args:
        path: The file path to determine media type for.
        
    Returns:
        The MIME type string, or 'application/octet-stream' if unknown.
    """
    suffix = path.suffix.lower()
    
    # Common media types mapping
    media_types = {
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".pptm": "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
        ".ppt": "application/vnd.ms-powerpoint",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
        ".xls": "application/vnd.ms-excel",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".doc": "application/msword",
        ".pdf": "application/pdf",
        ".zip": "application/zip",
        ".json": "application/json",
        ".xml": "application/xml",
        ".txt": "text/plain",
        ".csv": "text/csv",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".gif": "image/gif",
        ".svg": "image/svg+xml",
    }
    
    return media_types.get(suffix, "application/octet-stream")


def validate_download_path(path: Path, allowed_roots: Tuple[Path, ...] = ALLOWED_DOWNLOAD_ROOTS) -> bool:
    """Validate that a path is within allowed download directories.
    
    Args:
        path: The path to validate.
        allowed_roots: Tuple of allowed root directories.
        
    Returns:
        True if path is valid and within allowed directories.
    """
    if not path.exists():
        return False
    
    resolved_path = path.resolve()
    return any(is_subpath(resolved_path, root) for root in allowed_roots)
