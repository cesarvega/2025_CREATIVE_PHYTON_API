"""Shared path utilities for slide generation services."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Tuple

from app.config.settings import settings

_INVALID_FS_CHARS = r'[<>:"/\\|?*]'
_SANITIZE_PATTERN = re.compile(_INVALID_FS_CHARS)
_REPLACE_PATTERN = re.compile(r"[\s.]+")


def sanitize_folder_name(folder_name: str) -> str:
    """Sanitize a folder name for filesystem usage.

    Removes invalid characters, replaces whitespace/dots with single underscores,
    trims leading and trailing punctuation, and uppercases within a safe length.
    Hyphens are preserved: the frontend requests slide images using the raw
    display name, so folder names must keep them (e.g. "TRI-TAM").
    """
    sanitized = _SANITIZE_PATTERN.sub("", folder_name)
    sanitized = _REPLACE_PATTERN.sub("_", sanitized)
    sanitized = sanitized.strip("_. ")  # Also strip spaces
    if not sanitized:
        sanitized = "unnamed_project"
    return sanitized[:50].upper()


def get_project_base_dir(project_type: str) -> Path:
    """Return the base directory for a given project type.

    Args:
        project_type: Project type identifier (case-insensitive).

    Raises:
        ValueError: If an unknown project type is provided.
    """
    normalized = project_type.lower()
    return settings.get_base_dir_for_project_type(normalized)


def get_relative_slide_root(project_type: Optional[str]) -> str:
    """Return the web-visible root folder for slide assets.

    Accepts both canonical project types (e.g., 'bipresents', 'nw', 'dw') and
    common aliases such as 'bsr'.
    """
    if not project_type:
        return ""
    normalized = project_type.lower()
    # Map BSR to the bipresents directory structure
    if normalized in {settings.PROJECT_TYPE_BIPRESENTS, "bsr"}:
        return "bsr_slides"
    if normalized == settings.PROJECT_TYPE_NW or normalized == settings.PROJECT_TYPE_DW:
        return "nw_slides"
    return ""


def resolve_project_output(
    display_name: str,
    project_type: Optional[str],
    *,
    fallback_subdir: str = "temp_slides",
) -> Tuple[Path, bool]:
    """Resolve the filesystem directory for generated slide assets.

    Args:
        display_name: Presentation display name (will be sanitized).
        project_type: Optional project type.
        fallback_subdir: Sub-directory within ``settings.nw_files_dir`` to use when
            the project type is missing or invalid.

    Returns:
        A tuple containing (output_path, used_fallback).
    """
    folder_name = sanitize_folder_name(display_name)
    fallback_root = settings.nw_files_dir / fallback_subdir

    if project_type:
        try:
            base_dir = get_project_base_dir(project_type)
            return base_dir / folder_name, False
        except ValueError:
            pass

    return fallback_root / folder_name, True


def build_relative_slide_path(
    project_type: Optional[str],
    display_name: str,
    filename: str,
    *,
    subdir: Optional[str] = None,
    fallback_root: Optional[str] = None,
    cache_bust: Optional[int] = None,
) -> str:
    """Construct a relative URL path for slide assets.

    Args:
        project_type: Type of project (e.g., 'bipresents', 'nw')
        display_name: Display name for the project
        filename: Name of the file
        subdir: Optional subdirectory (e.g., 'Thumbnails')
        fallback_root: Fallback root if project_type is not recognized
        cache_bust: Optional timestamp for cache busting (prevents browser caching issues)

    Returns:
        URL path with optional cache busting query parameter
    """
    parts = []
    root = get_relative_slide_root(project_type)
    if root:
        parts.append(root)
    elif fallback_root:
        parts.append(fallback_root)
    parts.append(sanitize_folder_name(display_name))
    if subdir:
        parts.append(subdir)
    parts.append(filename)

    url = "/".join(parts)

    # Add cache busting query parameter if provided
    if cache_bust is not None:
        url = f"{url}?v={cache_bust}"

    return url


def get_nw_downloads_dir() -> Path:
    """Get the NW downloads directory.

    Returns:
        Path to NW_Files/downloads in the project directory

    Note: For now, using local project directory due to permission issues.
    In production, this should be configured to use C:/nw_files/downloads
    """
    # Always use local project directory for now
    return settings.app_dir / "NW_Files" / "downloads"


def get_bsr_downloads_dir() -> Path:
    """Get the BSR downloads directory.

    Returns:
        Path to NW_Files/BRS_files/downloads in the project directory
    """
    return settings.app_dir / "NW_Files" / "BRS_files" / "downloads"


def get_nw_analytics_dir() -> Path:
    """Get the NW Analytics directory.

    Returns:
        Path to NW_Files/NWAnalytics in the project directory
    """
    return settings.app_dir / "NW_Files" / "NWAnalytics"


def get_bsr_analytics_dir() -> Path:
    """Get the BSR Analytics directory.

    Returns:
        Path to NW_Files/BRS_files/BSR_Analytics in the project directory
    """
    return settings.app_dir / "NW_Files" / "BRS_files" / "BSR_Analytics"
