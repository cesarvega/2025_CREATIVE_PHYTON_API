"""Shared path utilities for slide generation services."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Tuple

from app.config.settings import settings

_INVALID_FS_CHARS = r'[<>:"/\\|?*]'
_SANITIZE_PATTERN = re.compile(_INVALID_FS_CHARS)
_REPLACE_PATTERN = re.compile(r"[\s\-.]+")


def sanitize_folder_name(folder_name: str) -> str:
    """Sanitize a folder name for filesystem usage.

    Matches the legacy behavior from PPTX conversions: removes invalid characters,
    replaces whitespace/special separators with single underscores, trims leading
    and trailing punctuation, and uppercases within a safe length.
    """
    sanitized = _SANITIZE_PATTERN.sub("", folder_name)
    sanitized = _REPLACE_PATTERN.sub("_", sanitized)
    sanitized = sanitized.strip("_.")
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
    """Return the web-visible root folder for slide assets."""
    if not project_type:
        return ""
    normalized = project_type.lower()
    if normalized == settings.PROJECT_TYPE_BIPRESENTS:
        return "bsr_slides"
    if normalized == settings.PROJECT_TYPE_NW:
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
) -> str:
    """Construct a relative URL path for slide assets."""
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
    return "/".join(parts)


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
