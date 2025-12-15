"""
PowerPoint to images conversion service.
"""

import gc
import os
import shutil
import time
import zipfile
from pathlib import Path
from typing import Any, Dict, List

# pylint: disable=E1101
import pythoncom
from win32com import client

from app.config.settings import settings
from app.utils.files_utils import FileUtils
from app.utils.logging_utils import get_logger
from app.utils.path_utils import (
    build_relative_slide_path,
    get_project_base_dir,
    get_relative_slide_root,
    resolve_project_output,
    sanitize_folder_name,
)

logger = get_logger(__name__)


def _force_delete_windows(path: Path, max_retries: int = 5, retry_delay: float = 0.5) -> bool:
    """
    Force delete a folder on Windows with retry logic to handle file locks.

    Windows file locking (especially with IIS) can prevent immediate deletion.
    This function implements multiple strategies:
    1. Try normal deletion
    2. Run garbage collection to release file handles
    3. Retry with exponential backoff
    4. Individual file deletion if folder deletion fails

    Args:
        path: Path to delete
        max_retries: Maximum number of retry attempts
        retry_delay: Initial delay between retries (doubles each retry)

    Returns:
        True if deletion succeeded, False otherwise
    """
    if not path.exists():
        logger.debug("Path does not exist, nothing to delete: %s", path)
        return True

    for attempt in range(max_retries):
        try:
            # Strategy 1: Try normal deletion
            if path.is_file():
                path.unlink()
            else:
                shutil.rmtree(path, ignore_errors=False)

            logger.debug("Successfully deleted on attempt %d: %s", attempt + 1, path)
            return True

        except PermissionError as e:
            if attempt < max_retries - 1:
                # Force garbage collection to release file handles
                gc.collect()

                # Calculate backoff delay
                current_delay = retry_delay * (2 ** attempt)
                logger.warning(
                    "Deletion failed (attempt %d/%d) for %s: %s. Retrying in %.2fs...",
                    attempt + 1,
                    max_retries,
                    path,
                    str(e),
                    current_delay
                )
                time.sleep(current_delay)
            else:
                # Last attempt failed, try to delete individual files
                logger.error(
                    "Failed to delete folder after %d attempts: %s. Attempting individual file deletion...",
                    max_retries,
                    path
                )
                return _force_delete_contents(path)

        except Exception as e:
            logger.error("Unexpected error deleting %s: %s", path, str(e), exc_info=True)
            return False

    return False


def _force_delete_contents(folder_path: Path) -> bool:
    """
    Attempt to delete all contents of a folder individually (last resort).

    On Windows/IIS, the folder itself may be locked but files can still be deleted.
    This function deletes all files and subdirectories, leaving the empty locked folder.

    Args:
        folder_path: Folder to clean

    Returns:
        True if most files were deleted (even if folder remains), False if major failures
    """
    if not folder_path.exists() or not folder_path.is_dir():
        return True

    deleted_count = 0
    failed_count = 0

    try:
        # First pass: Delete all files with retry for each file
        for item in folder_path.rglob("*"):
            if item.is_file():
                try:
                    # Try multiple times for each file with a short delay
                    for attempt in range(3):
                        try:
                            item.unlink()
                            deleted_count += 1
                            break
                        except PermissionError:
                            if attempt < 2:
                                time.sleep(0.1)
                            else:
                                raise
                except Exception as e:
                    logger.warning("Failed to delete file %s after retries: %s", item, str(e))
                    failed_count += 1

        # Second pass: Remove empty subdirectories (bottom-up)
        for item in sorted(folder_path.rglob("*"), key=lambda p: len(p.parts), reverse=True):
            if item.is_dir():
                try:
                    item.rmdir()
                    deleted_count += 1
                except Exception:
                    # Directory not empty or locked, that's OK
                    pass

        # Try to remove the root folder (may fail due to IIS lock, that's acceptable)
        try:
            folder_path.rmdir()
            logger.info("✅ Successfully removed folder after individual file deletion: %s", folder_path)
            return True
        except Exception as e:
            # Folder is locked but if we deleted files, consider it success
            if deleted_count > 0:
                logger.info(
                    "✅ Deleted %d files from locked folder: %s (folder remains but will be overwritten)",
                    deleted_count,
                    folder_path
                )
                return True  # SUCCESS: Files deleted even if folder remains
            else:
                logger.error(
                    "❌ Failed to delete any files from %s (deleted=%d, failed=%d)",
                    folder_path,
                    deleted_count,
                    failed_count
                )
                return False

    except Exception as e:
        logger.error("Error during individual file deletion for %s: %s", folder_path, str(e), exc_info=True)
        return deleted_count > 0  # Return True if we managed to delete at least some files


class PPTXService:
    """Service for converting PowerPoint presentations to images."""

    def __init__(self):
        self.image_format = settings.pptx_image_format

    def convert_pptx_to_images(
        self,
        file_content: bytes,
        filename: str,
        display_name: str,
        project_type: str,
        *,
        skip_cleanup: bool = False,
    ) -> Dict[str, Any]:
        """
        Convert PPTX file to images and extract slide titles.

        Args:
            file_content: PPTX file content as bytes
            filename: Original filename
            display_name: Display name to use as folder name
            project_type: Type of project ('bipresents' or 'nw')
            skip_cleanup: If True, keep existing files in the project folder (useful when
                converting a backup PPTX so we don't delete the originally uploaded Excel/PPTX).

        Returns:
            Dictionary with conversion results
        """
        # Resolve destination folder (sanitized) and ensure deterministic ID
        conversion_id = sanitize_folder_name(display_name)
        project_folder, used_fallback = resolve_project_output(
            display_name,
            project_type,
        )
        thumbnails_folder = project_folder / "Thumbnails"
        # Decide if thumbnails should be generated (re-enabled for all types to ensure UI images load)
        generate_thumbnails = True

        if used_fallback:
            logger.warning(
                "Project type '%s' does not map to a configured base directory; using fallback at %s",
                project_type,
                project_folder,
            )

        # Check if folder already exists and handle it
        if project_folder.exists():
            if skip_cleanup:
                logger.info(
                    "Project folder already exists: %s. Skipping cleanup to preserve existing files.",
                    project_folder,
                )
            else:
                logger.warning(
                    "Project folder already exists: %s. Cleaning contents for overwrite.",
                    project_folder,
                )
                # On Windows/IIS, folder may be locked but we can delete contents
                # Try to delete the folder completely first
                if not _force_delete_windows(project_folder):
                    # Folder deletion failed (locked by IIS), but _force_delete_windows
                    # will have cleaned the contents. The folder remains but is empty.
                    # This is acceptable - we'll reuse the existing folder structure.
                    logger.info(
                        "Folder remains locked but contents cleaned: %s. Reusing folder structure.",
                        project_folder
                    )

        # Create directories (exist_ok=True allows reusing locked folders)
        project_folder.mkdir(parents=True, exist_ok=True)
        if generate_thumbnails:
            thumbnails_folder.mkdir(parents=True, exist_ok=True)

        try:
            # Save the uploaded PowerPoint file in the project folder
            pptx_file_path = project_folder / filename

            with open(pptx_file_path, "wb") as buffer:
                buffer.write(file_content)

            logger.info(
                "PPTX file saved: %s (%s) for project: %s",
                pptx_file_path,
                FileUtils.format_file_size(len(file_content)),
                conversion_id,
            )

            # Convert PPTX to images and extract titles
            slide_results = self._pptx_to_images_with_titles(
                pptx_file_path, project_folder, thumbnails_folder, generate_thumbnails=generate_thumbnails
            )

            # Create response data
            image_urls = []
            thumbnail_urls = []
            titles = []

            # Generate cache busting timestamp to prevent browser caching issues
            # This ensures that when images are overwritten, browsers load the new versions
            cache_bust_timestamp = int(time.time() * 1000)  # Milliseconds since epoch

            # Generate public URLs for NW and BSR projects
            url_root = get_relative_slide_root(project_type)

            for result in slide_results:
                image_path = result["image_path"]
                thumbnail_path = result.get("thumbnail_path")
                title = result["title"]
                image_filename = Path(image_path).name
                thumbnail_filename = Path(thumbnail_path).name if thumbnail_path else None
                if url_root:
                    image_url = build_relative_slide_path(
                        project_type,
                        conversion_id,
                        image_filename,
                        cache_bust=cache_bust_timestamp,
                    )
                    if generate_thumbnails and thumbnail_filename:
                        thumbnail_url = build_relative_slide_path(
                            project_type,
                            conversion_id,
                            thumbnail_filename,
                            subdir="Thumbnails",
                            cache_bust=cache_bust_timestamp,
                        )
                else:
                    image_url = f"files/download/{project_type}/{conversion_id}/{image_filename}?v={cache_bust_timestamp}"
                    if generate_thumbnails and thumbnail_filename:
                        thumbnail_url = (
                            f"files/download/{project_type}/{conversion_id}/"
                            f"Thumbnails/{thumbnail_filename}?v={cache_bust_timestamp}"
                        )
                image_urls.append(image_url)
                if generate_thumbnails and thumbnail_filename:
                    thumbnail_urls.append(thumbnail_url)
                titles.append(title)

            return {
                "conversion_id": conversion_id,
                "project_type": project_type,
                "total_images": len(image_urls),
                "images": image_urls,
                "thumbnails": thumbnail_urls,
                "titles": titles,
                "pptx_file": (
                    f"files/download/{project_type}/{conversion_id}/{filename}"
                ),
                "message": (
                    f"Conversion successful for '{conversion_id}'. "
                    f"Generated {len(image_urls)} images."
                ),
            }

        except Exception as e:
            # Clean up directories on error
            self._cleanup_directories(project_folder)
            logger.error("Error in PPTX conversion: %s", str(e))
            raise

    def replace_project_images(
        self, file_content: bytes, filename: str, project_name: str, project_type: str
    ) -> Dict[str, Any]:
        """Replace slide images for an existing project folder.

        This writes the provided PPTX to the target project folder and exports
        each slide as JPEG into the project folder and its Thumbnails subfolder,
        overwriting existing images (001.jpg, 002.jpg, ...).

        Args:
            file_content: PPTX content bytes
            filename: Original uploaded filename (for reference)
            project_name: Existing project folder name
            project_type: 'nw' or 'bipresents'

        Returns:
            Dict with counts and relative paths to images and thumbnails
        """
        # Validate project_name to prevent path traversal
        if not project_name or project_name != Path(project_name).name:
            raise ValueError("Invalid project name")

        # Sanitize project_name to match actual folder name
        sanitized_project_name = sanitize_folder_name(project_name)

        # Resolve existing project folder
        base_dir = get_project_base_dir(project_type)
        project_folder = base_dir / sanitized_project_name
        thumbnails_folder = project_folder / "Thumbnails"
        # Always generate thumbnails to ensure UI thumbnails exist
        generate_thumbnails = True

        if not project_folder.exists():
            raise FileNotFoundError(
                f"Project folder not found: {project_folder} (sanitized from '{project_name}')"
            )

        # Ensure thumbnails folder exists only if generating thumbnails
        if generate_thumbnails:
            thumbnails_folder.mkdir(parents=True, exist_ok=True)

        # Remove existing JPGs to avoid stale files if slide count shrinks
        try:
            for f in project_folder.glob("*.jpg"):
                try:
                    f.unlink()
                except Exception:
                    pass
            if generate_thumbnails:
                for f in thumbnails_folder.glob("*.jpg"):
                    try:
                        f.unlink()
                    except Exception:
                        pass
        except Exception as e:
            logger.warning("Failed to clean old images in %s: %s", project_folder, str(e))

        # Save PPTX into the project folder (temporary name)
        pptx_file_path = project_folder / (Path(filename).name or "updated_layout.pptx")
        with open(pptx_file_path, "wb") as buffer:
            buffer.write(file_content)

        # Export slides using PowerPoint COM
        slide_results = self._pptx_to_images_with_titles(
            pptx_file_path, project_folder, thumbnails_folder, generate_thumbnails=generate_thumbnails
        )

        # Build relative URLs similar to convert flow
        image_urls: List[str] = []
        thumbnail_urls: List[str] = []
        titles: List[str] = []

        # Generate cache busting timestamp to prevent browser caching issues
        cache_bust_timestamp = int(time.time() * 1000)  # Milliseconds since epoch

        # Build relative URLs using sanitized folder name
        root = get_relative_slide_root(project_type)
        for result in slide_results:
            image_filename = Path(result["image_path"]).name
            thumbnail_filename = Path(result["thumbnail_path"]).name if result.get("thumbnail_path") else None
            if root:
                image_url = "/".join([root, sanitized_project_name, image_filename])
                image_url = f"{image_url}?v={cache_bust_timestamp}"
                image_urls.append(image_url)
                if generate_thumbnails and thumbnail_filename:
                    thumbnail_url = "/".join([root, sanitized_project_name, "Thumbnails", thumbnail_filename])
                    thumbnail_url = f"{thumbnail_url}?v={cache_bust_timestamp}"
                    thumbnail_urls.append(thumbnail_url)
            else:
                image_urls.append(str(Path(result["image_path"])) )
                if generate_thumbnails and thumbnail_filename:
                    thumbnail_urls.append(str(Path(result["thumbnail_path"])) )
            titles.append(result.get("title", ""))

        return {
            "project_name": sanitized_project_name,
            "project_type": project_type,
            "total_images": len(image_urls),
            "images": image_urls,
            "thumbnails": thumbnail_urls,
            "titles": titles,
            "message": (
                f"Replaced images for '{project_name}'. Generated {len(image_urls)} images."
            ),
        }

    def _pptx_to_images_with_titles(
        self, pptx_path: Path, project_folder: Path, thumbnails_folder: Path, generate_thumbnails: bool = True
    ) -> List[Dict[str, Any]]:
        """
        Convert each slide of the PPTX to an image using Microsoft PowerPoint.

        Args:
            pptx_path: Path to the PPTX file
            project_folder: Directory to save full-size images
            thumbnails_folder: Directory to save thumbnail images

        Returns:
            List of dictionaries containing image paths and slide titles
        """
        results = []

        # Convert to absolute paths
        pptx_path_abs = str(pptx_path.absolute())
        project_folder_abs = str(project_folder.absolute())
        thumbnails_folder_abs = str(thumbnails_folder.absolute()) if generate_thumbnails else ""

        # Initialize COM in this thread
        pythoncom.CoInitialize()

        powerpoint = None
        presentation = None

        try:
            # Start PowerPoint
            powerpoint = client.Dispatch("PowerPoint.Application")
            # PowerPoint visibility is handled automatically

            # Open the presentation
            presentation = powerpoint.Presentations.Open(pptx_path_abs)
            logger.info("Presentation opened with %d slides", presentation.Slides.Count)

            # Process each slide
            for i in range(1, presentation.Slides.Count + 1):
                # Generate filename with zero-padding
                img_filename = f"{i:03d}.jpg"  # 001.jpg, 002.jpg, etc.

                # Full-size image path
                img_path = os.path.join(project_folder_abs, img_filename)
                # Thumbnail image path (optional)
                thumbnail_path = os.path.join(thumbnails_folder_abs, img_filename) if generate_thumbnails else None

                # Export slide as full-size image
                slide = presentation.Slides(i)
                slide.Export(img_path, self.image_format)

                # Export slide as thumbnail (optional)
                if generate_thumbnails and thumbnail_path:
                    # For BSR, thumbnails have same quality as full-size images
                    slide.Export(thumbnail_path, self.image_format)

                # Extract slide title
                title = self._extract_slide_title(slide)

                if generate_thumbnails and thumbnail_path:
                    logger.info(
                        "Slide %d exported as %s and %s with title: %s",
                        i,
                        img_path,
                        thumbnail_path,
                        title,
                    )
                else:
                    logger.info(
                        "Slide %d exported as %s with title: %s",
                        i,
                        img_path,
                        title,
                    )

                result_item = {
                    "image_path": str(Path(img_path)),
                    "title": title,
                }
                if generate_thumbnails and thumbnail_path:
                    result_item["thumbnail_path"] = str(Path(thumbnail_path))
                results.append(result_item)

        except Exception as e:
            logger.error("Error processing PPTX with PowerPoint: %s", str(e))
            raise ValueError(f"Error processing PPTX with PowerPoint: {str(e)}") from e

        finally:
            # Clean up COM objects
            try:
                if presentation:
                    presentation.Close()
                    logger.info("Presentation closed")
            except Exception as close_error:
                logger.debug(
                    "Presentation already closed or error closing: %s", str(close_error)
                )

            try:
                if powerpoint:
                    powerpoint.Quit()
                    logger.info("PowerPoint closed")
            except Exception as quit_error:
                logger.debug(
                    "PowerPoint already closed or error quitting: %s", str(quit_error)
                )

            # Release COM resources
            pythoncom.CoUninitialize()

        return results

    def _extract_slide_title(self, slide) -> str:
        """Extract title from a PowerPoint slide."""
        title = "No Title"

        try:
            # Check if the slide has a title placeholder
            for shape in slide.Shapes:
                if shape.Type == 14:  # msoPlaceholder
                    if shape.PlaceholderFormat.Type == 1:  # ppPlaceholderTitle
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            title = shape.TextFrame.TextRange.Text.strip()
                            break
        except Exception as e:
            logger.warning("Error extracting slide title: %s", str(e))
            raise

        return title

    def get_conversion_folder(self, conversion_id: str, project_type: str) -> Path:
        """Get the project folder for a conversion."""
        base_dir = get_project_base_dir(project_type)
        return base_dir / conversion_id

    def get_image_path(
        self, conversion_id: str, image_name: str, project_type: str
    ) -> Path:
        """Get the path to a specific image."""
        base_dir = get_project_base_dir(project_type)
        return base_dir / conversion_id / image_name

    def get_thumbnail_path(
        self, conversion_id: str, image_name: str, project_type: str
    ) -> Path:
        """Get the path to a specific thumbnail."""
        base_dir = get_project_base_dir(project_type)
        return base_dir / conversion_id / "Thumbnails" / image_name

    def conversion_exists(self, conversion_id: str, project_type: str) -> bool:
        """Check if a conversion exists."""
        return self.get_conversion_folder(conversion_id, project_type).exists()

    def delete_conversion(self, conversion_id: str, project_type: str) -> bool:
        """Delete conversion files with Windows-compatible retry logic.

        Args:
            conversion_id: The conversion/display name (will be sanitized to match folder name)
            project_type: The project type

        Returns:
            True if folder was deleted, False if folder didn't exist or deletion failed
        """
        # CRITICAL: Sanitize conversion_id to match how folders are created
        # Folders are always created using sanitize_folder_name() in convert_pptx_to_images()
        sanitized_id = sanitize_folder_name(conversion_id)
        project_folder = self.get_conversion_folder(sanitized_id, project_type)

        if not project_folder.exists():
            logger.warning("⚠️ Project folder does not exist (already deleted or never created): %s", project_folder)
            return False

        logger.info("Deleting project folder: %s (from conversion_id: %s)", project_folder, conversion_id)

        # Use Windows-compatible deletion with retry logic
        deleted = _force_delete_windows(project_folder, max_retries=5, retry_delay=0.5)

        if deleted:
            logger.info("✅ Project folder deleted successfully: %s", project_folder)
        else:
            logger.error(
                "❌ Failed to delete project folder after multiple attempts: %s. "
                "Files may be locked by IIS or another process.",
                project_folder
            )

        return deleted

    def create_zip_archive(self, conversion_id: str, project_type: str) -> Path:
        """Create a ZIP archive of all files from a conversion."""

        project_folder = self.get_conversion_folder(conversion_id, project_type)

        if not project_folder.exists():
            raise FileNotFoundError(f"Conversion {conversion_id} not found")

        # Create ZIP file in the project folder
        zip_filename = f"{conversion_id}_slides.zip"
        zip_path = project_folder / zip_filename

        with zipfile.ZipFile(zip_path, "w") as zipf:
            # Add all files in the project folder
            for file in project_folder.rglob("*"):
                if (
                    file.is_file() and file != zip_path
                ):  # Don't include the zip file itself
                    # Get relative path for the archive
                    arcname = file.relative_to(project_folder)
                    zipf.write(file, arcname=arcname)

        logger.info("ZIP archive created: %s", zip_path)
        return zip_path

    def _cleanup_directories(self, *directories: Path) -> None:
        """Clean up directories in case of error."""
        for directory in directories:
            if directory.exists():
                try:
                    shutil.rmtree(directory)
                    logger.info("Cleaned up directory: %s", directory)
                except Exception as e:
                    logger.warning(
                        "Failed to clean up directory %s: %s", directory, str(e)
                    )
                    raise

    def cleanup_old_conversions(self, max_conversions: int = 20) -> None:
        """Clean up old conversion files to prevent disk space issues."""
        try:
            # Clean up both project types
            for project_type in [
                settings.PROJECT_TYPE_BIPRESENTS,
                settings.PROJECT_TYPE_NW,
            ]:
                base_dir = get_project_base_dir(project_type)

                if not base_dir.exists():
                    continue

                project_folders = [d for d in base_dir.iterdir() if d.is_dir()]

                if len(project_folders) > max_conversions:
                    project_folders.sort(key=lambda x: x.stat().st_mtime)
                    folders_to_delete = project_folders[:-max_conversions]

                    for folder in folders_to_delete:
                        shutil.rmtree(folder)
                        logger.info("Deleted old project folder: %s", folder)

        except Exception as e:
            logger.error("Error cleaning up old conversions: %s", str(e))
            raise

    def list_conversions(self, project_type: str = None) -> List[Dict[str, Any]]:
        """
        List all available conversions.

        Args:
            project_type: Optional project type filter

        Returns:
            List of conversion information
        """
        conversions = []

        project_types = (
            [project_type]
            if project_type
            else [settings.PROJECT_TYPE_BIPRESENTS, settings.PROJECT_TYPE_NW]
        )

        for ptype in project_types:
            try:
                base_dir = get_project_base_dir(ptype)
                if not base_dir.exists():
                    continue

                for project_folder in base_dir.iterdir():
                    if project_folder.is_dir():
                        # Count images (excluding thumbnails folder)
                        image_count = len(
                            [
                                f
                                for f in project_folder.iterdir()
                                if f.is_file() and f.suffix.lower() == ".jpg"
                            ]
                        )

                        # Get creation time
                        creation_time = project_folder.stat().st_mtime

                        conversions.append(
                            {
                                "conversion_id": project_folder.name,
                                "project_type": ptype,
                                "image_count": image_count,
                                "creation_time": creation_time,
                                "folder_path": str(project_folder),
                            }
                        )
            except (OSError, PermissionError, FileNotFoundError) as e:
                logger.error("Error listing conversions for %s: %s", ptype, str(e))

        return conversions


# Global PPTX service instance
pptx_service = PPTXService()
