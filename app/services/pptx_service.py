"""
PowerPoint to images conversion service.
"""

import os
import shutil
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
    sanitize_folder_name,
)

logger = get_logger(__name__)


class PPTXService:
    """Service for converting PowerPoint presentations to images."""

    def __init__(self):
        self.image_format = settings.pptx_image_format

    def convert_pptx_to_images(
        self, file_content: bytes, filename: str, display_name: str, project_type: str
    ) -> Dict[str, Any]:
        """
        Convert PPTX file to images and extract slide titles.

        Args:
            file_content: PPTX file content as bytes
            filename: Original filename
            display_name: Display name to use as folder name
            project_type: Type of project ('bipresents' or 'nw')

        Returns:
            Dictionary with conversion results
        """
        # Get the base directory for the project type
        base_dir = get_project_base_dir(project_type)

        # Use display_name as the conversion ID (folder name)
        conversion_id = display_name
        project_folder = base_dir / conversion_id
        thumbnails_folder = project_folder / "Thumbnails"

        # Check if folder already exists and handle it
        if project_folder.exists():
            logger.warning(
                "Project folder already exists: %s. It will be overwritten.",
                project_folder,
            )
            # Remove existing folder to avoid conflicts
            shutil.rmtree(project_folder)

        # Create directories
        project_folder.mkdir(parents=True, exist_ok=True)
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
                pptx_file_path, project_folder, thumbnails_folder
            )

            # Create response data
            image_urls = []
            thumbnail_urls = []
            titles = []

            # Generate public URLs for NW and BSR projects
            url_root = get_relative_slide_root(project_type)

            for result in slide_results:
                image_path = result["image_path"]
                thumbnail_path = result["thumbnail_path"]
                title = result["title"]
                image_filename = Path(image_path).name
                thumbnail_filename = Path(thumbnail_path).name
                if url_root:
                    image_url = build_relative_slide_path(
                        project_type,
                        conversion_id,
                        image_filename,
                    )
                    thumbnail_url = build_relative_slide_path(
                        project_type,
                        conversion_id,
                        thumbnail_filename,
                        subdir="Thumbnails",
                    )
                else:
                    image_url = f"files/download/{project_type}/{conversion_id}/{image_filename}"
                    thumbnail_url = (
                        f"files/download/{project_type}/{conversion_id}/"
                        f"Thumbnails/{thumbnail_filename}"
                    )
                image_urls.append(image_url)
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

    def _pptx_to_images_with_titles(
        self, pptx_path: Path, project_folder: Path, thumbnails_folder: Path
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
        thumbnails_folder_abs = str(thumbnails_folder.absolute())

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
                # Thumbnail image path
                thumbnail_path = os.path.join(thumbnails_folder_abs, img_filename)

                # Export slide as full-size image
                slide = presentation.Slides(i)
                slide.Export(img_path, self.image_format)

                # Export slide as thumbnail (smaller size)
                # Using a smaller resolution for thumbnails
                slide.Export(
                    thumbnail_path, self.image_format, 150, 113
                )  # Smaller dimensions for thumbnails

                # Extract slide title
                title = self._extract_slide_title(slide)

                logger.info(
                    "Slide %d exported as %s and %s with title: %s",
                    i,
                    img_path,
                    thumbnail_path,
                    title,
                )

                results.append(
                    {
                        "image_path": str(Path(img_path)),
                        "thumbnail_path": str(Path(thumbnail_path)),
                        "title": title,
                    }
                )

        except Exception as e:
            logger.error("Error processing PPTX with PowerPoint: %s", str(e))
            raise ValueError(f"Error processing PPTX with PowerPoint: {str(e)}") from e

        finally:
            # Clean up COM objects
            try:
                if presentation:
                    presentation.Close()
                    logger.info("Presentation closed")

                if powerpoint:
                    powerpoint.Quit()
                    logger.info("PowerPoint closed")

            except Exception as cleanup_error:
                logger.warning(
                    "Error during PowerPoint cleanup: %s", str(cleanup_error)
                )
                raise

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
        """Delete conversion files."""
        project_folder = self.get_conversion_folder(conversion_id, project_type)

        deleted = False

        if project_folder.exists():
            shutil.rmtree(project_folder)
            deleted = True
            logger.info("Project folder deleted: %s", project_folder)

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
