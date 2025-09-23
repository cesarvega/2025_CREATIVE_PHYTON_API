"""
File utilities for the Report Generator API
"""

import logging
import mimetypes
import os
import shutil
import time
from pathlib import Path
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


class FileUtils:
    """Utility class for file operations"""

    @staticmethod
    def get_media_type(suffix: str) -> str:
        """
        Get media type based on file extension

        Args:
            suffix: File extension (e.g., '.pptx', '.png')

        Returns:
            MIME type string
        """
        media_types = {
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".pdf": "application/pdf",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".gif": "image/gif",
            ".bmp": "image/bmp",
            ".svg": "image/svg+xml",
            ".txt": "text/plain",
            ".csv": "text/csv",
            ".json": "application/json",
            ".xml": "application/xml",
            ".zip": "application/zip",
        }

        # Try custom mapping first, then use mimetypes module
        media_type = media_types.get(suffix.lower())
        if media_type is None:
            media_type, _ = mimetypes.guess_type(f"file{suffix}")
            if media_type is None:
                media_type = "application/octet-stream"

        return media_type

    @staticmethod
    def format_file_size(size_bytes: int) -> str:
        """
        Format file size in human readable format

        Args:
            size_bytes: File size in bytes

        Returns:
            Formatted size string (e.g., "1.5MB")
        """
        if size_bytes == 0:
            return "0B"

        size_names = ["B", "KB", "MB", "GB", "TB"]
        i = 0
        size = float(size_bytes)

        while size >= 1024 and i < len(size_names) - 1:
            size /= 1024.0
            i += 1

        return f"{size:.1f}{size_names[i]}"

    @staticmethod
    def safe_delete_file(file_path: Path) -> bool:
        """
        Safely delete a file with error handling

        Args:
            file_path: Path to the file to delete

        Returns:
            True if file was deleted successfully, False otherwise
        """
        try:
            if file_path.exists() and file_path.is_file():
                os.remove(file_path)
                logger.info("File deleted successfully: %s", file_path)
                return True
            else:
                logger.warning("File not found or not a file: %s", file_path)
                return False
        except PermissionError:
            logger.error("Permission denied when deleting file: %s", file_path)
            return False
        except Exception as e: # pylint: disable=broad-exception-caught
            logger.error("Error deleting file %s: %s", file_path, str(e))
            return False

    @staticmethod
    def safe_delete_directory(dir_path: Path, recursive: bool = True) -> bool:
        """
        Safely delete a directory with error handling

        Args:
            dir_path: Path to the directory to delete
            recursive: Whether to delete recursively

        Returns:
            True if directory was deleted successfully, False otherwise
        """
        try:
            if dir_path.exists() and dir_path.is_dir():
                if recursive:
                    shutil.rmtree(dir_path)
                else:
                    dir_path.rmdir()
                logger.info("Directory deleted successfully: %s", dir_path)
                return True
            else:
                logger.warning("Directory not found: %s", dir_path)
                return False
        except PermissionError:
            logger.error("Permission denied when deleting directory: %s", dir_path)
            return False
        except OSError as e:
            logger.error("OS error when deleting directory %s: %s", dir_path, str(e))
            return False
        except Exception as e: # pylint: disable=broad-exception-caught
            logger.error("Error deleting directory %s: %s", dir_path, str(e))
            return False

    @staticmethod
    def cleanup_old_files(
        directory: Path, older_than_hours: int = 24
    ) -> Dict[str, int]:
        """
        Clean up old files from directory

        Args:
            directory: Directory to clean up
            older_than_hours: Delete files older than this many hours

        Returns:
            Dictionary with cleanup statistics
        """
        stats = {"deleted": 0, "errors": 0, "total_checked": 0}

        try:
            if not directory.exists():
                logger.warning("Directory does not exist: %s", directory)
                return stats

            current_time = time.time()
            cutoff_time = current_time - (older_than_hours * 3600)

            for file_path in directory.iterdir():
                stats["total_checked"] += 1

                if file_path.is_file():
                    try:
                        if file_path.stat().st_mtime < cutoff_time:
                            if FileUtils.safe_delete_file(file_path):
                                stats["deleted"] += 1
                            else:
                                stats["errors"] += 1
                    except Exception as e: # pylint: disable=broad-exception-caught
                        logger.error("Error checking file %s: %s", file_path, str(e))
                        stats["errors"] += 1
                        raise

            logger.info("Cleanup completed for %s: %s", directory, stats)

        except Exception as e: # pylint: disable=broad-exception-caught
            logger.error("Error during directory cleanup %s: %s", directory, str(e))
            stats["errors"] += 1
            raise

        return stats

    @staticmethod
    def get_file_info(file_path: Path) -> Optional[Dict]:
        """
        Get detailed information about a file

        Args:
            file_path: Path to the file

        Returns:
            Dictionary with file information or None if file doesn't exist
        """
        try:
            if not file_path.exists():
                return None

            stat = file_path.stat()

            return {
                "name": file_path.name,
                "path": str(file_path),
                "size": stat.st_size,
                "size_formatted": FileUtils.format_file_size(stat.st_size),
                "created": stat.st_ctime,
                "modified": stat.st_mtime,
                "accessed": stat.st_atime,
                "extension": file_path.suffix,
                "mime_type": FileUtils.get_media_type(file_path.suffix),
                "is_readable": os.access(file_path, os.R_OK),
                "is_writable": os.access(file_path, os.W_OK),
            }

        except Exception as e: # pylint: disable=broad-exception-caught
            logger.error("Error getting file info for %s: %s", file_path, str(e))
            return None

    @staticmethod
    def list_directory_files(
        directory: Path, pattern: str = "*", limit: Optional[int] = None
    ) -> List[Dict]:
        """
        List files in a directory with their information

        Args:
            directory: Directory to list
            pattern: File pattern to match (e.g., "*.docx")
            limit: Maximum number of files to return

        Returns:
            List of file information dictionaries
        """
        files = []

        try:
            if not directory.exists():
                return files

            for file_path in directory.glob(pattern):
                if file_path.is_file():
                    file_info = FileUtils.get_file_info(file_path)
                    if file_info:
                        files.append(file_info)

            # Sort by modification time (newest first)
            files.sort(key=lambda x: x["modified"], reverse=True)

            # Apply limit if specified
            if limit and len(files) > limit:
                files = files[:limit]

        except Exception as e:
            logger.error("Error listing directory %s: %s", directory, str(e))
            raise

        return files

    @staticmethod
    def ensure_directory(path: Path) -> bool:
        """
        Ensure a directory exists, create if it doesn't

        Args:
            path: Directory path to ensure

        Returns:
            True if directory exists or was created successfully
        """
        try:
            path.mkdir(parents=True, exist_ok=True)
            return True
        except Exception as e: # pylint: disable=broad-exception-caught
            logger.error("Error creating directory %s: %s", path, str(e))
            return False

    @staticmethod
    def is_safe_path(base_path: Path, target_path: Path) -> bool:
        """
        Check if target_path is safely within base_path (prevents directory traversal)

        Args:
            base_path: Base directory path
            target_path: Target file/directory path

        Returns:
            True if target_path is safe, False otherwise
        """
        try:
            # Resolve both paths to absolute paths
            base_resolved = base_path.resolve()
            target_resolved = target_path.resolve()

            # Check if target is within base
            return str(target_resolved).startswith(str(base_resolved))

        except Exception as e: # pylint: disable=broad-exception-caught
            logger.error("Error checking path safety: %s", str(e))
            return False
