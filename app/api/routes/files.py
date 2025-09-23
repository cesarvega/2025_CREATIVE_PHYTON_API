"""
Routes for file download and deletion.
"""

import os
import zipfile
from pathlib import Path

from fastapi import APIRouter, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse

from app.config.settings import settings
from app.services.pptx_service import pptx_service
from app.utils.files_utils import FileUtils
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/files", tags=["Files"])
logger = get_logger(__name__)


def get_directory_path(file_type: str) -> Path:
    """
    Gets the directory path based on file type (legacy function for backwards compatibility).

    Args:
        file_type: File type ('documents', 'images', 'powerpoint')

    Returns:
        Path of the corresponding directory

    Raises:
        HTTPException: If the file type is invalid
    """
    directory_map = {
        "documents": settings.output_dir,  # Legacy path
        "images": settings.output_dir,     # Legacy path
        "powerpoint": settings.powerpoint_dir,  # Legacy path
    }

    if file_type not in directory_map:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Invalid file type: {file_type}. "
                f"Valid types: {', '.join(directory_map.keys())}"
            ),
        )

    return directory_map[file_type]


def get_project_directory_path(project_type: str) -> Path:
    """
    Gets the base directory path for a specific project type.

    Args:
        project_type: Project type ('bipresents' or 'nw')

    Returns:
        Path of the corresponding project base directory

    Raises:
        HTTPException: If the project type is invalid
    """
    try:
        return settings.get_base_dir_for_project_type(project_type)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e


@router.get("/download/{project_type}/{conversion_id}/{filename:path}")
async def download_project_file(project_type: str, conversion_id: str, filename: str):
    """
    Downloads a file from a specific project conversion.

    Args:
        project_type: Project type ('bipresents' or 'nw')
        conversion_id: Conversion/project ID (folder name)
        filename: File name (can include subpath like 'Thumbnails/001.png')

    Returns:
        File for download
    """
    try:
        # Get project base directory
        base_directory = get_project_directory_path(project_type)
        
        # Build the full file path
        file_path = base_directory / conversion_id / filename

        # Validate the file doesn't escape the base directory for security
        resolved_path = file_path.resolve()
        base_resolved = base_directory.resolve()

        if not str(resolved_path).startswith(str(base_resolved)):
            raise HTTPException(status_code=400, detail="Invalid file path")

        if not file_path.exists():
            raise HTTPException(
                status_code=404,
                detail=f"File not found: {filename} in project {conversion_id}"
            )

        if not file_path.is_file():
            raise HTTPException(
                status_code=400,
                detail=f"Path is not a file: {filename}",
            )

        # Determine media type based on extension
        media_type = FileUtils.get_media_type(file_path.suffix)

        logger.info("Downloading file: %s from project %s (%s)", filename, conversion_id, project_type)

        return FileResponse(
            path=str(file_path),
            media_type=media_type,
            filename=file_path.name,  # Use just the filename, not the full path
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error downloading file %s from project %s: %s", filename, conversion_id, str(e))
        raise HTTPException(
            status_code=500, detail=f"Internal server error: {str(e)}"
        ) from e


@router.get("/download-all/{project_type}/{conversion_id}")
async def download_project_zip(project_type: str, conversion_id: str):
    """
    Downloads all files from a project as a ZIP archive.

    Args:
        project_type: Project type ('bipresents' or 'nw')
        conversion_id: Conversion/project ID (folder name)

    Returns:
        ZIP file for download
    """
    try:
        # Check if conversion exists
        if not pptx_service.conversion_exists(conversion_id, project_type):
            raise HTTPException(
                status_code=404,
                detail=f"Project '{conversion_id}' not found for type '{project_type}'"
            )

        # Create ZIP archive
        zip_path = pptx_service.create_zip_archive(conversion_id, project_type)
        
        if not zip_path.exists():
            raise HTTPException(
                status_code=500,
                detail="Failed to create ZIP archive"
            )

        logger.info("Downloading ZIP archive: %s for project %s (%s)", zip_path.name, conversion_id, project_type)

        return FileResponse(
            path=str(zip_path),
            media_type="application/zip",
            filename=f"{conversion_id}_slides.zip",
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error creating ZIP for project %s: %s", conversion_id, str(e))
        raise HTTPException(
            status_code=500, detail=f"Internal server error: {str(e)}"
        ) from e


@router.delete("/delete/{project_type}/{conversion_id}")
async def delete_project(project_type: str, conversion_id: str):
    """
    Deletes an entire project and all its files.

    Args:
        project_type: Project type ('bipresents' or 'nw')
        conversion_id: Conversion/project ID (folder name)

    Returns:
        Confirmation message
    """
    try:
        # Check if conversion exists
        if not pptx_service.conversion_exists(conversion_id, project_type):
            raise HTTPException(
                status_code=404,
                detail=f"Project '{conversion_id}' not found for type '{project_type}'"
            )

        # Delete the project
        deleted = pptx_service.delete_conversion(conversion_id, project_type)
        
        if not deleted:
            raise HTTPException(
                status_code=500,
                detail="Failed to delete project"
            )

        logger.info("Project deleted: %s (%s)", conversion_id, project_type)

        return {
            "message": f"Project '{conversion_id}' deleted successfully",
            "project_type": project_type,
            "conversion_id": conversion_id,
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error deleting project %s: %s", conversion_id, str(e))
        raise HTTPException(
            status_code=500, detail=f"Internal server error: {str(e)}"
        ) from e


@router.get("/list/{project_type}")
async def list_projects(project_type: str, limit: int = 100):
    """
    Lists all projects for a specific project type.

    Args:
        project_type: Project type ('bipresents' or 'nw')
        limit: Maximum number of projects to list

    Returns:
        List of projects
    """
    try:
        # Get conversions using the service method
        conversions = pptx_service.list_conversions(project_type)
        
        # Apply limit
        if limit:
            conversions = conversions[:limit]

        # Format the response
        projects = []
        for conv in conversions:
            project_folder = Path(conv["folder_path"])
            
            # Count different file types
            image_files = [f for f in project_folder.iterdir() if f.is_file() and f.suffix.lower() in ['.png', '.jpg', '.jpeg']]
            pptx_files = [f for f in project_folder.iterdir() if f.is_file() and f.suffix.lower() in ['.ppt', '.pptx']]
            
            # Check for thumbnails
            thumbnails_folder = project_folder / "Thumbnails"
            thumbnail_count = 0
            if thumbnails_folder.exists():
                thumbnail_files = [f for f in thumbnails_folder.iterdir() if f.is_file() and f.suffix.lower() in ['.png', '.jpg', '.jpeg']]
                thumbnail_count = len(thumbnail_files)

            projects.append({
                "conversion_id": conv["conversion_id"],
                "project_type": conv["project_type"],
                "total_images": len(image_files),
                "total_thumbnails": thumbnail_count,
                "total_pptx_files": len(pptx_files),
                "created_at": conv["creation_time"],
                "folder_path": conv["folder_path"],
                "download_zip_url": f"/files/download-all/{project_type}/{conv['conversion_id']}",
                "status_url": f"/pptx/status/{project_type}/{conv['conversion_id']}",
            })

        return {
            "projects": projects,
            "total": len(projects),
            "project_type": project_type,
            "limit_applied": limit,
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error listing projects for %s: %s", project_type, str(e))
        raise HTTPException(
            status_code=500, detail=f"Internal server error: {str(e)}"
        ) from e


@router.get("/list/all")
async def list_all_projects(limit: int = 100):
    """
    Lists all projects across all project types.

    Args:
        limit: Maximum number of projects to list

    Returns:
        List of all projects
    """
    try:
        # Get conversions from all project types
        all_conversions = pptx_service.list_conversions()
        
        # Apply limit
        if limit:
            all_conversions = all_conversions[:limit]

        # Format the response
        projects = []
        for conv in all_conversions:
            project_folder = Path(conv["folder_path"])
            
            # Count different file types
            image_files = [f for f in project_folder.iterdir() if f.is_file() and f.suffix.lower() in ['.png', '.jpg', '.jpeg']]
            pptx_files = [f for f in project_folder.iterdir() if f.is_file() and f.suffix.lower() in ['.ppt', '.pptx']]
            
            # Check for thumbnails
            thumbnails_folder = project_folder / "Thumbnails"
            thumbnail_count = 0
            if thumbnails_folder.exists():
                thumbnail_files = [f for f in thumbnails_folder.iterdir() if f.is_file() and f.suffix.lower() in ['.png', '.jpg', '.jpeg']]
                thumbnail_count = len(thumbnail_files)

            projects.append({
                "conversion_id": conv["conversion_id"],
                "project_type": conv["project_type"],
                "total_images": len(image_files),
                "total_thumbnails": thumbnail_count,
                "total_pptx_files": len(pptx_files),
                "created_at": conv["creation_time"],
                "folder_path": conv["folder_path"],
                "download_zip_url": f"/files/download-all/{conv['project_type']}/{conv['conversion_id']}",
                "status_url": f"/pptx/status/{conv['project_type']}/{conv['conversion_id']}",
            })

        return {
            "projects": projects,
            "total": len(projects),
            "limit_applied": limit,
        }

    except Exception as e:
        logger.error("Error listing all projects: %s", str(e))
        raise HTTPException(
            status_code=500, detail=f"Internal server error: {str(e)}"
        ) from e


@router.get("/info/{project_type}/{conversion_id}")
async def get_project_info(project_type: str, conversion_id: str):
    """
    Gets detailed information about a project.

    Args:
        project_type: Project type ('bipresents' or 'nw')
        conversion_id: Conversion/project ID (folder name)

    Returns:
        Project information
    """
    try:
        # Check if conversion exists
        if not pptx_service.conversion_exists(conversion_id, project_type):
            raise HTTPException(
                status_code=404,
                detail=f"Project '{conversion_id}' not found for type '{project_type}'"
            )

        project_folder = pptx_service.get_conversion_folder(conversion_id, project_type)
        thumbnails_folder = project_folder / "Thumbnails"
        
        # Get file information
        files = []
        total_size = 0
        
        # Main folder files
        for file_path in project_folder.iterdir():
            if file_path.is_file():
                stat = file_path.stat()
                total_size += stat.st_size
                files.append({
                    "name": file_path.name,
                    "path": file_path.name,
                    "folder": "root",
                    "size": stat.st_size,
                    "size_formatted": FileUtils.format_file_size(stat.st_size),
                    "created": stat.st_ctime,
                    "modified": stat.st_mtime,
                    "extension": file_path.suffix,
                    "download_url": f"/files/download/{project_type}/{conversion_id}/{file_path.name}"
                })
        
        # Thumbnails folder files
        if thumbnails_folder.exists():
            for file_path in thumbnails_folder.iterdir():
                if file_path.is_file():
                    stat = file_path.stat()
                    total_size += stat.st_size
                    files.append({
                        "name": file_path.name,
                        "path": f"Thumbnails/{file_path.name}",
                        "folder": "Thumbnails",
                        "size": stat.st_size,
                        "size_formatted": FileUtils.format_file_size(stat.st_size),
                        "created": stat.st_ctime,
                        "modified": stat.st_mtime,
                        "extension": file_path.suffix,
                        "download_url": f"/files/download/{project_type}/{conversion_id}/Thumbnails/{file_path.name}"
                    })

        # Project folder stats
        project_stat = project_folder.stat()

        return {
            "conversion_id": conversion_id,
            "project_type": project_type,
            "folder_path": str(project_folder),
            "total_files": len(files),
            "total_size": total_size,
            "total_size_formatted": FileUtils.format_file_size(total_size),
            "created_at": project_stat.st_ctime,
            "modified_at": project_stat.st_mtime,
            "files": files,
            "download_zip_url": f"/files/download-all/{project_type}/{conversion_id}",
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error getting project info for %s: %s", conversion_id, str(e))
        raise HTTPException(
            status_code=500, detail=f"Internal server error: {str(e)}"
        ) from e


# Legacy routes for backwards compatibility
@router.get("/download/{file_type}/{filename:path}")
async def download_file(file_type: str, filename: str):
    """
    Downloads a generated file (legacy route for backwards compatibility).

    Args:
        file_type: File type ('documents', 'images', 'powerpoint')
        filename: File name

    Returns:
        File for download
    """
    try:
        directory = get_directory_path(file_type)
        file_path = directory / filename

        # Validate the file doesn't escape the base directory for security
        resolved_path = file_path.resolve()
        base_directory = directory.resolve()

        if not str(resolved_path).startswith(str(base_directory)):
            raise HTTPException(status_code=400, detail="Invalid file path")

        if not file_path.is_file():
            raise HTTPException(
                status_code=400,
                detail=f"Path is not a file: {filename}",
            )

        # Determine media type based on extension
        media_type = FileUtils.get_media_type(file_path.suffix)

        logger.info("Downloading file (legacy): %s", file_path)

        return FileResponse(
            path=str(file_path),
            media_type=media_type,
            filename=filename,
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error downloading file %s: %s", filename, str(e))
        raise HTTPException(
            status_code=500, detail=f"Internal server error: {str(e)}"
        ) from e


@router.get("/types")
async def get_file_types():
    """
    Gets available file types and their directories.

    Returns:
        Dictionary with file types and their paths
    """
    try:
        file_types = {
            "project_types": {
                "bipresents": {
                    "path": str(settings.base_dir_bipresents),
                    "description": "BSR/NSR project files"
                },
                "nw": {
                    "path": str(settings.base_dir_nw),
                    "description": "NW project files"
                }
            },
            "legacy_types": {
                "documents": {
                    "path": str(settings.output_dir),
                    "description": "Generated document files (legacy)"
                },
                "images": {
                    "path": str(settings.output_dir),
                    "description": "Generated images and graphics (legacy)"
                },
                "powerpoint": {
                    "path": str(settings.powerpoint_dir),
                    "description": "PowerPoint files (legacy)"
                }
            }
        }

        return {
            "file_types": file_types,
            "total_project_types": len(file_types["project_types"]),
            "total_legacy_types": len(file_types["legacy_types"]),
        }

    except Exception as e:
        logger.error("Error getting file types: %s", str(e))
        raise HTTPException(
            status_code=500, detail=f"Internal server error: {str(e)}"
        ) from e