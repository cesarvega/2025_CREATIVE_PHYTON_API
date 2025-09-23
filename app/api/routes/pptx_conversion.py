"""
API routes for PowerPoint to images conversion.
"""

from fastapi import APIRouter, File, HTTPException, UploadFile, Form

from app.config.settings import settings
from app.models.response_models import PPTXConversionResponse
from app.services.pptx_service import pptx_service
from app.utils.files_utils import FileUtils
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

router = APIRouter(prefix="/pptx", tags=["PPTX Conversion"])


@router.post(
    "/convert/",
    summary="Convert PPTX to images and extract titles",
    response_model=PPTXConversionResponse,
)
async def convert_pptx_to_images(
    pptx_file: UploadFile = File(..., description="PPTX file to convert"),
    displayName: str = Form(..., description="Display name for the project folder"),
    projectType: str = Form(..., description="Project type: 'bipresents' or 'nw'")
):
    """
    Receives a PPTX file, displayName, and projectType, converts each slide to an image, and extracts slide titles.

    - **pptx_file**: PPTX file to process
    - **displayName**: Display name to use as folder name for storing images
    - **projectType**: Project type ('bipresents' or 'nw') to determine base directory

    Returns the paths to the generated images and thumbnails, and the titles of each slide.
    """
    try:
        if not pptx_file.filename:
            raise HTTPException(status_code=400, detail="No filename provided")

        # Verify that the file is a PPTX
        if not pptx_file.filename.lower().endswith(".pptx"):
            raise HTTPException(status_code=400, detail="File must be a PPTX")
        
        # Validate displayName
        if not displayName or not displayName.strip():
            raise HTTPException(status_code=400, detail="Display name is required")
        
        # Validate projectType
        valid_project_types = [settings.PROJECT_TYPE_BIPRESENTS, settings.PROJECT_TYPE_NW]
        if not projectType or projectType not in valid_project_types:
            raise HTTPException(
                status_code=400, 
                detail=f"Project type must be one of: {', '.join(valid_project_types)}"
            )
        
        # Clean displayName to be safe for folder names
        clean_display_name = pptx_service.sanitize_folder_name(displayName.strip())
        
        if not clean_display_name:
            raise HTTPException(status_code=400, detail="Invalid display name")

        # Read file content
        file_content = await pptx_file.read()

        logger.info(
            "Processing PPTX file: %s (%s) with display name: %s, project type: %s",
            pptx_file.filename,
            FileUtils.format_file_size(len(file_content)),
            clean_display_name,
            projectType
        )

        # Convert PPTX to images using displayName as folder name and projectType
        result = pptx_service.convert_pptx_to_images(
            file_content, 
            pptx_file.filename, 
            clean_display_name, 
            projectType
        )

        return PPTXConversionResponse(
            message=result["message"],
            conversion_id=result["conversion_id"],
            project_type=result["project_type"],
            total_images=result["total_images"],
            images=result["images"],
            thumbnails=result.get("thumbnails", []),
            titles=result["titles"],
            pptx_file=result.get("pptx_file", "")
        )

    except ValueError as ve:
        logger.error("Validation error: %s", str(ve))
        raise HTTPException(status_code=400, detail=str(ve)) from ve

    except Exception as e:
        logger.error("Error in PPTX conversion: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Conversion error: {str(e)}"
        ) from e


@router.get("/status/{project_type}/{conversion_id}", summary="Check conversion status")
async def check_conversion_status(project_type: str, conversion_id: str):
    """
    Check the status of a PowerPoint conversion.

    - **project_type**: The project type ('bipresents' or 'nw')
    - **conversion_id**: The ID (folder name) of the conversion to check
    """
    try:
        # Validate project type
        valid_project_types = [settings.PROJECT_TYPE_BIPRESENTS, settings.PROJECT_TYPE_NW]
        if project_type not in valid_project_types:
            raise HTTPException(
                status_code=400, 
                detail=f"Project type must be one of: {', '.join(valid_project_types)}"
            )

        exists = pptx_service.conversion_exists(conversion_id, project_type)

        if exists:
            conversion_folder = pptx_service.get_conversion_folder(conversion_id, project_type)
            thumbnails_folder = conversion_folder / "Thumbnails"
            
            # Get main images (exclude thumbnails folder)
            image_files = [f for f in conversion_folder.iterdir() 
                          if f.is_file() and f.suffix.lower() == '.png']
            
            # Get thumbnail images
            thumbnail_files = []
            if thumbnails_folder.exists():
                thumbnail_files = [f for f in thumbnails_folder.iterdir() 
                                 if f.is_file() and f.suffix.lower() == '.png']

            images = []
            thumbnails = []
            
            # Process main images
            for img_file in sorted(image_files):
                images.append({
                    "filename": img_file.name,
                    "url": f"/files/download/{project_type}/{conversion_id}/{img_file.name}",
                    "size": img_file.stat().st_size,
                })
            
            # Process thumbnails
            for thumb_file in sorted(thumbnail_files):
                thumbnails.append({
                    "filename": thumb_file.name,
                    "url": f"/files/download/{project_type}/{conversion_id}/Thumbnails/{thumb_file.name}",
                    "size": thumb_file.stat().st_size,
                })

            # Get PowerPoint file info
            pptx_files = [f for f in conversion_folder.iterdir() 
                         if f.is_file() and f.suffix.lower() == '.pptx']
            pptx_file_info = None
            if pptx_files:
                pptx_file = pptx_files[0]  # Should be only one
                pptx_file_info = {
                    "filename": pptx_file.name,
                    "url": f"/files/download/{project_type}/{conversion_id}/{pptx_file.name}",
                    "size": pptx_file.stat().st_size,
                }

            return {
                "conversion_id": conversion_id,
                "project_type": project_type,
                "exists": True,
                "total_images": len(images),
                "images": images,
                "thumbnails": thumbnails,
                "pptx_file": pptx_file_info,
                "zip_download_url": f"/files/download-all/{project_type}/{conversion_id}",
            }
        else:
            return {
                "conversion_id": conversion_id, 
                "project_type": project_type,
                "exists": False
            }

    except Exception as e:
        logger.error("Error checking conversion status: %s", str(e))
        raise HTTPException(
            status_code=500, detail=f"Error checking conversion status: {str(e)}"
        ) from e


@router.get("/list/", summary="List all conversions")
async def list_conversions(project_type: str = None):
    """
    List all available PowerPoint conversions.
    
    - **project_type**: Optional filter by project type ('bipresents' or 'nw')
    """
    try:
        # Validate project type if provided
        if project_type:
            valid_project_types = [settings.PROJECT_TYPE_BIPRESENTS, settings.PROJECT_TYPE_NW]
            if project_type not in valid_project_types:
                raise HTTPException(
                    status_code=400, 
                    detail=f"Project type must be one of: {', '.join(valid_project_types)}"
                )

        # Get conversions using the service method
        conversions_data = pptx_service.list_conversions(project_type)
        
        conversions = []
        for conv_data in conversions_data:
            conversion_id = conv_data["conversion_id"]
            ptype = conv_data["project_type"]
            
            conversions.append({
                "conversion_id": conversion_id,
                "project_type": ptype,
                "total_images": conv_data["image_count"],
                "created_at": conv_data["creation_time"],
                "modified_at": conv_data["creation_time"],  # Using creation_time as both
                "status_url": f"/pptx/status/{ptype}/{conversion_id}",
                "zip_download_url": f"/files/download-all/{ptype}/{conversion_id}",
                "folder_path": conv_data["folder_path"]
            })

        # Sort by creation time, newest first
        conversions.sort(key=lambda x: x["created_at"], reverse=True)

        return {
            "total_conversions": len(conversions), 
            "conversions": conversions,
            "filter_applied": project_type
        }

    except Exception as e:
        logger.error("Error listing conversions: %s", str(e))
        raise HTTPException(
            status_code=500, detail=f"Error listing conversions: {str(e)}"
        ) from e


@router.delete("/delete/{project_type}/{conversion_id}", summary="Delete a conversion")
async def delete_conversion(project_type: str, conversion_id: str):
    """
    Delete a PowerPoint conversion and all its files.

    - **project_type**: The project type ('bipresents' or 'nw')
    - **conversion_id**: The ID (folder name) of the conversion to delete
    """
    try:
        # Validate project type
        valid_project_types = [settings.PROJECT_TYPE_BIPRESENTS, settings.PROJECT_TYPE_NW]
        if project_type not in valid_project_types:
            raise HTTPException(
                status_code=400, 
                detail=f"Project type must be one of: {', '.join(valid_project_types)}"
            )

        # Check if conversion exists
        if not pptx_service.conversion_exists(conversion_id, project_type):
            raise HTTPException(
                status_code=404, 
                detail=f"Conversion '{conversion_id}' not found for project type '{project_type}'"
            )

        # Delete the conversion
        deleted = pptx_service.delete_conversion(conversion_id, project_type)
        
        if deleted:
            return {
                "message": f"Conversion '{conversion_id}' deleted successfully",
                "conversion_id": conversion_id,
                "project_type": project_type,
                "deleted": True
            }
        else:
            raise HTTPException(
                status_code=500, 
                detail="Failed to delete conversion"
            )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error deleting conversion: %s", str(e))
        raise HTTPException(
            status_code=500, detail=f"Error deleting conversion: {str(e)}"
        ) from e


@router.post("/create-zip/{project_type}/{conversion_id}", summary="Create ZIP archive of conversion")
async def create_zip_archive(project_type: str, conversion_id: str):
    """
    Create a ZIP archive containing all files from a conversion.

    - **project_type**: The project type ('bipresents' or 'nw')
    - **conversion_id**: The ID (folder name) of the conversion
    """
    try:
        # Validate project type
        valid_project_types = [settings.PROJECT_TYPE_BIPRESENTS, settings.PROJECT_TYPE_NW]
        if project_type not in valid_project_types:
            raise HTTPException(
                status_code=400, 
                detail=f"Project type must be one of: {', '.join(valid_project_types)}"
            )

        # Check if conversion exists
        if not pptx_service.conversion_exists(conversion_id, project_type):
            raise HTTPException(
                status_code=404, 
                detail=f"Conversion '{conversion_id}' not found for project type '{project_type}'"
            )

        # Create ZIP archive
        zip_path = pptx_service.create_zip_archive(conversion_id, project_type)
        
        return {
            "message": f"ZIP archive created successfully for '{conversion_id}'",
            "conversion_id": conversion_id,
            "project_type": project_type,
            "zip_filename": zip_path.name,
            "zip_url": f"/files/download/{project_type}/{conversion_id}/{zip_path.name}",
            "zip_path": str(zip_path)
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error creating ZIP archive: %s", str(e))
        raise HTTPException(
            status_code=500, detail=f"Error creating ZIP archive: {str(e)}"
        ) from e