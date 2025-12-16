"""Routes for presentation creation and assembly orchestration."""

import base64
import time
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, BackgroundTasks, Depends, File, HTTPException, UploadFile, Form, Query
from fastapi.responses import FileResponse
from urllib.parse import quote

from app.api.dependencies import (
    parse_build_metadata,
    parse_presentation_metadata,
    parse_simple_dw_metadata,
    parse_bsr_metadata,
    validate_excel_file_with_size,
)
from app.config.settings import settings
from app.models.presentation_models import (
    CreatePresentationMetadata,
    CreatePresentationRequest,
    CreatePresentationResponse,
    SimpleDWMetadata,
    PresentationBuildMetadata,
    PresentationBuildResponse,
    GenerateBackupRequest,
    GenerateBackupResponse,
    BSRCreatePresentationMetadata,
    BSRCreatePresentationResponse,
    ProjectCategoriesUpdate,
    UpdateCategoriesResponse,
)
from app.models.nw_reports_models import (
    CreateFeedbackTemplateRequest,
    CreateFeedbackTemplateResponse,
    DownloadResultsRequest,
    DownloadResultsResponse,
)
from app.models.bsr_reports_models import (
    BSRGenerateReportRequest,
    BSRGenerateReportResponse,
)
from app.services.bsr_report_orchestrator_service import bsr_report_orchestrator_service as bsr_report_orchestrator
from app.models.response_models import ReplaceProjectImagesResponse, CreateSimpleDWResponse
from app.models.task_models import TaskCreatedResponse, TaskStatusResponse, TaskListResponse
from app.services.presentation_service import presentation_service
from app.services.report_orchestrator_service import report_orchestrator_service
from app.services.feedback_template_generator import feedback_template_generator
from app.services.bi_guidelines_service import bi_guidelines_service
from app.services.pptx_service import pptx_service
from app.services.presentation_service import presentation_service
from app.models.presentation_models import CreatePresentationRequest
from app.utils.task_manager import task_manager, TaskStatus
from app.utils.concurrency_manager_v2 import concurrency_manager
from app.api.background_tasks import (
    _create_presentation_background,
    _generate_backup_background,
    _create_bsr_presentation_background,
    _download_results_background,
    _create_feedback_template_background,
    _generate_bsr_report_background,
    _build_presentation_files_background,
)
from app.config.db import get_connection_scope
from app.utils.download_utils import (
    build_api_download_url,
    create_download_token,
    decode_download_token,
    guess_media_type,
)
from pathlib import Path
from app.utils.logging_utils import get_logger
# OPTIMIZATION: Centralized error handling
from app.utils.error_handlers import handle_service_errors, ValidationError, validate_positive_integer
from app.api.dependencies import (
    validate_project_type,
    validate_pptx_file,
    validate_file_size,
)
# Removed old concurrency_utils import - using new concurrency system
from app.utils.concurrency_manager_v2 import concurrency_manager
from app.utils.task_queue import TaskPriority

router = APIRouter(prefix="/presentations", tags=["Presentation Creation"])
logger = get_logger(__name__)


# Helper functions for concurrency control
def check_concurrency_limit():
    """Check if system can accept new tasks.

    Raises:
        HTTPException: If queue is full
    """
    status = concurrency_manager.get_queue_status()
    queue_full = status["queued_count"] >= status["queue_capacity"]

    if queue_full:
        raise HTTPException(
            status_code=503,
            detail={
                "error": "Service temporarily unavailable",
                "message": "Task queue is full. Please try again later.",
                "queue_status": {
                    "active": status["active_count"],
                    "queued": status["queued_count"],
                    "capacity": status["queue_capacity"]
                }
            }
        )


def get_concurrency_status() -> dict:
    """Get current concurrency status.

    Returns:
        Dictionary with queue statistics
    """
    status = concurrency_manager.get_queue_status()
    return {
        "active_tasks": status["active_count"],
        "queued_tasks": status["queued_count"],
        "available_slots": status["available_slots"],
        "max_workers": status["max_workers"],
        "queue_capacity": status["queue_capacity"],
        "statistics": {
            "total_submitted": status["total_submitted"],
            "total_completed": status["total_completed"],
            "total_failed": status["total_failed"]
        }
    }


@router.post(
    "/create",
    response_model=TaskCreatedResponse,
    summary="Create a complete presentation from Excel + PPTX inputs with physical file generation (Background Task)",
    description=(
        "**[BACKGROUND TASK]** Upload Excel + PPTX to create a complete presentation.\n\n"
        "This endpoint starts a background task and returns immediately with a task_id. "
        "The presentation creation happens asynchronously, allowing other API requests to be processed.\n\n"
        "**Workflow:**\n"
        "1. Upload files → Get task_id immediately\n"
        "2. Poll `/api/presentations/tasks/{task_id}` to check status\n"
        "3. When status='completed', retrieve result with presentation_id\n\n"
        "**Processing Steps (Background):**\n"
        "1. **Checks if presentation exists** - Verifies project + display_name combination\n"
        "2. **Handles overwrite logic** - If exists and overwrite_existing=true, deletes old presentation\n"
        "3. **Processes Excel data** - Extracts names, categories, groups, rationales\n"
        "4. **Converts PPTX to images** - Generates JPG from each slide\n"
        "5. **Generates slide metadata** - Creates database records\n"
        "6. **Applies template rotation** - Rotates through specified templates\n"
        "7. **Generates physical PowerPoint** - Combines original + generated slides\n"
        "8. **Persists to database** - Saves to nw_Master and nw_Details\n\n"
        "**Overwrite Behavior:**\n"
        "- Set `overwrite_existing: true` in metadata to allow overwriting existing presentations\n"
        "- If presentation exists and overwrite_existing=false (default), returns 400 error\n"
        "- Frontend should first check with `/api/presentations/exists` and confirm with user\n"
        "- When overwriting, both database records AND physical files are deleted\n\n"
        "**Why Background?** PowerPoint generation is CPU-intensive and can take 30-120 seconds. "
        "Background processing ensures the API remains responsive."
    ),
    response_description=(
        "Task creation confirmation with task_id and status_url. "
        "Use the status_url to poll for completion and get the final result."
    ),
    responses={
        400: {
            "description": "Invalid request, missing files, or presentation already exists without overwrite confirmation.",
            "content": {
                "application/json": {
                    "examples": {
                        "invalid_file": {
                            "summary": "Invalid file type",
                            "value": {"detail": "Excel file must be .xlsx or .xls"}
                        },
                        "already_exists": {
                            "summary": "Presentation already exists (overwrite not confirmed)",
                            "value": {"detail": "Presentation already exists: NW_Project/NW_Display"}
                        }
                    }
                }
            },
        },
        422: {
            "description": "Metadata payload failed validation.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": [
                            {
                                "loc": ["metadata", "project"],
                                "msg": "Field required",
                                "type": "missing",
                            }
                        ]
                    }
                }
            },
        },
        500: {
            "description": "Unexpected error during orchestration.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to create presentation: PowerPoint automation unavailable"
                    }
                }
            },
        },
        503: {
            "description": "Service at capacity - too many concurrent tasks.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": {
                            "error": "Service temporarily unavailable",
                            "message": "El sistema ha alcanzado su capacidad máxima de procesamiento. Por favor, intente nuevamente en unos momentos.",
                            "message_en": "The system has reached its maximum processing capacity. Please try again in a few moments.",
                            "active_tasks": 2,
                            "max_concurrent_tasks": 2,
                            "retry_after_seconds": 30,
                            "action": "retry_later"
                        }
                    }
                }
            },
        },
    },
)
async def create_presentation(
    background_tasks: BackgroundTasks,
    metadata: CreatePresentationMetadata = Depends(parse_presentation_metadata),
    excel_file: UploadFile = File(..., description="Excel file (.xlsx or .xls)"),
    pptx_file: UploadFile = File(..., description="PowerPoint file (.pptx)"),
) -> TaskCreatedResponse:
    """Create a presentation by combining JSON metadata with uploaded template files.

    **Background Task Implementation:**
    This endpoint queues the presentation creation as a background task and returns
    immediately with a task_id. Clients should poll the task status endpoint to
    check progress and retrieve the final result.

    **Workflow:**
    1. Validates file uploads (Excel and PowerPoint)
    2. Creates a background task for processing
    3. Returns task_id and status_url immediately
    4. Processing happens asynchronously (Excel → PPTX → Database → Files)

    **Polling for Results:**
    Use GET /api/presentations/tasks/{task_id} to check status:
    - status='pending': Task queued
    - status='processing': Task running (check progress field)
    - status='completed': Task done (check result field for presentation_id)
    - status='failed': Task failed (check error field)
    """
    try:
        # Check concurrency limit before accepting the task
        check_concurrency_limit()
        # Validate Excel file
        if not excel_file.filename or not excel_file.filename.lower().endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="Excel file must be .xlsx or .xls")

        # Validate PPTX file
        if not pptx_file.filename or not pptx_file.filename.lower().endswith(".pptx"):
            raise HTTPException(status_code=400, detail="PPTX file must be .pptx")

        # Read file contents with size validation
        excel_content = await excel_file.read()
        pptx_content = await pptx_file.read()

        # Validate sizes
        max_excel_size = 10 * 1024 * 1024  # 10MB
        max_pptx_size = 50 * 1024 * 1024   # 50MB

        if len(excel_content) == 0:
            raise HTTPException(status_code=400, detail="Empty Excel file provided")
        if len(excel_content) > max_excel_size:
            raise HTTPException(status_code=413, detail="Excel file too large. Maximum size is 10MB")

        if len(pptx_content) == 0:
            raise HTTPException(status_code=400, detail="Empty PPTX file provided")
        if len(pptx_content) > max_pptx_size:
            raise HTTPException(status_code=413, detail="PPTX file too large. Maximum size is 50MB")

        # Create background task
        task_id = task_manager.create_task(
            task_type="create_presentation",
            description=f"Creating presentation for project: {metadata.project}, display: {metadata.display_name}",
            metadata={
                "project": metadata.project,
                "display_name": metadata.display_name,
                "excel_filename": excel_file.filename,
                "pptx_filename": pptx_file.filename,
            }
        )

        logger.info(
            "Created background task %s for presentation creation: project=%s, display_name=%s",
            task_id,
            metadata.project,
            metadata.display_name,
        )

        # Submit task to concurrency manager (new queue system)
        submission_result = await concurrency_manager.submit_task(
            task_id=task_id,
            task_type="create_presentation",
            func=_create_presentation_background,
            args=(task_id, metadata, excel_content, excel_file.filename, pptx_content, pptx_file.filename),
            priority=TaskPriority.NORMAL,
        )

        # Return task information with queue position
        task_info = task_manager.get_task(task_id)

        return TaskCreatedResponse(
            task_id=task_id,
            status=submission_result["status"],
            message=f"Presentation creation task {'started' if submission_result.get('will_start_immediately') else 'queued'} for project '{metadata.project}'",
            task_type="create_presentation",
            created_at=task_info["created_at"],
            status_url=f"/api/presentations/tasks/{task_id}",
            position=submission_result.get("position"),
            estimated_wait_seconds=submission_result.get("estimated_wait_seconds")
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error queueing presentation creation: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Failed to queue presentation creation: {str(e)}"
        ) from e


@router.get(
    "/concurrency-status",
    summary="Get current concurrency status",
    description=(
        "Get the current concurrency status of the system.\n\n"
        "This endpoint provides information about:\n"
        "- Number of currently active tasks\n"
        "- Maximum concurrent tasks allowed\n"
        "- Available slots for new tasks\n"
        "- Whether the system is at capacity\n\n"
        "Use this endpoint to check if the system can accept new tasks before submitting them."
    ),
)
async def get_concurrency_status_endpoint() -> dict:
    """Get current concurrency status.

    Returns:
        Dictionary with concurrency status information including:
        - active_tasks: Number of currently active tasks
        - max_concurrent_tasks: Maximum allowed concurrent tasks
        - available_slots: Number of available slots
        - at_capacity: Whether system is at capacity
        - utilization_percentage: Current utilization percentage
    """
    return get_concurrency_status()


@router.get(
    "/tasks/{task_id}",
    response_model=TaskStatusResponse,
    summary="Get background task status",
    description=(
        "Check the status of a background task (presentation creation, backup generation, etc.).\n\n"
        "**Status Values:**\n"
        "- `pending`: Task queued but not started\n"
        "- `processing`: Task currently running (check progress field for percentage)\n"
        "- `completed`: Task finished successfully (result field contains output)\n"
        "- `failed`: Task failed (error field contains error message)\n\n"
        "**Polling Recommendations:**\n"
        "- Poll every 2-5 seconds while status is 'pending' or 'processing'\n"
        "- Stop polling when status is 'completed' or 'failed'\n"
        "- Implement exponential backoff for long-running tasks"
    ),
)
async def get_task_status(task_id: str) -> TaskStatusResponse:
    """Get the status of a background task.

    Args:
        task_id: Task identifier returned by a background endpoint

    Returns:
        TaskStatusResponse with current status, progress, result, and error

    Raises:
        HTTPException 404: If task not found
    """
    task = task_manager.get_task(task_id)

    if not task:
        raise HTTPException(
            status_code=404,
            detail=f"Task {task_id} not found"
        )

    # Calculate current position in queue if task is pending
    # This ensures position is always up-to-date even after cancellations
    if task["status"] == "pending":
        queue_position = concurrency_manager.get_task_position(task_id)
        if queue_position is not None:
            # Task is in queue - calculate actual position
            pool_stats = concurrency_manager.worker_pool.get_stats()
            processing_workers = pool_stats["processing_workers"]
            actual_position = processing_workers + queue_position

            # Update task dict with current position info
            task["position"] = actual_position
            task["estimated_wait_seconds"] = actual_position * 90.0

    return TaskStatusResponse(**task)


@router.get(
    "/tasks",
    response_model=TaskListResponse,
    summary="List background tasks",
    description=(
        "Retrieve a list of background tasks with optional filtering.\n\n"
        "**Filters:**\n"
        "- `task_type`: Filter by task type (e.g., 'create_presentation', 'generate_backup')\n"
        "- `status`: Filter by status (pending, processing, completed, failed)\n"
        "- `limit`: Maximum number of tasks to return (default: 100)\n\n"
        "Tasks are sorted by creation time (newest first)."
    ),
)
async def list_tasks(
    task_type: Optional[str] = Query(None, description="Filter by task type"),
    status: Optional[str] = Query(None, description="Filter by status"),
    limit: int = Query(100, ge=1, le=1000, description="Maximum tasks to return"),
) -> TaskListResponse:
    """List background tasks with optional filtering.

    Args:
        task_type: Optional filter by task type
        status: Optional filter by status
        limit: Maximum number of tasks to return

    Returns:
        TaskListResponse with list of tasks and total count
    """
    # Convert status string to TaskStatus enum if provided
    status_filter = None
    if status:
        try:
            status_filter = TaskStatus(status)
        except ValueError:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid status: {status}. Must be one of: pending, processing, completed, failed"
            )

    tasks = task_manager.list_tasks(
        task_type=task_type,
        status=status_filter,
        limit=limit
    )

    return TaskListResponse(
        tasks=[TaskStatusResponse(**task) for task in tasks],
        total=len(tasks)
    )


@router.delete(
    "/tasks/{task_id}",
    summary="Cancel or delete a task",
    description="Cancel a queued/processing task or delete a completed/failed task.",
)
async def delete_task(task_id: str) -> dict:
    """Cancel or delete a task.

    Args:
        task_id: Task identifier

    Returns:
        Success message

    Raises:
        HTTPException 404: If task not found
        HTTPException 409: If task is processing and cannot be cancelled
    """
    # First, try to cancel if it's in queue
    cancelled = await concurrency_manager.cancel_task(task_id)

    if cancelled:
        return {"message": f"Task {task_id} cancelled successfully"}

    # If not cancelled, try to delete (for completed/failed tasks)
    deleted = task_manager.delete_task(task_id)

    if deleted:
        return {"message": f"Task {task_id} deleted successfully"}

    # Task not found anywhere
    raise HTTPException(
        status_code=404,
        detail=f"Task {task_id} not found"
    )


@router.post(
    "/replace-project-images/",
    summary="Replace slide images for an existing project",
    response_model=ReplaceProjectImagesResponse,
)
async def replace_project_images(
    powerpointFile: UploadFile | None = File(
        None, description="PowerPoint file (.pptx) - optional, uses latest from project root if omitted"
    ),
    excelFile: UploadFile | None = File(
        None, description="Excel file (.xlsx/.xls) - optional, uses latest from project root if omitted"
    ),
    page_number: int | None = Form(
        None,
        description="Page number (insert position). Optional; uses existing value if omitted.",
    ),
    project_name: str = Form(..., description="Existing project folder name"),
    project_type: str = Depends(validate_project_type),
):
    """Regenerate project using OpenXML backup flow with existing Excel in project root.

    - Saves the uploaded PPTX into the project root.
    - Looks up existing presentation by display_name.
    - Regenerates backup with OpenXML using the Excel already stored in the project folder.
    - Returns backup path and basic info (no legacy image replacement).
    """
    try:
        # Validate files if provided
        validated_ppt = None
        ppt_bytes = None
        if powerpointFile:
            validated_ppt = await validate_pptx_file(powerpointFile)
            ppt_bytes = await validate_file_size(validated_ppt, 50)

        validated_excel = None
        excel_bytes = None
        if excelFile:
            validated_excel = await validate_excel_file_with_size(excelFile)
            excel_bytes = validated_excel[1] if isinstance(validated_excel, tuple) else await excelFile.read()

        # Basic validation for project_name
        if not project_name or project_name.strip() != project_name or ("/" in project_name or "\\" in project_name):
            raise HTTPException(status_code=400, detail="Invalid project name")

        from app.utils.path_utils import resolve_project_output, sanitize_folder_name

        sanitized_display = sanitize_folder_name(project_name)

        # Locate presentation ID by display_name (original or sanitized) and fetch metadata
        # Try to find in appropriate table based on project_type
        with get_connection_scope(timeout=30) as cursor:
            # Determine which table to query based on project_type
            if project_type.lower() in ["bsr", "bipresents"]:
                table_name = "bsr_Master"
                cursor.execute(
                    """
                    SELECT TOP 1 PresentationId, Project, DisplayName,
                           '' as NameCandidateBGType, '' as NameCandidateBGName,
                           1 as NameCandidateStartingSlide, PresentationType,
                           UploadedBy, DisplayName as BSRDisplayName, 0 as isParticipantsVote,
                           1 as isWideScreenPPT, 0 as isAWSLinkReq, '' as NameCandidateFileName,
                           MainPptFileName
                    FROM [BI_GUIDELINES].[dbo].[bsr_Master]
                    WHERE DisplayName IN (?, ?)
                    ORDER BY PresentationId DESC
                    """,
                    (project_name, sanitized_display),
                )
            else:
                # NW, DW, NSR use nw_Master
                table_name = "nw_Master"
                cursor.execute(
                    """
                    SELECT TOP 1 PresentationId, Project, DisplayName,
                           NameCandidateBGType, NameCandidateBGName,
                           NameCandidateStartingSlide, PresentationType,
                           UploadedBy, BSRDisplayName, isParticipantsVote,
                           isWideScreenPPT, isAWSLinkReq, NameCandidateFileName,
                           MainPptFileName
                    FROM [BI_GUIDELINES].[dbo].[nw_Master]
                    WHERE DisplayName IN (?, ?)
                    ORDER BY PresentationId DESC
                    """,
                    (project_name, sanitized_display),
                )
            row = cursor.fetchone()

        if not row:
            raise HTTPException(
                status_code=404,
                detail=f"Presentation not found for project_name '{project_name}'",
            )

        presentation_id = int(row[0])
        project_db = row[1]
        display_name_db = row[2]
        bg_type_db = row[3] or "Default"
        bg_name_db = row[4] or "Default"
        page_number_db = row[5] or 1
        presentation_type_db = row[6] or "Normal"
        user_name_db = row[7] or ""
        mobile_link_bsr_db = row[8] or ""
        participant_vote_db = row[9] or 0
        is_wide_ppt_db = row[10] or 0
        is_aws_email_db = row[11] or 0
        excel_filename_db = row[12] or "original_data.xlsx"
        main_ppt_filename_db = row[13] or ""

        # Resolve project root
        output_base, _ = resolve_project_output(
            sanitized_display,
            project_type,
            fallback_subdir="generated_presentations",
        )
        output_base.mkdir(parents=True, exist_ok=True)

        # Select PPTX: uploaded -> else latest in root -> else DB
        ppt_path = None
        if ppt_bytes:
            safe_ppt_name = Path(validated_ppt.filename).name
            ppt_path = output_base / safe_ppt_name
            with open(ppt_path, "wb") as f:
                f.write(ppt_bytes)
        else:
            candidates = sorted(output_base.glob("*.pptx"), key=lambda f: f.stat().st_mtime, reverse=True)
            if candidates:
                ppt_path = candidates[0]
            elif main_ppt_filename_db:
                candidate = output_base / Path(main_ppt_filename_db).name
                if candidate.exists():
                    ppt_path = candidate
            if not ppt_path:
                raise HTTPException(status_code=404, detail="No PPTX found or provided")
            with open(ppt_path, "rb") as f:
                ppt_bytes = f.read()

        # Select Excel: uploaded -> else latest in root -> else DB/fallback
        # SPECIAL CASE: BSR, DW, NSR projects don't require Excel, only PPTX for image updates
        logger.info(f"Processing project_type: {project_type}")
        excel_path = None

        # Only require Excel for NW projects
        if project_type.lower() not in ["bsr", "bipresents", "dw", "nsr"]:
            # NW projects REQUIRE Excel
            if excel_bytes:
                safe_excel_name = Path(validated_excel[0].filename if isinstance(validated_excel, tuple) else excelFile.filename).name
                excel_path = output_base / safe_excel_name
                with open(excel_path, "wb") as f:
                    f.write(excel_bytes)
            else:
                candidates = sorted(output_base.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
                if candidates:
                    excel_path = candidates[0]
                else:
                    candidate = output_base / Path(excel_filename_db).name
                    if candidate.exists():
                        excel_path = candidate

            if not excel_path or not excel_path.exists():
                raise HTTPException(status_code=404, detail="No Excel file found or provided in project root (NW projects require Excel)")
            if not excel_bytes:
                with open(excel_path, "rb") as f:
                    excel_bytes = f.read()
        else:
            # BSR, DW, NSR: Excel is optional
            logger.info(f"{project_type.upper()} project - Excel is optional, will use simplified flow")
            if excel_bytes:
                safe_excel_name = Path(validated_excel[0].filename if isinstance(validated_excel, tuple) else excelFile.filename).name
                excel_path = output_base / safe_excel_name
                with open(excel_path, "wb") as f:
                    f.write(excel_bytes)

        # Determine page_number to use
        final_page_number = page_number if page_number is not None else page_number_db

        # ============ RECREATE FROM SCRATCH FLOW FOR BSR, NSR, DW PROJECTS ============
        # For these project types, recreate the entire project from scratch using the PPTX
        # This ensures the database is fully updated like a new project creation
        if project_type.lower() in ["bsr", "bipresents", "dw", "nsr"]:
            logger.info("%s project detected - recreating project from scratch with overwrite", project_type.upper())

            # Recreate project from scratch based on project type
            if project_type.lower() in ["bsr", "bipresents"]:
                # BSR projects: recreate using create_bsr_presentation
                logger.info("Recreating BSR project from scratch: %s / %s", project_db, display_name_db)

                result = presentation_service.create_bsr_presentation(
                    project_name=project_db,
                    display_name=display_name_db,
                    slide_number=final_page_number,
                    presentation_type=presentation_type_db,
                    user_name=user_name_db,
                    is_wide_ppt=is_wide_ppt_db,
                    pptx_content=ppt_bytes,
                    pptx_filename=ppt_path.name,
                    categories=None,  # No categories for replace flow
                    overwrite_existing=True,  # Always overwrite in replace flow
                    progress_callback=None,
                )

                logger.info("BSR project recreated successfully: ID=%d, slides=%d", result["presentation_id"], result["total_slides"])

                return ReplaceProjectImagesResponse(
                    message=f"BSR project recreated successfully from scratch",
                    project_name=project_name,
                    project_type=project_type,
                    total_images=result["total_slides"],
                    images=[],  # Images are in DB but not returned in this flow
                    thumbnails=[],
                    backup_file=None,
                    presentation_id=result["presentation_id"],
                )

            else:
                # NSR and DW projects: recreate without Excel requirement
                logger.info("Recreating %s project from scratch: %s / %s", project_type.upper(), project_db, display_name_db)

                # For DW, page_number is not needed, use 1 as default
                if project_type.lower() == "dw":
                    final_page_number = 1

                # Convert PPTX to images first
                from app.services.pptx_service import pptx_service
                pptx_result = pptx_service.convert_pptx_to_images(
                    file_content=ppt_bytes,
                    filename=ppt_path.name,
                    display_name=sanitized_display,
                    project_type=project_type,
                )

                images = pptx_result.get("images", [])
                logger.info("%s PPTX converted to %d images", project_type.upper(), len(images))

                # Delete existing presentation data to recreate from scratch
                with get_connection_scope(timeout=30) as cursor:
                    # Delete details first (foreign key constraint)
                    cursor.execute(
                        "DELETE FROM [BI_GUIDELINES].[dbo].[nw_Details] WHERE PresentationId = ?",
                        (presentation_id,)
                    )
                    # Delete master record
                    cursor.execute(
                        "DELETE FROM [BI_GUIDELINES].[dbo].[nw_Master] WHERE PresentationId = ?",
                        (presentation_id,)
                    )
                    cursor.connection.commit()
                    logger.info("Deleted existing %s project data for ID=%d", project_type.upper(), presentation_id)

                # Recreate master record using stored procedure
                with get_connection_scope(timeout=30) as cursor:
                    master_sql = """
                        EXEC [dbo].[nw_InsertPresentationMaster_sep2025]
                            @Project=?, @DisplayName=?, @MainPptFileName=?, @NameCandidateFileName=?,
                            @NameCandidateBGType=?, @NameCandidateBGName=?, @NameCandidateStartingSlide=?,
                            @PresentationType=?, @UploadedBy=?, @BSRDisplayName=?,
                            @isParticipantsVote=?, @isWideScreenPPT=?, @isAWSLinkReq=?;
                    """
                    cursor.execute(
                        master_sql,
                        (
                            project_db,                 # @Project
                            display_name_db,            # @DisplayName
                            ppt_path.name,              # @MainPptFileName
                            "",                         # @NameCandidateFileName (no Excel for NSR/DW)
                            bg_type_db,                 # @NameCandidateBGType
                            bg_name_db,                 # @NameCandidateBGName
                            final_page_number,          # @NameCandidateStartingSlide
                            presentation_type_db,       # @PresentationType
                            user_name_db,               # @UploadedBy
                            mobile_link_bsr_db,         # @BSRDisplayName
                            participant_vote_db,        # @isParticipantsVote
                            is_wide_ppt_db,             # @isWideScreenPPT
                            is_aws_email_db,            # @isAWSLinkReq
                        )
                    )

                    # Get the new presentation ID
                    presentation_id_new = None
                    while True:
                        try:
                            row = cursor.fetchone()
                            if row:
                                presentation_id_new = int(row[0])
                                logger.info("%s master record created: ID=%d", project_type.upper(), presentation_id_new)
                                break
                        except Exception:
                            if cursor.nextset():
                                continue
                            break

                    if not presentation_id_new:
                        raise HTTPException(status_code=500, detail="Failed to create master record")

                    cursor.connection.commit()

                # Insert detail records for each slide
                with get_connection_scope(timeout=30) as cursor:
                    detail_sql = """
                        EXEC [dbo].[nw_InsertPresentationDetail_copy]
                            @PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?,
                            @SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?,
                            @NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?,
                            @TemplateId=?, @NameSubGroup=?;
                    """

                    slide_idx = 1
                    is_dw_project = project_type.lower() == "dw"

                    if is_dw_project:
                        # DW: Insert all slides as Image (no NameSummary slide)
                        for pptx_idx in range(len(images)):
                            image_path = images[pptx_idx]
                            cursor.execute(
                                detail_sql,
                                (
                                    presentation_id_new,    # @PresentationId
                                    slide_idx,              # @SlideNumber (DB position)
                                    "Image",                # @SlideType
                                    image_path,             # @SlideBGFileName
                                    f"Slide {pptx_idx + 1}",  # @SlideDescription (PPTX slide number)
                                    "",                     # @NameGroup
                                    "",                     # @NameCategory
                                    "",                     # @Name
                                    "",                     # @NameRationale
                                    "",                     # @NameNotation
                                    "",                     # @KanaNames
                                    "",                     # @NameLogo
                                    0,                      # @TemplateId
                                    "",                     # @NameSubGroup
                                )
                            )
                            slide_idx += 1
                    else:
                        # NSR: Insert NameSummary slide at page_number position
                        # Insert slides BEFORE the summary position
                        for pptx_idx in range(min(final_page_number - 1, len(images))):
                            image_path = images[pptx_idx]
                            cursor.execute(
                                detail_sql,
                                (
                                    presentation_id_new,    # @PresentationId
                                    slide_idx,              # @SlideNumber (DB position)
                                    "Image",                # @SlideType
                                    image_path,             # @SlideBGFileName
                                    f"Slide {pptx_idx + 1}",  # @SlideDescription (PPTX slide number)
                                    "",                     # @NameGroup
                                    "",                     # @NameCategory
                                    "",                     # @Name
                                    "",                     # @NameRationale
                                    "",                     # @NameNotation
                                    "",                     # @KanaNames
                                    "",                     # @NameLogo
                                    0,                      # @TemplateId
                                    "",                     # @NameSubGroup
                                )
                            )
                            slide_idx += 1

                        # Insert SUMMARY slide at the specified page_number position
                        if final_page_number - 1 < len(images):
                            summary_image_path = images[final_page_number - 1]
                            cursor.execute(
                                detail_sql,
                                (
                                    presentation_id_new,    # @PresentationId
                                    slide_idx,              # @SlideNumber (DB position)
                                    "NameSummary",          # @SlideType
                                    summary_image_path,     # @SlideBGFileName
                                    "Brainstorm",           # @SlideDescription
                                    "",                     # @NameGroup
                                    "",                     # @NameCategory
                                    "",                     # @Name
                                    "",                     # @NameRationale
                                    "",                     # @NameNotation
                                    "",                     # @KanaNames
                                    "",                     # @NameLogo
                                    0,                      # @TemplateId
                                    "",                     # @NameSubGroup
                                )
                            )
                            slide_idx += 1

                        # Insert slides AFTER the summary position
                        for pptx_idx in range(final_page_number, len(images)):
                            image_path = images[pptx_idx]
                            cursor.execute(
                                detail_sql,
                                (
                                    presentation_id_new,    # @PresentationId
                                    slide_idx,              # @SlideNumber (DB position)
                                    "Image",                # @SlideType
                                    image_path,             # @SlideBGFileName
                                    f"Slide {pptx_idx + 1}",  # @SlideDescription (PPTX slide number)
                                    "",                     # @NameGroup
                                    "",                     # @NameCategory
                                    "",                     # @Name
                                    "",                     # @NameRationale
                                    "",                     # @NameNotation
                                    "",                     # @KanaNames
                                    "",                     # @NameLogo
                                    0,                      # @TemplateId
                                    "",                     # @NameSubGroup
                                )
                            )
                            slide_idx += 1

                    cursor.connection.commit()

                # Calculate total slides (DW: same as PPTX, NSR: PPTX + 1 NameSummary)
                total_slides_in_db = len(images) if is_dw_project else len(images) + 1
                summary_msg = "no NameSummary" if is_dw_project else "+ NameSummary=1"
                logger.info("%s project recreated successfully: ID=%d, slides=%d (PPTX=%d %s)",
                           project_type.upper(), presentation_id_new, total_slides_in_db, len(images), summary_msg)

                return ReplaceProjectImagesResponse(
                    message=f"{project_type.upper()} project recreated successfully from scratch",
                    project_name=project_name,
                    project_type=project_type,
                    total_images=total_slides_in_db,  # Total slides in DB (PPTX + NameSummary)
                    images=images,
                    thumbnails=pptx_result.get("thumbnails", []),
                    backup_file=None,
                    presentation_id=presentation_id_new,
                )

        # ============ STANDARD FLOW FOR NW PROJECTS (with Excel) ============
        # Update DB with filenames and page_number
        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                """
                UPDATE [BI_GUIDELINES].[dbo].[nw_Master]
                SET MainPptFileName = ?, NameCandidateFileName = ?, NameCandidateStartingSlide = ?
                WHERE PresentationId = ?
                """,
                (ppt_path.name, excel_path.name if excel_path else excel_filename_db, final_page_number, presentation_id),
            )
            cursor.connection.commit()

        # Regenerate backup (uses Excel from root + PPTX)
        backup_info = presentation_service.generate_backup_presentation(presentation_id)
        logger.info(
            "Backup regenerated for PresentationId=%d: %s",
            presentation_id,
            backup_info.get("printable_path"),
        )

        # Recreate presentation using selected Excel/PPTX
        create_result = None
        try:
            create_request = CreatePresentationRequest(
                project=project_db,
                display_name=display_name_db,
                background_type=bg_type_db,
                background_name=bg_name_db,
                page_number=final_page_number,
                presentation_type=presentation_type_db,
                user_name=user_name_db,
                mobile_link_bsr=mobile_link_bsr_db,
                participant_vote=participant_vote_db,
                is_wide_ppt=is_wide_ppt_db,
                is_aws_email=is_aws_email_db,
                excel_file=excel_bytes,
                excel_filename=excel_path.name,
                pptx_file=ppt_bytes,
                pptx_filename=ppt_path.name,
                create_backup=1,
                has_groups=True,
                test_name_order="Default",
                project_type=project_type,
                overwrite_existing=True,
            )

            create_result = presentation_service.create_presentation(create_request)
            logger.info(
                "Presentation recreated for project %s (ID from service: %s)",
                project_name,
                getattr(create_result, "presentation_id", None),
            )
        except Exception as create_exc:
            logger.warning(
                "Recreation using existing Excel/PPTX failed for project %s: %s",
                project_name,
                str(create_exc),
            )

        return ReplaceProjectImagesResponse(
            message="Project regenerated using OpenXML with existing/new files",
            project_name=project_name,
            project_type=project_type,
            total_images=0,
            images=[],
            thumbnails=[],
            backup_file=backup_info.get("printable_path"),
            presentation_id=getattr(create_result, "presentation_id", None) if create_result else presentation_id,
        )

    except FileNotFoundError as fnf:
        raise HTTPException(status_code=404, detail=str(fnf)) from fnf
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve)) from ve
    except Exception as e:
        logger.error("Error replacing project images: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500, detail=f"Replace images error: {str(e)}"
        ) from e


@router.post(
    "/create-simple-dw",
    summary="Create DW slide images from a PPTX using simplified metadata",
    response_model=CreateSimpleDWResponse,
)
async def create_simple_dw_presentation(
    metadata: SimpleDWMetadata = Depends(parse_simple_dw_metadata),
    powerpointFile: UploadFile = File(..., description="PowerPoint file (.pptx)"),
):
    """Simplified create endpoint for DW that only requires a PPTX.

    This endpoint reuses the image export logic from the presentations routes,
    saving all slide images into the existing DW project's folder:
    C:\\inetpub\\wwwroot\\nw2\\nw_slides\\{projectName}

    Notes:
    - projectName must match an existing folder under nw2/nw_slides
    - displayName and widePresentation are accepted for compatibility/logging
    - Files are exported as 001.jpg, 002.jpg, ... and Thumbnails/...
    """
    try:
        validated_file = await validate_pptx_file(powerpointFile)
        file_content = await validate_file_size(validated_file, 50)

        projectName = metadata.projectName
        displayName = metadata.displayName
        userName = metadata.userName
        slideType = (metadata.slideType or "Image").strip() or "Image"
        presentationType = (metadata.presentationType or "Design").strip() or "Design"

        # Basic validation for projectName (disallow path separators and trimmed changes)
        if (
            not projectName
            or projectName.strip() != projectName
            or ("/" in projectName or "\\" in projectName)
        ):
            raise HTTPException(status_code=400, detail="Invalid project name")

        logger.info(
            "DW create-simple request: projectName='%s', displayName='%s', user='%s', overwrite='%s'",
            projectName,
            displayName,
            userName,
            metadata.overwrite_existing,
        )

        # Check if presentation exists and handle overwrite logic
        existing_id = presentation_service._check_presentation_exists(projectName, displayName)
        was_overwritten = False
        overwrite_presentation_id = None

        if existing_id and not metadata.overwrite_existing:
            # Presentation exists and overwrite not confirmed
            logger.warning(
                "DW presentation already exists: %s/%s (ID: %d) - overwrite not confirmed",
                projectName,
                displayName,
                existing_id
            )
            raise HTTPException(
                status_code=400,
                detail=f"Presentation already exists: {projectName}/{displayName}"
            )

        if existing_id and metadata.overwrite_existing:
            # Presentation exists and overwrite confirmed
            logger.info(
                "Overwriting DW presentation: %s/%s (ID: %d) by user: %s",
                projectName,
                displayName,
                existing_id,
                userName
            )
            # Clean details and files but KEEP the master record (preserves ID)
            presentation_service._clean_presentation_for_overwrite(existing_id, settings.PROJECT_TYPE_DW)
            was_overwritten = True
            overwrite_presentation_id = existing_id

        # Generate images under displayName folder (nw_slides/[displayName]/001.jpg)
        result = pptx_service.convert_pptx_to_images(
            file_content,
            validated_file.filename,
            metadata.displayName,
            settings.PROJECT_TYPE_DW,
        )

        # After generating images, record slide details with SlideType='Design'
        # into nw_InsertPresentationDetail_copy for the matching presentation
        presentation_id: Optional[int] = None
        try:
            with get_connection_scope(timeout=30) as cursor:
                # If overwriting, use the preserved ID, otherwise check if master exists
                if overwrite_presentation_id:
                    presentation_id = overwrite_presentation_id
                    logger.info("Using preserved presentation ID for overwrite: %d", presentation_id)

                    # Update master record with new values
                    # Note: Only updating essential fields that are guaranteed to exist in nw_Master
                    update_sql = """
                        UPDATE [BI_GUIDELINES].[dbo].[nw_Master]
                        SET
                            MainPptFileName = ?,
                            PresentationType = ?,
                            UploadedBy = ?,
                            isWideScreenPPT = ?
                        WHERE PresentationId = ?
                    """
                    cursor.execute(
                        update_sql,
                        (
                            validated_file.filename or "",
                            presentationType,
                            userName or "",
                            1 if getattr(metadata, "widePresentation", False) else 0,
                            presentation_id,
                        )
                    )
                    logger.info("✅ DW master record updated: ID=%d", presentation_id)
                else:
                    # Resolve PresentationId from nw_Master via Project + DisplayName
                    cursor.execute(
                        (
                            "SELECT TOP 1 PresentationId "
                            "FROM [BI_GUIDELINES].[dbo].[nw_Master] "
                            "WHERE Project = ? AND DisplayName = ? "
                            "ORDER BY PresentationId DESC"
                        ),
                        (projectName, displayName),
                    )
                    row = cursor.fetchone()
                    if not row:
                        # If no master exists, create a minimal master record (Design/DW)
                        master_sql = (
                            "EXEC [dbo].[nw_InsertPresentationMaster_sep2025] "
                            "@Project=?, @DisplayName=?, @MainPptFileName=?, @NameCandidateFileName=?, "
                            "@NameCandidateBGType=?, @NameCandidateBGName=?, @NameCandidateStartingSlide=?, "
                            "@PresentationType=?, @UploadedBy=?, @BSRDisplayName=?, "
                            "@isParticipantsVote=?, @isWideScreenPPT=?, @isAWSLinkReq=?;"
                        )
                        master_params = (
                            projectName,
                            displayName,
                            validated_file.filename or "",
                            "",  # NameCandidateFileName not used in DW simple flow
                            "Default",  # BG type
                            "",  # BG name
                            1,  # Starting slide
                            presentationType,  # PresentationType from frontend (e.g., 'Design')
                            userName or "",
                            "",  # BSRDisplayName
                            0,  # isParticipantsVote
                            1 if getattr(metadata, "widePresentation", False) else 0,  # isWideScreenPPT
                            0,  # isAWSLinkReq
                        )

                        cursor.execute(master_sql, master_params)

                        # Try to fetch PresentationId from result set(s)
                        while True:
                            try:
                                created_row = cursor.fetchone()
                                if created_row:
                                    presentation_id = int(created_row[0])
                                    break
                            except Exception:
                                # Some drivers/SPs return no row; proceed to nextset
                                pass
                            if not cursor.nextset():
                                break

                        # Fallback lookup if SP did not return ID
                        if not presentation_id:
                            cursor.execute(
                                (
                                    "SELECT TOP 1 PresentationId "
                                    "FROM [BI_GUIDELINES].[dbo].[nw_Master] "
                                    "WHERE Project = ? AND DisplayName = ? "
                                    "ORDER BY PresentationId DESC"
                                ),
                                (projectName, displayName),
                            )
                            row2 = cursor.fetchone()
                            if not row2:
                                raise HTTPException(
                                    status_code=404,
                                    detail=(
                                        "Presentation master not found and could not be created"
                                    ),
                                )
                            presentation_id = int(row2[0])
                    else:
                        presentation_id = int(row[0])

                # Insert a detail row for each generated image with the requested SlideType
                insert_sql = (
                    "EXEC [dbo].[nw_InsertPresentationDetail_copy] "
                    "@PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?, "
                    "@SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?, "
                    "@NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?, "
                    "@TemplateId=?, @NameSubGroup=?;"
                )

                total = int(result.get("total_images", 0))
                for idx in range(1, total + 1):
                    # Use relative image path as background filename reference
                    # and store userName in SlideDescription for traceability
                    rel_path = result.get("images", [])[idx - 1] if result.get("images") else ""
                    cursor.execute(
                        insert_sql,
                        (
                            presentation_id,  # @PresentationId
                            idx,              # @SlideNumber
                            slideType,        # @SlideType
                            rel_path,         # @SlideBGFileName
                            "",              # @SlideDescription (left empty per DW spec)
                            "", "", "", "", "", "", "",  # group/category/name fields
                            0,                # @TemplateId
                            "",              # @NameSubGroup
                        ),
                    )
        except HTTPException:
            raise
        except Exception as db_exc:
            logger.error("Failed to insert DW detail records: %s", str(db_exc), exc_info=True)
            raise HTTPException(
                status_code=500,
                detail=f"Failed to record DW slide details: {str(db_exc)}",
            ) from db_exc

        return CreateSimpleDWResponse(
            message=result["message"] if not was_overwritten else "DW presentation overwritten successfully",
            project_name=result.get("project_name", projectName),
            display_name=displayName,
            project_type=result.get("project_type", settings.PROJECT_TYPE_DW),
            slide_type=slideType,
            user_name=userName,
            total_images=result["total_images"],
            images=result["images"],
            thumbnails=result.get("thumbnails", []),
            presentation_id=presentation_id,
            overwritten=was_overwritten,
        )

    except FileNotFoundError as fnf:
        raise HTTPException(status_code=404, detail=str(fnf)) from fnf
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve)) from ve
    except Exception as e:
        logger.error(
            "Error creating DW images from PPTX: %s", str(e), exc_info=True
        )
        raise HTTPException(
            status_code=500, detail=f"DW create-simple error: {str(e)}"
        ) from e


@router.post(
    "/create-bsr",
    summary="Create BSR presentation from PowerPoint file",
    response_model=TaskCreatedResponse,
    description=(
        "Create a new BSR (Board Sales Request) presentation by uploading a PowerPoint file.\n\n"
        "This endpoint:\n"
        "1. Checks if presentation already exists\n"
        "2. If exists and overwrite_existing=false (default), returns 400 error\n"
        "3. If exists and overwrite_existing=true, deletes old presentation and creates new one\n"
        "4. Validates display name hasn't been used by other projects (skipped if overwriting)\n"
        "5. Converts PowerPoint slides to images\n"
        "6. Reads slide titles from the PowerPoint\n"
        "7. Creates presentation record in database\n"
        "8. Inserts slide details including summary slide at specified position\n"
        "9. **(Optional)** Processes and inserts additional categories\n\n"
        "**Overwrite Behavior:**\n"
        "- Set `overwrite_existing: true` in metadata to allow overwriting existing presentations\n"
        "- When overwriting, the old presentation is deleted (including all details and categories)\n"
        "- Frontend should first check with `/api/presentations/bsr/exists` and confirm with user\n"
        "- Audit trail: deletion is logged with username for accountability\n\n"
        "**Slide Insertion Logic:**\n"
        "If the PowerPoint has 5 slides and slide_number=3:\n"
        "- Slide 1 (Image) -> Slide #1\n"
        "- Slide 2 (Image) -> Slide #2\n"
        "- Summary (NameSummary) -> Slide #3 (inserted at slide_number)\n"
        "- Slide 3 (Image) -> Slide #4\n"
        "- Slide 4 (Image) -> Slide #5\n"
        "- Slide 5 (Image) -> Slide #6\n"
        "Total: 6 slides (5 original + 1 summary)\n\n"
        "**Categories (Optional):**\n"
        "You can optionally add categories to the BSR project by including a `categories` object in metadata:\n"
        "- `add_categories`: Set to `true` to enable category processing\n"
        "- `mode`: Either `'single'` (1 category) or `'both'` (2 categories)\n"
        "- `category1`: Always required when `add_categories=true` (name + elements array)\n"
        "- `category2`: Required only when `mode='both'` (name + elements array)\n\n"
        "Example categories payload:\n"
        "```json\n"
        '{\n'
        '  "add_categories": true,\n'
        '  "mode": "both",\n'
        '  "category1": {\n'
        '    "name": "Region",\n'
        '    "elements": ["North America", "Europe", "Asia"]\n'
        '  },\n'
        '  "category2": {\n'
        '    "name": "Channel",\n'
        '    "elements": ["Retail", "Online", "Corporate"]\n'
        '  }\n'
        '}\n'
        "```\n\n"
        "**Data Storage:**\n"
        "- Master record: bsr_InsertPresentationMaster\n"
        "- Detail records: bsr_InsertPresentationDetail (one per slide)\n"
        "- Categories: BSR_CATEGORY and BSR_CATEGORY_ELEMENTS tables\n"
        "- Images stored in: BRS_slides/{ProjectName}/\n"
    ),
    responses={
        200: {
            "description": "BSR presentation created successfully",
            "content": {
                "application/json": {
                    "examples": {
                        "created": {
                            "summary": "New presentation created",
                            "value": {
                                "message": "BSR Presentation created successfully",
                                "presentation_id": 12345,
                                "total_slides": 6,
                                "processing_time_seconds": 8.45,
                                "categories_added": 2,
                                "overwritten": False
                            }
                        },
                        "overwritten": {
                            "summary": "Existing presentation overwritten",
                            "value": {
                                "message": "BSR Presentation overwritten successfully",
                                "presentation_id": 12346,
                                "total_slides": 6,
                                "processing_time_seconds": 9.12,
                                "categories_added": 2,
                                "overwritten": True
                            }
                        }
                    }
                }
            },
        },
        400: {
            "description": "Invalid request or presentation already exists without overwrite confirmation",
            "content": {
                "application/json": {
                    "examples": {
                        "already_exists": {
                            "summary": "Presentation already exists (overwrite not confirmed)",
                            "value": {
                                "detail": "BSR presentation already exists: BSR_Project/BSR_Presentation"
                            }
                        },
                        "display_name_used": {
                            "summary": "Display name already used by another project",
                            "value": {
                                "detail": "Display name already used by another project: OtherProject"
                            }
                        }
                    }
                }
            },
        },
        422: {
            "description": "Metadata validation failed",
            "content": {
                "application/json": {
                    "example": {
                        "detail": [
                            {
                                "loc": ["metadata", "slide_number"],
                                "msg": "Field required",
                                "type": "missing",
                            }
                        ]
                    }
                }
            },
        },
        500: {
            "description": "Unexpected error during creation",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to create BSR presentation: Database connection error"
                    }
                }
            },
        },
        503: {
            "description": "Service at capacity - too many concurrent tasks.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": {
                            "error": "Service temporarily unavailable",
                            "message": "El sistema ha alcanzado su capacidad máxima de procesamiento. Por favor, intente nuevamente en unos momentos.",
                            "active_tasks": 2,
                            "max_concurrent_tasks": 2,
                            "retry_after_seconds": 30
                        }
                    }
                }
            },
        },
    },
)
async def create_bsr_presentation(
    background_tasks: BackgroundTasks,
    metadata: BSRCreatePresentationMetadata = Depends(parse_bsr_metadata),
    pptx_file: UploadFile = File(..., description="PowerPoint file (.pptx)"),
) -> TaskCreatedResponse:
    """Create a BSR presentation from a PowerPoint file.

    This endpoint handles the complete BSR presentation creation workflow:
    1. Validation checks (existence, display name uniqueness)
    2. PowerPoint conversion to images
    3. Slide title extraction
    4. Database persistence

    Args:
        metadata: BSR presentation metadata (project_name, display_name, etc.)
        pptx_file: PowerPoint file to process

    Returns:
        BSRCreatePresentationResponse with presentation ID and statistics

    Raises:
        HTTPException 400: If validation fails or presentation already exists
        HTTPException 500: If creation process fails
    """
    try:
        # Check concurrency limit before accepting the task
        check_concurrency_limit()

        # Validate file
        if not pptx_file.filename or not pptx_file.filename.lower().endswith(".pptx"):
            raise HTTPException(status_code=400, detail="PPTX file must be .pptx")

        # Read file content
        pptx_content = await pptx_file.read()

        # Validate size
        max_pptx_size = 50 * 1024 * 1024
        if len(pptx_content) == 0:
            raise HTTPException(status_code=400, detail="Empty PPTX file provided")
        if len(pptx_content) > max_pptx_size:
            raise HTTPException(status_code=413, detail="PPTX file too large. Maximum size is 50MB")

        # Create background task
        task_id = task_manager.create_task(
            task_type="create_bsr_presentation",
            description=f"Creating BSR presentation: {metadata.project_name}/{metadata.display_name}",
            metadata={
                "project_name": metadata.project_name,
                "display_name": metadata.display_name,
                "pptx_filename": pptx_file.filename,
            }
        )

        logger.info(
            "Created background task %s for BSR presentation: project=%s, display=%s",
            task_id,
            metadata.project_name,
            metadata.display_name,
        )

        # Submit task to concurrency manager
        submission_result = await concurrency_manager.submit_task(
            task_id=task_id,
            task_type="create_bsr_presentation",
            func=_create_bsr_presentation_background,
            args=(task_id, metadata, pptx_content, pptx_file.filename),
            priority=TaskPriority.NORMAL,
        )

        # Return task information with queue position
        task_info = task_manager.get_task(task_id)

        return TaskCreatedResponse(
            task_id=task_id,
            status=submission_result["status"],
            message=f"BSR presentation creation task {'started' if submission_result.get('will_start_immediately') else 'queued'} for '{metadata.project_name}'",
            task_type="create_bsr_presentation",
            created_at=task_info["created_at"],
            status_url=f"/api/presentations/tasks/{task_id}",
            position=submission_result.get("position"),
            estimated_wait_seconds=submission_result.get("estimated_wait_seconds")
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error queueing BSR presentation creation: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to queue BSR presentation creation: {str(e)}"
        ) from e

@router.post(
    "/createTemplate",
    response_model=TaskCreatedResponse,
    summary="Assemble presentation files directly from an Excel candidate sheet",
    description=(
        "Generate PPTX artifacts using metadata and an Excel workbook without uploading a template file. "
        "The metadata defines template packs, slide ranges, and output behaviour. Optionally returns base64 "
        "content or download URLs for the generated files."
    ),
    response_description="Generated artifact metadata including slide range and file locations.",
    responses={
        400: {
            "description": "Invalid Excel file or missing templates on disk.",
            "content": {
                "application/json": {
                    "example": {"detail": "Excel file must be .xlsx or .xls"}
                }
            },
        },
        422: {
            "description": "Metadata payload failed validation (slide ranges, template pack, etc.).",
            "content": {
                "application/json": {
                    "example": {
                        "detail": [
                            {
                                "loc": ["metadata", "slide_end"],
                                "msg": "slide_end must be greater than or equal to slide_start",
                                "type": "value_error",
                            }
                        ]
                    }
                }
            },
        },
        500: {
            "description": "Unexpected error while composing the presentation files.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to assemble presentation from provided inputs."
                    }
                }
            },
        },
        503: {
            "description": "Service at capacity - too many concurrent tasks.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": {
                            "error": "Service temporarily unavailable",
                            "message": "El sistema ha alcanzado su capacidad máxima de procesamiento. Por favor, intente nuevamente en unos momentos.",
                            "active_tasks": 2,
                            "max_concurrent_tasks": 2,
                            "retry_after_seconds": 30
                        }
                    }
                }
            },
        },
    },
)
async def build_presentation_files(
    background_tasks: BackgroundTasks,
    metadata: PresentationBuildMetadata = Depends(parse_build_metadata),
    excel_file: UploadFile = File(..., description="Excel file (.xlsx or .xls)"),
) -> TaskCreatedResponse:
    try:
        # Check concurrency limit before accepting the task
        check_concurrency_limit()
        # Validate Excel file
        if not excel_file.filename or not excel_file.filename.lower().endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="Excel file must be .xlsx or .xls")

        # Read file content with size validation
        excel_content = await excel_file.read()

        # Validate size
        max_excel_size = 10 * 1024 * 1024  # 10MB
        if len(excel_content) == 0:
            raise HTTPException(status_code=400, detail="Empty Excel file provided")
        if len(excel_content) > max_excel_size:
            raise HTTPException(status_code=413, detail="Excel file too large. Maximum size is 10MB")

        # Create background task
        task_id = task_manager.create_task(
            task_type="build_presentation_files",
            description=f"Building presentation for: {metadata.project}/{metadata.display_name}",
            metadata={
                "project": metadata.project,
                "display_name": metadata.display_name,
                "excel_filename": excel_file.filename,
            }
        )

        logger.info(
            "Created background task %s for presentation build: project=%s, display_name=%s",
            task_id,
            metadata.project,
            metadata.display_name,
        )

        # Submit task to concurrency manager
        submission_result = await concurrency_manager.submit_task(
            task_id=task_id,
            task_type="build_presentation_files",
            func=_build_presentation_files_background,
            args=(task_id, metadata.model_dump(), excel_content, excel_file.filename),
            priority=TaskPriority.NORMAL,
        )

        # Return task information with queue position
        task_info = task_manager.get_task(task_id)

        return TaskCreatedResponse(
            task_id=task_id,
            status=submission_result["status"],
            message=f"Presentation build task {'started' if submission_result.get('will_start_immediately') else 'queued'} for '{metadata.project}'",
            task_type="build_presentation_files",
            created_at=task_info["created_at"],
            status_url=f"/api/presentations/tasks/{task_id}",
            position=submission_result.get("position"),
            estimated_wait_seconds=submission_result.get("estimated_wait_seconds")
        )

    except HTTPException:
        raise
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        logger.error("Error queueing presentation build: %s", exc, exc_info=True)
        raise HTTPException(
            status_code=500,
            detail="Failed to queue presentation build.",
        ) from exc


@router.get(
    "/exists",
    summary="Check if a presentation exists",
    response_description="Returns a boolean indicating if the presentation exists.",
    responses={
        200: {
            "description": "Check completed successfully.",
            "content": {
                "application/json": {
                    "example": {"exists": True}
                }
            },
        },
        500: {
            "description": "Database error during check.",
            "content": {
                "application/json": {
                    "example": {"detail": "Database error while checking for presentation."}
                }
            },
        },
    },
)
async def check_presentation_exists(
    project_name: str,
    display_name: str,
    exclude_id: Optional[int] = None,
) -> dict[str, bool]:
    """
    Check if a presentation with the given project name and display name already exists.
    An optional `exclude_id` can be provided to exclude a specific presentation ID
    from the check, which is useful for update operations.
    """
    try:
        logger.info(
            "Checking if presentation exists: project_name=%s, display_name=%s, exclude_id=%s",
            project_name,
            display_name,
            exclude_id
        )

        # Query database to check if presentation exists
        exists = presentation_service.presentation_exists(
            project_name=project_name,
            display_name=display_name,
            exclude_id=exclude_id
        )

        logger.info("Presentation exists check result: %s", exists)

        return {"exists": exists}

    except HTTPException:
        raise
    except Exception as e:
        logger.error(
            "Error checking presentation exists: %s",
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Database error while checking for presentation: {str(e)}"
        ) from e


@router.get(
    "/exists",
    summary="Check if a presentation exists (NW/NSR/DW)",
    response_description="Returns a boolean indicating if the presentation exists and its ID if found.",
    responses={
        200: {
            "description": "Check completed successfully.",
            "content": {
                "application/json": {
                    "examples": {
                        "exists": {
                            "summary": "Presentation exists",
                            "value": {"exists": True, "presentation_id": 12345}
                        },
                        "not_exists": {
                            "summary": "Presentation does not exist",
                            "value": {"exists": False, "presentation_id": None}
                        }
                    }
                }
            },
        },
        500: {
            "description": "Database error during check.",
            "content": {
                "application/json": {
                    "example": {"detail": "Database error while checking for presentation."}
                }
            },
        },
    },
)
async def check_presentation_exists(
    project_name: str = Query(..., description="Project name to check"),
    display_name: str = Query(..., description="Display name to check"),
) -> dict:
    """
    Check if a presentation with the given project name and display name already exists.

    This endpoint is for NW/NSR/DW presentations and returns both existence status
    and the presentation ID if found, which can be useful for the frontend to implement
    overwrite confirmation workflows.

    Args:
        project_name: The project name to check
        display_name: The display name to check

    Returns:
        dict: Contains 'exists' (bool) and 'presentation_id' (int or None)
    """
    try:
        logger.info(
            "Checking if presentation exists: project_name=%s, display_name=%s",
            project_name,
            display_name
        )

        # Use the service method to check existence
        presentation_id = presentation_service._check_presentation_exists(
            project_name=project_name,
            display_name=display_name
        )

        exists = presentation_id is not None

        logger.info(
            "Presentation exists check result: exists=%s, id=%s",
            exists,
            presentation_id
        )

        return {
            "exists": exists,
            "presentation_id": presentation_id
        }

    except Exception as e:
        logger.error(
            "Error checking presentation exists: %s",
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Database error while checking for presentation: {str(e)}"
        ) from e


@router.get(
    "/bsr/exists",
    summary="Check if a BSR presentation exists",
    response_description="Returns a boolean indicating if the BSR presentation exists and its ID if found.",
    responses={
        200: {
            "description": "Check completed successfully.",
            "content": {
                "application/json": {
                    "examples": {
                        "exists": {
                            "summary": "Presentation exists",
                            "value": {"exists": True, "presentation_id": 12345}
                        },
                        "not_exists": {
                            "summary": "Presentation does not exist",
                            "value": {"exists": False, "presentation_id": None}
                        }
                    }
                }
            },
        },
        500: {
            "description": "Database error during check.",
            "content": {
                "application/json": {
                    "example": {"detail": "Database error while checking for BSR presentation."}
                }
            },
        },
    },
)
async def check_bsr_presentation_exists(
    project_name: str = Query(..., description="BSR project name to check"),
    display_name: str = Query(..., description="BSR display name to check"),
) -> dict:
    """
    Check if a BSR presentation with the given project name and display name already exists.
    
    This endpoint is specifically for BSR presentations and returns both existence status
    and the presentation ID if found, which can be useful for the frontend to implement
    overwrite confirmation workflows.
    
    Args:
        project_name: The BSR project name to check
        display_name: The BSR display name to check
        
    Returns:
        dict: Contains 'exists' (bool) and 'presentation_id' (int or None)
    """
    try:
        logger.info(
            "Checking if BSR presentation exists: project_name=%s, display_name=%s",
            project_name,
            display_name
        )

        # Use the service method to check existence
        presentation_id = presentation_service._check_bsr_presentation_exists(
            project_name=project_name,
            display_name=display_name
        )
        
        exists = presentation_id is not None

        logger.info(
            "BSR presentation exists check result: exists=%s, id=%s",
            exists,
            presentation_id
        )

        return {
            "exists": exists,
            "presentation_id": presentation_id
        }

    except Exception as e:
        logger.error(
            "Error checking BSR presentation exists: %s",
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Database error while checking for BSR presentation: {str(e)}"
        ) from e


@router.get(
    "/files/{token}",
    summary="Download a generated presentation artifact",
    responses={
        200: {
            "description": "Binary contents of the requested presentation artifact.",
            "content": {"application/octet-stream": {"schema": {"type": "string", "format": "binary"}}},
        },
        400: {"description": "Invalid download token."},
        403: {"description": "Token does not reference an authorized file."},
        404: {"description": "Requested file not found."},
    },
)
async def download_generated_file(token: str) -> FileResponse:
    """Download a generated presentation file using a secure token.
    
    Args:
        token: The secure download token encoding the file path.
        
    Returns:
        FileResponse with the requested file.
    """
    file_path = decode_download_token(token)
    media_type = guess_media_type(file_path)

    # Build response and force attachment filename to avoid browsers/proxies
    # inferring a name based on the URL token or changing extensions.
    response = FileResponse(
        path=file_path,
        media_type=media_type,
        filename=file_path.name,  # best-effort for Starlette versions that support it
    )
    # Force Content-Disposition with RFC 5987 UTF-8 filename
    utf8_name = quote(file_path.name)
    response.headers["Content-Disposition"] = (
        f"attachment; filename=\"{file_path.name}\"; filename*=UTF-8''{utf8_name}"
    )
    # Prevent type sniffing that may alter extensions
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Content-Transfer-Encoding"] = "binary"
    return response


@router.post(
    "/test-slide-generation",
    summary="Test slide generation with multi-candidate layout",
    description=(
        "Test endpoint to verify slide generation logic in isolation. "
        "Allows testing the multi-candidate name layout without going through the full pipeline."
    ),
)
async def test_slide_generation(
    template_name: str = "template_default_withgroups2019.pptx",
    names: str = "John Doe##Jane Smith##Bob Johnson##Alice Williams##Charlie Brown##David Miller##Emma Davis##Frank Wilson##Grace Martinez##Henry Anderson##Ivy Thomas##Jack Taylor##Kelly Moore##Liam Jackson##Mia White##Noah Harris##Olivia Martin##Peter Thompson##Quinn Garcia##Rachel Robinson##Steve Clark##Tina Rodriguez##Uma Lewis##Victor Lee##Wendy Walker##Xavier Hall##Yara Allen##Zack Young",
) -> dict:
    """Test slide generation with specified template and names.

    Args:
        template_name: Name of the template file to use (must exist in templates directory)
        names: Delimited string of names to layout (use ## as delimiter)

    Returns:
        Result of the slide generation test including success status and file path
    """
    try:
        from app.services.pptx_builder_service import PPTXBuilderService, tokenize_delimited_block
        from app.utils.path_utils import resolve_project_output
        import pythoncom

        # Initialize COM for this thread
        pythoncom.CoInitialize()

        try:
            # Parse names
            names_list = tokenize_delimited_block(names)
            if not names_list:
                raise HTTPException(status_code=400, detail="No names provided")

            logger.info("Testing slide generation with %d names", len(names_list))

            # Initialize builder service
            templates_dir = Path(settings.templates_dir) if hasattr(settings, 'templates_dir') else Path("templates")
            subdir = "BackgroundDefaultTemplate"
            templates_dir = templates_dir / subdir
            template_path = templates_dir / template_name
            
            print(template_path)

            if not template_path.exists():
                raise HTTPException(
                    status_code=404,
                    detail=f"Template not found: {template_name}. Available templates are in {templates_dir}"
                )

            builder = PPTXBuilderService(
                templates_dir=templates_dir,
                base_template_path=template_path,
            )

            # Create output directory
            output_dir, _ = resolve_project_output("test_slide_generation", None)
            output_dir.mkdir(parents=True, exist_ok=True)

            # Generate a test presentation with just one slide
            from pptx import Presentation
            import win32com.client

            # Load template
            prs = Presentation(str(template_path))

            # Use first slide or add one
            if prs.slides:
                slide = prs.slides[0]
            else:
                slide = prs.slides.add_slide(prs.slide_layouts[5])

            # Save the presentation first
            test_output_path = output_dir / f"test_slide_{int(time.time())}.pptx"
            prs.save(str(test_output_path))

            # Now open with COM automation to test the layout logic
            from app.utils.com_manager import com_manager
            with com_manager.acquire("PowerPoint.Application Dispatch - Test"):
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                powerpoint.Visible = 1

            presentation = powerpoint.Presentations.Open(str(test_output_path.absolute()))
            com_slide = presentation.Slides(1)

            # Collect placeholder shapes (use Name Candidate for template_default_withgroups2019.pptx)
            placeholder_shapes = builder._collect_placeholder_shapes(com_slide, "Name Candidate")

            # Debug: List all text shapes in the slide
            logger.info("=== DIAGNOSTIC: Analyzing slide shapes ===")
            all_text_shapes = list(builder._iter_text_shapes(com_slide))
            logger.info("Total text shapes found: %d", len(all_text_shapes))

            for idx, shape in enumerate(all_text_shapes[:20]):  # Limit to first 20
                try:
                    text = shape.TextFrame.TextRange.Text
                    logger.info("  Shape %d: '%s'", idx + 1, text[:50] if text else "(empty)")
                except Exception as e:
                    logger.info("  Shape %d: (no text - %s)", idx + 1, str(e))

            if not placeholder_shapes:
                logger.warning("No 'Name Candidate' placeholder shapes found in template")
                warnings_list = ["No 'Name Candidate' placeholder shapes found in template. Check logs for all text shapes found."]
            else:
                logger.info("Found %d 'Name Candidate' placeholder shapes", len(placeholder_shapes))
                warnings_list = []

                # Log the placeholder shapes found
                for idx, pshape in enumerate(placeholder_shapes):
                    try:
                        ptext = pshape.TextFrame.TextRange.Text
                        logger.info("  Placeholder %d: '%s'", idx + 1, ptext[:30])
                    except:
                        pass

            # Test the simplified placeholder assignment
            builder._assign_names_to_placeholders(
                placeholder_shapes=placeholder_shapes or [],
                names=names_list,
                context="test slide",
            )
            success = True

            # Save and close properly
            try:
                presentation.Save()
            except Exception as e:
                logger.warning("Could not save presentation: %s", e)

            try:
                presentation.Close()
            except Exception as e:
                logger.warning("Could not close presentation: %s", e)

            try:
                powerpoint.Quit()
            except Exception as e:
                logger.warning("Could not quit PowerPoint: %s", e)

            logger.info("Test slide generation completed: success=%s", success)

            return {
                "success": success,
                "names_count": len(names_list),
                "names": names_list[:10],  # First 10 names for brevity
                "placeholder_count": len(placeholder_shapes) if placeholder_shapes else 0,
                "output_path": str(test_output_path),
                "warnings": warnings_list,
                "message": "Slide generation test completed. Check the output file to verify the layout."
            }

        finally:
            # Uninitialize COM
            pythoncom.CoUninitialize()

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error in test slide generation: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to test slide generation: {str(e)}"
        ) from e


@router.post(
    "/test-table-layout",
    summary="Test table layout with simple list of names",
    description="Simple endpoint to test dynamic table layout with just a list of names",
)
@handle_service_errors  # OPTIMIZATION: Centralized error handling
async def test_table_layout(names: list[str]) -> dict:
    """Test table layout with a simple list of names.

    Args:
        names: List of candidate names

    Returns:
        Result of the table generation including success status and file path
    """
    try:
        from app.services.pptx_builder_service import PPTXBuilderService
        from app.utils.path_utils import resolve_project_output
        import pythoncom

        # Initialize COM for this thread
        pythoncom.CoInitialize()

        try:
            if not names:
                raise HTTPException(status_code=400, detail="No names provided")

            logger.info("Testing table layout with %d names", len(names))

            # Initialize builder service
            templates_dir = Path(settings.templates_dir) if hasattr(settings, 'templates_dir') else Path("templates")
            subdir = "BackgroundDefaultTemplate"
            templates_dir = templates_dir / subdir
            template_name = "template_default_withgroups2019.pptx"
            template_path = templates_dir / template_name

            if not template_path.exists():
                raise HTTPException(
                    status_code=404,
                    detail=f"Template not found: {template_name}"
                )

            builder = PPTXBuilderService(
                templates_dir=templates_dir,
                base_template_path=template_path,
            )

            # Create output directory
            output_dir, _ = resolve_project_output("test_table_layout", None)
            output_dir.mkdir(parents=True, exist_ok=True)

            # Generate test presentation
            from pptx import Presentation
            import win32com.client

            # Load template
            prs = Presentation(str(template_path))

            # Use first slide or add one
            if prs.slides:
                slide = prs.slides[0]
            else:
                slide = prs.slides.add_slide(prs.slide_layouts[5])

            # Save the presentation first
            test_output_path = output_dir / f"table_test_{int(time.time())}.pptx"
            prs.save(str(test_output_path))

            # Open with COM automation
            from app.utils.com_manager import com_manager
            with com_manager.acquire("PowerPoint.Application Dispatch - Test"):
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                powerpoint.Visible = 1

            presentation = powerpoint.Presentations.Open(str(test_output_path.absolute()))
            com_slide = presentation.Slides(1)

            # Collect placeholder shapes
            placeholder_shapes = builder._collect_placeholder_shapes(com_slide, "Name Candidate")

            logger.info("Found %d placeholder shapes", len(placeholder_shapes))

            # Test the simplified placeholder assignment
            cleaned_names = [n for n in names if n]
            builder._assign_names_to_placeholders(
                placeholder_shapes=placeholder_shapes or [],
                names=cleaned_names,
                context="test table",
            )
            success = True

            # Save and close properly
            try:
                presentation.Save()
            except Exception as e:
                logger.warning("Could not save: %s", e)

            try:
                presentation.Close()
            except Exception as e:
                logger.warning("Could not close: %s", e)

            try:
                powerpoint.Quit()
            except Exception as e:
                logger.warning("Could not quit: %s", e)

            logger.info("Table layout test completed: success=%s", success)

            return {
                "success": success,
                "names_count": len(names),
                "output_path": str(test_output_path),
                "message": "Table layout test completed. Check the output file."
            }

        finally:
            pythoncom.CoUninitialize()

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error in test table layout: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to test table layout: {str(e)}"
        ) from e


@router.post(
    "/download-results",
    response_model=TaskCreatedResponse,
    summary="Download NW results (generates Excel and Word in a ZIP)",
    description=(
        "Always generates two files for the specified presentation and packages them in a ZIP file:\n\n"
        "- **Excel**: Results workbook (retained, new names, explore/avoid, notes) and, when applicable, votes and participants sheets (internal logic).\n"
        "- **Word**: NW report based on a template, with tables and chart.\n\n"
        "**Output**: ZIP file that contains both reports (Excel and Word).\n\n"
        "Minimum input: only `presentation_id`. The endpoint resolves which sheets to include and returns a download token for the ZIP file."
    ),
    response_description=(
        "Generation status with path, filename, and a download token for the ZIP. "
        "Use `/presentations/files/{token}` to download it. The ZIP contains both files (Excel and Word)."
    ),
    responses={
        200: {
            "description": "Reports generated in a ZIP file (Excel and Word).",
            "content": {
                "application/json": {
                    "example": {
                        "success": True,
                        "message": "Generated 2 reports (Excel and Word) in ZIP archive",
                        "file_path": "C:/.../NW_Files/downloads/TestPresentation_Reports_20250122_143022.zip",
                        "file_name": "TestPresentation_Reports_20250122_143022.zip",
                        "download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                        "excel_download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                        "word_download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                        "excel_file_name": "TestPresentation_20250122_143022.xlsx",
                        "word_file_name": "TestPresentation_20250122_143022.doc",
                        "report_type": "excel",
                        "presentation_id": 12345,
                        "generated_at": "2025-01-22T14:30:22.123456",
                        "warnings": []
                    }
                }
            },
        },
        400: {
            "description": "Invalid request parameters.",
            "content": {
                "application/json": {
                    "example": {"detail": "Presentation with ID 12345 not found"}
                }
            },
        },
        404: {
            "description": "Required template file not found.",
            "content": {
                "application/json": {
                    "example": {"detail": "Word template not found at: C:/Templates/Nomenclature Workshop Report_Template_new_xxx.doc"}
                }
            },
        },
        500: {
            "description": "Unexpected error during report generation.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to generate report: Database connection error"
                    }
                }
            },
        },
        503: {
            "description": "Service at capacity - too many concurrent tasks.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": {
                            "error": "Service temporarily unavailable",
                            "message": "El sistema ha alcanzado su capacidad máxima de procesamiento. Por favor, intente nuevamente en unos momentos.",
                            "active_tasks": 2,
                            "max_concurrent_tasks": 2,
                            "retry_after_seconds": 30
                        }
                    }
                }
            },
        },
    },
)
@handle_service_errors  # OPTIMIZATION: Centralized error handling
async def download_results(background_tasks: BackgroundTasks, request: DownloadResultsRequest) -> TaskCreatedResponse:
    """Generate and download NW presentation results.

    This endpoint orchestrates the generation of various NW report types:

    **Excel Report (report_type: "excel"):**
    - Generates a complete workbook with multiple sheets
    - Includes retained names, newly created names, roots/concepts, and notes
    - Optionally includes voting data and participant information
    - Data is retrieved from stored procedures:
      - nw_dlRetainedNames_withRecraft
      - nw_CombineNewNames
      - nw_dlRootsOrConceptsToExplore
      - nw_dlRootsOrConceptsToAvoid
      - nw_dlOpenNotes
      - nw_Votesbygroups (if include_votes=true)
      - nw_VotedParticipants (if include_participants=true)

    **Analytics Report (report_type: "analytics"):**
    - Generates from NWAnalytics.xlsx template
    - Fills Project-Specific and Region-Specific sheets
    - Data from stored procedures:
      - NW_ProjectAnalytics
      - NW_RegionSpecificAnalytics

    **Word Report (report_type: "word"):**
    - Generates a Word document from a template
    - Replaces placeholders (cover, headers, footers, text boxes)
    - Tablas pobladas por tipo (Positive, Neutral, Reconsider, New Names, Explore)
    - Inserta grÃ¡fico de torta (By The Numbers)
    - Procedimientos utilizados:
      - nw_wdValuesToReplace (placeholders)
      - nw_wdGetResults_Phonetics / nw_CombineNewNames

    **BSR Report (report_type: "bsr"):** [Not Yet Implemented]
    - Would generate BSR-specific Excel and Word documents
    - Would use stored procedures:
      - bsr_GetExcelReport
      - Plus direct queries to bsr_ProjectConcepts and BSR_PageComments tables

    Args:
        request: Download results request containing:
            - presentation_id: The presentation to generate report for
            - report_type: Type of report (excel, word, analytics, bsr)
            - summary_type: Optional, for Word phonetics reports
            - mobile_link: Optional, for BSR downloads
            - include_votes: Whether to include votes sheet
            - include_participants: Whether to include participants sheet

    Returns:
        DownloadResultsResponse with:
            - success: Boolean indicating if generation succeeded
            - message: Descriptive message
            - file_path: Full path to generated file
            - file_name: Name of generated file
            - download_token: Secure token for downloading the file
            - report_type: Type of report generated
            - presentation_id: ID of presentation
            - generated_at: Timestamp of generation
            - warnings: List of any warnings during generation

    Raises:
        HTTPException 400: If presentation ID is invalid or not found
        HTTPException 404: If required template files are missing
        HTTPException 500: If report generation fails
        HTTPException 501: If report type is not yet implemented
    """
    try:
        # Check concurrency limit before accepting the task
        check_concurrency_limit()

        # Create background task
        task_id = task_manager.create_task(
            task_type="download_results",
            description=f"Generating reports for presentation_id: {request.presentation_id}",
            metadata={"presentation_id": request.presentation_id}
        )

        logger.info(
            "Created background task %s for report generation: presentation_id=%d",
            task_id,
            request.presentation_id,
        )

        # Submit task to concurrency manager
        submission_result = await concurrency_manager.submit_task(
            task_id=task_id,
            task_type="download_results",
            func=_download_results_background,
            args=(task_id, request.presentation_id),
            priority=TaskPriority.NORMAL,
        )

        # Return task information with queue position
        task_info = task_manager.get_task(task_id)

        return TaskCreatedResponse(
            task_id=task_id,
            status=submission_result["status"],
            message=f"Report generation task {'started' if submission_result.get('will_start_immediately') else 'queued'} for presentation {request.presentation_id}",
            task_type="download_results",
            created_at=task_info["created_at"],
            status_url=f"/api/presentations/tasks/{task_id}",
            position=submission_result.get("position"),
            estimated_wait_seconds=submission_result.get("estimated_wait_seconds")
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error queueing report generation: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to queue report generation: {str(e)}"
        ) from e

@router.post(
    "/create-feedback-template",
    response_model=TaskCreatedResponse,
    summary="Create NW Feedback Template document (InputDocumentRationales)",
    description=(
        "Generate a complete Feedback Template document from the InputDocumentRationales template.\n\n"
        "This endpoint replicates the 'Create Feedback Template' button functionality from the original app. "
        "It generates a fully populated Word document with:\n\n"
        "**Document Generation Process:**\n"
        "1. **Template Selection** - Automatically selects the correct InputDocumentRationales template:\n"
        "   - `InputDocumentRationales.doc` for Normal presentations\n"
        "   - `InputDocumentRationales_Phonetics.doc` for Phonetics presentations\n"
        "   - `InputDocumentRationales_Katakana.doc` for Katakana presentations\n\n"
        "2. **Placeholder Replacement** - Replaces all placeholders throughout the document:\n"
        "   - `<Client>`, `<ProjectName>`, `<DateNormalCase>`, etc.\n"
        "   - Searches in body, headers, footers, text boxes, and shapes\n"
        "   - Data from `nw_wdValuesToReplace` stored procedure\n\n"
        "3. **Table Population** - Fills all feedback tables with project data:\n"
        "   - **Positive Names** table (from `Positive` or `Positive_Phonetics`)\n"
        "   - **Neutral Names** table (from `Neutral` or `Negative_Phonetics`)\n"
        "   - **Reconsider Names** table (from `Reconsider` or `Reconsider_Phonetics`)\n"
        "   - **Newly Created Names** table (from `nw_CombineNewNames`)\n"
        "   - **Roots/Concepts to Explore** table (from `Explore`)\n"
        "   - Uses Open Sans 12pt font, rationales in 10pt\n"
        "   - Handles HTML content in rationale fields\n"
        "   - Processes name groups (## or $$ delimiters)\n\n"
        "4. **Pie Chart Generation** - Creates and inserts a pie chart:\n"
        "   - Uses data from `ByTheNumbers` summary type\n"
        "   - Shows Positive/Neutral/Reconsider distribution\n"
        "   - Corporate colors: Green (92,184,92), Brown (183,122,51), Purple (153,0,76)\n"
        "   - Generated in Excel and pasted into PieChart bookmark\n\n"
        "5. **Cleanup** - Removes all bookmarks from the final document\n\n"
        "**Data Sources:**\n"
        "- `nw_wdValuesToReplace` - Placeholder values (client, project, date, etc.)\n"
        "- `nw_wdGetResults_Phonetics` - Name results by category (Positive, Neutral, Reconsider, Explore)\n"
        "- `nw_CombineNewNames` - Newly created names with categories and rationales\n\n"
        "**Output Format:**\n"
        "- Word 97-2003 Document (.doc)\n"
        "- Saved to NW_Downloads directory\n"
        "- Filename: `Feedback_Template_{DisplayName}_{timestamp}.doc`\n\n"
        "The generated document is ready for client delivery and contains all project feedback data."
    ),
    response_description=(
        "Generation status with file path, filename, and secure download token. "
        "Use the download token with the /presentations/files/{token} endpoint to retrieve the file."
    ),
    responses={
        200: {
            "description": "Feedback template generated successfully.",
            "content": {
                "application/json": {
                    "example": {
                        "success": True,
                        "message": "Feedback template generated successfully",
                        "file_path": "C:/output/NW_Downloads/Feedback_Template_MyProject_20250117_103045.doc",
                        "file_name": "Feedback_Template_MyProject_20250117_103045.doc",
                        "download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                        "presentation_id": 12345,
                        "presentation_type": "Phonetics",
                        "generated_at": "2025-01-17T10:30:45.123456",
                        "warnings": []
                    }
                }
            },
        },
        400: {
            "description": "Invalid presentation ID or presentation not found.",
            "content": {
                "application/json": {
                    "example": {"detail": "Presentation with ID 12345 not found"}
                }
            },
        },
        404: {
            "description": "Required template file not found.",
            "content": {
                "application/json": {
                    "example": {"detail": "Feedback template not found: C:/Templates/InputDocumentRationales_Phonetics.doc"}
                }
            },
        },
        500: {
            "description": "Unexpected error during document generation.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to generate feedback template: Word COM automation error"
                    }
                }
            },
        },
        503: {
            "description": "Service at capacity - too many concurrent tasks.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": {
                            "error": "Service temporarily unavailable",
                            "message": "El sistema ha alcanzado su capacidad máxima de procesamiento. Por favor, intente nuevamente en unos momentos.",
                            "active_tasks": 2,
                            "max_concurrent_tasks": 2,
                            "retry_after_seconds": 30
                        }
                    }
                }
            },
        },
    },
)
async def create_feedback_template(background_tasks: BackgroundTasks, request: CreateFeedbackTemplateRequest) -> TaskCreatedResponse:
    """Generate a Feedback Template document for an NW presentation.

    This endpoint creates a complete InputDocumentRationales document populated with all
    feedback data for a specific presentation. The document includes:

    - Client and project metadata in headers/footers
    - Positive, Neutral, and Reconsider name tables
    - Newly created names table
    - Roots/concepts to explore table
    - Pie chart showing name distribution

    The template used is automatically selected based on the presentation type
    (Normal, Phonetics, or Katakana).

    Args:
        request: CreateFeedbackTemplateRequest with presentation_id

    Returns:
        CreateFeedbackTemplateResponse with:
            - success: Boolean indicating if generation succeeded
            - message: Descriptive message
            - file_path: Full path to generated Word file
            - file_name: Name of generated file
            - download_token: Secure token for downloading the file
            - presentation_id: ID of presentation
            - presentation_type: Type of presentation (Normal, Phonetics, Katakana)
            - generated_at: Timestamp of generation
            - warnings: List of any warnings during generation

    Raises:
        HTTPException 400: If presentation ID is invalid or not found
        HTTPException 404: If required template files are missing
        HTTPException 500: If document generation fails
    """
    try:
        # Check concurrency limit before accepting the task
        check_concurrency_limit()

        # Get presentation info first
        presentation_info = bi_guidelines_service.get_project_info(request.presentation_id)

        if not presentation_info:
            raise HTTPException(
                status_code=404,
                detail=f"Presentation with ID {request.presentation_id} not found"
            )

        display_name = presentation_info.get("DisplayName", "Unknown")

        # Create background task
        task_id = task_manager.create_task(
            task_type="create_feedback_template",
            description=f"Creating feedback template for: {display_name}",
            metadata={
                "presentation_id": request.presentation_id,
                "display_name": display_name
            }
        )

        logger.info(
            "Created background task %s for feedback template: presentation_id=%d",
            task_id,
            request.presentation_id,
        )

        # Submit task to concurrency manager
        submission_result = await concurrency_manager.submit_task(
            task_id=task_id,
            task_type="create_feedback_template",
            func=_create_feedback_template_background,
            args=(task_id, request.presentation_id, display_name),
            priority=TaskPriority.NORMAL,
        )

        # Return task information with queue position
        task_info = task_manager.get_task(task_id)

        return TaskCreatedResponse(
            task_id=task_id,
            status=submission_result["status"],
            message=f"Feedback template task {'started' if submission_result.get('will_start_immediately') else 'queued'} for presentation {request.presentation_id}",
            task_type="create_feedback_template",
            created_at=task_info["created_at"],
            status_url=f"/api/presentations/tasks/{task_id}",
            position=submission_result.get("position"),
            estimated_wait_seconds=submission_result.get("estimated_wait_seconds")
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error queueing feedback template: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to queue feedback template: {str(e)}"
        ) from e

@router.post(
    "/backup",
    response_model=TaskCreatedResponse,
    summary="Generate backup PowerPoint from saved files",
    description=(
        "Generate a complete PowerPoint backup presentation from previously saved Excel and PPTX files.\n\n"
        "This endpoint allows users to generate the backup PowerPoint file at a later time, "
        "even if they didn't select the 'create_backup' option during the initial presentation creation.\n\n"
        "**How it works:**\n"
        "1. Retrieves presentation metadata from the database using the presentation_id\n"
        "2. Loads the saved original Excel file (saved during initial creation)\n"
        "3. Loads the saved original PowerPoint template file\n"
        "4. Processes the Excel data and generates slides\n"
        "5. Creates a complete PowerPoint backup combining:\n"
        "   - Original PPTX slides (before page_number)\n"
        "   - Dynamically generated slides from Excel data\n"
        "   - Original PPTX slides (after generated content)\n\n"
        "**Requirements:**\n"
        "- The presentation must exist in the database\n"
        "- Original Excel file must have been saved (automatic since this update)\n"
        "- Original PowerPoint template must exist\n\n"
        "**Returns:** Complete PowerPoint backup file with download token"
    ),
    response_description=(
        "Generation status with file path, filename, total slides, and secure download token. "
        "Use the download token with the /presentations/files/{token} endpoint to retrieve the file."
    ),
    responses={
        200: {
            "description": "Backup generation response (success or missing files)",
            "content": {
                "application/json": {
                    "examples": {
                        "success": {
                            "summary": "Backup generated successfully",
                            "value": {
                                "success": True,
                                "message": "Backup presentation generated successfully",
                                "presentation_id": 12345,
                                "printable_path": "C:/inetpub/wwwroot/nw2/nw_slides/TestProject/Presentations/backup_20250123_143022.pptx",
                                "file_name": "backup_20250123_143022.pptx",
                                "download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                                "total_slides": "24",
                                "warnings": "None",
                                "missing_files": []
                            }
                        },
                        "missing_files": {
                            "summary": "Required files missing",
                            "value": {
                                "success": False,
                                "message": "Cannot generate backup: 2 required file(s) missing. Please contact IT to upload the missing files.",
                                "presentation_id": 12345,
                                "printable_path": None,
                                "file_name": None,
                                "download_token": None,
                                "total_slides": None,
                                "warnings": "Missing 2 required file(s)",
                                "missing_files": [
                                    {
                                        "file_type": "Excel",
                                        "expected_path": "C:/inetpub/wwwroot/nw2/nw_slides/TestProject/original_data.xlsx",
                                        "instructions": "Please upload the original Excel file to: C:/inetpub/wwwroot/nw2/nw_slides/TestProject/original_data.xlsx"
                                    },
                                    {
                                        "file_type": "PowerPoint",
                                        "expected_path": "C:/inetpub/wwwroot/nw2/nw_slides/TestProject/template.pptx",
                                        "instructions": "Please upload the original PowerPoint template to: C:/inetpub/wwwroot/nw2/nw_slides/TestProject/template.pptx"
                                    }
                                ]
                            }
                        }
                    }
                }
            },
        },
        404: {
            "description": "Presentation not found in database",
            "content": {
                "application/json": {
                    "example": {"detail": "Presentation with ID 12345 not found"}
                }
            },
        },
        500: {
            "description": "Unexpected error during backup generation",
            "content": {
                "application/json": {
                    "example": {"detail": "Failed to generate backup: PowerPoint automation error"}
                }
            },
        },
        503: {
            "description": "Service at capacity - too many concurrent tasks.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": {
                            "error": "Service temporarily unavailable",
                            "message": "El sistema ha alcanzado su capacidad máxima de procesamiento. Por favor, intente nuevamente en unos momentos.",
                            "active_tasks": 2,
                            "max_concurrent_tasks": 2,
                            "retry_after_seconds": 30
                        }
                    }
                }
            },
        },
    },
)
async def generate_backup(background_tasks: BackgroundTasks, request: GenerateBackupRequest) -> TaskCreatedResponse:
    """Generate a backup PowerPoint presentation from saved files.

    This endpoint enables deferred backup generation. Users who didn't create
    a backup during initial presentation creation can generate it later using
    this endpoint.

    Args:
        request: GenerateBackupRequest with presentation_id

    Returns:
        GenerateBackupResponse with:
            - success: Boolean indicating if generation succeeded
            - message: Descriptive message
            - presentation_id: ID of the presentation
            - printable_path: Full path to generated .pptx file
            - file_name: Name of the generated file
            - download_token: Secure token for downloading the file
            - total_slides: Total number of slides in the presentation
            - warnings: Any warnings encountered during generation

    Raises:
        HTTPException 404: If presentation or required files not found
        HTTPException 500: If backup generation fails
    """
    try:
        # Check concurrency limit before accepting the task
        check_concurrency_limit()

        # Create background task
        task_id = task_manager.create_task(
            task_type="generate_backup",
            description=f"Generating backup for presentation_id: {request.presentation_id}",
            metadata={"presentation_id": request.presentation_id}
        )

        logger.info(
            "Created background task %s for backup generation: presentation_id=%d",
            task_id,
            request.presentation_id,
        )

        # Submit task to concurrency manager
        submission_result = await concurrency_manager.submit_task(
            task_id=task_id,
            task_type="generate_backup",
            func=_generate_backup_background,
            args=(task_id, request.presentation_id),
            priority=TaskPriority.NORMAL,
        )

        # Return task information with queue position
        task_info = task_manager.get_task(task_id)

        return TaskCreatedResponse(
            task_id=task_id,
            status=submission_result["status"],
            message=f"Backup generation task {'started' if submission_result.get('will_start_immediately') else 'queued'} for presentation {request.presentation_id}",
            task_type="generate_backup",
            created_at=task_info["created_at"],
            status_url=f"/api/presentations/tasks/{task_id}",
            position=submission_result.get("position"),
            estimated_wait_seconds=submission_result.get("estimated_wait_seconds")
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error queueing backup generation: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to queue backup generation: {str(e)}"
        ) from e


@router.get(
    "/download-presentation/{presentation_id}",
    summary="Download generated PowerPoint presentation file",
    description=(
        "Download the physical PowerPoint file (.pptx) generated for a specific presentation.\n\n"
        "This endpoint retrieves the PowerPoint file that was created when `create_backup=1` "
        "was specified during presentation creation.\n\n"
        "**Requirements:**\n"
        "- The presentation must exist in the database\n"
        "- A PowerPoint file must have been generated (create_backup=1)\n"
        "- The file must exist on the server\n\n"
        "**Returns:** Direct file download of the .pptx file"
    ),
    responses={
        200: {
            "description": "PowerPoint file download",
            "content": {"application/vnd.openxmlformats-officedocument.presentationml.presentation": {}},
        },
        404: {"description": "Presentation not found or file doesn't exist"},
        500: {"description": "Server error"},
    },
)
async def download_presentation_file(presentation_id: int) -> FileResponse:
    """Download the generated PowerPoint file for a presentation.

    Args:
        presentation_id: The ID of the presentation to download

    Returns:
        FileResponse with the PowerPoint file

    Raises:
        HTTPException: 404 if presentation or file not found, 500 on server error
    """
    try:
        logger.info("Downloading presentation file for ID: %d", presentation_id)

        # Get presentation info from database
        with get_connection_scope(timeout=30) as cursor:
            cursor.execute(
                """
                SELECT
                    DisplayName,
                    MainPptFileName,
                    Project
                FROM [BI_GUIDELINES].[dbo].[nw_Master]
                WHERE PresentationId = ?
                """,
                (presentation_id,)
            )

            row = cursor.fetchone()

            if not row:
                logger.warning("Presentation ID %d not found", presentation_id)
                raise HTTPException(
                    status_code=404,
                    detail=f"Presentation with ID {presentation_id} not found"
                )

            display_name = row.DisplayName
            main_ppt_filename = row.MainPptFileName
            project = row.Project

            if not main_ppt_filename:
                logger.warning(
                    "Presentation ID %d has no PowerPoint file (create_backup was not enabled)",
                    presentation_id
                )
                raise HTTPException(
                    status_code=404,
                    detail=f"No PowerPoint file available for presentation ID {presentation_id}. "
                           f"The file was not generated (create_backup=0)."
                )

        # Construct file path using resolve_project_output (same as backup generation)
        from pathlib import Path as PathLib
        from app.utils.path_utils import resolve_project_output

        # Clean display_name if it contains path separators
        clean_display_name = PathLib(display_name).name if "/" in display_name or "\\" in display_name else display_name

        # Clean main_ppt_filename if it contains path separators (extract just the filename)
        # Some old records may have full paths like "files/download/NW/Project/file.pptx"
        clean_ppt_filename = PathLib(main_ppt_filename).name

        # Build expected backup filename pattern: backup_[displayname]_*.pptx
        expected_prefix = f"backup_{clean_display_name}_"
        expected_suffix = ".pptx"

        # Resolve the output base directory
        output_base, _ = resolve_project_output(
            clean_display_name,
            "NW",  # Default to NW project type
            fallback_subdir="generated_presentations",
        )

        # Try to find the file in multiple locations:
        # 1. Exact path from DB if it's absolute
        # 2. Root directory (new OpenXML backup location)
        # 3. Presentations subdirectory (legacy location)
        # 4. Fallback to any backup_*.pptx in root (newest file)

        file_path = None
        search_locations = []

        # Location 0: Absolute path if provided
        try:
            from pathlib import Path as PathLib2
            maybe_abs = PathLib2(main_ppt_filename)
            if maybe_abs.is_absolute():
                search_locations.append(("absolute", maybe_abs))
        except Exception:
            pass

        # Location 1: Root directory (new location for backup files)
        root_path = output_base / clean_ppt_filename
        search_locations.append(("root", root_path))

        # Location 2: Presentations subdirectory (legacy location)
        presentations_path = output_base / "Presentations" / clean_ppt_filename
        search_locations.append(("Presentations", presentations_path))

        logger.info(
            "Looking for PowerPoint file: display_name=%s, clean_display_name=%s, output_base=%s, filename_from_db=%s, clean_filename=%s",
            display_name,
            clean_display_name,
            output_base,
            main_ppt_filename,
            clean_ppt_filename
        )

        # Try each location in order
        for location_name, candidate_path in search_locations:
            logger.info("Checking %s: %s", location_name, candidate_path)
            if candidate_path.exists():
                file_path = candidate_path
                logger.info("Found PowerPoint file in %s: %s", location_name, file_path)
                break

        # If still not found, try fallback: find any backup_<displayname>_*.pptx in root (use newest)
        if not file_path:
            logger.warning("File not found in standard locations. Searching for backup_<displayname>_*.pptx in root...")
            backup_files = list(output_base.glob(f"{expected_prefix}*{expected_suffix}"))
            if backup_files:
                # Sort by modification time, newest first
                backup_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
                file_path = backup_files[0]
                logger.warning(
                    "Using most recent backup file: %s (expected: %s, presentation_id=%d)",
                    file_path.name,
                    clean_ppt_filename,
                    presentation_id
                )

        # If STILL not found, list available files and raise error
        if not file_path:
            # List all files in the root to help diagnose
            available_in_root = []

            if output_base.exists():
                available_in_root = [f.name for f in output_base.glob("*.pptx")]

            logger.error(
                "PowerPoint file not found. Expected prefix: %s (presentation_id=%d)\n"
                "  Available in root: %s",
                expected_prefix,
                presentation_id,
                available_in_root,
            )

            raise HTTPException(
                status_code=404,
                detail=f"PowerPoint file not found on server. Expected: {clean_ppt_filename}. "
                       f"Available in root: {available_in_root}. Available in Presentations/: {available_in_presentations}"
            )

        # Get file info and final download name (use actual file name on disk)
        file_size = file_path.stat().st_size
        media_type = guess_media_type(file_path)
        download_name = file_path.name

        logger.info(
            "Serving PowerPoint file: %s (size: %d bytes, presentation_id: %d)",
            download_name,
            file_size,
            presentation_id
        )

        # Create FileResponse with explicit Content-Disposition header (frontend expects filename here)
        response = FileResponse(
            path=file_path,
            media_type=media_type,
            headers={
                # Keep it simple for the frontend: only filename, no extra encoding params
                "Content-Disposition": f"attachment; filename={download_name}",
                "X-Content-Type-Options": "nosniff",
                "Content-Transfer-Encoding": "binary",
            },
        )

        return response

    except HTTPException:
        raise

    except Exception as e:
        logger.error(
            "Error downloading presentation file for ID %d: %s",
            presentation_id,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Failed to download presentation file: {str(e)}"
        ) from e

@router.post(
    "/generate-bsr-report",
    response_model=TaskCreatedResponse,
    summary="Generate BSR Excel and Word reports",
    description=(
        "Generate downloadable Excel and Word reports for a BSR presentation.\n\n"
        "**Excel Report**: Contains name candidates with categories\n"
        "**Word Report**: Contains slide notes, attributes, and key concepts\n\n"
        "Files are saved to: C:\\inetpub\\wwwroot\\CreativePythonAPI\\NW_Files\\downloads\\{DisplayName}.{ext}"
    ),
    responses={
        503: {
            "description": "Service at capacity - too many concurrent tasks.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": {
                            "error": "Service temporarily unavailable",
                            "message": "El sistema ha alcanzado su capacidad máxima de procesamiento. Por favor, intente nuevamente en unos momentos.",
                            "active_tasks": 2,
                            "max_concurrent_tasks": 2,
                            "retry_after_seconds": 30
                        }
                    }
                }
            },
        },
    },
)
async def generate_bsr_report(background_tasks: BackgroundTasks, request: BSRGenerateReportRequest) -> TaskCreatedResponse:
    """Generate BSR reports (Excel + Word)."""
    try:
        # Check concurrency limit before accepting the task
        check_concurrency_limit()
        # Create background task
        task_id = task_manager.create_task(
            task_type="generate_bsr_report",
            description=f"Generating BSR reports for: {request.display_name or request.presentation_id}",
            metadata={
                "presentation_id": request.presentation_id,
                "display_name": request.display_name
            }
        )

        logger.info(
            "Created background task %s for BSR report: presentation_id=%d",
            task_id,
            request.presentation_id,
        )

        # Submit task to concurrency manager
        submission_result = await concurrency_manager.submit_task(
            task_id=task_id,
            task_type="generate_bsr_report",
            func=_generate_bsr_report_background,
            args=(task_id, request.presentation_id, request.display_name),
            priority=TaskPriority.NORMAL,
        )

        # Return task information with queue position
        task_info = task_manager.get_task(task_id)

        return TaskCreatedResponse(
            task_id=task_id,
            status=submission_result["status"],
            message=f"BSR report generation task {'started' if submission_result.get('will_start_immediately') else 'queued'} for presentation {request.presentation_id}",
            task_type="generate_bsr_report",
            created_at=task_info["created_at"],
            status_url=f"/api/presentations/tasks/{task_id}",
            position=submission_result.get("position"),
            estimated_wait_seconds=submission_result.get("estimated_wait_seconds")
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error queueing BSR report: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to queue BSR report: {str(e)}"
        ) from e

@router.put(
    "/{presentation_id}/categories",
    summary="Update categories for a BSR project",
    response_model=UpdateCategoriesResponse,
    description=(
        "Update or delete categories for an existing BSR/NSR project.\n\n"
        "**Equivalente VB.NET:** FnInsertCategory_FromUpdate()\n\n"
        "**Process:**\n"
        "1. Check if project exists\n"
        "2. Check if categories already exist\n"
        "3. If categories exist and force_update=false, return 409 (confirmation required)\n"
        "4. Delete existing categories and elements\n"
        "5. Insert new categories based on mode\n\n"
        "**Modes:**\n"
        "- `single`: Insert only category1\n"
        "- `both`: Insert category1 and category2\n"
        "- `none`: Delete all categories (no insertion)\n\n"
        "**Query Parameters:**\n"
        "- `force_update`: If false, requires user confirmation when categories exist\n\n"
        "**Example Payloads:**\n\n"
        "Update to 2 categories:\n"
        "```json\n"
        '{\n'
        '  "mode": "both",\n'
        '  "category1": {\n'
        '    "name": "Sales Region",\n'
        '    "elements": ["Americas", "EMEA", "APAC"]\n'
        '  },\n'
        '  "category2": {\n'
        '    "name": "Product Type",\n'
        '    "elements": ["Software", "Hardware", "Services"]\n'
        '  }\n'
        '}\n'
        "```\n\n"
        "Delete all categories:\n"
        "```json\n"
        '{\n'
        '  "mode": "none"\n'
        '}\n'
        "```\n"
    ),
    responses={
        200: {
            "description": "Categories updated successfully",
            "content": {
                "application/json": {
                    "example": {
                        "message": "Categories updated successfully",
                        "categories_updated": 2,
                        "deleted_count": 2
                    }
                }
            },
        },
        404: {
            "description": "Project not found",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Project 12345 not found"
                    }
                }
            },
        },
        409: {
            "description": "Categories exist - confirmation required",
            "content": {
                "application/json": {
                    "example": {
                        "detail": {
                            "message": "Categories already exist for this project",
                            "existing_count": 2,
                            "action_required": "Add ?force_update=true to update categories"
                        }
                    }
                }
            },
        },
        422: {
            "description": "Validation error",
            "content": {
                "application/json": {
                    "example": {
                        "detail": [
                            {
                                "loc": ["body", "mode"],
                                "msg": "mode must be 'single', 'both', or 'none'",
                                "type": "value_error"
                            }
                        ]
                    }
                }
            },
        },
    },
)
async def update_project_categories(
    presentation_id: int,
    categories_data: ProjectCategoriesUpdate,
    force_update: bool = Query(
        False,
        description="Force update without confirmation when categories exist"
    ),
) -> UpdateCategoriesResponse:
    """
    Update categories for an existing BSR/NSR project.

    Equivalente VB.NET: FnInsertCategory_FromUpdate() en clsData.vb

    Args:
        presentation_id: ID of the BSR project
        categories_data: Category update payload (mode, category1, category2)
        force_update: If true, bypasses confirmation when categories exist

    Returns:
        UpdateCategoriesResponse with message, categories_updated, deleted_count

    Raises:
        HTTPException 404: If project doesn't exist
        HTTPException 409: If categories exist and force_update=false
        HTTPException 422: If validation fails
        HTTPException 500: If update operation fails
    """
    try:
        # Step 1: Verify project exists
        logger.info("Checking if BSR project %d exists", presentation_id)
        project_exists = presentation_service.check_bsr_project_exists(presentation_id)

        if not project_exists:
            raise HTTPException(
                status_code=404,
                detail=f"Project {presentation_id} not found"
            )

        # Step 2: Check if categories already exist
        existing_count = presentation_service.count_project_categories(presentation_id)

        logger.info(
            "Project %d has %d existing categories (force_update=%s)",
            presentation_id,
            existing_count,
            force_update
        )

        # Step 3: If categories exist and no force_update, require confirmation
        if existing_count > 0 and not force_update:
            raise HTTPException(
                status_code=409,
                detail={
                    "message": "Categories already exist for this project",
                    "existing_count": existing_count,
                    "action_required": "Add ?force_update=true to update categories"
                }
            )

        # Step 4: Update categories
        logger.info(
            "Updating categories for project %d with mode='%s'",
            presentation_id,
            categories_data.mode
        )

        result = presentation_service.update_project_categories(
            presentation_id=presentation_id,
            categories_data=categories_data.model_dump()
        )

        logger.info(
            "Categories update completed for project %d: %s",
            presentation_id,
            result
        )

        return UpdateCategoriesResponse(
            message=result["message"],
            categories_updated=result["categories_updated"],
            deleted_count=result["deleted_count"]
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(
            "Error updating categories for project %d: %s",
            presentation_id,
            str(e),
            exc_info=True
        )
        raise HTTPException(
            status_code=500,
            detail=f"Failed to update categories: {str(e)}"
        ) from e

