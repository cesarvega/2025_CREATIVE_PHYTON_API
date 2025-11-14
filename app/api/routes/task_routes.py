"""Routes for task queue management and monitoring."""

from typing import Optional
from fastapi import APIRouter, HTTPException, Query

from app.utils.task_manager import task_manager, TaskStatus
from app.utils.concurrency_manager_v2 import concurrency_manager
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

router = APIRouter(prefix="/tasks", tags=["Task Management"])


@router.get(
    "/{task_id}",
    summary="Get detailed task status",
    description="Get detailed status information for a specific task including progress, result, and error details."
)
async def get_task_status(task_id: str):
    """Get detailed task status.

    Args:
        task_id: Task identifier

    Returns:
        Task status information
    """
    task = task_manager.get_task(task_id)

    if not task:
        raise HTTPException(status_code=404, detail=f"Task {task_id} not found")

    # Calculate current position in queue if task is pending
    position = None
    estimated_wait_seconds = None

    if task["status"] == "pending":
        queue_position = concurrency_manager.get_task_position(task_id)
        if queue_position is not None:
            # Task is in queue - calculate actual position
            pool_stats = concurrency_manager.worker_pool.get_stats()
            processing_workers = pool_stats["processing_workers"]
            position = processing_workers + queue_position
            estimated_wait_seconds = position * 90.0

    return {
        "task_id": task["task_id"],
        "task_type": task["task_type"],
        "status": task["status"],
        "progress": task["progress"],
        "description": task.get("description"),
        "result": task.get("result"),
        "error": task.get("error"),
        "created_at": task["created_at"],
        "started_at": task.get("started_at"),
        "completed_at": task.get("completed_at"),
        "metadata": task.get("metadata", {}),
        "position": position,
        "estimated_wait_seconds": estimated_wait_seconds,
    }


@router.delete(
    "/{task_id}",
    summary="Cancel a queued task",
    description="Cancel a task if it's still in the queue. Returns error if task is already processing."
)
async def cancel_task(task_id: str):
    """Cancel a task in the queue.

    Args:
        task_id: Task identifier

    Returns:
        Cancellation status
    """
    cancelled = await concurrency_manager.cancel_task(task_id)

    if cancelled:
        return {
            "success": True,
            "message": f"Task {task_id} cancelled successfully",
            "task_id": task_id,
        }
    else:
        # Check if task exists
        task = task_manager.get_task(task_id)
        if not task:
            raise HTTPException(status_code=404, detail=f"Task {task_id} not found")

        # Task exists but cannot be cancelled (probably processing)
        raise HTTPException(
            status_code=409,
            detail=f"Task {task_id} is already processing and cannot be cancelled"
        )


@router.get(
    "/queue/status",
    summary="Get queue status",
    description="Get global queue status including active tasks, queued tasks, and available worker slots."
)
async def get_queue_status():
    """Get global queue status.

    Returns:
        Queue statistics and status
    """
    status = concurrency_manager.get_queue_status()

    return {
        "active_count": status["active_count"],
        "queued_count": status["queued_count"],
        "available_slots": status["available_slots"],
        "max_workers": status["max_workers"],
        "queue_capacity": status["queue_capacity"],
        "statistics": {
            "total_submitted": status["total_submitted"],
            "total_completed": status["total_completed"],
            "total_failed": status["total_failed"],
        },
    }


@router.get(
    "/queue/position/{task_id}",
    summary="Get task position in queue",
    description="Get the current position of a task in the queue and estimated wait time."
)
async def get_task_position(task_id: str):
    """Get task position in queue.

    Args:
        task_id: Task identifier

    Returns:
        Position and estimated wait time
    """
    # Get position in queue (0-indexed within queue only)
    queue_position = concurrency_manager.get_task_position(task_id)

    if queue_position is None:
        # Task not in queue - check if it exists
        task = task_manager.get_task(task_id)
        if not task:
            raise HTTPException(status_code=404, detail=f"Task {task_id} not found")

        # Task exists but not in queue (probably processing or completed)
        return {
            "task_id": task_id,
            "in_queue": False,
            "status": task["status"],
            "message": f"Task is not in queue (status: {task['status']})",
        }

    # Task is in queue - calculate actual position including processing tasks
    pool_stats = concurrency_manager.worker_pool.get_stats()
    processing_workers = pool_stats["processing_workers"]

    # Actual position = number of tasks ahead (processing + queued before this one)
    actual_position = processing_workers + queue_position

    # Estimate wait time based on actual position
    estimated_wait = actual_position * 90.0  # Assume 90 seconds per task

    return {
        "task_id": task_id,
        "in_queue": True,
        "position": actual_position,
        "queue_position": queue_position,
        "processing_tasks": processing_workers,
        "estimated_wait_seconds": estimated_wait,
        "estimated_wait_minutes": round(estimated_wait / 60, 1),
    }


@router.get(
    "/queue/history",
    summary="Get task history",
    description="Get recent task history (last N tasks) with filtering options."
)
async def get_task_history(
    limit: int = Query(100, ge=1, le=500, description="Maximum number of tasks to return"),
    task_type: Optional[str] = Query(None, description="Filter by task type"),
    status: Optional[TaskStatus] = Query(None, description="Filter by status"),
):
    """Get task history.

    Args:
        limit: Maximum number of tasks to return
        task_type: Optional task type filter
        status: Optional status filter

    Returns:
        List of recent tasks
    """
    tasks = task_manager.list_tasks(
        task_type=task_type,
        status=status,
        limit=limit,
    )

    return {
        "total": len(tasks),
        "limit": limit,
        "filters": {
            "task_type": task_type,
            "status": status.value if status else None,
        },
        "tasks": tasks,
    }


@router.post(
    "/queue/cleanup",
    summary="Cleanup old tasks",
    description="Remove old completed/failed tasks from the system to free up memory."
)
async def cleanup_old_tasks(
    max_age_hours: int = Query(24, ge=1, le=168, description="Maximum age in hours")
):
    """Cleanup old tasks.

    Args:
        max_age_hours: Tasks older than this will be removed

    Returns:
        Number of tasks cleaned up
    """
    cleaned = await concurrency_manager.cleanup_old_tasks(max_age_hours=max_age_hours)

    return {
        "success": True,
        "cleaned_count": cleaned,
        "max_age_hours": max_age_hours,
        "message": f"Cleaned up {cleaned} tasks older than {max_age_hours} hours",
    }
