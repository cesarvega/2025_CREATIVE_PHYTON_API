"""Task status management for background operations.

This module provides a simple in-memory task status tracking system for long-running
background tasks. For production with multiple workers, consider using Redis or a database.
"""

import uuid
from datetime import datetime
from typing import Dict, Optional, Any
from enum import Enum
from threading import Lock

from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class TaskStatus(str, Enum):
    """Task status enumeration."""
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"


class TaskManager:
    """Manages background task status tracking."""

    def __init__(self):
        """Initialize task manager with thread-safe storage."""
        self._tasks: Dict[str, Dict[str, Any]] = {}
        self._lock = Lock()

    def create_task(
        self,
        task_type: str,
        description: Optional[str] = None,
        metadata: Optional[Dict[str, Any]] = None
    ) -> str:
        """Create a new task and return its ID.

        Args:
            task_type: Type of task (e.g., 'create_presentation', 'generate_backup')
            description: Human-readable task description
            metadata: Additional metadata to store with the task

        Returns:
            task_id: Unique task identifier
        """
        task_id = str(uuid.uuid4())

        with self._lock:
            self._tasks[task_id] = {
                "task_id": task_id,
                "task_type": task_type,
                "status": TaskStatus.PENDING,
                "description": description or task_type,
                "progress": 0,
                "result": None,
                "error": None,
                "metadata": metadata or {},
                "created_at": datetime.now().isoformat(),
                "started_at": None,
                "completed_at": None,
            }

        logger.info("Created task %s (type: %s)", task_id, task_type)
        return task_id

    def update_status(
        self,
        task_id: str,
        status: TaskStatus,
        progress: Optional[int] = None,
        result: Optional[Any] = None,
        error: Optional[str] = None
    ) -> None:
        """Update task status.

        Args:
            task_id: Task identifier
            status: New task status
            progress: Progress percentage (0-100)
            result: Task result (when completed)
            error: Error message (when failed)
        """
        with self._lock:
            if task_id not in self._tasks:
                logger.warning("Attempted to update non-existent task: %s", task_id)
                return

            task = self._tasks[task_id]
            task["status"] = status

            if progress is not None:
                task["progress"] = min(100, max(0, progress))

            if result is not None:
                task["result"] = result

            if error is not None:
                task["error"] = error

            # Update timestamps
            if status == TaskStatus.PROCESSING and task["started_at"] is None:
                task["started_at"] = datetime.now().isoformat()

            if status in (TaskStatus.COMPLETED, TaskStatus.FAILED):
                task["completed_at"] = datetime.now().isoformat()
                task["progress"] = 100 if status == TaskStatus.COMPLETED else task["progress"]

        logger.info("Updated task %s: status=%s, progress=%s", task_id, status, progress)

    def get_task(self, task_id: str) -> Optional[Dict[str, Any]]:
        """Get task information by ID.

        Args:
            task_id: Task identifier

        Returns:
            Task information dictionary or None if not found
        """
        with self._lock:
            return self._tasks.get(task_id)

    def delete_task(self, task_id: str) -> bool:
        """Delete a task from the manager.

        Args:
            task_id: Task identifier

        Returns:
            True if task was deleted, False if not found
        """
        with self._lock:
            if task_id in self._tasks:
                del self._tasks[task_id]
                logger.info("Deleted task %s", task_id)
                return True
            return False

    def list_tasks(
        self,
        task_type: Optional[str] = None,
        status: Optional[TaskStatus] = None,
        limit: int = 100
    ) -> list[Dict[str, Any]]:
        """List tasks with optional filtering.

        Args:
            task_type: Filter by task type
            status: Filter by status
            limit: Maximum number of tasks to return

        Returns:
            List of task dictionaries
        """
        with self._lock:
            tasks = list(self._tasks.values())

        # Apply filters
        if task_type:
            tasks = [t for t in tasks if t["task_type"] == task_type]

        if status:
            tasks = [t for t in tasks if t["status"] == status]

        # Sort by creation time (newest first)
        tasks.sort(key=lambda t: t["created_at"], reverse=True)

        return tasks[:limit]

    def cleanup_old_tasks(self, max_age_hours: int = 24) -> int:
        """Clean up old completed/failed tasks.

        Args:
            max_age_hours: Maximum age in hours for completed tasks

        Returns:
            Number of tasks cleaned up
        """
        from datetime import timedelta

        cutoff_time = datetime.now() - timedelta(hours=max_age_hours)
        cleaned = 0

        with self._lock:
            tasks_to_delete = []

            for task_id, task in self._tasks.items():
                if task["status"] in (TaskStatus.COMPLETED, TaskStatus.FAILED):
                    completed_at = task.get("completed_at")
                    if completed_at:
                        task_time = datetime.fromisoformat(completed_at)
                        if task_time < cutoff_time:
                            tasks_to_delete.append(task_id)

            for task_id in tasks_to_delete:
                del self._tasks[task_id]
                cleaned += 1

        if cleaned > 0:
            logger.info("Cleaned up %d old tasks", cleaned)

        return cleaned


# Global task manager instance
task_manager = TaskManager()
