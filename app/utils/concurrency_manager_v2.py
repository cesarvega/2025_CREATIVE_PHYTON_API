"""Concurrency manager - orchestrates queue and workers."""

import asyncio
import time
import threading
from typing import Any, Callable, Dict, Optional
from datetime import datetime, timedelta

from app.config.settings import settings
from app.utils.logging_utils import get_logger
from app.utils.task_manager import TaskStatus, task_manager
from app.utils.task_queue import TaskQueue, TaskPriority
from app.utils.task_worker import WorkerPool

logger = get_logger(__name__)


class ConcurrencyManager:
    """Coordinates task queue and worker pool."""

    def __init__(
        self,
        max_workers: int = None,
        queue_size: int = None,
        timeout_seconds: int = None,
    ):
        """Initialize concurrency manager.

        Args:
            max_workers: Maximum concurrent workers (default from settings)
            queue_size: Maximum queue size (default from settings)
            timeout_seconds: Task timeout (default from settings)
        """
        self.max_workers = max_workers or settings.concurrency_max_workers
        self.queue_size = queue_size or settings.concurrency_queue_size
        self.timeout_seconds = timeout_seconds or settings.task_timeout_seconds

        # Initialize components
        self.task_queue = TaskQueue(max_size=self.queue_size)
        self.worker_pool = WorkerPool(
            num_workers=self.max_workers,
            task_queue=self.task_queue,
            status_callback=self._on_worker_status_change,
            progress_callback=self._on_task_progress,
            timeout_seconds=self.timeout_seconds,
        )

        # Statistics
        self._total_tasks_submitted = 0
        self._total_tasks_completed = 0
        self._total_tasks_failed = 0

        logger.info(
            "ConcurrencyManager initialized: max_workers=%d, queue_size=%d, timeout=%ds",
            self.max_workers, self.queue_size, self.timeout_seconds
        )

    def start(self):
        """Start the concurrency manager (start worker pool)."""
        logger.info("Starting ConcurrencyManager")
        self.worker_pool.start()
        logger.info("ConcurrencyManager started")

    def stop(self, wait: bool = True, timeout: float = 10.0):
        """Stop the concurrency manager (stop worker pool).

        Args:
            wait: Wait for workers to finish
            timeout: Maximum wait time per worker
        """
        logger.info("Stopping ConcurrencyManager")
        self.worker_pool.stop(wait=wait, timeout=timeout)
        logger.info("ConcurrencyManager stopped")

    async def submit_task(
        self,
        task_id: str,
        task_type: str,
        func: Callable,
        args: tuple = (),
        kwargs: dict = None,
        priority: TaskPriority = TaskPriority.NORMAL,
        metadata: Dict[str, Any] = None,
    ) -> Dict[str, Any]:
        """Submit a task for execution.

        Args:
            task_id: Unique task identifier
            task_type: Type of task
            func: Function to execute
            args: Positional arguments
            kwargs: Keyword arguments
            priority: Task priority
            metadata: Additional metadata

        Returns:
            Dictionary with task submission info

        Raises:
            ValueError: If queue is full or task already exists
        """
        if kwargs is None:
            kwargs = {}
        if metadata is None:
            metadata = {}

        logger.info("Submitting task %s (type=%s, priority=%s)", task_id, task_type, priority.name)

        # Check if there are available worker slots
        pool_stats = self.worker_pool.get_stats()
        idle_workers = pool_stats["idle_workers"]
        processing_workers = pool_stats["processing_workers"]

        # Add to queue first
        queue_position = self.task_queue.enqueue(
            task_id=task_id,
            task_type=task_type,
            func=func,
            args=args,
            kwargs=kwargs,
            priority=priority,
            metadata=metadata,
        )

        self._total_tasks_submitted += 1

        if idle_workers > 0:
            # Task will be processed immediately
            logger.info("Task %s will be processed immediately (idle workers: %d)", task_id, idle_workers)

            return {
                "task_id": task_id,
                "status": "queued",
                "position": queue_position,
                "will_start_immediately": True,
                "estimated_wait_seconds": 0,
            }
        else:
            # Task will wait in queue
            # Position = number of tasks ahead = processing + queued before this task
            actual_position = processing_workers + queue_position

            logger.info(
                "Task %s will wait in queue (no idle workers, position=%d, processing=%d, queue_pos=%d)",
                task_id, actual_position, processing_workers, queue_position
            )

            # Estimate wait time based on actual position
            estimated_wait = self._estimate_wait_time(actual_position)

            return {
                "task_id": task_id,
                "status": "queued",
                "position": actual_position,
                "will_start_immediately": False,
                "estimated_wait_seconds": estimated_wait,
            }

    async def cancel_task(self, task_id: str) -> bool:
        """Cancel a task if it's still in queue.

        Args:
            task_id: Task identifier

        Returns:
            True if cancelled, False if not found or already processing
        """
        # Check if task is currently being processed
        active_tasks = self.worker_pool.get_active_tasks()
        if task_id in active_tasks:
            logger.warning("Cannot cancel task %s - already processing", task_id)
            return False

        # Try to cancel from queue
        cancelled = self.task_queue.cancel(task_id)

        if cancelled:
            logger.info("Task %s cancelled from queue", task_id)

            # Update task status
            task_manager.update_status(
                task_id,
                TaskStatus.FAILED,
                error="Task cancelled by user"
            )

            return True

        return False

    def get_task_position(self, task_id: str) -> Optional[int]:
        """Get current position of a task in queue.

        Args:
            task_id: Task identifier

        Returns:
            Position (0-indexed) or None if not in queue
        """
        return self.task_queue.get_position(task_id)

    def get_queue_status(self) -> Dict[str, Any]:
        """Get current queue status.

        Returns:
            Dictionary with queue statistics
        """
        queue_stats = self.task_queue.get_stats()
        pool_stats = self.worker_pool.get_stats()

        return {
            "active_count": pool_stats["processing_workers"],
            "queued_count": queue_stats["queued_count"],
            "available_slots": pool_stats["idle_workers"],
            "max_workers": self.max_workers,
            "queue_capacity": self.queue_size,
            "total_submitted": self._total_tasks_submitted,
            "total_completed": self._total_tasks_completed,
            "total_failed": self._total_tasks_failed,
        }

    def get_task_history(self, limit: int = 100) -> list[Dict[str, Any]]:
        """Get recent task history.

        Args:
            limit: Maximum number of tasks to return

        Returns:
            List of task information
        """
        return task_manager.list_tasks(limit=limit)

    def _estimate_wait_time(self, position: int) -> float:
        """Estimate wait time for a task at given position.

        Args:
            position: Number of tasks ahead (0 = process immediately, 1+ = wait)

        Returns:
            Estimated wait time in seconds
        """
        # Simple estimation: assume each task takes average time
        # For PowerPoint tasks, assume 90 seconds average
        # position represents number of tasks ahead, so multiply directly
        avg_task_time = 90.0
        return position * avg_task_time

    def _on_worker_status_change(self, worker_id: int, status: str, current_task_id: Optional[str]):
        """Callback when worker status changes.

        Args:
            worker_id: Worker identifier
            status: New status (idle, processing, stopped)
            current_task_id: Current task ID if processing
        """
        logger.debug("Worker %d status changed: %s (task: %s)", worker_id, status, current_task_id)

    def _on_task_progress(self, task_id: str, event_name: str, data: Dict[str, Any]):
        """Callback for task progress updates.

        Args:
            task_id: Task identifier
            event_name: Event name
            data: Event data
        """
        logger.debug("Task %s progress: %s", task_id, event_name)

        # Update statistics
        if event_name == "task:completed":
            self._total_tasks_completed += 1
        elif event_name == "task:failed":
            self._total_tasks_failed += 1

    async def cleanup_old_tasks(self, max_age_hours: int = 24) -> int:
        """Clean up old completed/failed tasks.

        Args:
            max_age_hours: Maximum age in hours

        Returns:
            Number of tasks cleaned up
        """
        logger.info("Cleaning up tasks older than %d hours", max_age_hours)
        cleaned = task_manager.cleanup_old_tasks(max_age_hours=max_age_hours)
        logger.info("Cleaned up %d old tasks", cleaned)
        return cleaned


# Global concurrency manager instance
concurrency_manager = ConcurrencyManager()
