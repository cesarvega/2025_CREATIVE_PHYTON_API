"""Task queue management with priority and persistence."""

import asyncio
import time
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Any, Callable, Dict, List, Optional
from threading import Lock

from app.utils.logging_utils import get_logger
from app.utils.task_manager import TaskStatus

logger = get_logger(__name__)


class TaskPriority(int, Enum):
    """Task priority levels (lower number = higher priority)."""
    HIGH = 1
    NORMAL = 2
    LOW = 3


@dataclass(order=True)
class QueuedTask:
    """Represents a task in the queue with priority ordering."""

    # Priority field comes first for ordering
    priority: int = field(compare=True)

    # Other fields not used in comparison
    task_id: str = field(compare=False)
    task_type: str = field(compare=False)
    func: Callable = field(compare=False)
    args: tuple = field(compare=False, default_factory=tuple)
    kwargs: dict = field(compare=False, default_factory=dict)
    created_at: float = field(compare=False, default_factory=time.time)
    metadata: Dict[str, Any] = field(compare=False, default_factory=dict)


class TaskQueue:
    """Thread-safe task queue with priority and position tracking."""

    def __init__(self, max_size: int = 100):
        """Initialize task queue.

        Args:
            max_size: Maximum number of tasks allowed in queue
        """
        self._queue: List[QueuedTask] = []
        self._lock = Lock()
        self._max_size = max_size
        self._task_index: Dict[str, QueuedTask] = {}  # Fast task lookup by ID
        logger.info("TaskQueue initialized with max_size=%d", max_size)

    def enqueue(
        self,
        task_id: str,
        task_type: str,
        func: Callable,
        args: tuple = (),
        kwargs: dict = None,
        priority: TaskPriority = TaskPriority.NORMAL,
        metadata: Dict[str, Any] = None,
    ) -> int:
        """Add a task to the queue.

        Args:
            task_id: Unique task identifier
            task_type: Type of task
            func: Function to execute
            args: Positional arguments for func
            kwargs: Keyword arguments for func
            priority: Task priority
            metadata: Additional task metadata

        Returns:
            Position in queue (0-indexed)

        Raises:
            ValueError: If queue is full or task already exists
        """
        if kwargs is None:
            kwargs = {}
        if metadata is None:
            metadata = {}

        with self._lock:
            if len(self._queue) >= self._max_size:
                raise ValueError(f"Queue is full (max size: {self._max_size})")

            if task_id in self._task_index:
                raise ValueError(f"Task {task_id} already exists in queue")

            task = QueuedTask(
                priority=priority.value,
                task_id=task_id,
                task_type=task_type,
                func=func,
                args=args,
                kwargs=kwargs,
                metadata=metadata,
            )

            self._queue.append(task)
            self._task_index[task_id] = task

            # Sort queue by priority (lower number = higher priority)
            self._queue.sort()

            position = self._get_position_unsafe(task_id)

            logger.info(
                "Task %s enqueued at position %d (priority=%s, type=%s)",
                task_id, position, priority.name, task_type
            )

            return position

    def dequeue(self) -> Optional[QueuedTask]:
        """Remove and return the highest priority task from queue.

        Returns:
            QueuedTask if queue is not empty, None otherwise
        """
        with self._lock:
            if not self._queue:
                return None

            task = self._queue.pop(0)
            del self._task_index[task.task_id]

            logger.info(
                "Task %s dequeued (priority=%d, type=%s, waited=%.1fs)",
                task.task_id, task.priority, task.task_type,
                time.time() - task.created_at
            )

            return task

    def cancel(self, task_id: str) -> bool:
        """Cancel a task in the queue.

        Args:
            task_id: Task identifier

        Returns:
            True if task was cancelled, False if not found
        """
        with self._lock:
            if task_id not in self._task_index:
                return False

            task = self._task_index[task_id]
            self._queue.remove(task)
            del self._task_index[task_id]

            logger.info("Task %s cancelled from queue", task_id)
            return True

    def get_position(self, task_id: str) -> Optional[int]:
        """Get the current position of a task in the queue.

        Args:
            task_id: Task identifier

        Returns:
            Position (0-indexed) or None if not found
        """
        with self._lock:
            return self._get_position_unsafe(task_id)

    def _get_position_unsafe(self, task_id: str) -> Optional[int]:
        """Get position without acquiring lock (internal use only)."""
        try:
            return next(
                i for i, task in enumerate(self._queue)
                if task.task_id == task_id
            )
        except StopIteration:
            return None

    def get_task(self, task_id: str) -> Optional[QueuedTask]:
        """Get task by ID without removing it.

        Args:
            task_id: Task identifier

        Returns:
            QueuedTask or None if not found
        """
        with self._lock:
            return self._task_index.get(task_id)

    def get_stats(self) -> Dict[str, Any]:
        """Get queue statistics.

        Returns:
            Dictionary with queue stats
        """
        with self._lock:
            return {
                "queued_count": len(self._queue),
                "max_size": self._max_size,
                "available_slots": self._max_size - len(self._queue),
                "tasks_by_priority": self._count_by_priority(),
            }

    def _count_by_priority(self) -> Dict[str, int]:
        """Count tasks by priority (internal use only)."""
        counts = {p.name: 0 for p in TaskPriority}
        for task in self._queue:
            priority_name = TaskPriority(task.priority).name
            counts[priority_name] += 1
        return counts

    def list_tasks(self, limit: int = 100) -> List[Dict[str, Any]]:
        """List all tasks in queue.

        Args:
            limit: Maximum number of tasks to return

        Returns:
            List of task information dictionaries
        """
        with self._lock:
            return [
                {
                    "task_id": task.task_id,
                    "task_type": task.task_type,
                    "priority": TaskPriority(task.priority).name,
                    "position": i,
                    "created_at": datetime.fromtimestamp(task.created_at).isoformat(),
                    "wait_time_seconds": time.time() - task.created_at,
                    "metadata": task.metadata,
                }
                for i, task in enumerate(self._queue[:limit])
            ]

    def size(self) -> int:
        """Get current queue size.

        Returns:
            Number of tasks in queue
        """
        with self._lock:
            return len(self._queue)

    def is_empty(self) -> bool:
        """Check if queue is empty.

        Returns:
            True if queue is empty
        """
        with self._lock:
            return len(self._queue) == 0

    def is_full(self) -> bool:
        """Check if queue is full.

        Returns:
            True if queue is at max capacity
        """
        with self._lock:
            return len(self._queue) >= self._max_size

    def clear(self) -> int:
        """Clear all tasks from queue.

        Returns:
            Number of tasks cleared
        """
        with self._lock:
            count = len(self._queue)
            self._queue.clear()
            self._task_index.clear()
            logger.info("Cleared %d tasks from queue", count)
            return count
