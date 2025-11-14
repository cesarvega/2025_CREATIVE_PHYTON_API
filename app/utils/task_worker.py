"""Worker pool for executing background tasks with progress tracking."""

import threading
import time
from typing import Callable, Optional
from datetime import datetime

from app.utils.logging_utils import get_logger
from app.utils.task_manager import TaskStatus, task_manager
from app.utils.task_queue import TaskQueue

logger = get_logger(__name__)


class TaskWorker:
    """Worker thread that consumes tasks from queue and executes them."""

    def __init__(
        self,
        worker_id: int,
        task_queue: TaskQueue,
        status_callback: Optional[Callable] = None,
        progress_callback: Optional[Callable] = None,
        timeout_seconds: int = 600,
    ):
        """Initialize task worker.

        Args:
            worker_id: Unique worker identifier
            task_queue: Queue to consume tasks from
            status_callback: Callback when worker status changes
            progress_callback: Callback for progress updates
            timeout_seconds: Task execution timeout
        """
        self.worker_id = worker_id
        self.task_queue = task_queue
        self.status_callback = status_callback
        self.progress_callback = progress_callback
        self.timeout_seconds = timeout_seconds

        self._thread: Optional[threading.Thread] = None
        self._running = False
        self._current_task_id: Optional[str] = None
        self._stop_event = threading.Event()

        logger.info("TaskWorker %d initialized", worker_id)

    def start(self):
        """Start the worker thread."""
        if self._running:
            logger.warning("Worker %d already running", self.worker_id)
            return

        self._running = True
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._worker_loop,
            name=f"TaskWorker-{self.worker_id}",
            daemon=True,
        )
        self._thread.start()
        logger.info("Worker %d started", self.worker_id)

        if self.status_callback:
            self.status_callback(self.worker_id, "idle", None)

    def stop(self, wait: bool = True, timeout: float = 5.0):
        """Stop the worker thread.

        Args:
            wait: Wait for worker to finish current task
            timeout: Maximum time to wait (seconds)
        """
        if not self._running:
            return

        logger.info("Stopping worker %d (wait=%s)", self.worker_id, wait)
        self._running = False
        self._stop_event.set()

        if wait and self._thread and self._thread.is_alive():
            self._thread.join(timeout=timeout)
            if self._thread.is_alive():
                logger.warning("Worker %d did not stop within timeout", self.worker_id)

        logger.info("Worker %d stopped", self.worker_id)

        if self.status_callback:
            self.status_callback(self.worker_id, "stopped", None)

    def _worker_loop(self):
        """Main worker loop - continuously process tasks from queue."""
        logger.info("Worker %d entering main loop", self.worker_id)

        while self._running:
            try:
                # Try to get a task from queue (non-blocking)
                task = self.task_queue.dequeue()

                if task is None:
                    # No tasks available, wait a bit
                    self._stop_event.wait(timeout=0.5)
                    continue

                # Mark as processing IMMEDIATELY after dequeue to prevent race condition
                self._current_task_id = task.task_id

                # Execute the task
                self._execute_task(task)

            except Exception as exc:
                logger.error("Worker %d encountered error in loop: %s", self.worker_id, exc, exc_info=True)
                # Clear task ID on error
                self._current_task_id = None
                time.sleep(1)  # Brief pause before retrying

        logger.info("Worker %d exiting main loop", self.worker_id)

    def _execute_task(self, task):
        """Execute a single task with timeout and error handling.

        Args:
            task: QueuedTask to execute
        """
        # _current_task_id is already set in _worker_loop to prevent race condition
        start_time = time.time()

        logger.info(
            "Worker %d executing task %s (type=%s, priority=%d)",
            self.worker_id, task.task_id, task.task_type, task.priority
        )

        # Notify status change
        if self.status_callback:
            self.status_callback(self.worker_id, "processing", task.task_id)

        # Update task status to processing
        task_manager.update_status(
            task.task_id,
            TaskStatus.PROCESSING,
            progress=0
        )

        # Emit task:processing event
        if self.progress_callback:
            self.progress_callback(task.task_id, "task:processing", {
                "task_id": task.task_id,
                "started_at": datetime.now().isoformat(),
                "worker_id": self.worker_id,
            })

        try:
            # Execute the task function
            task.func(*task.args, **task.kwargs)

            elapsed = time.time() - start_time

            logger.info(
                "Worker %d completed task %s in %.2fs",
                self.worker_id, task.task_id, elapsed
            )

            # Task function handles its own status updates and events
            # No need to update here unless task failed to do so

        except Exception as exc:
            elapsed = time.time() - start_time
            error_msg = f"Task execution failed: {str(exc)}"

            logger.error(
                "Worker %d failed task %s after %.2fs: %s",
                self.worker_id, task.task_id, elapsed, exc, exc_info=True
            )

            # Update task status to failed
            task_manager.update_status(
                task.task_id,
                TaskStatus.FAILED,
                error=error_msg
            )

            # Emit task:failed event
            if self.progress_callback:
                self.progress_callback(task.task_id, "task:failed", {
                    "task_id": task.task_id,
                    "error": error_msg,
                    "error_code": "EXECUTION_ERROR",
                    "elapsed_seconds": elapsed,
                })

        finally:
            self._current_task_id = None

            # Notify worker is back to idle
            if self.status_callback:
                self.status_callback(self.worker_id, "idle", None)

    def get_status(self) -> dict:
        """Get current worker status.

        Returns:
            Dictionary with worker status information
        """
        return {
            "worker_id": self.worker_id,
            "running": self._running,
            "current_task_id": self._current_task_id,
            "status": "processing" if self._current_task_id else ("idle" if self._running else "stopped"),
        }

    def is_idle(self) -> bool:
        """Check if worker is idle (not processing a task).

        Returns:
            True if worker is idle
        """
        return self._running and self._current_task_id is None


class WorkerPool:
    """Pool of worker threads for parallel task execution."""

    def __init__(
        self,
        num_workers: int,
        task_queue: TaskQueue,
        status_callback: Optional[Callable] = None,
        progress_callback: Optional[Callable] = None,
        timeout_seconds: int = 600,
    ):
        """Initialize worker pool.

        Args:
            num_workers: Number of worker threads
            task_queue: Shared task queue
            status_callback: Callback for worker status changes
            progress_callback: Callback for progress updates
            timeout_seconds: Task execution timeout
        """
        self.num_workers = num_workers
        self.task_queue = task_queue
        self.status_callback = status_callback
        self.progress_callback = progress_callback
        self.timeout_seconds = timeout_seconds

        self.workers: list[TaskWorker] = []

        logger.info("WorkerPool initialized with %d workers", num_workers)

    def start(self):
        """Start all workers in the pool."""
        logger.info("Starting worker pool with %d workers", self.num_workers)

        for i in range(self.num_workers):
            worker = TaskWorker(
                worker_id=i,
                task_queue=self.task_queue,
                status_callback=self.status_callback,
                progress_callback=self.progress_callback,
                timeout_seconds=self.timeout_seconds,
            )
            worker.start()
            self.workers.append(worker)

        logger.info("Worker pool started with %d workers", len(self.workers))

    def stop(self, wait: bool = True, timeout: float = 10.0):
        """Stop all workers in the pool.

        Args:
            wait: Wait for workers to finish
            timeout: Maximum time to wait per worker
        """
        logger.info("Stopping worker pool (%d workers)", len(self.workers))

        for worker in self.workers:
            worker.stop(wait=wait, timeout=timeout)

        self.workers.clear()
        logger.info("Worker pool stopped")

    def get_stats(self) -> dict:
        """Get worker pool statistics.

        Returns:
            Dictionary with pool stats
        """
        idle_count = sum(1 for w in self.workers if w.is_idle())
        processing_count = sum(1 for w in self.workers if w._current_task_id is not None)

        return {
            "total_workers": len(self.workers),
            "idle_workers": idle_count,
            "processing_workers": processing_count,
            "workers": [w.get_status() for w in self.workers],
        }

    def get_active_tasks(self) -> list[str]:
        """Get list of currently processing task IDs.

        Returns:
            List of task IDs
        """
        return [
            w._current_task_id
            for w in self.workers
            if w._current_task_id is not None
        ]
