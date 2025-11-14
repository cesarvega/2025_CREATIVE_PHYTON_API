"""COM concurrency manager to handle Office automation threading."""

import threading
from contextlib import contextmanager
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class COMConcurrencyManager:
    """Manages concurrent access to Office COM automation.

    Office COM automation (Word, Excel, PowerPoint) has limitations when
    multiple threads try to create instances simultaneously. This manager
    uses a semaphore to limit concurrent COM operations.
    """

    def __init__(self, max_concurrent: int = 3):
        """Initialize the COM concurrency manager.

        Args:
            max_concurrent: Maximum number of concurrent COM operations allowed.
                           Default is 3, which is a safe number for most systems.
        """
        self._semaphore = threading.Semaphore(max_concurrent)
        self._max_concurrent = max_concurrent
        self._active_count = 0
        self._count_lock = threading.Lock()
        logger.info("COM Concurrency Manager initialized: max_concurrent=%d", max_concurrent)
    
    @contextmanager
    def acquire(self, operation_name: str = "COM operation"):
        """Context manager to acquire COM access.

        Usage:
            with com_manager.acquire("Creating PowerPoint"):
                # Your COM code here
                app = win32com.client.Dispatch("PowerPoint.Application")
                ...

        Args:
            operation_name: Name of the operation for logging purposes
        """
        logger.debug("Waiting to acquire COM semaphore for: %s", operation_name)
        self._semaphore.acquire()
        try:
            with self._count_lock:
                self._active_count += 1
            logger.debug("COM semaphore acquired for: %s (active: %d/%d)",
                        operation_name, self._active_count, self._max_concurrent)
            yield
        finally:
            with self._count_lock:
                self._active_count -= 1
            self._semaphore.release()
            logger.debug("COM semaphore released for: %s (active: %d/%d)",
                        operation_name, self._active_count, self._max_concurrent)

    def get_active_count(self) -> int:
        """Get the current number of active COM operations.

        Returns:
            Number of currently active COM operations
        """
        with self._count_lock:
            return self._active_count

    def get_max_concurrent(self) -> int:
        """Get the maximum number of concurrent COM operations allowed.

        Returns:
            Maximum concurrent operations limit
        """
        return self._max_concurrent

    def is_at_capacity(self) -> bool:
        """Check if the manager is at maximum capacity.

        Returns:
            True if at capacity, False if slots are available
        """
        with self._count_lock:
            return self._active_count >= self._max_concurrent


# Global COM manager instance
# Limit to 2 concurrent COM operations (configured for system capacity)
com_manager = COMConcurrencyManager(max_concurrent=2)
