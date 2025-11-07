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
            logger.debug("COM semaphore acquired for: %s", operation_name)
            yield
        finally:
            self._semaphore.release()
            logger.debug("COM semaphore released for: %s", operation_name)


# Global COM manager instance
# Limit to 3 concurrent COM operations (safe default for Office automation)
com_manager = COMConcurrencyManager(max_concurrent=3)
