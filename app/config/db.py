"""Database configuration and helper utilities."""

from __future__ import annotations

from contextlib import contextmanager
from typing import Generator, Optional

import pyodbc

from app.config.settings import settings
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class DatabaseConnectionError(Exception):
    """Raised when a database connection cannot be established."""


class DatabaseTransactionError(Exception):
    """Raised when a database transaction fails."""


def get_connection_string() -> str:
    """Return the SQL Server connection string from settings."""
    return settings.sql_connection_string


def create_connection(timeout: Optional[int] = None) -> pyodbc.Connection:
    """Create and return a new database connection.

    Args:
        timeout: Optional timeout (in seconds) for the connection attempt.

    Raises:
        DatabaseConnectionError: If the connection cannot be established.

    Returns:
        An open ``pyodbc.Connection`` instance.
    """
    connection_string = get_connection_string()
    try:
        kwargs = {}
        if timeout is not None:
            kwargs["timeout"] = timeout

        connection = pyodbc.connect(  # pylint: disable=c-extension-no-member
            connection_string,
            **kwargs,
        )
        logger.debug("Database connection established.")
        return connection
    except pyodbc.Error as exc:  # pylint: disable=c-extension-no-member
        logger.error("Failed to establish database connection: %s", exc)
        raise DatabaseConnectionError("Could not connect to database") from exc


@contextmanager
def get_connection_scope(
    *,
    autocommit: bool = False,
    timeout: Optional[int] = None,
) -> Generator[pyodbc.Cursor, None, None]:
    """Provide a cursor within a managed connection scope.

    This context manager handles connection lifecycle, transactions,
    and proper cleanup automatically.

    Args:
        autocommit: Whether to commit automatically when exiting the context.
                   If False, will commit on success or rollback on error.
        timeout: Optional timeout (in seconds) for establishing the connection.

    Yields:
        A ``pyodbc.Cursor`` ready for queries.

    Raises:
        DatabaseConnectionError: If the connection cannot be established.
        DatabaseTransactionError: If the transaction fails.

    Example:
        >>> with get_connection_scope() as cursor:
        ...     cursor.execute("SELECT * FROM table")
        ...     results = cursor.fetchall()
    """
    connection: Optional[pyodbc.Connection] = None
    cursor: Optional[pyodbc.Cursor] = None
    
    try:
        connection = create_connection(timeout=timeout)
        connection.autocommit = autocommit
        cursor = connection.cursor()
        
        yield cursor
        
        # Commit if not in autocommit mode and no exception occurred
        if not autocommit and connection:
            connection.commit()
            logger.debug("Database transaction committed successfully.")
            
    except pyodbc.Error as exc:  # pylint: disable=c-extension-no-member
        # Rollback on error if not in autocommit mode
        if connection and not autocommit:
            connection.rollback()
            logger.warning("Database transaction rolled back due to error.")
        
        logger.error("Database operation failed: %s", exc, exc_info=True)
        raise DatabaseTransactionError(f"Database operation failed: {str(exc)}") from exc
        
    finally:
        # Always close cursor and connection
        if cursor:
            cursor.close()
        if connection:
            connection.close()
            logger.debug("Database connection closed.")


@contextmanager
def get_db_connection(
    *,
    autocommit: bool = False,
    timeout: Optional[int] = None,
) -> Generator[pyodbc.Connection, None, None]:
    """Provide a managed database connection.

    This context manager yields the connection directly instead of a cursor,
    allowing for more control when needed.

    Args:
        autocommit: Whether to commit automatically when exiting the context.
        timeout: Optional timeout (in seconds) for establishing the connection.

    Yields:
        An open ``pyodbc.Connection`` instance.

    Raises:
        DatabaseConnectionError: If the connection cannot be established.
        DatabaseTransactionError: If the transaction fails.

    Example:
        >>> with get_db_connection() as conn:
        ...     cursor = conn.cursor()
        ...     cursor.execute("SELECT * FROM table")
    """
    connection: Optional[pyodbc.Connection] = None
    
    try:
        connection = create_connection(timeout=timeout)
        connection.autocommit = autocommit
        
        yield connection
        
        # Commit if not in autocommit mode and no exception occurred
        if not autocommit and connection:
            connection.commit()
            logger.debug("Database transaction committed successfully.")
            
    except pyodbc.Error as exc:  # pylint: disable=c-extension-no-member
        # Rollback on error if not in autocommit mode
        if connection and not autocommit:
            connection.rollback()
            logger.warning("Database transaction rolled back due to error.")
        
        logger.error("Database operation failed: %s", exc, exc_info=True)
        raise DatabaseTransactionError(f"Database operation failed: {str(exc)}") from exc
        
    finally:
        if connection:
            connection.close()
            logger.debug("Database connection closed.")


def test_connection(timeout: Optional[int] = 5) -> bool:
    """Test if a database connection can be established.
    
    Args:
        timeout: Connection timeout in seconds.
        
    Returns:
        True if connection successful, False otherwise.
    """
    try:
        connection = create_connection(timeout=timeout)
        connection.close()
        logger.info("Database connection test successful.")
        return True
    except DatabaseConnectionError:
        logger.warning("Database connection test failed.")
        return False
