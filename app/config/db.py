"""Database configuration and helper utilities."""

from __future__ import annotations

import logging
from contextlib import contextmanager
from typing import Generator, Optional

import pyodbc

from app.config.settings import settings

logger = logging.getLogger(__name__)


class DatabaseConnectionError(Exception):
	"""Raised when a database connection cannot be established."""


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
		logger.exception("Failed to establish database connection: %s", exc)
		raise DatabaseConnectionError("Could not connect to database") from exc


@contextmanager
def get_connection_scope(
	*,
	autocommit: bool = False,
	timeout: Optional[int] = None,
) -> Generator[pyodbc.Cursor, None, None]:
	"""Provide a cursor within a managed connection scope.

	Args:
		autocommit: Whether to commit automatically when exiting the context.
		timeout: Optional timeout (in seconds) for establishing the connection.

	Yields:
		A ``pyodbc.Cursor`` ready for queries.

	Raises:
		DatabaseConnectionError: If the connection cannot be established.
	"""

	connection: Optional[pyodbc.Connection] = None
	try:
		connection = create_connection(timeout=timeout)
		connection.autocommit = autocommit
		cursor = connection.cursor()
		yield cursor
		if not autocommit:
			connection.commit()
	except pyodbc.Error as exc:  # pylint: disable=c-extension-no-member
		if connection and not autocommit:
			connection.rollback()
		logger.exception("Database operation failed: %s", exc)
		raise
	finally:
		if connection is not None:
			connection.close()


def test_connection(timeout: Optional[int] = 5) -> bool:
	"""Test if a database connection can be established."""

	try:
		connection = create_connection(timeout=timeout)
		connection.close()
		return True
	except DatabaseConnectionError:
		return False
