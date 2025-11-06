"""Database utility functions to reduce code duplication."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from app.config.db import get_connection_scope
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


def execute_sp_single_result(
    sp_name: str,
    params: Tuple[Any, ...],
    timeout: int = 30,
    use_daymaster: bool = False
) -> Optional[Dict[str, Any]]:
    """Execute a stored procedure and return a single result row as a dictionary.

    OPTIMIZATION: Reusable function to eliminate code duplication across services.
    Used in ~15+ places with identical logic.

    Args:
        sp_name: Fully qualified stored procedure name (e.g., "[BI_GUIDELINES].[dbo].[sp_name]")
        params: Tuple of parameters to pass to the SP
        timeout: Query timeout in seconds
        use_daymaster: If True, use DayMaster database connection

    Returns:
        Dictionary with column_name: value pairs, or None if no results

    Example:
        >>> result = execute_sp_single_result(
        ...     "[BI_GUIDELINES].[dbo].[nw_IsParticipantVoted]",
        ...     (presentation_id,)
        ... )
        >>> has_voted = result.get("HasVoted", False) if result else False
    """
    try:
        with get_connection_scope(timeout=timeout, use_daymaster=use_daymaster) as cursor:
            cursor.execute(f"{{CALL {sp_name}({', '.join(['?'] * len(params))})}}", params)

            row = cursor.fetchone()
            if row and cursor.description:
                columns = [column[0] for column in cursor.description]
                return dict(zip(columns, row))

            return None

    except Exception as e:
        logger.error("Error executing SP %s: %s", sp_name, str(e), exc_info=True)
        return None


def execute_sp_multiple_results(
    sp_name: str,
    params: Tuple[Any, ...],
    timeout: int = 30,
    use_daymaster: bool = False,
    return_all_resultsets: bool = False
) -> Tuple[Optional[List[str]], Optional[List[Any]]]:
    """Execute a stored procedure and return all result rows.

    OPTIMIZATION: Handles SPs that return multiple result sets (common pattern in this codebase).
    Reduces ~200 lines of duplicated code across services.

    Args:
        sp_name: Fully qualified stored procedure name
        params: Tuple of parameters to pass to the SP
        timeout: Query timeout in seconds
        use_daymaster: If True, use DayMaster database connection
        return_all_resultsets: If True, iterate through all result sets and return the last one with data

    Returns:
        Tuple of (columns, rows) where columns is list of column names and rows is list of row tuples

    Example:
        >>> columns, rows = execute_sp_multiple_results(
        ...     "[BI_GUIDELINES].[dbo].[nw_dlRetainedNames_withRecraft]",
        ...     (presentation_id,),
        ...     return_all_resultsets=True
        ... )
        >>> for row in rows:
        ...     data = dict(zip(columns, row))
    """
    try:
        with get_connection_scope(timeout=timeout, use_daymaster=use_daymaster) as cursor:
            cursor.execute(f"{{CALL {sp_name}({', '.join(['?'] * len(params))})}}", params)

            columns = None
            rows = None
            result_set_num = 0

            while True:
                if cursor.description is not None:
                    result_set_num += 1
                    temp_columns = [column[0] for column in cursor.description]
                    temp_rows = cursor.fetchall()

                    # Use result sets that have data OR if we have no data yet
                    if temp_rows or columns is None:
                        columns = temp_columns
                        rows = temp_rows
                        logger.debug(
                            "SP %s result set %d: %d columns, %d rows",
                            sp_name, result_set_num, len(columns), len(rows)
                        )

                    if not return_all_resultsets and temp_rows:
                        # Return first result set with data
                        break

                # Try to move to next result set
                if not cursor.nextset():
                    break

            if columns is None:
                logger.warning("No results from SP %s", sp_name)
                return None, None

            if rows is None:
                rows = []

            logger.debug("SP %s returned %d rows from %d result sets", sp_name, len(rows), result_set_num)
            return columns, rows

    except Exception as e:
        logger.error("Error executing SP %s: %s", sp_name, str(e), exc_info=True)
        return None, None


def rows_to_dicts(columns: List[str], rows: List[Any]) -> List[Dict[str, Any]]:
    """Convert rows to list of dictionaries using column names as keys.

    OPTIMIZATION: Common utility pattern used throughout the codebase.

    Args:
        columns: List of column names
        rows: List of row tuples

    Returns:
        List of dictionaries with column_name: value pairs

    Example:
        >>> columns = ["Name", "Value"]
        >>> rows = [("Alice", 100), ("Bob", 200)]
        >>> result = rows_to_dicts(columns, rows)
        >>> # [{"Name": "Alice", "Value": 100}, {"Name": "Bob", "Value": 200}]
    """
    if not columns or not rows:
        return []

    return [dict(zip(columns, row)) for row in rows]


def safe_get_first(items: List[Any], default: Any = None) -> Any:
    """Safely get the first item from a list or return default.

    OPTIMIZATION: Common pattern to avoid IndexError checks.

    Args:
        items: List to get first item from
        default: Default value if list is empty

    Returns:
        First item or default value
    """
    return items[0] if items else default


def execute_query(
    query: str,
    params: Optional[Tuple[Any, ...]] = None,
    timeout: int = 30,
    use_daymaster: bool = False,
    fetch_one: bool = False
) -> Tuple[Optional[List[str]], Optional[List[Any]]]:
    """Execute a raw SQL query (not a stored procedure).

    OPTIMIZATION: Centralized query execution with consistent error handling.

    Args:
        query: SQL query string
        params: Optional tuple of query parameters
        timeout: Query timeout in seconds
        use_daymaster: If True, use DayMaster database connection
        fetch_one: If True, fetch only one row

    Returns:
        Tuple of (columns, rows) or (columns, single_row) if fetch_one=True
    """
    try:
        with get_connection_scope(timeout=timeout, use_daymaster=use_daymaster) as cursor:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)

            if cursor.description is None:
                return None, None

            columns = [column[0] for column in cursor.description]

            if fetch_one:
                row = cursor.fetchone()
                return columns, row
            else:
                rows = cursor.fetchall()
                return columns, rows

    except Exception as e:
        logger.error("Error executing query: %s", str(e), exc_info=True)
        return None, None
