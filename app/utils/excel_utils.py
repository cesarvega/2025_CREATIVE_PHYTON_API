"""Excel utility functions to reduce code duplication."""

from __future__ import annotations

from typing import Any, List

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

def merge_recraft_columns(
    columns: List[str],
    rows: List,
) -> tuple[List[str], List[list]]:
    """Merge TestnameImagePath data into recraft and remove TestnameImagePath column.

    Must be called BEFORE group expansion (expand_grouped_data).

    Two scenarios:

    Scenario A - TestnameImagePath has recraft data:
        TestnameImagePath contains per-name recraft values delimited by ##
        (e.g. "True##False##True"). Copy the string into the recraft column.
        If Name uses $$ delimiter, convert ## to $$ so the subsequent expansion
        splits correctly.

    Scenario B - TestnameImagePath is empty but Name is grouped:
        The recraft column has a single BIT value (e.g. "True") but Name is
        grouped ("A##B##C"). Replicate the recraft value for each name so that
        expansion produces one value per row (e.g. "True##True##True").

    Args:
        columns: Column names from the stored procedure.
        rows: Data rows (list of tuples/lists).

    Returns:
        Tuple of (cleaned_columns, cleaned_rows) with TestnameImagePath
        removed and its group recraft values transferred to the recraft column.
    """
    # Find column indices (case-insensitive)
    tip_idx = -1
    recraft_idx = -1
    name_idx = -1
    for idx, col in enumerate(columns):
        col_lower = col.lower() if col else ''
        if col_lower == 'testnameimagepath':
            tip_idx = idx
        elif col_lower == 'recraft':
            recraft_idx = idx
        elif col_lower in ('name', 'newname', 'namestoexplore', 'namestoavoid', 'candidate'):
            name_idx = idx

    # Nothing to do if TestnameImagePath is not present
    if tip_idx == -1:
        return columns, rows

    cleaned_rows = []
    for row in rows:
        row_list = list(row)

        if recraft_idx != -1:
            tip_value = row_list[tip_idx]
            name_value = row_list[name_idx] if name_idx != -1 else None

            # Detect the delimiter used by Name
            name_delimiter = None
            name_count = 1
            if name_value and isinstance(name_value, str):
                if '##' in name_value:
                    name_delimiter = '##'
                    name_count = len(name_value.split('##'))
                elif '$$' in name_value:
                    name_delimiter = '$$'
                    name_count = len(name_value.split('$$'))

            if tip_value and isinstance(tip_value, str) and tip_value.strip():
                tip_stripped = tip_value.strip()
                # Scenario A: TestnameImagePath has recraft data
                parts = [p.strip().lower() for p in tip_stripped.split('##')]
                is_recraft_data = all(p in ('true', 'false', '1', '0', '') for p in parts)
                if is_recraft_data:
                    # If Name uses $$, convert ## to $$ so expansion works
                    if name_delimiter == '$$':
                        tip_stripped = tip_stripped.replace('##', '$$')
                    row_list[recraft_idx] = tip_stripped
            elif name_delimiter and name_count > 1:
                # Scenario B: TestnameImagePath is empty but Name is grouped
                # Replicate the single recraft value for each name
                recraft_value = row_list[recraft_idx]
                recraft_str = str(recraft_value) if recraft_value is not None else 'False'
                row_list[recraft_idx] = name_delimiter.join([recraft_str] * name_count)

        # Remove TestnameImagePath column from the row
        row_list.pop(tip_idx)
        cleaned_rows.append(row_list)

    # Remove TestnameImagePath from column names
    cleaned_columns = list(columns)
    cleaned_columns.pop(tip_idx)

    return cleaned_columns, cleaned_rows


def _strip_namegroup_prefix(columns: List[str], rows: List[list]) -> List[list]:
    """Remove the sort-order prefix (e.g. 'A|', 'B|', 'C|') from NameGroup values.

    The database stores NameGroup as 'B|Prescreen Survivors' where the letter
    is used for ordering. This strips the prefix for cleaner report output.

    Args:
        columns: Column names.
        rows: Data rows (list of lists, modified in place).

    Returns:
        The same rows list with NameGroup values cleaned.
    """
    ng_idx = -1
    for idx, col in enumerate(columns):
        if col and col.lower() == 'namegroup':
            ng_idx = idx
            break

    if ng_idx == -1:
        return rows

    for row in rows:
        val = row[ng_idx]
        if val and isinstance(val, str) and '|' in val:
            row[ng_idx] = val.split('|', 1)[1]

    return rows


def _clean_orphaned_delimiters(columns: List[str], rows: List[list]) -> List[list]:
    """Clean ## or $$ delimiters from non-name columns when Name has no delimiter.

    This handles an edge case where individual slide names were split from a
    group but other columns (NameRanking, recraft, etc.) still contain the
    original grouped values with ## delimiters.  In that situation, keep only
    the first value from the delimited string.

    Args:
        columns: Column names.
        rows: Data rows (list of lists, modified in place).

    Returns:
        The same rows list with orphaned delimiters cleaned.
    """
    name_idx = find_column_index(
        columns,
        ['name', 'newname', 'namestoexplore', 'namestoavoid', 'candidate']
    )
    if name_idx == -1:
        return rows

    for row in rows:
        name_val = row[name_idx]
        # If Name has ## or $$, this is a proper group row - leave it for expansion
        if name_val and isinstance(name_val, str) and ('##' in name_val or '$$' in name_val):
            continue

        # Name has no delimiter, so clean any ## or $$ in other columns
        for col_idx, val in enumerate(row):
            if col_idx == name_idx:
                continue
            if val and isinstance(val, str):
                if '##' in val:
                    row[col_idx] = val.split('##')[0].strip()
                elif '$$' in val:
                    row[col_idx] = val.split('$$')[0].strip()

    return rows


def clean_report_data(
    columns: List[str],
    rows: List,
) -> tuple[List[str], List[list]]:
    """Apply all report data cleaning transformations.

    Combines all data cleaning steps that run BEFORE group expansion:
    1. Merge TestnameImagePath into recraft and remove the column
    2. Strip sort-order prefix from NameGroup (e.g. 'B|Name' -> 'Name')
    3. Clean orphaned ## delimiters in non-name columns

    Args:
        columns: Column names from the stored procedure.
        rows: Data rows (list of tuples/lists).

    Returns:
        Tuple of (cleaned_columns, cleaned_rows).
    """
    columns, rows = merge_recraft_columns(columns, rows)
    rows = _strip_namegroup_prefix(columns, rows)
    rows = _clean_orphaned_delimiters(columns, rows)

    return columns, rows


# Standard header styling constants
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")


def write_header_row(ws: Worksheet, headers: List[str]) -> None:
    """Write and format header row with consistent styling.

    OPTIMIZATION: Eliminates duplicated header formatting code across 10+ files.

    Args:
        ws: Worksheet object
        headers: List of header strings

    Example:
        >>> from openpyxl import Workbook
        >>> wb = Workbook()
        >>> ws = wb.active
        >>> write_header_row(ws, ["Name", "Email", "Status"])
    """
    for col_num, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT


def auto_size_columns(ws: Worksheet, max_width: int = 100, min_width: int = 8) -> None:
    """Auto-size columns based on content with min/max constraints.

    OPTIMIZATION: Reusable function used in 6+ places with identical logic.

    Args:
        ws: Worksheet object
        max_width: Maximum column width (default: 100)
        min_width: Minimum column width (default: 8)

    Example:
        >>> from openpyxl import Workbook
        >>> wb = Workbook()
        >>> ws = wb.active
        >>> ws['A1'] = "Very Long Header Text That Needs Auto Sizing"
        >>> auto_size_columns(ws)
    """
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)

        for cell in column:
            try:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            except Exception as e:
                logger.debug("Error calculating cell length: %s", e)
                pass

        adjusted_width = max(min(max_length + 2, max_width), min_width)
        ws.column_dimensions[column_letter].width = adjusted_width


def expand_grouped_data(
    row_list: List[Any],
    name_col_idx: int,
    columns: List[str]
) -> List[List[Any]]:
    """Expand rows that contain grouped data with ## or $$ delimiters.

    OPTIMIZATION: Centralizes the group expansion logic used in Excel and Word report generators.
    Reduces ~150 lines of duplicated code.

    Args:
        row_list: List of values from a single row
        name_col_idx: Index of the column containing the primary name
        columns: List of all column names

    Returns:
        List of expanded rows (may be just the original row if no grouping)

    Example:
        >>> row = ["Name1##Name2##Name3", "Value1##Value2##Value3", 100]
        >>> columns = ["Name", "Description", "Count"]
        >>> expanded = expand_grouped_data(row, 0, columns)
        >>> # Returns: [
        >>> #   ["Name1", "Value1", 100],
        >>> #   ["Name2", "Value2", 100],
        >>> #   ["Name3", "Value3", 100]
        >>> # ]
    """
    name_value = row_list[name_col_idx]

    # Check if name contains group delimiters
    if not (name_value and isinstance(name_value, str) and ('##' in name_value or '$$' in name_value)):
        # No grouping, return original row
        return [row_list]

    # Determine delimiter
    delimiter = '##' if '##' in name_value else '$$'
    alt_delimiter = '$$' if delimiter == '##' else '##'

    # Split the primary name column and filter out empty values
    names = [n.strip() for n in name_value.split(delimiter)]
    num_items = len(names)

    logger.debug("Expanding row with delimiter '%s' into %d items", delimiter, num_items)

    # Split ALL columns that contain the same delimiter OR the alternate delimiter
    split_columns = []
    for col_idx, col_value in enumerate(row_list):
        if col_value and isinstance(col_value, str):
            if delimiter in col_value:
                # Split this column by the primary delimiter
                parts = [p.strip() for p in col_value.split(delimiter)]
                while len(parts) < num_items:
                    parts.append('')
                split_columns.append((col_idx, parts))
            elif alt_delimiter in col_value:
                # Column uses the other delimiter (e.g. ## when Name uses $$)
                parts = [p.strip() for p in col_value.split(alt_delimiter)]
                while len(parts) < num_items:
                    parts.append('')
                split_columns.append((col_idx, parts))
            else:
                # No delimiter, repeat value
                split_columns.append((col_idx, [col_value] * num_items))
        else:
            # Non-string or empty, repeat value
            split_columns.append((col_idx, [col_value] * num_items))

    # Create expanded rows, SKIP completely empty rows
    expanded_rows = []
    for item_idx in range(num_items):
        expanded_row = row_list.copy()

        # Update all split columns with their respective values
        for col_idx, parts in split_columns:
            value = parts[item_idx] if item_idx < len(parts) else ''
            # Clean up the value (strip whitespace)
            if isinstance(value, str):
                value = value.strip()
            expanded_row[col_idx] = value

        # Check if the primary name column is empty
        # If the name is empty, skip this entire row
        if not expanded_row[name_col_idx]:
            logger.debug("Skipping empty row at item_idx=%d", item_idx)
            continue

        expanded_rows.append(expanded_row)

    return expanded_rows


def find_column_index(columns: List[str], possible_names: List[str]) -> int:
    """Find column index by checking multiple possible column names (case-insensitive).

    OPTIMIZATION: Common pattern for handling variations in column naming.

    Args:
        columns: List of column names
        possible_names: List of possible names to search for (case-insensitive)

    Returns:
        Index of the found column, or -1 if not found

    Example:
        >>> columns = ["UserName", "Email", "Status"]
        >>> idx = find_column_index(columns, ["username", "user_name", "name"])
        >>> # Returns: 0
    """
    for idx, col in enumerate(columns):
        col_lower = col.lower() if col else ''
        if col_lower in [name.lower() for name in possible_names]:
            return idx
    return -1


def write_data_rows(
    ws: Worksheet,
    columns: List[str],
    rows: List[Any],
    start_row: int = 2,
    expand_grouped: bool = True
) -> int:
    """Write data rows to worksheet with optional group expansion.

    OPTIMIZATION: Consolidates the common pattern of writing rows to Excel.

    Args:
        ws: Worksheet object
        columns: List of column names
        rows: List of row tuples from database
        start_row: Starting row number (default: 2, after header)
        expand_grouped: If True, expand rows with ## or $$ delimiters

    Returns:
        Number of rows written

    Example:
        >>> from openpyxl import Workbook
        >>> wb = Workbook()
        >>> ws = wb.active
        >>> columns = ["Name", "Value"]
        >>> rows = [("Alice", 100), ("Bob##Charlie", 200)]
        >>> count = write_data_rows(ws, columns, rows, expand_grouped=True)
        >>> # count = 3 (Alice, Bob, Charlie)
    """
    # Find name column for expansion
    name_col_idx = -1
    if expand_grouped:
        name_col_idx = find_column_index(
            columns,
            ['name', 'newname', 'namestoexplore', 'namestoavoid', 'candidate']
        )

    current_row = start_row
    for row in rows:
        row_list = list(row)

        # Check if we need to expand this row
        if expand_grouped and name_col_idx >= 0:
            expanded_rows = expand_grouped_data(row_list, name_col_idx, columns)
            for expanded_row in expanded_rows:
                for col_num, value in enumerate(expanded_row, start=1):
                    ws.cell(row=current_row, column=col_num, value=value)
                current_row += 1
        else:
            # No expansion needed, write normally
            for col_num, value in enumerate(row_list, start=1):
                ws.cell(row=current_row, column=col_num, value=value)
            current_row += 1

    return current_row - start_row  # Return number of rows written
