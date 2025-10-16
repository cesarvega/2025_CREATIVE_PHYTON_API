"""Service for generating NW Excel reports."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from app.services.nw_reports_service import nw_reports_service
from app.utils.logging_utils import get_logger
from app.utils.nw_data_utils import (
    convert_vote_to_text,
    sanitize_filename,
    split_grouped_names,
)
from app.utils.path_utils import get_nw_downloads_dir

logger = get_logger(__name__)


class ExcelReportGenerator:
    """Generator for NW Excel reports."""

    def __init__(self):
        """Initialize the Excel report generator."""
        self.workbook: Optional[Workbook] = None

    def _create_sheet_from_sp(
        self,
        sheet_name: str,
        sp_name: str,
        presentation_id: int,
        expand_grouped_names: bool = True
    ) -> None:
        """Create a sheet directly from stored procedure results.

        Args:
            sheet_name: Name for the worksheet
            sp_name: Name of the stored procedure
            presentation_id: Presentation ID to pass to SP
            expand_grouped_names: If True, split grouped names (with ## or $$) into separate rows
        """
        logger.debug("Creating sheet '%s' from SP '%s'", sheet_name, sp_name)

        ws = self.workbook.create_sheet(sheet_name)

        try:
            from app.config.db import get_connection_scope

            with get_connection_scope(timeout=30) as cursor:
                # Execute stored procedure
                cursor.execute(
                    f"{{CALL [BI_GUIDELINES].[dbo].[{sp_name}](?)}}",
                    (presentation_id,)
                )

                # Some SPs return multiple result sets (e.g., nw_CombineNewNames)
                # We need to iterate through all result sets and use the last one with data
                columns = None
                rows = None
                result_set_num = 0

                while True:
                    if cursor.description is not None:
                        # This result set has columns, save it
                        result_set_num += 1
                        temp_columns = [column[0] for column in cursor.description]
                        temp_rows = cursor.fetchall()

                        # Only use result sets that have data OR if we have no data yet
                        if temp_rows or columns is None:
                            columns = temp_columns
                            rows = temp_rows
                            logger.debug("SP %s result set %d: %d columns, %d rows",
                                       sp_name, result_set_num, len(columns), len(rows))

                    # Try to move to next result set
                    if not cursor.nextset():
                        break

                # Check if we have results
                if columns is None or cursor.description is None and result_set_num == 0:
                    logger.warning("No results from SP %s", sp_name)
                    headers = ["No Data"]
                    self._write_header_row(ws, headers)
                    return

                if not rows:
                    logger.warning("SP %s returned no rows for presentation_id=%d", sp_name, presentation_id)
                    self._write_header_row(ws, columns)
                    # Write a note that no data was found
                    ws.cell(row=2, column=1, value="No data available")
                    return

                logger.info("SP %s returned %d rows", sp_name, len(rows))

                # Write headers
                self._write_header_row(ws, columns)

                # Find columns that might contain delimited data
                # We'll look for the primary "Name" column first
                name_col_idx = None
                for idx, col in enumerate(columns):
                    col_lower = col.lower() if col else ''
                    if col_lower in ['name', 'newname', 'namestoexplore', 'namestoavoid']:
                        name_col_idx = idx
                        logger.debug("Found name column '%s' at index %d", col, idx)
                        break

                if name_col_idx is None:
                    logger.warning("No name column found in columns: %s", columns)

                # Write data rows with expansion if needed
                current_row = 2
                for row in rows:
                    # Convert row to list for easier manipulation
                    row_list = list(row)

                    # Check if we need to expand this row
                    if expand_grouped_names and name_col_idx is not None:
                        name_value = row_list[name_col_idx]

                        # Check if name contains group delimiters
                        if name_value and isinstance(name_value, str) and ('##' in name_value or '$$' in name_value):
                            # Determine delimiter
                            delimiter = '##' if '##' in name_value else '$$'

                            # Split the primary name column and filter out empty values
                            names = [n.strip() for n in name_value.split(delimiter)]

                            # Count how many names we have (including empty strings for positioning)
                            num_items = len(names)

                            logger.debug("Expanding row with delimiter '%s' into %d items", delimiter, num_items)

                            # Now split ALL columns that contain the same delimiter
                            split_columns = []
                            for col_idx, col_value in enumerate(row_list):
                                if col_value and isinstance(col_value, str) and delimiter in col_value:
                                    # Split this column
                                    parts = [p.strip() for p in col_value.split(delimiter)]
                                    # Pad with empty strings if needed to match num_items
                                    while len(parts) < num_items:
                                        parts.append('')
                                    split_columns.append((col_idx, parts))
                                else:
                                    # This column doesn't have delimiter, will be repeated
                                    split_columns.append((col_idx, [col_value] * num_items))

                            # Write a row for each expanded item, but SKIP completely empty rows
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

                                # Write the expanded row
                                for col_num, value in enumerate(expanded_row, start=1):
                                    ws.cell(row=current_row, column=col_num, value=value)

                                current_row += 1
                        else:
                            # No grouping, write normally
                            for col_num, value in enumerate(row_list, start=1):
                                ws.cell(row=current_row, column=col_num, value=value)
                            current_row += 1
                    else:
                        # No expansion needed, write normally
                        for col_num, value in enumerate(row_list, start=1):
                            ws.cell(row=current_row, column=col_num, value=value)
                        current_row += 1

                # Auto-size columns
                self._auto_size_columns(ws)

                logger.info("Sheet '%s' created with %d rows (expanded)", sheet_name, current_row - 2)

        except Exception as e:
            logger.error("Error creating sheet '%s' from SP '%s': %s", sheet_name, sp_name, e, exc_info=True)
            # Create empty sheet on error
            headers = ["Error"]
            self._write_header_row(ws, headers)
            ws.cell(row=2, column=1, value=f"Error: {str(e)}")

    def generate_excel_report(
        self,
        presentation_id: int,
        project_name: str,
        display_name: str,
        include_votes: bool = False,
        include_participants: bool = False,
    ) -> Path:
        """Generate complete Excel report for NW presentation.

        Args:
            presentation_id: Presentation ID to generate report for
            project_name: Project name
            display_name: Display name for the presentation
            include_votes: Whether to include votes sheet (if None, auto-detect)
            include_participants: Whether to include participants sheet (if None, auto-detect)

        Returns:
            Path to generated Excel file
        """
        logger.info(
            "Starting Excel report generation for presentation_id=%d",
            presentation_id,
        )

        # Check if presentation has participant voting enabled
        # This determines if we create 5 or 7 sheets
        has_participants = nw_reports_service.check_has_participants(presentation_id)
        logger.info(
            "Presentation %d has participant voting: %s",
            presentation_id,
            has_participants
        )

        # If include_votes/include_participants not explicitly set, use auto-detection
        if include_votes is False and has_participants == 1:
            include_votes = True
        if include_participants is False and has_participants == 1:
            include_participants = True

        # Create new workbook
        self.workbook = Workbook()

        # Remove default sheet
        if "Sheet" in self.workbook.sheetnames:
            del self.workbook["Sheet"]

        # Generate each sheet directly from stored procedures
        # This preserves the exact column structure from the database
        # Sheets 1-5: Always included
        self._create_sheet_from_sp("Retained Names", "nw_dlRetainedNames_withRecraft", presentation_id)
        self._create_sheet_from_sp("Newly Created Names", "nw_CombineNewNames", presentation_id)
        self._create_sheet_from_sp("Roots Or Concepts To Explore", "nw_dlRootsOrConceptsToExplore", presentation_id)
        self._create_sheet_from_sp("Roots Or Concepts To Avoid", "nw_dlRootsOrConceptsToAvoid", presentation_id)
        self._create_sheet_from_sp("Notes", "nw_dlOpenNotes", presentation_id)

        # Sheets 6-7: Only if participant voting is enabled
        if include_votes:
            self._create_sheet_from_sp("NW_Votes", "nw_Votesbygroups", presentation_id)

        if include_participants:
            self._create_sheet_from_sp("Participants", "nw_VotedParticipants", presentation_id)

        # Save workbook to NW downloads directory
        # Format: [DisplayName].xlsx
        output_dir = get_nw_downloads_dir()
        output_dir.mkdir(parents=True, exist_ok=True)

        safe_display_name = sanitize_filename(display_name)
        filename = f"{safe_display_name}.xlsx"
        output_path = output_dir / filename

        self.workbook.save(str(output_path))
        logger.info("Excel report saved to: %s", output_path)

        return output_path

    def _create_retained_names_sheet(self, presentation_id: int) -> None:
        """Create 'Retained Names' sheet."""
        logger.debug("Creating Retained Names sheet")

        ws = self.workbook.create_sheet("Retained Names")

        # Get data directly from SP without processing
        retained_names_data = nw_reports_service.get_retained_names(presentation_id)

        if not retained_names_data:
            # Create empty sheet with just headers
            headers = ["Name"]
            self._write_header_row(ws, headers)
            logger.info("Retained Names sheet created (empty - no data)")
            return

        # Since we don't know the exact column structure from SP,
        # we'll write the data as-is from the stored procedure
        # The SP should return all columns needed
        try:
            with nw_reports_service._get_raw_sp_data(presentation_id, "nw_dlRetainedNames_withRecraft") as (columns, rows):
                if not columns or not rows:
                    headers = ["Name"]
                    self._write_header_row(ws, headers)
                    logger.info("Retained Names sheet created (empty)")
                    return

                # Write headers from SP columns
                self._write_header_row(ws, columns)

                # Write data rows
                for row_num, row in enumerate(rows, start=2):
                    for col_num, value in enumerate(row, start=1):
                        ws.cell(row=row_num, column=col_num, value=value)

                # Auto-size columns
                self._auto_size_columns(ws)

                logger.info("Retained Names sheet created with %d rows", len(rows))

        except Exception as e:
            logger.error("Error creating Retained Names sheet: %s", e)
            # Fallback: create sheet with basic headers
            headers = ["Name"]
            self._write_header_row(ws, headers)
            logger.info("Retained Names sheet created (fallback due to error)")

    def _create_newly_created_names_sheet(self, presentation_id: int) -> None:
        """Create 'Newly Created Names' sheet."""
        logger.debug("Creating Newly Created Names sheet")

        ws = self.workbook.create_sheet("Newly Created Names")

        # Headers
        headers = ["Name", "Category", "Rationale"]
        self._write_header_row(ws, headers)

        # Get data
        new_names = nw_reports_service.get_newly_created_names(presentation_id)

        # Write data rows
        row_num = 2
        for item in new_names:
            # Handle grouped names
            names = split_grouped_names(item.name)

            for name in names:
                ws.cell(row=row_num, column=1, value=name)
                ws.cell(row=row_num, column=2, value=item.category or "")
                ws.cell(row=row_num, column=3, value=item.rationale or "")
                row_num += 1

        # Auto-size columns
        self._auto_size_columns(ws)

        logger.info("Newly Created Names sheet created with %d rows", row_num - 2)

    def _create_roots_to_explore_sheet(self, presentation_id: int) -> None:
        """Create 'Roots Or Concepts To Explore' sheet."""
        logger.debug("Creating Roots To Explore sheet")

        ws = self.workbook.create_sheet("Roots Or Concepts To Explore")

        # Headers
        headers = ["Concept", "Description"]
        self._write_header_row(ws, headers)

        # Get data
        concepts = nw_reports_service.get_roots_to_explore(presentation_id)

        # Write data rows
        row_num = 2
        for item in concepts:
            ws.cell(row=row_num, column=1, value=item.concept)
            ws.cell(row=row_num, column=2, value=item.description or "")
            row_num += 1

        # Auto-size columns
        self._auto_size_columns(ws)

        logger.info("Roots To Explore sheet created with %d rows", row_num - 2)

    def _create_roots_to_avoid_sheet(self, presentation_id: int) -> None:
        """Create 'Roots Or Concepts To Avoid' sheet."""
        logger.debug("Creating Roots To Avoid sheet")

        ws = self.workbook.create_sheet("Roots Or Concepts To Avoid")

        # Headers
        headers = ["Concept", "Description"]
        self._write_header_row(ws, headers)

        # Get data
        concepts = nw_reports_service.get_roots_to_avoid(presentation_id)

        # Write data rows
        row_num = 2
        for item in concepts:
            ws.cell(row=row_num, column=1, value=item.concept)
            ws.cell(row=row_num, column=2, value=item.description or "")
            row_num += 1

        # Auto-size columns
        self._auto_size_columns(ws)

        logger.info("Roots To Avoid sheet created with %d rows", row_num - 2)

    def _create_notes_sheet(self, presentation_id: int) -> None:
        """Create 'Notes' sheet."""
        logger.debug("Creating Notes sheet")

        ws = self.workbook.create_sheet("Notes")

        # Headers
        headers = ["Note", "Created By", "Created Date"]
        self._write_header_row(ws, headers)

        # Get data
        notes = nw_reports_service.get_open_notes(presentation_id)

        # Write data rows
        row_num = 2
        for item in notes:
            ws.cell(row=row_num, column=1, value=item.note_text)
            ws.cell(row=row_num, column=2, value=item.created_by or "")

            # Format date
            date_str = ""
            if item.created_date:
                if isinstance(item.created_date, datetime):
                    date_str = item.created_date.strftime("%Y-%m-%d %H:%M:%S")
                else:
                    date_str = str(item.created_date)

            ws.cell(row=row_num, column=3, value=date_str)
            row_num += 1

        # Auto-size columns
        self._auto_size_columns(ws)

        logger.info("Notes sheet created with %d rows", row_num - 2)

    def _create_votes_sheet(self, presentation_id: int) -> None:
        """Create 'NW_Votes' sheet."""
        logger.debug("Creating Votes sheet")

        ws = self.workbook.create_sheet("NW_Votes")

        # Headers
        headers = ["Group", "Positive", "Neutral", "Negative", "Total"]
        self._write_header_row(ws, headers)

        # Get data
        votes = nw_reports_service.get_votes_by_groups(presentation_id)

        # Write data rows
        row_num = 2
        for item in votes:
            ws.cell(row=row_num, column=1, value=item.group_name)
            ws.cell(row=row_num, column=2, value=item.positive_votes)
            ws.cell(row=row_num, column=3, value=item.neutral_votes)
            ws.cell(row=row_num, column=4, value=item.negative_votes)
            ws.cell(row=row_num, column=5, value=item.total_votes)
            row_num += 1

        # Auto-size columns
        self._auto_size_columns(ws)

        logger.info("Votes sheet created with %d rows", row_num - 2)

    def _create_participants_sheet(self, presentation_id: int) -> None:
        """Create 'Participants' sheet."""
        logger.debug("Creating Participants sheet")

        ws = self.workbook.create_sheet("Participants")

        # Headers
        headers = ["Participant Name", "Email", "Voted Date"]
        self._write_header_row(ws, headers)

        # Get data
        participants = nw_reports_service.get_voted_participants(presentation_id)

        # Write data rows
        row_num = 2
        for item in participants:
            ws.cell(row=row_num, column=1, value=item.participant_name)
            ws.cell(row=row_num, column=2, value=item.email or "")

            # Format date
            date_str = ""
            if item.voted_date:
                if isinstance(item.voted_date, datetime):
                    date_str = item.voted_date.strftime("%Y-%m-%d %H:%M:%S")
                else:
                    date_str = str(item.voted_date)

            ws.cell(row=row_num, column=3, value=date_str)
            row_num += 1

        # Auto-size columns
        self._auto_size_columns(ws)

        logger.info("Participants sheet created with %d rows", row_num - 2)

    def _write_header_row(self, ws, headers: list) -> None:
        """Write and format header row.

        Args:
            ws: Worksheet object
            headers: List of header strings
        """
        # Header styling
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")

        for col_num, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

    def _auto_size_columns(self, ws, max_width: int = 100) -> None:
        """Auto-size columns based on content.

        Args:
            ws: Worksheet object
            max_width: Maximum column width
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
                except:
                    pass

            adjusted_width = min(max_length + 2, max_width)
            ws.column_dimensions[column_letter].width = adjusted_width


# Global service instance
excel_report_generator = ExcelReportGenerator()
