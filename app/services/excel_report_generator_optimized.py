"""Optimized service for generating NW Excel reports with parallel execution."""

from __future__ import annotations

import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from app.services.nw_reports_service import nw_reports_service
from app.utils.logging_utils import get_logger
from app.utils.nw_data_utils import sanitize_filename
from app.utils.path_utils import get_nw_downloads_dir

logger = get_logger(__name__)


class ExcelReportGeneratorOptimized:
    """Optimized generator for NW Excel reports with parallel stored procedure execution."""

    def __init__(self):
        """Initialize the Excel report generator."""
        self.workbook: Optional[Workbook] = None

    def generate_excel_report(
        self,
        presentation_id: int,
        project_name: str,
        display_name: str,
        include_votes: bool = False,
        include_participants: bool = False,
    ) -> Path:
        """Generate complete Excel report for NW presentation (OPTIMIZED VERSION).

        This optimized version:
        1. Executes stored procedures in parallel using ThreadPoolExecutor
        2. Reduces total generation time by running independent SPs concurrently
        3. Adds detailed timing logs to identify bottlenecks

        Args:
            presentation_id: Presentation ID to generate report for
            project_name: Project name
            display_name: Display name for the presentation
            include_votes: Whether to include votes sheet
            include_participants: Whether to include participants sheet

        Returns:
            Path to generated Excel file
        """
        start_time = time.time()
        logger.info(
            "Starting OPTIMIZED Excel report generation for presentation_id=%d",
            presentation_id,
        )

        # Check if presentation has participant voting enabled
        has_participants = nw_reports_service.check_has_participants(presentation_id)
        logger.info(
            "Presentation %d has participant voting: %s",
            presentation_id,
            has_participants
        )

        # Auto-detection logic
        if include_votes is False and has_participants == 1:
            include_votes = True
        if include_participants is False and has_participants == 1:
            include_participants = True

        # Create new workbook
        self.workbook = Workbook()

        # Remove default sheet
        if "Sheet" in self.workbook.sheetnames:
            del self.workbook["Sheet"]

        # Define all sheets to generate
        core_sheets = [
            ("Retained Names", "nw_dlRetainedNames_withRecraft"),
            ("Newly Created Names", "nw_CombineNewNames"),
            ("Roots Or Concepts To Explore", "nw_dlRootsOrConceptsToExplore"),
            ("Roots Or Concepts To Avoid", "nw_dlRootsOrConceptsToAvoid"),
            ("Notes", "nw_dlOpenNotes"),
        ]

        optional_sheets = []
        if include_votes:
            optional_sheets.append(("NW_Votes", "nw_Votesbygroups"))
        if include_participants:
            optional_sheets.append(("Participants", "nw_VotedParticipants"))

        all_sheets = core_sheets + optional_sheets

        # OPTIMIZATION 1: Execute all stored procedures in parallel
        parallel_start = time.time()
        logger.info("Fetching data from %d stored procedures in parallel...", len(all_sheets))

        sheet_data = self._fetch_all_sheets_parallel(all_sheets, presentation_id)

        parallel_time = time.time() - parallel_start
        logger.info("Parallel fetch completed in %.2f seconds", parallel_time)

        # OPTIMIZATION 2: Create sheets from pre-fetched data
        creation_start = time.time()
        for sheet_name, sp_name in all_sheets:
            if sheet_name in sheet_data:
                columns, rows = sheet_data[sheet_name]
                self._create_sheet_from_data(sheet_name, columns, rows, sp_name, presentation_id)
            else:
                logger.warning("No data for sheet '%s', skipping", sheet_name)

        # Create Participant Votes sheet if needed
        if include_participants:
            self._create_participant_votes_sheet(presentation_id)

        creation_time = time.time() - creation_start
        logger.info("Sheet creation completed in %.2f seconds", creation_time)

        # Save workbook
        save_start = time.time()
        output_dir = get_nw_downloads_dir()
        output_dir.mkdir(parents=True, exist_ok=True)

        safe_display_name = sanitize_filename(display_name)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{safe_display_name}_{timestamp}.xlsx"
        output_path = output_dir / filename

        self.workbook.save(str(output_path))
        save_time = time.time() - save_start

        total_time = time.time() - start_time
        logger.info(
            "Excel report saved to: %s (Total: %.2fs, Fetch: %.2fs, Create: %.2fs, Save: %.2fs)",
            output_path, total_time, parallel_time, creation_time, save_time
        )

        return output_path

    def _fetch_all_sheets_parallel(
        self,
        sheets: List[Tuple[str, str]],
        presentation_id: int
    ) -> dict:
        """Fetch data from multiple stored procedures in parallel.

        Args:
            sheets: List of (sheet_name, sp_name) tuples
            presentation_id: Presentation ID

        Returns:
            Dictionary mapping sheet_name to (columns, rows) tuples
        """
        sheet_data = {}

        # Use ThreadPoolExecutor to run SPs in parallel
        # Max 5 concurrent threads to avoid overwhelming the database
        with ThreadPoolExecutor(max_workers=5) as executor:
            # Submit all tasks
            future_to_sheet = {
                executor.submit(
                    self._fetch_sheet_data,
                    sp_name,
                    presentation_id
                ): sheet_name
                for sheet_name, sp_name in sheets
            }

            # Collect results as they complete
            for future in as_completed(future_to_sheet):
                sheet_name = future_to_sheet[future]
                try:
                    columns, rows = future.result(timeout=60)  # 60 second timeout per SP
                    sheet_data[sheet_name] = (columns, rows)
                    logger.debug("Fetched data for sheet '%s': %d rows", sheet_name, len(rows) if rows else 0)
                except Exception as e:
                    logger.error("Error fetching data for sheet '%s': %s", sheet_name, e, exc_info=True)
                    # Store empty data so we can still create the sheet
                    sheet_data[sheet_name] = (["Error"], [])

        return sheet_data

    def _fetch_sheet_data(
        self,
        sp_name: str,
        presentation_id: int
    ) -> Tuple[List[str], List]:
        """Fetch data from a single stored procedure.

        Args:
            sp_name: Stored procedure name
            presentation_id: Presentation ID

        Returns:
            Tuple of (columns, rows)
        """
        from app.config.db import get_connection_scope

        sp_start = time.time()

        with get_connection_scope(timeout=60) as cursor:  # Increased timeout
            cursor.execute(
                f"{{CALL [BI_GUIDELINES].[dbo].[{sp_name}](?)}}",
                (presentation_id,)
            )

            # Handle multiple result sets
            columns = None
            rows = None
            result_set_num = 0

            while True:
                if cursor.description is not None:
                    result_set_num += 1
                    temp_columns = [column[0] for column in cursor.description]
                    temp_rows = cursor.fetchall()

                    if temp_rows or columns is None:
                        columns = temp_columns
                        rows = temp_rows

                if not cursor.nextset():
                    break

            if columns is None:
                logger.warning("No results from SP %s", sp_name)
                columns = ["No Data"]
                rows = []

            if rows is None:
                rows = []

        sp_time = time.time() - sp_start
        logger.debug("SP %s completed in %.2f seconds (%d rows)", sp_name, sp_time, len(rows))

        return columns, rows

    def _create_sheet_from_data(
        self,
        sheet_name: str,
        columns: List[str],
        rows: List,
        sp_name: str,
        presentation_id: int,
        expand_grouped_names: bool = True
    ) -> None:
        """Create a sheet from pre-fetched data.

        Args:
            sheet_name: Name for the worksheet
            columns: Column names
            rows: Data rows
            sp_name: Stored procedure name (for logging)
            presentation_id: Presentation ID (for logging)
            expand_grouped_names: If True, split grouped names (with ## or $$) into separate rows
        """
        logger.debug("Creating sheet '%s' with %d rows", sheet_name, len(rows))

        ws = self.workbook.create_sheet(sheet_name)

        # Write headers
        self._write_header_row(ws, columns)

        if not rows:
            ws.cell(row=2, column=1, value="No data available")
            return

        # Find name column for expansion
        name_col_idx = None
        if expand_grouped_names:
            for idx, col in enumerate(columns):
                col_lower = col.lower() if col else ''
                if col_lower in ['name', 'newname', 'namestoexplore', 'namestoavoid']:
                    name_col_idx = idx
                    break

        # Write data rows with expansion if needed
        current_row = 2
        for row in rows:
            row_list = list(row)

            # Check if we need to expand this row
            if expand_grouped_names and name_col_idx is not None:
                name_value = row_list[name_col_idx]

                if name_value and isinstance(name_value, str) and ('##' in name_value or '$$' in name_value):
                    delimiter = '##' if '##' in name_value else '$$'
                    names = [n.strip() for n in name_value.split(delimiter)]
                    num_items = len(names)

                    # Split ALL columns that contain the same delimiter
                    split_columns = []
                    for col_idx, col_value in enumerate(row_list):
                        if col_value and isinstance(col_value, str) and delimiter in col_value:
                            parts = [p.strip() for p in col_value.split(delimiter)]
                            while len(parts) < num_items:
                                parts.append('')
                            split_columns.append((col_idx, parts))
                        else:
                            split_columns.append((col_idx, [col_value] * num_items))

                    # Write a row for each expanded item, but SKIP completely empty rows
                    for item_idx in range(num_items):
                        expanded_row = row_list.copy()

                        for col_idx, parts in split_columns:
                            value = parts[item_idx] if item_idx < len(parts) else ''
                            if isinstance(value, str):
                                value = value.strip()
                            expanded_row[col_idx] = value

                        # Skip if the name is empty
                        if not expanded_row[name_col_idx]:
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

        logger.debug("Sheet '%s' created with %d rows", sheet_name, current_row - 2)

    def _create_participant_votes_sheet(self, presentation_id: int) -> None:
        """Create Participant Votes sheet (optimized version).

        Args:
            presentation_id: Presentation ID
        """
        logger.debug("Creating Participant Votes sheet for presentation_id=%d", presentation_id)

        ws = self.workbook.create_sheet("Participant Votes")

        try:
            from app.config.db import get_connection_scope

            with get_connection_scope(timeout=60) as cursor:  # Increased timeout
                try:
                    cursor.execute(
                        "{CALL [BI_GUIDELINES].[dbo].[nw_ParticipantVotesByName](?)}",
                        (presentation_id,)
                    )

                    if cursor.description is None:
                        logger.warning("No results from nw_ParticipantVotesByName SP")
                        self._write_header_row(ws, ["Participant", "Name", "Vote"])
                        ws.cell(row=2, column=1, value="No data available")
                        return

                    columns = [column[0] for column in cursor.description]
                    rows = cursor.fetchall()

                    if not rows:
                        logger.warning("nw_ParticipantVotesByName returned no rows")
                        self._write_header_row(ws, ["Participant", "Name", "Vote"])
                        ws.cell(row=2, column=1, value="No data available")
                        return

                    # Convert to dict and create matrix
                    data = [dict(zip(columns, row)) for row in rows]

                    # Get unique participants and names
                    participants = []
                    participant_ids = set()
                    for item in data:
                        pid = item.get('ParticipantId') or item.get('participant_id')
                        if pid and pid not in participant_ids:
                            participant_ids.add(pid)
                            participants.append({
                                'id': pid,
                                'name': item.get('ParticipantName') or item.get('participant_name') or f"Participant {pid}"
                            })

                    names = list(set(item.get('Name') or item.get('name') or '' for item in data if item.get('Name') or item.get('name')))
                    names.sort()

                    # Create header row
                    headers = ['Participant'] + names
                    self._write_header_row(ws, headers)

                    # Create vote lookup
                    vote_lookup = {}
                    for item in data:
                        pid = item.get('ParticipantId') or item.get('participant_id')
                        name = item.get('Name') or item.get('name')
                        vote = item.get('Vote') or item.get('vote')

                        # Convert vote value
                        if vote is not None:
                            if isinstance(vote, (int, float)):
                                if vote == 1:
                                    vote = 'Positive'
                                elif vote == 0:
                                    vote = 'Neutral'
                                elif vote == -1:
                                    vote = 'Negative'

                        if pid and name:
                            vote_lookup[(pid, name)] = vote

                    # Write data rows
                    current_row = 2
                    for participant in participants:
                        ws.cell(row=current_row, column=1, value=participant['name'])

                        for col_idx, name in enumerate(names, start=2):
                            vote = vote_lookup.get((participant['id'], name), '')
                            ws.cell(row=current_row, column=col_idx, value=vote if vote else '')

                        current_row += 1

                    # Auto-size columns
                    self._auto_size_columns(ws)

                    logger.info("Participant Votes sheet created with %d participants and %d names",
                               len(participants), len(names))

                except Exception as e:
                    logger.error("Error creating Participant Votes sheet: %s", str(e))
                    self._write_header_row(ws, ["Message"])
                    ws.cell(row=2, column=1, value=f"Error creating participant votes: {str(e)}")

        except Exception as e:
            logger.error("Error in _create_participant_votes_sheet: %s", str(e), exc_info=True)
            self._write_header_row(ws, ["Error"])
            ws.cell(row=2, column=1, value=f"Error: {str(e)}")

    def _write_header_row(self, ws, headers: list) -> None:
        """Write and format header row."""
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")

        for col_num, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

    def _auto_size_columns(self, ws, max_width: int = 100) -> None:
        """Auto-size columns based on content."""
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
excel_report_generator_optimized = ExcelReportGeneratorOptimized()
