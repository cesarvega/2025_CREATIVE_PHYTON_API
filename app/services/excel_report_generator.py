"""Service for generating NW Excel reports."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Optional

from openpyxl import Workbook

from app.services.nw_reports_service import nw_reports_service
from app.utils.logging_utils import get_logger
from app.utils.nw_data_utils import sanitize_filename
from app.utils.path_utils import get_nw_downloads_dir
# OPTIMIZATION: Use centralized utilities to reduce code duplication
from app.utils.db_utils import execute_sp_multiple_results
from app.utils.excel_utils import write_header_row, auto_size_columns, write_data_rows, clean_report_data

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

        OPTIMIZED: Now uses db_utils and excel_utils for cleaner, reusable code.

        Args:
            sheet_name: Name for the worksheet
            sp_name: Name of the stored procedure
            presentation_id: Presentation ID to pass to SP
            expand_grouped_names: If True, split grouped names (with ## or $$) into separate rows
        """
        logger.debug("Creating sheet '%s' from SP '%s'", sheet_name, sp_name)

        ws = self.workbook.create_sheet(sheet_name)

        try:
            # OPTIMIZATION: Use centralized db_utils instead of manual SP execution
            columns, rows = execute_sp_multiple_results(
                f"[BI_GUIDELINES].[dbo].[{sp_name}]",
                (presentation_id,),
                timeout=30,
                return_all_resultsets=True
            )

            # Check if we have results
            if columns is None:
                logger.warning("No results from SP %s", sp_name)
                write_header_row(ws, ["No Data"])
                return

            # Clean report data (merge recraft, strip NameGroup prefix)
            columns, rows = clean_report_data(columns, rows)

            if not rows:
                logger.warning("SP %s returned no rows for presentation_id=%d", sp_name, presentation_id)
                write_header_row(ws, columns)
                ws.cell(row=2, column=1, value="No data available")
                return

            logger.info("SP %s returned %d rows", sp_name, len(rows))

            # OPTIMIZATION: Use excel_utils for headers, data writing, and auto-sizing
            write_header_row(ws, columns)
            rows_written = write_data_rows(ws, columns, rows, expand_grouped=expand_grouped_names)
            auto_size_columns(ws)

            logger.info("Sheet '%s' created with %d rows (expanded)", sheet_name, rows_written)

        except Exception as e:
            logger.error("Error creating sheet '%s' from SP '%s': %s", sheet_name, sp_name, e, exc_info=True)
            # Create empty sheet on error
            write_header_row(ws, ["Error"])
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

        # Sheets 6-8: Only if participant voting is enabled
        if include_votes:
            self._create_sheet_from_sp("NW_Votes", "nw_Votesbygroups", presentation_id)

        if include_participants:
            self._create_sheet_from_sp("Participants", "nw_VotedParticipants", presentation_id)
            self._create_participant_votes_sheet(presentation_id)

        # Save workbook to NW downloads directory
        # Format: [DisplayName]_[Timestamp].xlsx
        output_dir = get_nw_downloads_dir()
        output_dir.mkdir(parents=True, exist_ok=True)

        safe_display_name = sanitize_filename(display_name)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{safe_display_name}_{timestamp}.xlsx"
        output_path = output_dir / filename

        self.workbook.save(str(output_path))
        logger.info("Excel report saved to: %s", output_path)

        return output_path

    def _create_participant_votes_sheet(self, presentation_id: int) -> None:
        """Create Participant Votes sheet showing how each participant voted for each name.

        This creates a matrix with:
        - Rows: Participants
        - Columns: Names
        - Values: Positive, Neutral, Negative (or blank if no vote)

        Args:
            presentation_id: Presentation ID to generate votes for
        """
        logger.debug("Creating Participant Votes sheet for presentation_id=%d", presentation_id)

        ws = self.workbook.create_sheet("Participant Votes")

        try:
            from app.config.db import get_connection_scope

            with get_connection_scope(timeout=30) as cursor:
                # Try to call a stored procedure that returns participant votes
                # If it doesn't exist, we'll catch the error and create an informative message
                try:
                    cursor.execute(
                        "{CALL [BI_GUIDELINES].[dbo].[nw_ParticipantVotesByName](?)}",
                        (presentation_id,)
                    )

                    # Fetch results
                    if cursor.description is None:
                        logger.warning("No results from nw_ParticipantVotesByName SP")
                        write_header_row(ws, ["Participant", "Name", "Vote"])
                        ws.cell(row=2, column=1, value="No data available")
                        return

                    columns = [column[0] for column in cursor.description]
                    rows = cursor.fetchall()

                    if not rows:
                        logger.warning("nw_ParticipantVotesByName returned no rows")
                        write_header_row(ws, ["Participant", "Name", "Vote"])
                        ws.cell(row=2, column=1, value="No data available")
                        return

                    # Convert rows to dictionary format
                    data = []
                    for row in rows:
                        data.append(dict(zip(columns, row)))

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

                    # Create header row with participant name in first column and all names
                    headers = ['Participant'] + names
                    write_header_row(ws, headers)

                    # Create a lookup dictionary for quick access
                    vote_lookup = {}
                    for item in data:
                        pid = item.get('ParticipantId') or item.get('participant_id')
                        name = item.get('Name') or item.get('name')
                        # Fix: Use if/else to handle 0 value correctly (0 is falsy in Python)
                        vote = item.get('Vote') if 'Vote' in item else item.get('vote')

                        # Convert vote value to text if numeric
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
                    auto_size_columns(ws)

                    logger.info("Participant Votes sheet created with %d participants and %d names",
                               len(participants), len(names))

                except Exception as sp_error:
                    # If stored procedure doesn't exist, log and create a message
                    logger.warning("Stored procedure nw_ParticipantVotesByName not found or error: %s", str(sp_error))

                    # Try alternative: query the database directly for vote data
                    # This query structure is a best guess - may need adjustment based on actual schema
                    try:
                        query = """
                            SELECT
                                p.ParticipantId,
                                p.ParticipantName,
                                n.Name,
                                v.Vote
                            FROM [BI_GUIDELINES].[dbo].[NW_Participants] p
                            INNER JOIN [BI_GUIDELINES].[dbo].[NW_ParticipantVotes] v
                                ON p.ParticipantId = v.ParticipantId
                            INNER JOIN [BI_GUIDELINES].[dbo].[NW_Names] n
                                ON v.NameId = n.NameId
                            WHERE p.PresentationId = ?
                            ORDER BY p.ParticipantName, n.Name
                        """
                        cursor.execute(query, (presentation_id,))

                        if cursor.description is None:
                            raise Exception("Query returned no results")

                        columns = [column[0] for column in cursor.description]
                        rows = cursor.fetchall()

                        if not rows:
                            write_header_row(ws, ["Participant", "Name", "Vote"])
                            ws.cell(row=2, column=1, value="No participant votes found for this presentation")
                            return

                        # Process the query results similar to SP results above
                        data = []
                        for row in rows:
                            data.append(dict(zip(columns, row)))

                        # Get unique participants and names
                        participants = []
                        participant_ids = set()
                        for item in data:
                            pid = item.get('ParticipantId')
                            if pid and pid not in participant_ids:
                                participant_ids.add(pid)
                                participants.append({
                                    'id': pid,
                                    'name': item.get('ParticipantName') or f"Participant {pid}"
                                })

                        names = list(set(item.get('Name') or '' for item in data if item.get('Name')))
                        names.sort()

                        # Create header row
                        headers = ['Participant'] + names
                        write_header_row(ws, headers)

                        # Create vote lookup
                        vote_lookup = {}
                        for item in data:
                            pid = item.get('ParticipantId')
                            name = item.get('Name')
                            vote = item.get('Vote')

                            # Convert vote value to text if numeric
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
                        auto_size_columns(ws)

                        logger.info("Participant Votes sheet created with %d participants and %d names",
                                   len(participants), len(names))

                    except Exception as query_error:
                        logger.error("Error querying participant votes directly: %s", str(query_error))
                        write_header_row(ws, ["Message"])
                        ws.cell(row=2, column=1, value="Unable to retrieve participant votes. Please contact your database administrator to create the nw_ParticipantVotesByName stored procedure.")
                        ws.cell(row=3, column=1, value=f"Technical details: {str(sp_error)}")

        except Exception as e:
            logger.error("Error creating Participant Votes sheet: %s", str(e), exc_info=True)
            write_header_row(ws, ["Error"])
            ws.cell(row=2, column=1, value=f"Error: {str(e)}")

    # OPTIMIZATION: Removed _write_header_row and _auto_size_columns methods
    # Now using centralized excel_utils.write_header_row() and excel_utils.auto_size_columns()
    # This eliminates ~40 lines of duplicated code


# Global service instance
excel_report_generator = ExcelReportGenerator()
