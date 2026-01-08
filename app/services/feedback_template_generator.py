"""Service for generating NW Feedback Template documents using InputDocumentRationales templates."""

from __future__ import annotations
import html
import re
import time
import pythoncom
import win32com.client
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from app.config.settings import settings
from app.constants import TemplateFilename
from app.models.nw_reports_models import SummaryType
from app.services.bi_guidelines_service import bi_guidelines_service
from app.services.nw_reports_service import nw_reports_service
from app.utils.logging_utils import get_logger
from app.utils.nw_data_utils import sanitize_filename
from app.utils.path_utils import get_nw_downloads_dir

logger = get_logger(__name__)

# Constants for Word COM
WD_REPLACE_ALL = 2
WD_FIND_CONTINUE = 1
WD_FORMAT_DOCUMENT = 0  # .doc format
WD_CELL = 12  # wdCell
WD_FORMAT_HTML = 8  # wdFormatHTML


def retry_com_operation(operation, max_retries=3, delay=0.5):
    """Retry COM operation if it fails due to busy application.
    
    Args:
        operation: Callable to execute
        max_retries: Maximum number of retry attempts
        delay: Delay in seconds between retries
        
    Returns:
        Result of the operation
        
    Raises:
        Last exception if all retries fail
    """
    last_exception = None
    for attempt in range(max_retries):
        try:
            return operation()
        except pythoncom.com_error as e:
            # Check if it's a "busy" error (-2147417846)
            if e.args[0] == -2147417846:
                last_exception = e
                if attempt < max_retries - 1:
                    logger.debug("COM busy (attempt %d/%d), retrying in %.1fs...", attempt + 1, max_retries, delay)
                    time.sleep(delay)
                else:
                    logger.error("COM operation failed after %d attempts", max_retries)
                    raise
            else:
                # Different error, don't retry
                raise
    
    if last_exception:
        raise last_exception


class FeedbackTemplateGenerator:
    """Generate NW Feedback Template documents from InputDocumentRationales templates."""

    def __init__(self):
        """Initialize the feedback template generator."""
        self.word_app: Optional[win32com.client.CDispatch] = None
        self.excel_app: Optional[win32com.client.CDispatch] = None
        self.template_path: Optional[Path] = None

    def generate_feedback_template(
        self,
        presentation_id: int,
        display_name: str,
        progress_callback=None,
    ) -> Path:
        """Generate complete Feedback Template document for NW presentation.

        This replicates the "Create Feedback Template" functionality from the original app.

        Steps:
        1. Determine presentation type (Normal, Phonetics, Katakana) and select template
        2. Copy template to output location
        3. Replace placeholders with values from nw_wdValuesToReplace
        4. Populate tables with data from nw_wdGetResults_Phonetics
        5. Generate and insert pie chart from ByTheNumbers data
        6. Clean up bookmarks
        7. Save document

        Args:
            presentation_id: Presentation ID to generate template for
            display_name: Display name for the presentation
            progress_callback: Optional callback function to report progress (int 0-100)

        Returns:
            Path to generated Word file
        """
        logger.info(
            "Starting Feedback Template generation for presentation_id=%d",
            presentation_id,
        )

        # Get presentation info to determine type (30-40%)
        if progress_callback:
            progress_callback(30)
        presentation_info = bi_guidelines_service.get_project_info(presentation_id)
        if not presentation_info:
            raise ValueError(f"Presentation with ID {presentation_id} not found")

        presentation_type = presentation_info.get("PresentationType", "Normal")
        logger.info("Presentation type: %s", presentation_type)

        # Determine template path based on presentation type (40-45%)
        if progress_callback:
            progress_callback(40)
        self.template_path = self._get_template_path(presentation_type)
        if not self.template_path.exists():
            raise FileNotFoundError(f"Feedback template not found: {self.template_path}")

        logger.info("Using template: %s", self.template_path)

        # Initialize COM for this thread
        pythoncom.CoInitialize()

        # Initialize Word application (45-50%)
        if progress_callback:
            progress_callback(45)
        
        doc = None
        try:
            # Initialize Word with proper COM handling
            logger.info("Initializing Word application...")
            from app.utils.com_manager import com_manager
            with com_manager.acquire("Word.Application Dispatch"):
                try:
                    self.word_app = win32com.client.gencache.EnsureDispatch("Word.Application")
                except Exception as e:
                    logger.warning("EnsureDispatch failed, trying Dispatch: %s", str(e))
                    self.word_app = win32com.client.Dispatch("Word.Application")

                self.word_app.Visible = False
                self.word_app.DisplayAlerts = 0  # wdAlertsNone
                logger.info("Word application initialized successfully")

            # Open template - ensure proper COM object handling
            template_path_str = str(self.template_path.absolute())
            logger.info("Opening template file: %s", template_path_str)
            
            # Use positional parameters to ensure proper COM dispatch
            # Parameters: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, etc.
            doc = self.word_app.Documents.Open(
                FileName=template_path_str,
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False
            )
            
            # Give Word time to fully load the document
            time.sleep(0.5)
            
            # Verify document opened successfully
            if not doc:
                raise RuntimeError("Documents.Open returned None")
            
            # Check if we got the actual document object (with retry for busy COM object)
            def verify_doc():
                doc_name = doc.Name
                logger.info("Template opened successfully: %s", doc_name)
                return doc_name
            
            try:
                retry_com_operation(verify_doc, max_retries=5, delay=1.0)
            except AttributeError as attr_err:
                logger.error("Failed to access document properties: %s", str(attr_err))
                logger.error("Object type: %s", type(doc))
                raise RuntimeError(f"Document object invalid. Got: {type(doc)}")


            # PHASE 1: Replace placeholders in document (50-70%)
            if progress_callback:
                progress_callback(50)
            logger.info("Phase 1: Replacing placeholders")
            replacements = nw_reports_service.get_word_report_replacements(presentation_id)
            logger.info("Retrieved %d replacement values", len(replacements))

            for replacement in replacements:
                if replacement.placeholder and replacement.value:
                    self._replace_text_in_doc(
                        doc,
                        replacement.placeholder,
                        replacement.value
                    )

            # Format header with correct font size and style
            logger.info("Formatting header")
            self._format_document_header(doc)

            # PHASE 2: Populate tables with results (70-80%)
            if progress_callback:
                progress_callback(70)
            logger.info("Phase 2: Populating tables")
            is_phonetics = presentation_type.lower() == "phonetics"
            self._populate_feedback_tables(doc, presentation_id, is_phonetics)

            # PHASE 3: Get "By The Numbers" data for pie chart (80-85%)
            if progress_callback:
                progress_callback(80)
            logger.info("Phase 3: Generating pie chart")
            by_the_numbers_data = nw_reports_service.get_word_report_results_phonetics(
                presentation_id,
                SummaryType.BY_THE_NUMBERS
            )

            # Generate and insert Pie Chart
            self._generate_and_insert_pie_chart(doc, by_the_numbers_data)

            # PHASE 4: Clean up bookmarks (85-90%)
            if progress_callback:
                progress_callback(85)
            logger.info("Phase 4: Cleaning up bookmarks")
            self._cleanup_bookmarks(doc)

            # Save document to output directory (90-95%)
            if progress_callback:
                progress_callback(90)
            output_dir = get_nw_downloads_dir()
            output_dir.mkdir(parents=True, exist_ok=True)

            safe_display_name = sanitize_filename(display_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Feedback_Template_{safe_display_name}_{timestamp}.doc"
            output_path = output_dir / filename

            # Save as Word 97-2003 format (.doc) with retry logic
            def save_document():
                doc.SaveAs2(str(output_path.absolute()), FileFormat=WD_FORMAT_DOCUMENT)
                logger.info("Document saved successfully")
            
            retry_com_operation(save_document, max_retries=5, delay=1.0)
            
            # Close document with retry
            def close_document():
                doc.Close(SaveChanges=False)
                logger.debug("Document closed after saving")
            
            retry_com_operation(close_document, max_retries=3, delay=0.5)
            doc = None  # Mark as closed

            logger.info("Feedback template saved to: %s", output_path)

            return output_path

        except Exception as e:
            logger.error("Error generating Feedback Template: %s", str(e), exc_info=True)
            raise

        finally:
            # Cleanup document and applications
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                    logger.debug("Document closed")
                except Exception as e:
                    logger.warning("Error closing document: %s", str(e))
            
            if self.excel_app:
                try:
                    self.excel_app.Quit()
                except Exception as e:
                    logger.warning("Error closing Excel application: %s", str(e))

            if self.word_app:
                try:
                    self.word_app.Quit()
                except Exception as e:
                    logger.warning("Error closing Word application: %s", str(e))
            
            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.warning("Error uninitializing COM: %s", str(e))

    def _get_template_path(self, presentation_type: str) -> Path:
        """Get the path to the InputDocumentRationales template based on presentation type.

        Args:
            presentation_type: Type of presentation (Normal, Phonetics, Katakana, etc.)

        Returns:
            Path to template file
        """
        templates_dir = settings.app_dir / "templates"

        # Map presentation type to template filename
        if presentation_type.lower() == "phonetics":
            template_name = TemplateFilename.WORD_PHONETICS
        elif presentation_type.lower() in ["katakana", "katakana_bigjap"]:
            template_name = TemplateFilename.WORD_KATAKANA
        else:
            # Default to normal template (Normal, Nonproprietary, Tagline, etc.)
            template_name = TemplateFilename.WORD_RATIONALES

        template_path = templates_dir / template_name
        logger.debug("Using Feedback template: %s", template_path)
        return template_path

    def _replace_text_in_doc(self, doc, find_text: str, replace_text: str) -> None:
        """Replace text everywhere in the Word document.

        Covers all story ranges (main body, headers/footers, text boxes) and shapes.

        Args:
            doc: Word Document COM object
            find_text: Placeholder to find (e.g., "<Client>")
            replace_text: Replacement value
        """
        try:
            replace_value = str(replace_text) if replace_text else ""
            total_hits = 0

            # 1) Replace in sections (body, headers, footers)
            try:
                for section in doc.Sections:
                    # Body content
                    try:
                        rng = section.Range
                        if self._execute_find_replace(rng, find_text, replace_value):
                            total_hits += 1
                    except Exception:
                        pass

                    # Headers
                    for header in section.Headers:
                        try:
                            rng = header.Range
                            if self._execute_find_replace(rng, find_text, replace_value):
                                total_hits += 1
                        except Exception:
                            pass

                    # Footers
                    for footer in section.Footers:
                        try:
                            rng = footer.Range
                            if self._execute_find_replace(rng, find_text, replace_value):
                                total_hits += 1
                        except Exception:
                            pass
            except Exception as e:
                logger.debug("Section iteration issue: %s", str(e))

            # 2) Replace in all StoryRanges
            try:
                for sr in doc.StoryRanges:
                    current = sr
                    while current is not None:
                        if self._execute_find_replace(current, find_text, replace_value):
                            total_hits += 1
                        try:
                            current = current.NextStoryRange
                        except Exception:
                            current = None
            except Exception as e:
                logger.debug("StoryRanges iteration issue: %s", str(e))

            # 3) Replace inside Shapes in document body
            try:
                for shape in doc.Shapes:
                    total_hits += self._replace_in_shape(shape, find_text, replace_value)
            except Exception as e:
                logger.debug("Error iterating doc.Shapes: %s", str(e))

            # 4) Replace inside Shapes within headers/footers
            try:
                for section in doc.Sections:
                    for header in section.Headers:
                        try:
                            for shape in header.Shapes:
                                total_hits += self._replace_in_shape(shape, find_text, replace_value)
                        except Exception:
                            pass
                    for footer in section.Footers:
                        try:
                            for shape in footer.Shapes:
                                total_hits += self._replace_in_shape(shape, find_text, replace_value)
                        except Exception:
                            pass
            except Exception as e:
                logger.debug("Section Shapes iteration issue: %s", str(e))

            if total_hits > 0:
                logger.debug("Replaced '%s' -> '%s' (%d occurrences)", find_text, replace_value[:50], total_hits)
            else:
                logger.debug("Placeholder '%s' NOT found in document", find_text)

        except Exception as e:
            logger.warning("Error replacing text '%s': %s", find_text, str(e))

    def _replace_in_shape(self, shape, find_text: str, replace_text: str) -> int:
        """Replace text inside a Shape/TextFrame, recursing into group items."""
        hits = 0
        try:
            if getattr(shape, "HasTextFrame", 0):
                tf = shape.TextFrame
                if getattr(tf, "HasText", 0):
                    try:
                        rng = tf.TextRange
                        if self._execute_find_replace(rng, find_text, replace_text):
                            hits += 1
                    except Exception:
                        pass

            # Recurse into grouped shapes
            try:
                group_items = getattr(shape, "GroupItems", None)
                if group_items is not None:
                    for sub in group_items:
                        hits += self._replace_in_shape(sub, find_text, replace_text)
            except Exception:
                pass
        except Exception:
            pass
        return hits

    def _execute_find_replace(self, range_obj, find_text: str, replace_text: str) -> int:
        """Execute find and replace on a specific range."""
        replacements_made = 0
        try:
            try:
                range_obj.SetRange(range_obj.Start, range_obj.End)
            except Exception:
                pass

            find_object = range_obj.Find
            find_object.ClearFormatting()
            find_object.Replacement.ClearFormatting()

            result = find_object.Execute(
                FindText=find_text,
                MatchCase=False,
                MatchWholeWord=False,
                MatchWildcards=False,
                MatchSoundsLike=False,
                MatchAllWordForms=False,
                Forward=True,
                Wrap=WD_FIND_CONTINUE,
                Format=False,
                ReplaceWith=replace_text,
                Replace=WD_REPLACE_ALL
            )

            if bool(result):
                replacements_made = 1
        except Exception as e:
            logger.debug("Error in find/replace for range: %s", str(e))

        return replacements_made

    def _format_document_header(self, doc) -> None:
        """Format document header with correct font size (8.5pt) and bold, and add logo if available.

        Args:
            doc: Word Document COM object
        """
        try:
            logger.info("Formatting document header")

            # Iterate through all sections in the document
            for section_idx in range(1, doc.Sections.Count + 1):
                section = doc.Sections(section_idx)

                # Get the header for this section
                # wdHeaderFooterPrimary = 1
                WD_HEADER_FOOTER_PRIMARY = 1
                try:
                    header = section.Headers(WD_HEADER_FOOTER_PRIMARY)

                    # Check if logo file exists
                    logo_path = settings.app_dir / "templates" / "bi_logo.png"
                    if logo_path.exists():
                        try:
                            # Insert logo at the beginning of header if table exists
                            if header.Range.Tables.Count > 0:
                                first_table = header.Range.Tables(1)
                                if first_table.Rows.Count > 0:
                                    first_cell = first_table.Rows(1).Cells(1)
                                    # Set vertical alignment for the cell (center)
                                    # wdCellAlignVerticalCenter = 1
                                    first_cell.VerticalAlignment = 1
                                    # Insert logo inline shape at the very beginning
                                    logo_range = first_cell.Range
                                    # Move to start of cell (before any text)
                                    logo_range.Collapse(Direction=1)  # wdCollapseStart = 1
                                    # Insert picture
                                    logo_shape = logo_range.InlineShapes.AddPicture(
                                        FileName=str(logo_path.absolute()),
                                        LinkToFile=False,
                                        SaveWithDocument=True
                                    )
                                    # Set logo dimensions: width=0.67", height=0.38"
                                    # 0.67 inches = 48.24 points, 0.38 inches = 27.36 points
                                    logo_shape.Width = 48.24
                                    logo_shape.Height = 27.36
                                    # Add space after logo
                                    logo_range.Collapse(Direction=0)  # wdCollapseEnd = 0
                                    logo_range.InsertAfter("  ")
                                    logger.info("Logo inserted in header table")
                        except Exception as logo_error:
                            logger.warning("Error inserting logo: %s", str(logo_error))
                    else:
                        logger.debug("Logo file not found at: %s", logo_path)

                    # Format all text in header to 8.5pt and bold
                    header_range = header.Range
                    # Format font
                    header_range.Font.Size = 8.5
                    header_range.Font.Bold = True
                    logger.debug("Formatted header section %d", section_idx)

                except Exception as header_error:
                    logger.warning("Error formatting header in section %d: %s", section_idx, str(header_error))

            logger.info("Header formatting completed")

        except Exception as e:
            logger.warning("Error formatting header: %s", str(e))

    def _populate_feedback_tables(self, doc, presentation_id: int, is_phonetics: bool) -> None:
        """Populate the feedback table using the correct stored procedure.

        Uses nw_wdToIndicateFeedback SP which returns data already formatted for the Word table.
        This SP returns columns: ' ', 'Test Name', 'Rationale', 'Positive', 'Neutral', 'Negative', 'Comments/Suggestions'

        Args:
            doc: Word Document object
            presentation_id: Presentation ID
            is_phonetics: Not used anymore, kept for compatibility
        """
        try:
            # Access tables with retry logic
            def get_table_count():
                return doc.Tables.Count

            total_tables = retry_com_operation(get_table_count, max_retries=5, delay=1.0)
            logger.info("Document has %d tables total", total_tables)

            if total_tables < 1:
                logger.error("No tables found in document!")
                return

            # Get the main feedback table (should be table 1)
            def get_main_table():
                return doc.Tables(1)

            main_table = retry_com_operation(get_main_table, max_retries=5, delay=1.0)
            logger.info("Found main table - Rows: %d, Columns: %d", main_table.Rows.Count, main_table.Columns.Count)

            # Call the correct SP: nw_wdToIndicateFeedback
            # This SP returns all the data in the correct format for the Word table
            logger.info("Fetching feedback data using nw_wdToIndicateFeedback")

            from app.config.db import get_connection_scope

            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_wdToIndicateFeedback](?)}",
                    (presentation_id,)
                )

                # Get the result set
                if cursor.description is None:
                    logger.warning("nw_wdToIndicateFeedback returned no result set")
                    return

                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()

                logger.info("Retrieved %d rows from nw_wdToIndicateFeedback", len(rows))
                logger.info("Columns: %s", columns)

                if not rows:
                    logger.warning("No data to populate - SP returned empty results")
                    return

                # Populate the table with the data
                self._populate_feedback_table_from_sp(main_table, columns, rows)

        except Exception as e:
            logger.error("Error populating feedback tables: %s", str(e), exc_info=True)

    def _populate_feedback_table_from_sp(self, table, columns, rows) -> None:
        """Populate feedback table directly from SP results.

        The nw_wdToIndicateFeedback SP returns data in this format:
        Col 0: ' ' (row number)
        Col 1: 'Test Name'
        Col 2: 'Rationale'
        Col 3: 'Positive' (empty for client)
        Col 4: 'Neutral' (empty for client)
        Col 5: 'Negative' (empty for client)
        Col 6: 'Comments/Suggestions' (empty for client)

        But the Word template table has this structure:
        Col 1: # (row number)
        Col 2: Test Name
        Col 3: Pronunciation (left empty - SP doesn't return this field)
        Col 4: Rationale
        Col 5: Positive (empty checkbox)
        Col 6: Neutral (empty checkbox)
        Col 7: Negative (empty checkbox)
        Col 8: Comments/Suggestions (empty)

        Args:
            table: Word Table object
            columns: List of column names from SP
            rows: List of row tuples from SP
        """
        try:
            logger.info("Starting to populate table with %d rows from SP", len(rows))
            logger.info("Table before clear - Rows: %d, Columns: %d", table.Rows.Count, table.Columns.Count)

            def split_grouped_text(value: object) -> List[str]:
                """Split grouped fields (## or $$) into individual entries."""
                if value is None:
                    return []
                text = html.unescape(str(value))
                delimiter = None
                if "##" in text:
                    delimiter = "##"
                elif "$$" in text:
                    delimiter = "$$"

                if not delimiter:
                    cleaned = text.strip()
                    return [cleaned] if cleaned else []

                return [part.strip() for part in text.split(delimiter) if part and part.strip()]

            def expand_grouped_rows(row_dict: dict) -> List[dict]:
                """Expand SP row into multiple rows when Test Name/Rationale contain '##'."""
                test_name_parts = split_grouped_text(row_dict.get("Test Name"))
                rationale_parts = split_grouped_text(row_dict.get("Rationale"))

                # No grouping: return as-is (but with basic unescape/strip)
                if len(test_name_parts) <= 1 and len(rationale_parts) <= 1:
                    return [
                        {
                            "Test Name": (test_name_parts[0] if test_name_parts else "").strip(),
                            "Rationale": (rationale_parts[0] if rationale_parts else "").strip(),
                        }
                    ]

                # Align list lengths
                count = max(len(test_name_parts), len(rationale_parts), 1)
                if not test_name_parts:
                    test_name_parts = [""] * count
                if not rationale_parts:
                    rationale_parts = [""] * count
                if len(test_name_parts) == 1 and count > 1:
                    test_name_parts = test_name_parts * count
                if len(rationale_parts) == 1 and count > 1:
                    rationale_parts = rationale_parts * count
                while len(test_name_parts) < count:
                    test_name_parts.append(test_name_parts[-1] if test_name_parts else "")
                while len(rationale_parts) < count:
                    rationale_parts.append(rationale_parts[-1] if rationale_parts else "")

                expanded = []
                for i in range(count):
                    expanded.append(
                        {
                            "Test Name": (test_name_parts[i] or "").strip(),
                            "Rationale": (rationale_parts[i] or "").strip(),
                        }
                    )
                return expanded

            # Delete all rows except header (row 1)
            rows_deleted = 0
            while table.Rows.Count > 1:
                table.Rows(2).Delete()
                rows_deleted += 1

            logger.info("Cleared %d rows from table, now has %d rows", rows_deleted, table.Rows.Count)

            # Expand grouped rows (##) into 1 row per entry
            expanded_rows: List[dict] = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                expanded_rows.extend(expand_grouped_rows(row_dict))

            logger.info("Expanded %d SP rows into %d table rows (group delimiter: ##/$$)", len(rows), len(expanded_rows))

            # Add rows for each result from SP (after expansion)
            for idx, row_dict in enumerate(expanded_rows):
                try:
                    # Add new row to table
                    new_row = table.Rows.Add()
                    new_row.AllowBreakAcrossPages = 0

                    # Ensure white background for body rows
                    self._force_row_white_background(new_row)

                    # The SP already provides the data in the correct format
                    # We just need to copy it to the Word table cells

                    # Set vertical alignment for all cells in the row (center)
                    # wdCellAlignVerticalCenter = 1
                    for cell_idx in range(1, new_row.Cells.Count + 1):
                        new_row.Cells(cell_idx).VerticalAlignment = 1

                    # Column 1: Row number - CENTER aligned
                    if new_row.Cells.Count >= 1:
                        cell_range = new_row.Cells(1).Range
                        cell_range.Text = str(idx + 1)
                        cell_range.Font.Size = 10
                        cell_range.Font.Bold = False
                        cell_range.Font.Name = "Calibri"
                        cell_range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
                        new_row.Cells(1).VerticalAlignment = 1  # wdCellAlignVerticalCenter

                    # Column 2: Test Name - LEFT aligned, BOLD
                    if new_row.Cells.Count >= 2:
                        cell_range = new_row.Cells(2).Range
                        cell_range.Text = str(row_dict.get('Test Name') or '')
                        cell_range.Font.Size = 10
                        cell_range.Font.Bold = True  # Test Name should be bold
                        cell_range.Font.Name = "Calibri"
                        cell_range.ParagraphFormat.Alignment = 0  # wdAlignParagraphLeft
                        new_row.Cells(2).VerticalAlignment = 1  # wdCellAlignVerticalCenter

                    # Column 3: Pronunciation - LEFT aligned
                    if new_row.Cells.Count >= 3:
                        cell_range = new_row.Cells(3).Range
                        cell_range.Text = str(row_dict.get('Pronunciation') or '')
                        cell_range.Font.Size = 10
                        cell_range.Font.Bold = False
                        cell_range.Font.Name = "Calibri"
                        cell_range.ParagraphFormat.Alignment = 0  # wdAlignParagraphLeft
                        new_row.Cells(3).VerticalAlignment = 1  # wdCellAlignVerticalCenter

                    # Column 4: Rationale - LEFT aligned
                    if new_row.Cells.Count >= 4:
                        cell_range = new_row.Cells(4).Range
                        cell_range.Text = str(row_dict.get('Rationale') or '')
                        cell_range.Font.Size = 10
                        cell_range.Font.Bold = False
                        cell_range.Font.Name = "Calibri"
                        cell_range.ParagraphFormat.Alignment = 0  # wdAlignParagraphLeft
                        new_row.Cells(4).VerticalAlignment = 1  # wdCellAlignVerticalCenter

                    # Columns 5-8: Positive, Neutral, Negative, Comments/Suggestions - CENTER aligned
                    # These are left empty for the client to fill
                    for col_idx in range(5, min(9, new_row.Cells.Count + 1)):
                        if new_row.Cells.Count >= col_idx:
                            cell_range = new_row.Cells(col_idx).Range
                            cell_range.Text = ""
                            cell_range.Font.Size = 10
                            cell_range.Font.Bold = False
                            cell_range.Font.Name = "Calibri"
                            cell_range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
                            new_row.Cells(col_idx).VerticalAlignment = 1  # wdCellAlignVerticalCenter

                    # Ensure row font is 10pt and not bolded by style inheritance
                    try:
                        new_row.Range.Font.Size = 10
                        new_row.Range.Font.Bold = False
                    except Exception:
                        pass

                    logger.debug("Added row %d: %s", idx + 1, row_dict.get('Test Name', ''))

                except Exception as row_error:
                    logger.error("Error adding row %d: %s", idx + 1, str(row_error), exc_info=True)

            logger.info("Successfully populated table with %d rows", len(expanded_rows))

        except Exception as e:
            logger.error("Error populating feedback table from SP: %s", str(e), exc_info=True)

    def _populate_feedback_table(self, table, results: List, table_name: str, summary_type: SummaryType) -> None:
        """Populate a feedback table with results.

        Based on the screenshots, the InputDocumentRationales table structure is:
        Col 1: # (row number)
        Col 2: Test Name (with special formatting - underlined portions in parentheses)
        Col 3: Pronunciation (phonetics)
        Col 4: Rationale
        Col 5: Positive (empty checkbox)
        Col 6: Neutral (empty checkbox)
        Col 7: Negative (empty checkbox)
        Col 8: Comments/Suggestions (empty)

        This is a CLIENT FEEDBACK form, so checkboxes are left empty for them to fill.

        Args:
            table: Word Table object
            results: List of WordReportResult objects
            table_name: Name of the table
            summary_type: Type of summary (to determine column structure)
        """
        try:
            logger.info("Starting to populate table '%s' with %d results", table_name, len(results))
            logger.info("Table before clear - Rows: %d, Columns: %d", table.Rows.Count, table.Columns.Count)

            # Delete all rows except header (row 1)
            rows_deleted = 0
            while table.Rows.Count > 1:
                table.Rows(2).Delete()
                rows_deleted += 1

            logger.info("Cleared %d rows from table '%s', now has %d rows", rows_deleted, table_name, table.Rows.Count)

            # Keep template formatting for fonts/sizes and header
            # Body rows will explicitly force white background per requirement

            # Row counter for numbering
            row_number = 1

            # Add rows for each result
            logger.info("Processing %d results to add to table", len(results))
            for idx, result in enumerate(results):
                logger.debug("Processing result #%d: name='%s'", idx + 1, result.name)
                # Handle grouped names (split by ## or $$)
                names_to_add = []
                if result.name and ('##' in result.name or '$$' in result.name):
                    delimiter = '##' if '##' in result.name else '$$'
                    names_to_add = [n.strip() for n in result.name.split(delimiter) if n.strip()]
                else:
                    names_to_add = [result.name] if result.name else [""]

                # Process each name (grouped names get separate rows)
                for name_idx, name in enumerate(names_to_add):
                    logger.debug("Adding name %d/%d: '%s' (row_number=%d)", name_idx + 1, len(names_to_add), name, row_number)

                    try:
                        new_row = table.Rows.Add()
                        logger.debug("Row added successfully. Table now has %d rows", table.Rows.Count)
                        new_row.AllowBreakAcrossPages = 0

                        # Ensure body row background is white (header remains as in template)
                        self._force_row_white_background(new_row)

                        # Column 1: Row number
                        if new_row.Cells.Count >= 1:
                            new_row.Cells(1).Range.Text = str(row_number)
                            logger.debug("Set row number to: %d", row_number)

                        # Column 2: Test Name with special formatting
                        # Format: MainName (C) (CHB) where portions in () are underlined
                        if new_row.Cells.Count >= 2:
                            self._format_test_name(new_row.Cells(2).Range, name, result.category)
                            logger.debug("Formatted test name: '%s' with category '%s'", name, result.category)

                        # Column 3: Pronunciation (for phonetics presentations)
                        if new_row.Cells.Count >= 3:
                            pronunciation_text = result.pronunciation or ""
                            new_row.Cells(3).Range.Text = pronunciation_text
                            logger.debug("Set pronunciation: '%s'", pronunciation_text[:30] if pronunciation_text else "(empty)")

                        # Column 4: Rationale (DO NOT modify font size - keep template format)
                        if new_row.Cells.Count >= 4:
                            cell_range = new_row.Cells(4).Range

                            rationale_text = result.rationale or ""

                            # Check if rationale contains HTML
                            if self._is_html_content(rationale_text):
                                logger.debug("Rationale contains HTML, inserting formatted content")
                                # Clean HTML and insert formatted content
                                self._insert_html_content(cell_range, rationale_text)
                            else:
                                cell_range.Text = rationale_text
                                logger.debug("Set rationale (plain text): '%s'", rationale_text[:30] if rationale_text else "(empty)")

                        # Columns 5-7: Positive, Neutral, Negative checkboxes (left EMPTY for client)
                        # These are already formatted in the template, don't modify

                        # Column 8: Comments/Suggestions (left EMPTY for client)
                        # Already formatted in template, don't modify

                        row_number += 1
                        logger.debug("Successfully populated row %d", row_number - 1)

                    except Exception as cell_error:
                        logger.error("Error populating cell in '%s' for name '%s': %s", table_name, name, str(cell_error), exc_info=True)

            logger.info("Successfully populated table '%s' with %d rows", table_name, len(results))

        except Exception as e:
            logger.error("Error populating feedback table '%s': %s", table_name, str(e), exc_info=True)

    def _force_row_white_background(self, row) -> None:
        """Force white background for all cells in the given row.

        Leaves the header row unchanged because we only call this on newly
        created body rows. Tries both ColorIndex and RGB assignments for
        compatibility across Word versions.
        """
        try:
            # Word constants (fallback numeric values)
            WD_COLOR_INDEX_WHITE = 8  # wdWhite
            WD_TEXTURE_NONE = 0       # wdTextureNone
            WD_COLOR_WHITE = 0xFFFFFF # RGB white

            cell_count = row.Cells.Count if hasattr(row, 'Cells') else 0
            for i in range(1, cell_count + 1):
                try:
                    cell = row.Cells(i)
                    shading = getattr(cell, 'Shading', None)
                    if shading is None:
                        continue
                    # Remove any texture/pattern
                    try:
                        shading.Texture = WD_TEXTURE_NONE
                    except Exception:
                        pass
                    # Prefer color index if supported
                    try:
                        shading.BackgroundPatternColorIndex = WD_COLOR_INDEX_WHITE
                    except Exception:
                        # Fallback to explicit RGB color
                        try:
                            shading.BackgroundPatternColor = WD_COLOR_WHITE
                        except Exception:
                            pass
                except Exception:
                    # Continue applying to remaining cells even if one fails
                    continue
        except Exception as e:
            logger.debug("Could not set row background to white: %s", str(e))

    def _format_test_name(self, cell_range, name: str, category: Optional[str]) -> None:
        """Format test name with category codes.

        Simply inserts the name with category codes in parentheses.
        Does NOT apply underline or any other formatting - keeps template formatting.

        Example: "HealthNav (C)" or "MyOrbis (CB) (DE)"

        Args:
            cell_range: Word Range object for the cell
            name: The candidate name
            category: Category code (e.g., "C", "CHB", "DE", "FR")
        """
        try:
            # Build the full text with category codes
            if category:
                # Parse category codes if needed
                category_parts = self._parse_category_codes(category)

                # Build text: "Name (Code1) (Code2)"
                codes_text = " ".join([f"({code})" for code in category_parts])
                full_text = f"{name} {codes_text}"
            else:
                full_text = name

            # Simply set the text - DO NOT apply any formatting
            # The template's existing formatting will be preserved
            cell_range.Text = full_text

        except Exception as e:
            logger.warning("Error formatting test name: %s", str(e))
            # Fallback: just set plain text
            cell_range.Text = f"{name} ({category})" if category else name

    def _parse_category_codes(self, category: str) -> List[str]:
        """Parse category string into individual codes.

        Examples:
        - "C" -> ["C"]
        - "CHB" -> ["C", "CHB"] or just ["CHB"] depending on convention
        - "DE" -> ["DE"]
        - "C (DE)" -> ["C", "DE"]

        Args:
            category: Category string from database

        Returns:
            List of category codes
        """
        if not category:
            return []

        # If already in parentheses format, extract codes
        # Pattern: "(C) (DE) (FR)"
        if '(' in category and ')' in category:
            codes = []
            import re
            matches = re.findall(r'\(([^)]+)\)', category)
            return matches

        # Otherwise, treat as single code
        return [category.strip()]

    def _is_html_content(self, text: str) -> bool:
        """Check if text contains HTML tags.

        Args:
            text: Text to check

        Returns:
            True if HTML detected, False otherwise
        """
        if not text:
            return False
        # Simple check for HTML tags
        return '<' in text and '>' in text

    def _insert_html_content(self, cell_range, html_text: str) -> None:
        """Insert HTML-formatted content into a cell.

        Args:
            cell_range: Word Range object for the cell
            html_text: HTML content to insert
        """
        try:
            # Clean up inverted tildes and other special characters
            cleaned_html = self._clean_html_for_word(html_text)

            # Create a temporary document to parse HTML
            temp_doc = self.word_app.Documents.Add()
            temp_range = temp_doc.Range()

            # Paste HTML content
            temp_range.Text = cleaned_html
            # Try to interpret as HTML (Word will auto-format)
            # Note: Word COM doesn't have direct HTML paste, so we use a workaround
            # Copy formatted content
            temp_range.Copy()

            # Paste into target cell
            cell_range.Select()
            cell_range.PasteSpecial()

            # Close temp document
            temp_doc.Close(False)

        except Exception as e:
            logger.warning("Error inserting HTML content, falling back to plain text: %s", str(e))
            # Fallback: strip HTML tags and use plain text
            plain_text = self._strip_html_tags(html_text)
            cell_range.Text = plain_text

    def _clean_html_for_word(self, html_text: str) -> str:
        """Clean HTML content for Word insertion.

        Args:
            html_text: HTML text to clean

        Returns:
            Cleaned HTML text
        """
        # Remove inverted tildes and other problematic characters
        cleaned = html_text.replace('`', "'")
        # Unescape HTML entities
        cleaned = html.unescape(cleaned)
        return cleaned

    def _strip_html_tags(self, html_text: str) -> str:
        """Strip HTML tags from text.

        Args:
            html_text: HTML text

        Returns:
            Plain text without HTML tags
        """
        # Simple regex to remove HTML tags
        clean = re.compile('<.*?>')
        plain_text = re.sub(clean, '', html_text)
        # Unescape HTML entities
        plain_text = html.unescape(plain_text)
        return plain_text

    def _generate_and_insert_pie_chart(self, doc, by_the_numbers: List) -> None:
        """Generate pie chart in Excel and insert into Word document.

        Uses corporate colors:
        - Positive: RGB(92, 184, 92) - Green
        - Neutral: RGB(183, 122, 51) - Brown/Orange
        - Reconsider: RGB(153, 0, 76) - Purple

        Args:
            doc: Word Document object
            by_the_numbers: List with counts from ByTheNumbers SP
        """
        try:
            if not by_the_numbers:
                logger.warning("No data for pie chart")
                return

            # Extract counts
            counts = by_the_numbers[0] if by_the_numbers else None
            if not counts:
                return

            positive = getattr(counts, 'positive_count', 0) or 0
            neutral = getattr(counts, 'neutral_count', 0) or 0
            reconsider = getattr(counts, 'reconsider_count', 0) or 0

            logger.info("Pie chart data: Positive=%d, Neutral=%d, Reconsider=%d", positive, neutral, reconsider)

            # Create Excel instance
            from app.utils.com_manager import com_manager
            with com_manager.acquire("Excel.Application Dispatch"):
                self.excel_app = win32com.client.Dispatch("Excel.Application")
                self.excel_app.Visible = False
                self.excel_app.DisplayAlerts = False

            try:
                # Create workbook
                workbook = self.excel_app.Workbooks.Add()
                sheet = workbook.Worksheets(1)

                # Write data for chart
                sheet.Cells(1, 1).Value = "Category"
                sheet.Cells(1, 2).Value = "Count"
                sheet.Cells(2, 1).Value = "Positive"
                sheet.Cells(2, 2).Value = positive
                sheet.Cells(3, 1).Value = "Neutral"
                sheet.Cells(3, 2).Value = neutral
                sheet.Cells(4, 1).Value = "Reconsider"
                sheet.Cells(4, 2).Value = reconsider

                # Create pie chart
                chart = sheet.ChartObjects().Add(50, 50, 400, 300)
                chart.Chart.ChartType = 5  # xlPie
                chart.Chart.SetSourceData(sheet.Range("A1:B4"))
                chart.Chart.HasTitle = False
                chart.Chart.HasLegend = True

                # Set corporate colors (in BGR format for COM)
                if chart.Chart.SeriesCollection().Count > 0:
                    series = chart.Chart.SeriesCollection(1)
                    if series.Points().Count >= 1:
                        # RGB(92, 184, 92) -> BGR = 0x5CB85C
                        series.Points(1).Format.Fill.ForeColor.RGB = 0x5CB85C  # Green
                    if series.Points().Count >= 2:
                        # RGB(183, 122, 51) -> BGR = 0x337AB7
                        series.Points(2).Format.Fill.ForeColor.RGB = 0x337AB7  # Brown/Orange
                    if series.Points().Count >= 3:
                        # RGB(153, 0, 76) -> BGR = 0x4C0099
                        series.Points(3).Format.Fill.ForeColor.RGB = 0x4C0099  # Purple

                # Copy chart to clipboard
                chart.Copy()

                # Paste into Word at PieChart bookmark
                def check_and_paste_chart():
                    if doc.Bookmarks.Exists("PieChart"):
                        bookmark = doc.Bookmarks("PieChart")
                        bookmark.Range.Paste()
                        logger.info("Pie chart inserted successfully")
                        return True
                    else:
                        logger.warning("PieChart bookmark not found in document")
                        return False
                
                retry_com_operation(check_and_paste_chart, max_retries=5, delay=1.0)

                workbook.Close(False)

            finally:
                # Close Excel
                try:
                    self.excel_app.Quit()
                finally:
                    self.excel_app = None

        except Exception as e:
            logger.error("Error generating pie chart: %s", str(e), exc_info=True)

    def _cleanup_bookmarks(self, doc) -> None:
        """Remove all bookmarks from the document.

        Args:
            doc: Word Document object
        """
        try:
            def get_bookmark_count():
                return doc.Bookmarks.Count
            
            bookmark_count = retry_com_operation(get_bookmark_count, max_retries=5, delay=1.0)

            # Delete bookmarks in reverse order to avoid index issues
            for i in range(bookmark_count, 0, -1):
                try:
                    def delete_bookmark():
                        doc.Bookmarks(i).Delete()
                    retry_com_operation(delete_bookmark, max_retries=3, delay=0.5)
                except:
                    pass

            logger.info("Cleaned up %d bookmarks", bookmark_count)

        except Exception as e:
            logger.warning("Error cleaning up bookmarks: %s", str(e))


# Global service instance
feedback_template_generator = FeedbackTemplateGenerator()
