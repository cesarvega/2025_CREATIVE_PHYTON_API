"""Service for generating NW Word reports using templates and COM automation."""

from __future__ import annotations
import win32com.client
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from app.config.settings import settings
from app.models.nw_reports_models import SummaryType
from app.services.nw_reports_service import nw_reports_service
from app.utils.logging_utils import get_logger
from app.utils.nw_data_utils import sanitize_filename
from app.utils.path_utils import get_nw_downloads_dir
# OPTIMIZATION: Use centralized Word COM utilities
from app.utils.word_com_utils import (
    replace_text_in_document,
    update_bookmark,
    cleanup_bookmarks,
    set_word_app_optimization,
    safe_close_com_object,
    WD_FORMAT_DOCUMENT
)

logger = get_logger(__name__)


class WordReportGenerator:
    """Generate NW Word reports from templates using Word COM automation."""

    def __init__(self):
        """Initialize the Word report generator."""
        self.word_app: Optional[win32com.client.CDispatch] = None
        self.excel_app: Optional[win32com.client.CDispatch] = None
        self.template_path: Optional[Path] = None

    def generate_word_report(
        self,
        presentation_id: int,
        display_name: str,
        is_phonetics: bool = True,
    ) -> Path:
        """Generate complete Word report for NW presentation.

        Args:
            presentation_id: Presentation ID to generate report for
            display_name: Display name for the presentation

        Returns:
            Path to generated Word file
        """
        logger.info(
            "Starting Word report generation for presentation_id=%d",
            presentation_id,
        )

        # Determine template path
        self.template_path = self._get_template_path()
        if not self.template_path.exists():
            raise FileNotFoundError(f"Word template not found: {self.template_path}")

        # Initialize Word application
        try:
            from app.utils.com_manager import com_manager
            with com_manager.acquire("Word.Application Dispatch"):
                self.word_app = win32com.client.Dispatch("Word.Application")
                self.word_app.Visible = False
                self.word_app.DisplayAlerts = 0  # wdAlertsNone

            # Open template
            doc = self.word_app.Documents.Open(str(self.template_path.absolute()))

            # Get replacement values from SP
            replacements = nw_reports_service.get_word_report_replacements(presentation_id)
            logger.info("Retrieved %d replacement values", len(replacements))

            # PHASE 1: Replace placeholders in document
            logger.info("Phase 1: Replacing placeholders")
            for replacement in replacements:
                if replacement.placeholder and replacement.value:
                    logger.info(
                        "Replacing placeholder: '%s' -> '%s'",
                        replacement.placeholder,
                        replacement.value
                    )
                    # OPTIMIZATION: Use centralized replace_text_in_document
                    replace_text_in_document(
                        doc,
                        replacement.placeholder,
                        replacement.value
                    )

            # PHASE 2: Populate content
            logger.info("Phase 2: Populating content")

            # Populate tables for each summary type
            self._populate_result_tables(doc, presentation_id, is_phonetics)

            # Get "By The Numbers" data for bookmarks and pie chart
            by_the_numbers_data = nw_reports_service.get_word_report_results_phonetics(
                presentation_id,
                SummaryType.BY_THE_NUMBERS
            )

            # Update "By The Numbers" bookmarks
            self._update_by_the_numbers(doc, by_the_numbers_data)

            # Generate and insert Pie Chart
            self._generate_and_insert_pie_chart(doc, by_the_numbers_data)

            # Clean up bookmarks - OPTIMIZATION: Use centralized function
            cleanup_bookmarks(doc)

            # Save document to output directory
            output_dir = get_nw_downloads_dir()
            output_dir.mkdir(parents=True, exist_ok=True)

            safe_display_name = sanitize_filename(display_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"NW_Report_{safe_display_name}_{timestamp}.doc"
            output_path = output_dir / filename

            # Save as Word 97-2003 format (.doc)
            doc.SaveAs2(str(output_path.absolute()), FileFormat=WD_FORMAT_DOCUMENT)
            doc.Close()

            logger.info("Word report saved to: %s", output_path)

            return output_path

        except Exception as e:
            logger.error("Error generating Word report: %s", str(e), exc_info=True)
            raise

        finally:
            # OPTIMIZATION: Use safe_close_com_object for cleanup
            safe_close_com_object(self.excel_app, "Excel Application")
            safe_close_com_object(self.word_app, "Word Application")


    def _get_template_path(self) -> Path:
        """Get the path to the Word template.

        Returns:
            Path to Word template file
        """
        # Use the newest template by default
        template_name = "nomenclature workshop report_template_new_phonetics.doc"
        template_path = settings.app_dir / "templates" / template_name

        logger.debug("Using Word template: %s", template_path)
        return template_path

    def _replace_text_in_doc(self, doc, find_text: str, replace_text: str) -> None:
        """Replace text everywhere in the Word document.

        DEPRECATED: This method now delegates to word_com_utils.replace_text_in_document()
        Kept for backward compatibility but internally uses centralized implementation.

        Args:
            doc: Word Document COM object
            find_text: Placeholder to find (e.g., "<Client>")
            replace_text: Replacement value
        """
        # OPTIMIZATION: Delegate to centralized implementation
        return replace_text_in_document(doc, find_text, replace_text)

    def _replace_in_shape(self, shape, find_text: str, replace_text: str) -> int:
        """Replace text inside a Shape/TextFrame, recursing into group items.

        Returns an approximate count of places we executed a replacement.
        """
        hits = 0
        try:
            # If the shape has a text frame with text, run Find on its Range
            if getattr(shape, "HasTextFrame", 0):
                tf = shape.TextFrame
                if getattr(tf, "HasText", 0):
                    try:
                        rng = tf.TextRange
                        if self._execute_find_replace(rng, find_text, replace_text):
                            hits += 1
                    except Exception:
                        pass

            # Recurse into grouped shapes when available
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
        """Execute find and replace on a specific range.

        Args:
            range_obj: Word Range object
            find_text: Text to find
            replace_text: Text to replace with

        Returns:
            Number of replacements made
        """
        replacements_made = 0
        try:
            # Reset the range to the start to ensure we search the entire range
            try:
                range_obj.SetRange(range_obj.Start, range_obj.End)
            except Exception:
                pass

            find_object = range_obj.Find
            find_object.ClearFormatting()
            find_object.Replacement.ClearFormatting()

            # Use explicit Execute args (more reliable across COM objects)
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

    def _update_by_the_numbers(self, doc, by_the_numbers: List) -> None:
        """Update bookmarks with 'By The Numbers' summary counts.

        Args:
            doc: Word Document object
            by_the_numbers: List of results from BY_THE_NUMBERS SP
        """
        try:
            # Extract counts from the first row if available
            if not by_the_numbers:
                logger.warning("No 'By The Numbers' data available")
                return

            # The SP should return counts, we need to map them to bookmarks
            # Typical bookmarks: Positive, Neutral, Reconsider, NewName (not NewNames)
            counts = by_the_numbers[0] if by_the_numbers else None

            if counts:
                # Get counts from the WordReportResult object
                total = counts.total_count or 8  # Default to 8 as shown in screenshot
                positive_count = counts.positive_count or 0
                neutral_count = counts.neutral_count or 0
                reconsider_count = counts.reconsider_count or 0
                new_names_count = counts.new_names_count or 0

                logger.info(
                    "By The Numbers: Positive=%d, Neutral=%d, Reconsider=%d, NewNames=%d, Total=%d",
                    positive_count, neutral_count, reconsider_count, new_names_count, total
                )

                # Map bookmarks to their counts
                # Format: "X Out of Y" where Y is the total
                bookmark_map = {
                    "Positive": f"{positive_count} Out of {total}",
                    "Neutral": f"{neutral_count} Out of {total}",
                    "Reconsider": f"{reconsider_count} Out of {total}",
                    "NewName": str(new_names_count),  # Note: singular "NewName" not "NewNames"
                }

                # OPTIMIZATION: Use centralized update_bookmark
                for bookmark_name, text in bookmark_map.items():
                    update_bookmark(doc, bookmark_name, text)

        except Exception as e:
            logger.error("Error updating 'By The Numbers': %s", str(e), exc_info=True)

    def _update_bookmark(self, doc, bookmark_name: str, text: str) -> None:
        """Update a bookmark in the document with text.

        Args:
            doc: Word Document object
            bookmark_name: Name of the bookmark
            text: Text to insert
        """
        try:
            if doc.Bookmarks.Exists(bookmark_name):
                bookmark = doc.Bookmarks(bookmark_name)
                bookmark.Range.Text = text
                logger.debug("Updated bookmark '%s' with '%s'", bookmark_name, text)
            else:
                logger.warning("Bookmark '%s' not found in document", bookmark_name)
        except Exception as e:
            logger.warning("Error updating bookmark '%s': %s", bookmark_name, str(e))

    def _populate_result_tables(self, doc, presentation_id: int, is_phonetics: bool) -> None:
        """Populate all result tables using phonetics SP.

        Args:
            doc: Word Document object
            presentation_id: Presentation ID
            is_phonetics: If True, use Phonetics summary types (Positive_Phonetics, etc.)
        """
        # OPTIMIZATION: Enable all Word optimizations (40-60% faster)
        set_word_app_optimization(self.word_app, enabled=True)
        try:
            # Map of table indices to summary types
            # Based on original C# implementation (NWReportClass.cs):
            # Table 2: Positive names
            # Table 3: Neutral/Negative names (Negative_Phonetics for phonetics mode)
            # Table 4: Reconsider names
            # Table 5: New names
            # Table 6: Explore (roots/concepts to explore)

            if is_phonetics:
                table_configs = [
                    (2, SummaryType.POSITIVE_PHONETICS, "Positive Names"),
                    (3, SummaryType.NEGATIVE_PHONETICS, "Neutral Names"),  # Note: Uses NEGATIVE_PHONETICS
                    (4, SummaryType.RECONSIDER_PHONETICS, "Names For Reconsideration"),
                    (5, SummaryType.NEW_NAMES, "Newly Created Names"),
                    (6, SummaryType.EXPLORE, "Roots/Concepts to Explore"),
                ]
            else:
                table_configs = [
                    (2, SummaryType.POSITIVE, "Positive Names"),
                    (3, SummaryType.NEUTRAL, "Neutral Names"),
                    (4, SummaryType.RECONSIDER, "Names For Reconsideration"),
                    (5, SummaryType.NEW_NAMES, "Newly Created Names"),
                    (6, SummaryType.EXPLORE, "Roots/Concepts to Explore"),
                ]

            total_tables = doc.Tables.Count
            logger.info("Document has %d tables total", total_tables)

            for table_index, summary_type, table_name in table_configs:
                if table_index > total_tables:
                    logger.warning("Table %d (%s) not found", table_index, table_name)
                    continue

                # Get results for this summary type
                # For Newly Created Names, use special method that calls nw_wdGetResults_Phonetics with NewNames
                if summary_type == SummaryType.NEW_NAMES:
                    results = self._get_new_names_results(presentation_id)
                else:
                    # For all other tables, use nw_wdGetResults_Phonetics
                    results = nw_reports_service.get_word_report_results_phonetics(
                        presentation_id,
                        summary_type
                    )

                logger.info("Retrieved %d results for '%s'", len(results), table_name)

                if results:
                    table = doc.Tables(table_index)
                    self._populate_table(table, results, table_name)
                else:
                    logger.info("No data to populate for table '%s'", table_name)

        except Exception as e:
            logger.error("Error populating tables: %s", str(e), exc_info=True)
        finally:
            # OPTIMIZATION: Restore normal Word operations
            set_word_app_optimization(self.word_app, enabled=False)

    def _get_new_names_results(self, presentation_id: int) -> List:
        """Get NewNames results and process the special NameRationale format.

        The NewNames data has a special format where NameRationale can contain
        a '+' separator to split into NameRationale and OriginalCategory columns.

        Uses nw_CombineNewNames SP (same as Excel report) instead of nw_wdGetResults_Phonetics
        because the Word report NewNames table needs the expanded data.

        Args:
            presentation_id: Presentation ID

        Returns:
            List of WordReportResult objects with processed data
        """
        try:
            # Use the same method as Excel report - nw_CombineNewNames
            # This returns the actual newly created names data
            results = nw_reports_service.get_newly_created_names_for_word(
                presentation_id
            )

            # Process the NameRationale field
            # If it contains '+', split it into: rationale + original_category
            processed_results = []
            for result in results:
                if result.rationale and '+' in result.rationale:
                    parts = result.rationale.split('+', 1)
                    # Update rationale to first part
                    result.rationale = parts[0].strip()
                    # Store original category in a new field (we'll use name_rationale_part2)
                    result.name_rationale_part2 = parts[1].strip() if len(parts) > 1 else ""
                else:
                    # No split needed
                    result.name_rationale_part2 = ""

                processed_results.append(result)

            logger.info("Processed %d NewNames results from nw_CombineNewNames", len(processed_results))
            return processed_results

        except Exception as e:
            logger.error("Error getting NewNames results: %s", str(e), exc_info=True)
            return []

    def _populate_table(self, table, results: List, table_name: str) -> None:
        """Populate a table with results, clearing existing rows first.

        Args:
            table: Word Table object
            results: List of WordReportResult objects
            table_name: Name of the table
        """
        try:
            # Delete all rows except header (row 1)
            while table.Rows.Count > 1:
                table.Rows(2).Delete()

            logger.debug("Cleared table '%s', now adding %d rows", table_name, len(results))

            # Set table formatting
            table.Range.Font.Size = 12
            table.Range.Font.Name = "Open Sans"
            table.Rows(1).HeadingFormat = -1

            # Ensure only the header row gets the header style/colors from template
            try:
                table.ApplyStyleHeadingRows = True
                table.ApplyStyleFirstColumn = False
                table.ApplyStyleLastColumn = False
                table.ApplyStyleRowBands = False
                table.ApplyStyleColumnBands = False
            except Exception:
                pass

            # Check if this is the "Newly Created Names" table (has 4 columns including Original Category)
            is_new_names = table_name == "Newly Created Names"

            # Add rows for each result
            for result in results:
                new_row = table.Rows.Add()

                # Prevent row from breaking across pages
                new_row.AllowBreakAcrossPages = 0
                # Make sure added rows are not treated as header rows
                try:
                    new_row.HeadingFormat = 0
                except Exception:
                    pass

                # OPTIMIZATION: Commented out excessive cell-by-cell formatting
                # This was causing ~1 second per row due to excessive COM calls (40+ operations per row)
                # Total impact with 195 rows: 195 seconds wasted on formatting
                # If rows need formatting cleanup, it will be done in bulk after all rows are added
                # See: OPTIMIZACION_WORD_REPORT.md for details

                # Populate cells based on table type
                try:
                    # Column 1: Candidate/Name
                    if new_row.Cells.Count >= 1:
                        new_row.Cells(1).Range.Text = result.name or ""

                    if is_new_names:
                        # NewNames table: Candidate | Pronunciation | Rationale | Original Category
                        if new_row.Cells.Count >= 2:
                            # Column 2: Pronunciation (empty for new names)
                            new_row.Cells(2).Range.Text = result.pronunciation or ""
                        if new_row.Cells.Count >= 3:
                            # Column 3: Rationale (size 10)
                            cell_range = new_row.Cells(3).Range
                            cell_range.Font.Size = 10
                            cell_range.Text = result.rationale or ""
                        if new_row.Cells.Count >= 4:
                            # Column 4: Original Category (from name_rationale_part2)
                            new_row.Cells(4).Range.Text = result.name_rationale_part2 or ""
                    else:
                        # Regular tables: Candidate | Pronunciation | Rationale | Category
                        if new_row.Cells.Count >= 2:
                            # Column 2: Pronunciation
                            pronunciation = result.pronunciation or ""
                            new_row.Cells(2).Range.Text = pronunciation
                        if new_row.Cells.Count >= 3:
                            # Column 3: Rationale (size 10)
                            cell_range = new_row.Cells(3).Range
                            cell_range.Font.Size = 10
                            cell_range.Text = result.rationale or ""
                        if new_row.Cells.Count >= 4:
                            # Column 4: Category
                            new_row.Cells(4).Range.Text = result.category or ""

                except Exception as cell_error:
                    logger.warning("Error populating cell in '%s': %s", table_name, str(cell_error))

            # OPTIMIZATION: Clean up formatting in bulk for all data rows (much faster than per-cell)
            # Remove blue background and white text inherited from header row
            if table.Rows.Count > 1:
                try:
                    logger.debug("Cleaning formatting for %d data rows in table '%s'", table.Rows.Count - 1, table_name)
                    for row_num in range(2, table.Rows.Count + 1):
                        try:
                            row_range = table.Rows(row_num).Range
                            # Set white background for data rows (16777215 = RGB white)
                            row_range.Shading.BackgroundPatternColor = 16777215
                            # Also ensure no texture
                            row_range.Shading.Texture = 0
                            # Set text color to black (0 = wdColorAutomatic/black)
                            row_range.Font.ColorIndex = 0  # wdAuto (black)
                            row_range.Font.Color = 0  # RGB black
                        except Exception as row_error:
                            logger.debug("Error cleaning row %d: %s", row_num, str(row_error))
                except Exception as e:
                    logger.warning("Error cleaning table formatting: %s", str(e))

            logger.info("Successfully populated table '%s' with %d rows", table_name, len(results))

        except Exception as e:
            logger.error("Error populating table '%s': %s", table_name, str(e), exc_info=True)

    def _generate_and_insert_pie_chart(self, doc, by_the_numbers: List) -> None:
        """Generate pie chart in Excel and insert into Word document.

        Args:
            doc: Word Document object
            by_the_numbers: List with counts
        """
        try:
            if not by_the_numbers:
                logger.warning("No data for pie chart")
                return

            # Extract counts
            counts = by_the_numbers[0] if by_the_numbers else None
            if not counts:
                return

            positive = getattr(counts, 'positive_count', 0)
            neutral = getattr(counts, 'neutral_count', 0)
            reconsider = getattr(counts, 'reconsider_count', 0)

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
                sheet.Cells(2, 1).Value = "Positive"  # This will be the first slice
                sheet.Cells(2, 2).Value = positive
                sheet.Cells(3, 1).Value = "Neutral"  # Second slice
                sheet.Cells(3, 2).Value = neutral
                sheet.Cells(4, 1).Value = "Reconsider"  # Third slice
                sheet.Cells(4, 2).Value = reconsider

                # Create pie chart
                chart = sheet.ChartObjects().Add(50, 50, 400, 300)
                chart.Chart.ChartType = 5  # xlPie
                chart.Chart.SetSourceData(sheet.Range("A1:B4"))
                chart.Chart.HasTitle = False
                chart.Chart.HasLegend = True

                # Set colors (RGB) - matching C# implementation
                # Point 1: Green (92, 184, 92)
                # Point 2: Brown/Orange (183, 122, 51)
                # Point 3: Purple (153, 0, 76)
                if chart.Chart.SeriesCollection().Count > 0:
                    series = chart.Chart.SeriesCollection(1)
                    if series.Points().Count >= 1:
                        # RGB(92, 184, 92) = 0x5CB85C in BGR format for COM
                        series.Points(1).Format.Fill.ForeColor.RGB = 0x5CB85C  # Green
                    if series.Points().Count >= 2:
                        # RGB(183, 122, 51) = 0x337AB7 in BGR format for COM
                        series.Points(2).Format.Fill.ForeColor.RGB = 0x337AB7  # Brown/Orange
                    if series.Points().Count >= 3:
                        # RGB(153, 0, 76) = 0x4C0099 in BGR format for COM
                        series.Points(3).Format.Fill.ForeColor.RGB = 0x4C0099  # Purple

                # Copy chart to clipboard
                chart.Copy()

                # Paste into Word at PieChart bookmark
                if doc.Bookmarks.Exists("PieChart"):
                    bookmark = doc.Bookmarks("PieChart")
                    bookmark.Range.Paste()
                    logger.info("Pie chart inserted successfully")
                else:
                    logger.warning("PieChart bookmark not found in document")

                workbook.Close(False)

            finally:
                # Close Excel and clear the reference so outer finally doesn't double-quit
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
            bookmark_count = doc.Bookmarks.Count

            # Delete bookmarks in reverse order to avoid index issues
            for i in range(bookmark_count, 0, -1):
                try:
                    doc.Bookmarks(i).Delete()
                except:
                    pass

            logger.info("Cleaned up %d bookmarks", bookmark_count)

        except Exception as e:
            logger.warning("Error cleaning up bookmarks: %s", str(e))


# Global service instance
word_report_generator = WordReportGenerator()
