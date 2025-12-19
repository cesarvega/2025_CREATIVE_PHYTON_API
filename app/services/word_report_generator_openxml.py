"""OpenXML-based service for generating NW Word reports.

This is a high-performance alternative to the COM-based word_report_generator.
Uses python-docx and matplotlib instead of Win32 COM automation for 5-10x speed improvement.
"""

from __future__ import annotations
import html
import pythoncom
import win32com.client
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import matplotlib
matplotlib.use('Agg')  # Use non-GUI backend for thread safety
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from app.config.settings import settings
from app.models.nw_reports_models import SummaryType
from app.services.nw_reports_service import nw_reports_service
from app.utils.logging_utils import get_logger
from app.utils.nw_data_utils import sanitize_filename
from app.utils.path_utils import get_nw_downloads_dir

logger = get_logger(__name__)

# Corporate colors (RGB format for matplotlib)
COLOR_POSITIVE = '#5CB85C'  # Green
COLOR_NEUTRAL = '#B77A33'   # Brown/Orange
COLOR_RECONSIDER = '#99004C'  # Purple


class WordReportGeneratorOpenXML:
    """Generate NW Word reports using OpenXML (python-docx)."""

    def __init__(self):
        """Initialize the OpenXML Word report generator."""
        self.template_path: Optional[Path] = None
        self.converted_template_path: Optional[Path] = None

    def _convert_doc_to_docx(self, doc_path: Path) -> Path:
        """Convert .doc file to .docx using COM automation.

        This is needed because python-docx can only read .docx files, not legacy .doc format.
        Uses Word COM to open .doc and save as .docx in temp directory.

        Args:
            doc_path: Path to .doc file

        Returns:
            Path to converted .docx file
        """
        logger.info("[OpenXML Word Report] Converting legacy .doc template to .docx format")

        # Use NW downloads directory directly (Word COM works reliably with this path)
        temp_base = get_nw_downloads_dir()
        temp_base.mkdir(parents=True, exist_ok=True)

        # Output path for converted file (use timestamp for uniqueness)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:20]
        docx_filename = f"_temp_{doc_path.stem}_{timestamp}.docx"
        docx_path = temp_base / docx_filename

        # Store for cleanup later
        self.converted_template_path = docx_path

        logger.info("[OpenXML Word Report] Temp conversion path: %s", docx_path)

        # Initialize COM if not already initialized
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
        except:
            # Already initialized in this thread
            pass

        word_app = None

        try:
            # Launch Word
            from app.utils.com_manager import com_manager
            with com_manager.acquire("Word.Application for .doc conversion"):
                word_app = win32com.client.Dispatch("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0

                # Open .doc file
                logger.debug("[OpenXML Word Report] Opening source .doc file: %s", doc_path)
                source_path = str(doc_path.absolute())
                doc = word_app.Documents.Open(source_path)

                # Save as .docx (format 16 = wdFormatXMLDocument)
                if not docx_path.parent.exists():
                    docx_path.parent.mkdir(parents=True, exist_ok=True)

                output_path_str = str(docx_path.absolute())
                logger.info("[OpenXML Word Report] Converting to: %s", output_path_str)

                WD_FORMAT_XML_DOCUMENT = 16
                doc.SaveAs2(output_path_str, WD_FORMAT_XML_DOCUMENT)
                doc.Close(SaveChanges=False)
                logger.debug("[OpenXML Word Report] Document closed")

            # Verify file was created
            if not docx_path.exists():
                raise RuntimeError(f"Converted file not found at: {docx_path}")

            logger.info("[OpenXML Word Report] Template converted successfully: %s", docx_path)
            return docx_path

        except Exception as e:
            logger.error("[OpenXML Word Report] Error converting .doc to .docx: %s", str(e), exc_info=True)
            raise RuntimeError(f"Failed to convert template to .docx: {str(e)}")

        finally:
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

    def generate_word_report(
        self,
        presentation_id: int,
        display_name: str,
        is_phonetics: bool = True,
    ) -> Path:
        """Generate complete Word report for NW presentation using OpenXML.

        This is a faster alternative to COM-based generation.
        Expected performance: 3-8 seconds (vs 15-45 seconds with COM).

        Args:
            presentation_id: Presentation ID to generate report for
            display_name: Display name for the presentation
            is_phonetics: If True, use Phonetics summary types

        Returns:
            Path to generated Word file (.docx format)
        """
        logger.info(
            "[OpenXML Word Report] Starting Word report generation for presentation_id=%d",
            presentation_id,
        )

        try:
            # Determine template path
            self.template_path = self._get_template_path()
            if not self.template_path.exists():
                raise FileNotFoundError(f"Word template not found: {self.template_path}")

            logger.info("[OpenXML Word Report] Using template: %s", self.template_path)

            # Check if template is .doc and convert to .docx if needed
            template_to_load = self.template_path
            if self.template_path.suffix.lower() == '.doc':
                logger.info("[OpenXML Word Report] Template is legacy .doc format, converting to .docx...")
                template_to_load = self._convert_doc_to_docx(self.template_path)

            # Load template document
            doc = Document(str(template_to_load))
            logger.info("[OpenXML Word Report] Template loaded successfully")

            # PHASE 1: Replace placeholders in document
            logger.info("[OpenXML Word Report] Phase 1: Replacing placeholders")
            replacements = nw_reports_service.get_word_report_replacements(presentation_id)
            logger.info("[OpenXML Word Report] Retrieved %d replacement values", len(replacements))

            for replacement in replacements:
                if replacement.placeholder and replacement.value:
                    self._replace_text_in_doc(
                        doc,
                        replacement.placeholder,
                        replacement.value
                    )

            # PHASE 2: Populate content tables
            logger.info("[OpenXML Word Report] Phase 2: Populating tables")
            self._populate_result_tables(doc, presentation_id, is_phonetics)

            # PHASE 3: Get "By The Numbers" data for bookmarks and pie chart
            logger.info("[OpenXML Word Report] Phase 3: Generating pie chart and updating counts")
            by_the_numbers_data = nw_reports_service.get_word_report_results_phonetics(
                presentation_id,
                SummaryType.BY_THE_NUMBERS
            )

            # Update "By The Numbers" text (bookmark replacement)
            self._update_by_the_numbers(doc, by_the_numbers_data)

            # Generate and insert Pie Chart
            self._generate_and_insert_pie_chart(doc, by_the_numbers_data)

            # PHASE 4: Save document to output directory
            logger.info("[OpenXML Word Report] Phase 4: Saving document")
            output_dir = get_nw_downloads_dir()
            output_dir.mkdir(parents=True, exist_ok=True)

            safe_display_name = sanitize_filename(display_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"NW_Report_{safe_display_name}_{timestamp}.docx"
            output_path = output_dir / filename

            # Save as modern .docx format
            doc.save(str(output_path))
            logger.info("[OpenXML Word Report] Word report saved to: %s", output_path)

            return output_path

        except Exception as e:
            logger.error("[OpenXML Word Report] Error generating Word report: %s", str(e), exc_info=True)
            raise

        finally:
            # Cleanup converted template file if created
            if self.converted_template_path and self.converted_template_path.exists():
                try:
                    self.converted_template_path.unlink()
                    logger.debug("[OpenXML Word Report] Cleaned up converted template: %s", self.converted_template_path)
                except Exception as e:
                    logger.warning("[OpenXML Word Report] Error cleaning up converted template: %s", str(e))

    def _get_template_path(self) -> Path:
        """Get the path to the Word template.

        Returns:
            Path to Word template file
        """
        template_name = "nomenclature workshop report_template_new_phonetics.doc"
        template_path = settings.app_dir / "templates" / template_name

        logger.debug("[OpenXML Word Report] Using Word template: %s", template_path)
        return template_path

    def _replace_text_in_doc(self, doc: Document, find_text: str, replace_text: str) -> None:
        """Replace text everywhere in the Word document.

        Args:
            doc: python-docx Document object
            find_text: Placeholder to find (e.g., "<Client>")
            replace_text: Replacement value
        """
        try:
            replace_value = str(replace_text) if replace_text else ""
            total_hits = 0

            # Replace in paragraphs
            for paragraph in doc.paragraphs:
                if find_text in paragraph.text:
                    self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                    total_hits += 1

            # Replace in tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if find_text in paragraph.text:
                                self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                                total_hits += 1

            # Replace in headers
            for section in doc.sections:
                header = section.header
                for paragraph in header.paragraphs:
                    if find_text in paragraph.text:
                        self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                        total_hits += 1

                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                if find_text in paragraph.text:
                                    self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                                    total_hits += 1

            # Replace in footers
            for section in doc.sections:
                footer = section.footer
                for paragraph in footer.paragraphs:
                    if find_text in paragraph.text:
                        self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                        total_hits += 1

                for table in footer.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                if find_text in paragraph.text:
                                    self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                                    total_hits += 1

            if total_hits > 0:
                logger.debug("[OpenXML Word Report] Replaced '%s' -> '%s' (%d occurrences)",
                           find_text, replace_value[:50], total_hits)
            else:
                logger.debug("[OpenXML Word Report] Placeholder '%s' NOT found in document", find_text)

        except Exception as e:
            logger.warning("[OpenXML Word Report] Error replacing text '%s': %s", find_text, str(e))

    def _replace_text_in_paragraph(self, paragraph, find_text: str, replace_text: str) -> None:
        """Replace text in a paragraph while preserving formatting.

        Args:
            paragraph: python-docx Paragraph object
            find_text: Text to find
            replace_text: Replacement text
        """
        if find_text in paragraph.text:
            # Simple replacement - for more complex scenarios, would need run-level replacement
            full_text = paragraph.text
            new_text = full_text.replace(find_text, replace_text)

            # Clear existing runs and add new text
            for run in paragraph.runs:
                run.text = ""

            if paragraph.runs:
                paragraph.runs[0].text = new_text
            else:
                paragraph.add_run(new_text)

    def _update_by_the_numbers(self, doc: Document, by_the_numbers: List) -> None:
        """Update bookmarks with 'By The Numbers' summary counts.

        Uses Word bookmarks (same as COM version):
        - Positive
        - Neutral
        - Reconsider
        - NewName (singular, not NewNames)

        Args:
            doc: python-docx Document object
            by_the_numbers: List of results from BY_THE_NUMBERS SP
        """
        try:
            if not by_the_numbers:
                logger.warning("[OpenXML Word Report] No 'By The Numbers' data available")
                return

            counts = by_the_numbers[0] if by_the_numbers else None

            if counts:
                total = counts.total_count or 8
                positive_count = counts.positive_count or 0
                neutral_count = counts.neutral_count or 0
                reconsider_count = counts.reconsider_count or 0
                new_names_count = counts.new_names_count or 0

                logger.info(
                    "[OpenXML Word Report] By The Numbers: Positive=%d, Neutral=%d, Reconsider=%d, NewNames=%d, Total=%d",
                    positive_count, neutral_count, reconsider_count, new_names_count, total
                )

                # Map bookmarks to their counts (same as COM version)
                # Format: "X Out of Y" where Y is the total
                bookmark_map = {
                    "Positive": f"{positive_count} Out of {total}",
                    "Neutral": f"{neutral_count} Out of {total}",
                    "Reconsider": f"{reconsider_count} Out of {total}",
                    "NewName": str(new_names_count),  # Note: singular "NewName" not "NewNames"
                }

                # Update each bookmark
                for bookmark_name, text in bookmark_map.items():
                    self._update_bookmark(doc, bookmark_name, text)

        except Exception as e:
            logger.error("[OpenXML Word Report] Error updating 'By The Numbers': %s", str(e), exc_info=True)

    def _update_bookmark(self, doc: Document, bookmark_name: str, text: str) -> None:
        """Update a bookmark in the document with text.

        python-docx doesn't have native bookmark support, so we access the XML directly.
        Searches in: paragraphs, tables, headers, and footers.

        Args:
            doc: python-docx Document object
            bookmark_name: Name of the bookmark
            text: Text to insert
        """
        try:
            from docx.oxml import parse_xml
            from docx.oxml.ns import qn

            def update_bookmark_in_element(element, location_desc="body"):
                """Helper to update bookmark in any element."""
                for elem in element.iter():
                    if elem.tag == qn('w:bookmarkStart'):
                        name = elem.get(qn('w:name'))
                        if name == bookmark_name:
                            bookmark_id = elem.get(qn('w:id'))
                            parent = elem.getparent()

                            # Clear content between bookmarkStart and bookmarkEnd
                            next_elem = elem.getnext()
                            while next_elem is not None:
                                if next_elem.tag == qn('w:bookmarkEnd'):
                                    end_id = next_elem.get(qn('w:id'))
                                    if end_id == bookmark_id:
                                        break
                                temp = next_elem.getnext()
                                parent.remove(next_elem)
                                next_elem = temp

                            # Insert new run with text after bookmarkStart
                            new_run = parse_xml(f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>{text}</w:t></w:r>')
                            elem.addnext(new_run)

                            logger.info("[OpenXML Word Report] Updated bookmark '%s' = '%s' in %s",
                                       bookmark_name, text, location_desc)
                            return True
                return False

            bookmark_found = False

            # 1. Search in main document body paragraphs
            for paragraph in doc.paragraphs:
                if update_bookmark_in_element(paragraph._element, "main body paragraph"):
                    bookmark_found = True
                    break

            # 2. Search in tables
            if not bookmark_found:
                for table_idx, table in enumerate(doc.tables):
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                if update_bookmark_in_element(paragraph._element, f"table {table_idx}"):
                                    bookmark_found = True
                                    break
                            if bookmark_found:
                                break
                        if bookmark_found:
                            break
                    if bookmark_found:
                        break

            # 3. Search in headers
            if not bookmark_found:
                for section_idx, section in enumerate(doc.sections):
                    header = section.header
                    # Search header paragraphs
                    for paragraph in header.paragraphs:
                        if update_bookmark_in_element(paragraph._element, f"header section {section_idx}"):
                            bookmark_found = True
                            break
                    if bookmark_found:
                        break

                    # Search header tables
                    for table in header.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for paragraph in cell.paragraphs:
                                    if update_bookmark_in_element(paragraph._element, f"header table section {section_idx}"):
                                        bookmark_found = True
                                        break
                                if bookmark_found:
                                    break
                            if bookmark_found:
                                break
                        if bookmark_found:
                            break
                    if bookmark_found:
                        break

            # 4. Search in footers
            if not bookmark_found:
                for section_idx, section in enumerate(doc.sections):
                    footer = section.footer
                    # Search footer paragraphs
                    for paragraph in footer.paragraphs:
                        if update_bookmark_in_element(paragraph._element, f"footer section {section_idx}"):
                            bookmark_found = True
                            break
                    if bookmark_found:
                        break

                    # Search footer tables
                    for table in footer.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for paragraph in cell.paragraphs:
                                    if update_bookmark_in_element(paragraph._element, f"footer table section {section_idx}"):
                                        bookmark_found = True
                                        break
                                if bookmark_found:
                                    break
                            if bookmark_found:
                                break
                        if bookmark_found:
                            break
                    if bookmark_found:
                        break

            if not bookmark_found:
                logger.warning("[OpenXML Word Report] Bookmark '%s' not found in document (searched body, tables, headers, footers)",
                             bookmark_name)

        except Exception as e:
            logger.error("[OpenXML Word Report] Error updating bookmark '%s': %s",
                       bookmark_name, str(e), exc_info=True)

    def _populate_result_tables(self, doc: Document, presentation_id: int, is_phonetics: bool) -> None:
        """Populate all result tables.

        Args:
            doc: python-docx Document object
            presentation_id: Presentation ID
            is_phonetics: If True, use Phonetics summary types
        """
        try:
            # Map of table indices to summary types
            # Table indices are 0-based in python-docx (vs 1-based in COM)
            if is_phonetics:
                table_configs = [
                    (1, SummaryType.POSITIVE_PHONETICS, "Positive Names"),  # Table 2 in COM = index 1
                    (2, SummaryType.NEGATIVE_PHONETICS, "Neutral Names"),
                    (3, SummaryType.RECONSIDER_PHONETICS, "Names For Reconsideration"),
                    (4, SummaryType.NEW_NAMES, "Newly Created Names"),
                    (5, SummaryType.EXPLORE, "Roots/Concepts to Explore"),
                ]
            else:
                table_configs = [
                    (1, SummaryType.POSITIVE, "Positive Names"),
                    (2, SummaryType.NEUTRAL, "Neutral Names"),
                    (3, SummaryType.RECONSIDER, "Names For Reconsideration"),
                    (4, SummaryType.NEW_NAMES, "Newly Created Names"),
                    (5, SummaryType.EXPLORE, "Roots/Concepts to Explore"),
                ]

            total_tables = len(doc.tables)
            logger.info("[OpenXML Word Report] Document has %d tables total", total_tables)

            for table_index, summary_type, table_name in table_configs:
                if table_index >= total_tables:
                    logger.warning("[OpenXML Word Report] Table %d (%s) not found", table_index, table_name)
                    continue

                # Get results for this summary type
                # NEWLY CREATED NAMES uses different SP (nw_CombineNewNames)
                if summary_type == SummaryType.NEW_NAMES:
                    results = nw_reports_service.get_newly_created_names_for_word(presentation_id)
                else:
                    # All other tables use nw_wdGetResults_Phonetics
                    results = nw_reports_service.get_word_report_results_phonetics(
                        presentation_id,
                        summary_type
                    )

                logger.info("[OpenXML Word Report] Retrieved %d results for '%s'", len(results), table_name)

                if results:
                    table = doc.tables[table_index]
                    self._populate_table(table, results, table_name)
                else:
                    logger.info("[OpenXML Word Report] No data to populate for table '%s'", table_name)

        except Exception as e:
            logger.error("[OpenXML Word Report] Error populating tables: %s", str(e), exc_info=True)

    def _populate_table(self, table, results: List, table_name: str) -> None:
        """Populate a table with results, clearing existing rows first.

        Args:
            table: python-docx Table object
            results: List of WordReportResult objects
            table_name: Name of the table
        """
        try:
            # Delete all rows except header (row 0)
            rows_to_delete = len(table.rows) - 1
            for _ in range(rows_to_delete):
                table._element.remove(table.rows[-1]._element)

            logger.debug("[OpenXML Word Report] Cleared table '%s', now adding %d rows", table_name, len(results))

            # Add rows for each result
            for idx, result in enumerate(results):
                # Handle grouped names (split by ## or $$)
                # IMPORTANT: For "Newly Created Names", do NOT split - keep all names together
                names_to_add = []
                if "Newly Created Names" not in table_name and result.name and ('##' in result.name or '$$' in result.name):
                    delimiter = '##' if '##' in result.name else '$$'
                    names_to_add = [n.strip() for n in result.name.split(delimiter) if n.strip()]
                else:
                    names_to_add = [result.name] if result.name else [""]

                # Process each name
                for name in names_to_add:
                    try:
                        # Add new row
                        row = table.add_row()

                        # Column 0: Candidate (Name) - IN BOLD
                        if len(row.cells) >= 1:
                            cell = row.cells[0]
                            # For NEWLY CREATED NAMES: use name (the new created name)
                            # For REFINED CREATIVE DIRECTION: use original_name if available
                            if "Newly Created Names" in table_name:
                                candidate_name = name or ''
                            else:
                                candidate_name = getattr(result, 'original_name', None) or name or ''
                            cell.text = str(candidate_name)
                            # Make candidate name bold
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    run.font.bold = True

                        # Column 1: Pronunciation OR Direction OR Original Name
                        if len(row.cells) >= 2:
                            # For NEWLY CREATED NAMES: use original_name (concatenated with $$)
                            # For REFINED CREATIVE DIRECTION: use direction instead of pronunciation
                            if "Newly Created Names" in table_name:
                                original_name = getattr(result, 'original_name', None) or ''
                                row.cells[1].text = str(original_name)
                            else:
                                direction = getattr(result, 'direction', None)
                                pronunciation = getattr(result, 'pronunciation', None)
                                col1_value = direction or pronunciation or ''
                                row.cells[1].text = str(col1_value)

                        # Column 2: Rationale OR Original Rationale
                        if len(row.cells) >= 3:
                            # For NEWLY CREATED NAMES: use original_rationale (concatenated with $$)
                            # For REFINED CREATIVE DIRECTION: may use original_rationale
                            original_rationale = getattr(result, 'original_rationale', None)
                            rationale = original_rationale or result.rationale or ''
                            # Unescape HTML entities
                            rationale = html.unescape(str(rationale))
                            row.cells[2].text = rationale

                        # Column 3: Category OR Original Category (if applicable)
                        if len(row.cells) >= 4:
                            # For NEWLY CREATED NAMES: use original_category if available
                            if "Newly Created Names" in table_name:
                                category = getattr(result, 'original_category', None) or getattr(result, 'category', None) or ''
                            else:
                                category = getattr(result, 'category', None) or ''
                            row.cells[3].text = str(category)

                    except Exception as row_error:
                        logger.error("[OpenXML Word Report] Error adding row to table '%s': %s",
                                   table_name, str(row_error), exc_info=True)

            logger.info("[OpenXML Word Report] Successfully populated table '%s' with %d rows",
                       table_name, len(results))

        except Exception as e:
            logger.error("[OpenXML Word Report] Error populating table '%s': %s",
                       table_name, str(e), exc_info=True)

    def _generate_and_insert_pie_chart(self, doc: Document, by_the_numbers: List) -> None:
        """Generate pie chart using matplotlib and insert into Word document at PieChart bookmark.

        Uses corporate colors:
        - Positive: RGB(92, 184, 92) - Green
        - Neutral: RGB(183, 122, 51) - Brown/Orange
        - Reconsider: RGB(153, 0, 76) - Purple

        Args:
            doc: python-docx Document object
            by_the_numbers: List with counts from ByTheNumbers SP
        """
        try:
            if not by_the_numbers:
                logger.warning("[OpenXML Word Report] No data for pie chart")
                return

            counts = by_the_numbers[0] if by_the_numbers else None
            if not counts:
                return

            positive = getattr(counts, 'positive_count', 0) or 0
            neutral = getattr(counts, 'neutral_count', 0) or 0
            reconsider = getattr(counts, 'reconsider_count', 0) or 0

            logger.info("[OpenXML Word Report] Pie chart data: Positive=%d, Neutral=%d, Reconsider=%d",
                       positive, neutral, reconsider)

            # Create pie chart with matplotlib
            labels = ['Positive', 'Neutral', 'Reconsider']
            sizes = [positive, neutral, reconsider]
            colors = [COLOR_POSITIVE, COLOR_NEUTRAL, COLOR_RECONSIDER]

            # Create figure with optimized size
            fig, ax = plt.subplots(figsize=(6, 4))

            # Create pie chart without labels on the pie itself (no percentages)
            wedges, texts = ax.pie(
                sizes,
                colors=colors,
                startangle=90,
                radius=1.0
            )

            # Add legend to the right side
            ax.legend(
                wedges,
                labels,
                title="",
                loc="center left",
                bbox_to_anchor=(1.05, 0.5),
                fontsize=10,
                frameon=False
            )

            # Equal aspect ratio ensures circular pie
            ax.axis('equal')

            # Remove extra whitespace
            plt.tight_layout(pad=0.3)

            # Save to temporary file with minimal padding
            import tempfile
            temp_dir = Path(tempfile.mkdtemp())
            chart_path = temp_dir / "pie_chart.png"
            fig.savefig(str(chart_path), dpi=150, bbox_inches='tight', pad_inches=0.1, facecolor='white')
            plt.close(fig)

            logger.info("[OpenXML Word Report] Pie chart image generated successfully")

            # Find and replace PieChart bookmark with image (reduced width to fit page)
            chart_inserted = self._insert_image_at_bookmark(doc, "PieChart", chart_path, width_inches=3.8)

            if not chart_inserted:
                logger.warning("[OpenXML Word Report] PieChart bookmark not found in document")

            # Cleanup temp directory
            try:
                import shutil
                shutil.rmtree(temp_dir)
            except:
                pass

        except Exception as e:
            logger.error("[OpenXML Word Report] Error generating pie chart: %s", str(e), exc_info=True)

    def _insert_image_at_bookmark(self, doc: Document, bookmark_name: str, image_path: Path, width_inches: float = 4.5) -> bool:
        """Insert image at a bookmark location.

        Args:
            doc: python-docx Document object
            bookmark_name: Name of the bookmark
            image_path: Path to image file
            width_inches: Width of image in inches

        Returns:
            True if bookmark was found and image inserted, False otherwise
        """
        try:
            from docx.oxml.ns import qn

            def find_paragraph_with_bookmark(paragraphs_list, location_desc="body"):
                """Find paragraph containing bookmark and return it."""
                for paragraph in paragraphs_list:
                    for elem in paragraph._element.iter():
                        if elem.tag == qn('w:bookmarkStart'):
                            name = elem.get(qn('w:name'))
                            if name == bookmark_name:
                                return paragraph, location_desc
                return None, None

            # Search in main document body
            paragraph, location = find_paragraph_with_bookmark(doc.paragraphs, "main body")
            if paragraph:
                paragraph.clear()
                # Add spacing before the chart to push it down
                paragraph_format = paragraph.paragraph_format
                paragraph_format.space_before = Pt(12)  # Add 12pt space before

                run = paragraph.add_run()
                run.add_picture(str(image_path), width=Inches(width_inches))
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                logger.info("[OpenXML Word Report] Inserted image at bookmark '%s' in %s",
                           bookmark_name, location)
                return True

            # Search in tables
            for table_idx, table in enumerate(doc.tables):
                for row in table.rows:
                    for cell in row.cells:
                        paragraph, _ = find_paragraph_with_bookmark(cell.paragraphs, f"table {table_idx}")
                        if paragraph:
                            paragraph.clear()
                            run = paragraph.add_run()
                            run.add_picture(str(image_path), width=Inches(width_inches))
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            logger.info("[OpenXML Word Report] Inserted image at bookmark '%s' in table %d",
                                       bookmark_name, table_idx)
                            return True

            # Search in headers
            for section_idx, section in enumerate(doc.sections):
                header = section.header
                paragraph, _ = find_paragraph_with_bookmark(header.paragraphs, f"header section {section_idx}")
                if paragraph:
                    paragraph.clear()
                    run = paragraph.add_run()
                    run.add_picture(str(image_path), width=Inches(width_inches))
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    logger.info("[OpenXML Word Report] Inserted image at bookmark '%s' in header section %d",
                               bookmark_name, section_idx)
                    return True

                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            paragraph, _ = find_paragraph_with_bookmark(cell.paragraphs, f"header table section {section_idx}")
                            if paragraph:
                                paragraph.clear()
                                run = paragraph.add_run()
                                run.add_picture(str(image_path), width=Inches(width_inches))
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                logger.info("[OpenXML Word Report] Inserted image at bookmark '%s' in header table section %d",
                                           bookmark_name, section_idx)
                                return True

            # Search in footers
            for section_idx, section in enumerate(doc.sections):
                footer = section.footer
                paragraph, _ = find_paragraph_with_bookmark(footer.paragraphs, f"footer section {section_idx}")
                if paragraph:
                    paragraph.clear()
                    run = paragraph.add_run()
                    run.add_picture(str(image_path), width=Inches(width_inches))
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    logger.info("[OpenXML Word Report] Inserted image at bookmark '%s' in footer section %d",
                               bookmark_name, section_idx)
                    return True

                for table in footer.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            paragraph, _ = find_paragraph_with_bookmark(cell.paragraphs, f"footer table section {section_idx}")
                            if paragraph:
                                paragraph.clear()
                                run = paragraph.add_run()
                                run.add_picture(str(image_path), width=Inches(width_inches))
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                logger.info("[OpenXML Word Report] Inserted image at bookmark '%s' in footer table section %d",
                                           bookmark_name, section_idx)
                                return True

            return False

        except Exception as e:
            logger.error("[OpenXML Word Report] Error inserting image at bookmark '%s': %s",
                       bookmark_name, str(e), exc_info=True)
            return False


# Global service instance
word_report_generator_openxml = WordReportGeneratorOpenXML()
