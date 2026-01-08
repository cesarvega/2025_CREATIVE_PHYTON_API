"""OpenXML-based service for generating NW Feedback Template documents.

This is a high-performance alternative to the COM-based feedback_template_generator.
Uses python-docx and matplotlib instead of Win32 COM automation for 5-10x speed improvement.
"""

from __future__ import annotations
import html
import io
import re
import tempfile
import pythoncom
import win32com.client
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Dict, Any

import matplotlib
matplotlib.use('Agg')  # Use non-GUI backend for thread safety
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.shared import qn as qn_helper

from app.config.settings import settings
from app.constants import TemplateFilename
from app.models.nw_reports_models import SummaryType
from app.services.bi_guidelines_service import bi_guidelines_service
from app.services.nw_reports_service import nw_reports_service
from app.utils.logging_utils import get_logger
from app.utils.nw_data_utils import sanitize_filename
from app.utils.path_utils import get_nw_downloads_dir

logger = get_logger(__name__)

# Corporate colors (RGB format for matplotlib)
COLOR_POSITIVE = '#5CB85C'  # Green
COLOR_NEUTRAL = '#B77A33'   # Brown/Orange
COLOR_RECONSIDER = '#99004C'  # Purple


class FeedbackTemplateGeneratorOpenXML:
    """Generate NW Feedback Template documents using OpenXML (python-docx)."""

    def __init__(self):
        """Initialize the OpenXML feedback template generator."""
        self.template_path: Optional[Path] = None
        self.temp_dir: Optional[Path] = None
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
        logger.info("[OpenXML] Converting legacy .doc template to .docx format")

        # Use NW downloads directory directly (Word COM works reliably with this path)
        # Avoid subdirectories to prevent permission issues
        temp_base = get_nw_downloads_dir()
        temp_base.mkdir(parents=True, exist_ok=True)

        # Output path for converted file (use timestamp for uniqueness)
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:20]  # Include microseconds for uniqueness
        docx_filename = f"_temp_{doc_path.stem}_{timestamp}.docx"
        docx_path = temp_base / docx_filename

        # Store for cleanup later
        self.converted_template_path = docx_path

        logger.info("[OpenXML] Temp conversion path: %s", docx_path)

        # Initialize COM if not already initialized
        # CoInitialize is safe to call multiple times (increments refcount)
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
                logger.debug("[OpenXML] Opening source .doc file: %s", doc_path)
                source_path = str(doc_path.absolute())
                doc = word_app.Documents.Open(source_path)

                # Save as .docx (format 16 = wdFormatXMLDocument)
                # Ensure output directory exists
                if not docx_path.parent.exists():
                    docx_path.parent.mkdir(parents=True, exist_ok=True)
                    logger.debug("[OpenXML] Created output directory: %s", docx_path.parent)

                # Convert path to absolute string (Word COM expects full paths)
                output_path_str = str(docx_path.absolute())
                logger.info("[OpenXML] Converting to: %s", output_path_str)

                # Save using Word COM
                # Format 16 = wdFormatXMLDocument (.docx)
                WD_FORMAT_XML_DOCUMENT = 16

                try:
                    doc.SaveAs2(output_path_str, WD_FORMAT_XML_DOCUMENT)
                    logger.info("[OpenXML] SaveAs2 completed successfully")
                except Exception as save_error:
                    logger.error("[OpenXML] SaveAs2 failed: %s", str(save_error))
                    logger.error("[OpenXML] Attempted path: %s", output_path_str)
                    logger.error("[OpenXML] Parent exists: %s", docx_path.parent.exists())
                    logger.error("[OpenXML] Parent is dir: %s", docx_path.parent.is_dir())
                    raise

                doc.Close(SaveChanges=False)
                logger.debug("[OpenXML] Document closed")

            # Verify file was created
            if not docx_path.exists():
                raise RuntimeError(f"Converted file not found at: {docx_path}")

            logger.info("[OpenXML] Template converted successfully: %s", docx_path)
            return docx_path

        except Exception as e:
            logger.error("[OpenXML] Error converting .doc to .docx: %s", str(e), exc_info=True)
            raise RuntimeError(f"Failed to convert template to .docx: {str(e)}")

        finally:
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            # Only uninitialize if we initialized it
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

    def generate_feedback_template(
        self,
        presentation_id: int,
        display_name: str,
        progress_callback=None,
    ) -> Path:
        """Generate complete Feedback Template document for NW presentation using OpenXML.

        This is a faster, more reliable alternative to COM-based generation.
        Expected performance: 3-5 seconds (vs 15-30 seconds with COM).

        Steps:
        1. Determine presentation type (Normal, Phonetics, Katakana) and select template
        2. Load template using python-docx
        3. Replace placeholders with values from nw_wdValuesToReplace
        4. Populate tables with data from nw_wdToIndicateFeedback
        5. Generate and insert pie chart using matplotlib
        6. Clean up bookmarks (if supported)
        7. Save document

        Args:
            presentation_id: Presentation ID to generate template for
            display_name: Display name for the presentation
            progress_callback: Optional callback function to report progress (int 0-100)

        Returns:
            Path to generated Word file (.docx format)
        """
        logger.info(
            "[OpenXML] Starting Feedback Template generation for presentation_id=%d",
            presentation_id,
        )

        try:
            # Get presentation info to determine type (30-40%)
            if progress_callback:
                progress_callback(30)
            presentation_info = bi_guidelines_service.get_project_info(presentation_id)
            if not presentation_info:
                raise ValueError(f"Presentation with ID {presentation_id} not found")

            presentation_type = presentation_info.get("PresentationType", "Normal")
            logger.info("[OpenXML] Presentation type: %s", presentation_type)

            # Determine template path based on presentation type (40-45%)
            if progress_callback:
                progress_callback(40)
            self.template_path = self._get_template_path(presentation_type)
            if not self.template_path.exists():
                raise FileNotFoundError(f"Feedback template not found: {self.template_path}")

            logger.info("[OpenXML] Using template: %s", self.template_path)

            # Check if template is .doc (legacy format) and convert to .docx if needed (45-48%)
            if progress_callback:
                progress_callback(45)

            template_to_load = self.template_path
            if self.template_path.suffix.lower() == '.doc':
                logger.info("[OpenXML] Template is legacy .doc format, converting to .docx...")
                template_to_load = self._convert_doc_to_docx(self.template_path)
                self.converted_template_path = template_to_load
                if progress_callback:
                    progress_callback(48)

            # Load template document (48-50%)
            doc = Document(str(template_to_load))
            logger.info("[OpenXML] Template loaded successfully")

            # PHASE 1: Replace placeholders in document (50-70%)
            if progress_callback:
                progress_callback(50)
            logger.info("[OpenXML] Phase 1: Replacing placeholders")
            replacements = nw_reports_service.get_word_report_replacements(presentation_id)
            logger.info("[OpenXML] Retrieved %d replacement values", len(replacements))

            for replacement in replacements:
                if replacement.placeholder and replacement.value:
                    self._replace_text_in_doc(
                        doc,
                        replacement.placeholder,
                        replacement.value
                    )

            if progress_callback:
                progress_callback(70)

            # Format header with correct font size and style
            logger.info("[OpenXML] Formatting header")
            self._format_document_header(doc)

            # PHASE 2: Populate tables with results (70-80%)
            logger.info("[OpenXML] Phase 2: Populating tables")
            self._populate_feedback_tables(doc, presentation_id)

            if progress_callback:
                progress_callback(80)

            # PHASE 3: Get "By The Numbers" data for pie chart (80-85%)
            logger.info("[OpenXML] Phase 3: Generating pie chart")
            by_the_numbers_data = nw_reports_service.get_word_report_results_phonetics(
                presentation_id,
                SummaryType.BY_THE_NUMBERS
            )

            # Generate and insert Pie Chart
            self._generate_and_insert_pie_chart(doc, by_the_numbers_data)

            if progress_callback:
                progress_callback(85)

            # PHASE 4: Clean up (bookmarks not fully supported in python-docx)
            logger.info("[OpenXML] Phase 4: Cleanup")
            # Note: python-docx has limited bookmark support, skip for now

            # Save document to output directory (90-95%)
            if progress_callback:
                progress_callback(90)
            output_dir = get_nw_downloads_dir()
            output_dir.mkdir(parents=True, exist_ok=True)

            safe_display_name = sanitize_filename(display_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Feedback_Template_{safe_display_name}_{timestamp}.docx"
            output_path = output_dir / filename

            # Save as modern .docx format
            doc.save(str(output_path))
            logger.info("[OpenXML] Document saved successfully to: %s", output_path)

            if progress_callback:
                progress_callback(100)

            return output_path

        except Exception as e:
            logger.error("[OpenXML] Error generating Feedback Template: %s", str(e), exc_info=True)
            raise

        finally:
            # Cleanup converted template file if created
            if self.converted_template_path and self.converted_template_path.exists():
                try:
                    self.converted_template_path.unlink()
                    logger.debug("[OpenXML] Cleaned up converted template: %s", self.converted_template_path)
                except Exception as e:
                    logger.warning("[OpenXML] Error cleaning up converted template: %s", str(e))

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
        logger.debug("[OpenXML] Using Feedback template: %s", template_path)
        return template_path

    def _replace_text_in_doc(self, doc: Document, find_text: str, replace_text: str) -> None:
        """Replace text everywhere in the Word document.

        Covers paragraphs, tables, headers, and footers.

        Args:
            doc: python-docx Document object
            find_text: Placeholder to find (e.g., "<Client>")
            replace_text: Replacement value
        """
        try:
            replace_value = str(replace_text) if replace_text else ""
            total_hits = 0

            # 1) Replace in paragraphs (main body)
            for paragraph in doc.paragraphs:
                if find_text in paragraph.text:
                    self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                    total_hits += 1

            # 2) Replace in tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if find_text in paragraph.text:
                                self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                                total_hits += 1

            # 3) Replace in headers
            for section in doc.sections:
                header = section.header
                for paragraph in header.paragraphs:
                    if find_text in paragraph.text:
                        self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                        total_hits += 1

                # Replace in header tables
                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                if find_text in paragraph.text:
                                    self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                                    total_hits += 1

            # 4) Replace in footers
            for section in doc.sections:
                footer = section.footer
                for paragraph in footer.paragraphs:
                    if find_text in paragraph.text:
                        self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                        total_hits += 1

                # Replace in footer tables
                for table in footer.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                if find_text in paragraph.text:
                                    self._replace_text_in_paragraph(paragraph, find_text, replace_value)
                                    total_hits += 1

            if total_hits > 0:
                logger.debug("[OpenXML] Replaced '%s' -> '%s' (%d occurrences)",
                           find_text, replace_value[:50], total_hits)
            else:
                logger.debug("[OpenXML] Placeholder '%s' NOT found in document", find_text)

        except Exception as e:
            logger.warning("[OpenXML] Error replacing text '%s': %s", find_text, str(e))

    def _replace_text_in_paragraph(self, paragraph, find_text: str, replace_text: str) -> None:
        """Replace text in a paragraph while preserving formatting.

        Args:
            paragraph: python-docx Paragraph object
            find_text: Text to find
            replace_text: Replacement text
        """
        # Simple replacement preserving paragraph-level formatting
        # For more complex scenarios (preserving run-level formatting), would need more sophisticated approach
        if find_text in paragraph.text:
            # Get original formatting
            original_runs = list(paragraph.runs)

            # Replace text
            full_text = paragraph.text
            new_text = full_text.replace(find_text, replace_text)

            # Clear existing runs
            for run in original_runs:
                run.text = ""

            # Add new text in first run (or create new run if none exist)
            if paragraph.runs:
                paragraph.runs[0].text = new_text
            else:
                paragraph.add_run(new_text)

    def _format_document_header(self, doc: Document) -> None:
        """Format document header with correct font size (8.5pt) and bold, and add logo if available.

        Args:
            doc: python-docx Document object
        """
        try:
            logger.info("[OpenXML] Formatting document header")

            for section in doc.sections:
                header = section.header

                # Insert logo at the beginning of the header if logo file exists
                logo_path = settings.app_dir / "templates" / "bi_logo.png"
                if logo_path.exists():
                    try:
                        # Check if header has a table (common structure for Word headers)
                        if header.tables:
                            # Insert logo in first cell of first table
                            first_table = header.tables[0]
                            if first_table.rows:
                                first_cell = first_table.rows[0].cells[0]
                                # Set vertical alignment for the cell (center)
                                first_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                                # Get the first paragraph in the cell
                                if first_cell.paragraphs:
                                    logo_para = first_cell.paragraphs[0]
                                    # Save existing text
                                    existing_text = logo_para.text
                                    # Clear the paragraph
                                    logo_para.clear()
                                    
                                    # CRITICAL: Set paragraph spacing to zero
                                    logo_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                    logo_para.paragraph_format.space_before = Pt(0)
                                    logo_para.paragraph_format.space_after = Pt(0)
                                    logo_para.paragraph_format.line_spacing = 1.0
                                    
                                    # Set paragraph to align elements at bottom (baseline alignment)
                                    try:
                                        pPr = logo_para._element.get_or_add_pPr()
                                        # Remove any existing textAlignment
                                        for textAlign in pPr.findall(qn('w:textAlignment')):
                                            pPr.remove(textAlign)
                                        # Add textAlignment = bottom (aligns at baseline)
                                        textAlignment = OxmlElement('w:textAlignment')
                                        textAlignment.set(qn('w:val'), 'bottom')
                                        pPr.append(textAlignment)
                                    except Exception as e:
                                        logger.debug("[OpenXML] Could not set baseline alignment: %s", str(e))
                                    
                                    # Add logo inline - height matched to text size
                                    logo_run = logo_para.add_run()
                                    # Very small logo: 12pt height (same as text size approximately)
                                    picture = logo_run.add_picture(str(logo_path), height=Pt(12))
                                    
                                    # Add non-breaking space after logo for better spacing
                                    space_run = logo_para.add_run("\u00A0\u00A0")  # Two non-breaking spaces
                                    space_run.font.size = Pt(8.5)
                                    space_run.font.bold = True
                                    
                                    # Add back the existing text with proper formatting
                                    text_run = logo_para.add_run(existing_text)
                                    text_run.font.size = Pt(8.5)
                                    text_run.font.bold = True
                                    
                                    logger.info("[OpenXML] Logo inserted inline at 12pt height with baseline alignment")
                        else:
                            # Insert logo in first paragraph if no table
                            if header.paragraphs:
                                first_para = header.paragraphs[0]
                                # Save existing text
                                existing_text = first_para.text
                                # Clear the paragraph
                                first_para.clear()
                                # Set paragraph alignment and spacing
                                first_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                first_para.paragraph_format.space_before = Pt(0)
                                first_para.paragraph_format.space_after = Pt(0)
                                # Add logo as first run (inline with text)
                                logo_run = first_para.add_run()
                                # Set logo dimensions: width=0.67", height=0.38"
                                logo_run.add_picture(str(logo_path), width=Inches(0.67), height=Inches(0.38))
                                # Add small space after logo
                                space_run = first_para.add_run("  ")
                                space_run.font.size = Pt(8.5)
                                space_run.font.bold = True
                                # Add back the existing text with proper formatting
                                text_run = first_para.add_run(existing_text)
                                text_run.font.size = Pt(8.5)
                                text_run.font.bold = True
                                logger.info("[OpenXML] Logo inserted in header paragraph")
                    except Exception as logo_error:
                        logger.warning("[OpenXML] Error inserting logo: %s", str(logo_error))
                else:
                    logger.debug("[OpenXML] Logo file not found at: %s", logo_path)

                # Format all paragraphs in header
                for paragraph in header.paragraphs:
                    for run in paragraph.runs:
                        # Skip image runs (don't format logo)
                        if hasattr(run, '_element') and run._element.xpath('.//w:drawing'):
                            continue
                        # Set font size to 8.5pt
                        run.font.size = Pt(8.5)
                        # Set bold
                        run.font.bold = True
                        logger.debug("[OpenXML] Formatted header paragraph: %s", paragraph.text[:50])

                # Format text in header tables (if any)
                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    # Skip image runs (don't format logo)
                                    if hasattr(run, '_element') and run._element.xpath('.//w:drawing'):
                                        continue
                                    # Set font size to 8.5pt
                                    run.font.size = Pt(8.5)
                                    # Set bold
                                    run.font.bold = True
                                    logger.debug("[OpenXML] Formatted header table cell: %s", paragraph.text[:50])

            logger.info("[OpenXML] Header formatting completed")

        except Exception as e:
            logger.warning("[OpenXML] Error formatting header: %s", str(e))

    def _populate_feedback_tables(self, doc: Document, presentation_id: int) -> None:
        """Populate the feedback table using the nw_wdToIndicateFeedback SP.

        Args:
            doc: python-docx Document object
            presentation_id: Presentation ID
        """
        try:
            if not doc.tables:
                logger.error("[OpenXML] No tables found in document!")
                return

            # Get the main feedback table (should be first table)
            main_table = doc.tables[0]
            logger.info("[OpenXML] Found main table - Rows: %d, Columns: %d",
                       len(main_table.rows), len(main_table.columns))

            # Call the SP: nw_wdToIndicateFeedback
            logger.info("[OpenXML] Fetching feedback data using nw_wdToIndicateFeedback")

            from app.config.db import get_connection_scope

            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "{CALL [BI_GUIDELINES].[dbo].[nw_wdToIndicateFeedback](?)}",
                    (presentation_id,)
                )

                # Get the result set
                if cursor.description is None:
                    logger.warning("[OpenXML] nw_wdToIndicateFeedback returned no result set")
                    return

                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()

                logger.info("[OpenXML] Retrieved %d rows from nw_wdToIndicateFeedback", len(rows))
                logger.info("[OpenXML] Columns: %s", columns)

                if not rows:
                    logger.warning("[OpenXML] No data to populate - SP returned empty results")
                    return

                # Populate the table with the data
                self._populate_feedback_table_from_sp(main_table, columns, rows)

        except Exception as e:
            logger.error("[OpenXML] Error populating feedback tables: %s", str(e), exc_info=True)

    def _populate_feedback_table_from_sp(self, table, columns: List[str], rows: List) -> None:
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
            table: python-docx Table object
            columns: List of column names from SP
            rows: List of row tuples from SP
        """
        try:
            logger.info("[OpenXML] Starting to populate table with %d rows from SP", len(rows))

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

            # Delete all rows except header (row 0)
            # Keep header row, delete from index 1 onwards
            rows_to_delete = len(table.rows) - 1
            for _ in range(rows_to_delete):
                table._element.remove(table.rows[-1]._element)

            logger.info("[OpenXML] Cleared %d rows from table, now has %d rows",
                       rows_to_delete, len(table.rows))

            # Expand grouped rows (##) into 1 row per entry
            expanded_rows: List[dict] = []
            for row in rows:
                row_dict = dict(zip(columns, row))
                expanded_rows.extend(expand_grouped_rows(row_dict))

            logger.info("[OpenXML] Expanded %d SP rows into %d table rows (group delimiter: ##/$$)",
                       len(rows), len(expanded_rows))

            # Add rows for each result from SP (after expansion)
            for idx, row_dict in enumerate(expanded_rows):
                try:
                    # Add new row to table
                    new_row = table.add_row()

                    # Set minimum row height to make vertical centering visible
                    tr = new_row._element
                    trPr = tr.get_or_add_trPr()
                    
                    # Remove existing height if present
                    for existing_height in trPr.findall(qn('w:trHeight')):
                        trPr.remove(existing_height)
                    
                    # Set row height to at least 0.4 inches (~1 cm) to make centering visible
                    trHeight = OxmlElement('w:trHeight')
                    trHeight.set(qn('w:val'), '576')  # 576 twips = 0.4 inches
                    trHeight.set(qn('w:hRule'), 'atLeast')  # Allow expansion if needed
                    trPr.append(trHeight)

                    # Column 1: Row number - CENTER aligned
                    cell = new_row.cells[0]
                    cell.text = str(idx + 1)
                    para = cell.paragraphs[0]
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.bold = False
                        run.font.name = "Calibri"
                    # Apply vertical center using XML
                    self._set_cell_vertical_center(cell)

                    # Column 2: Test Name (Candidate) - IN BOLD, CENTER ALIGNED
                    cell = new_row.cells[1]
                    cell.text = str(row_dict.get('Test Name') or '')
                    para = cell.paragraphs[0]
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.bold = True
                        run.font.name = "Calibri"
                    # Apply vertical center using XML
                    self._set_cell_vertical_center(cell)

                    # Column 3: Pronunciation - CENTER ALIGNED
                    cell = new_row.cells[2]
                    cell.text = str(row_dict.get('Pronunciation') or '')
                    para = cell.paragraphs[0]
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.bold = False
                        run.font.name = "Calibri"
                    # Apply vertical center using XML
                    self._set_cell_vertical_center(cell)

                    # Column 4: Rationale - CENTER ALIGNED
                    cell = new_row.cells[3]
                    cell.text = str(row_dict.get('Rationale') or '')
                    para = cell.paragraphs[0]
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.bold = False
                        run.font.name = "Calibri"
                    # Apply vertical center using XML
                    self._set_cell_vertical_center(cell)

                    # Columns 5-8: Positive, Neutral, Negative, Comments/Suggestions - CENTER aligned
                    for col_idx in range(4, min(8, len(new_row.cells))):
                        cell = new_row.cells[col_idx]
                        cell.text = ""
                        para = cell.paragraphs[0]
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        # Ensure font formatting even for empty cells
                        if para.runs:
                            for run in para.runs:
                                run.font.size = Pt(10)
                                run.font.bold = False
                                run.font.name = "Calibri"
                        else:
                            # Add a run with proper formatting for empty cells
                            run = para.add_run()
                            run.font.size = Pt(10)
                            run.font.bold = False
                            run.font.name = "Calibri"
                        # Apply vertical center using XML
                        self._set_cell_vertical_center(cell)

                    # Set white background for ALL cells using XML
                    for cell in new_row.cells:
                        tcPr = cell._element.get_or_add_tcPr()
                        for shd in tcPr.findall(qn('w:shd')):
                            tcPr.remove(shd)
                        shading_elm = OxmlElement('w:shd')
                        shading_elm.set(qn('w:fill'), 'FFFFFF')
                        tcPr.append(shading_elm)

                    logger.debug("[OpenXML] Added row %d: %s", idx + 1, row_dict.get('Test Name', ''))

                except Exception as row_error:
                    logger.error("[OpenXML] Error adding row %d: %s", idx + 1, str(row_error), exc_info=True)

            logger.info("[OpenXML] Successfully populated table with %d rows", len(expanded_rows))

        except Exception as e:
            logger.error("[OpenXML] Error populating feedback table from SP: %s", str(e), exc_info=True)

    def _set_cell_vertical_center(self, cell) -> None:
        """Set cell vertical alignment to center using XML manipulation.

        Args:
            cell: python-docx Cell object
        """
        try:
            tc = cell._element
            tcPr = tc.get_or_add_tcPr()

            # Remove existing vAlign elements
            for vAlign in tcPr.findall(qn('w:vAlign')):
                tcPr.remove(vAlign)

            # Add new vAlign element at the beginning for higher precedence
            vAlign = OxmlElement('w:vAlign')
            vAlign.set(qn('w:val'), 'center')
            tcPr.insert(0, vAlign)

            # Set cell margins to provide space for vertical centering
            tcMar = tcPr.find(qn('w:tcMar'))
            if tcMar is None:
                tcMar = OxmlElement('w:tcMar')
                tcPr.append(tcMar)
            else:
                # Clear existing margins
                for child in list(tcMar):
                    tcMar.remove(child)
            
            # Set balanced top/bottom margins (100 twips = ~7pt padding)
            for margin_name in ['top', 'bottom']:
                margin = OxmlElement(f'w:{margin_name}')
                margin.set(qn('w:w'), '100')  # 100 twips = ~7pt
                margin.set(qn('w:type'), 'dxa')
                tcMar.append(margin)
            
            # Smaller left/right margins
            for margin_name in ['left', 'right']:
                margin = OxmlElement(f'w:{margin_name}')
                margin.set(qn('w:w'), '50')  # 50 twips = ~3.5pt
                margin.set(qn('w:type'), 'dxa')
                tcMar.append(margin)

        except Exception as e:
            logger.debug("[OpenXML] Error setting vertical alignment: %s", str(e))

    def _set_cell_white_background(self, cell) -> None:
        """Set cell background to white using XML manipulation.

        Args:
            cell: python-docx Cell object
        """
        try:
            tcPr = cell._element.get_or_add_tcPr()

            # Remove existing shading if present
            for shd in tcPr.findall(qn('w:shd')):
                tcPr.remove(shd)

            # Add white background
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), 'FFFFFF')
            tcPr.append(shading_elm)

        except Exception as e:
            logger.debug("[OpenXML] Error setting white background: %s", str(e))

    def _copy_cell_style(self, source_cell, target_cell, is_header: bool = False) -> None:
        """Copy cell styling from source to target cell.

        Args:
            source_cell: Source cell to copy style from
            target_cell: Target cell to apply style to
            is_header: Whether this is a header cell (keep bold) or body cell (remove bold)
        """
        try:
            # Set background to white for body cells
            if not is_header:
                tcPr = target_cell._element.get_or_add_tcPr()

                # Remove existing shading if present to avoid duplicates
                existing_shd = tcPr.findall(qn('w:shd'))
                for shd in existing_shd:
                    tcPr.remove(shd)

                # Add white background
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:fill'), 'FFFFFF')  # White background
                tcPr.append(shading_elm)

            # Note: Vertical alignment is set separately by calling code
            # Don't set it here to avoid conflicts

            # Copy paragraph formatting
            if source_cell.paragraphs and target_cell.paragraphs:
                source_para = source_cell.paragraphs[0]
                target_para = target_cell.paragraphs[0]

                # Set left alignment for body cells
                if not is_header:
                    target_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    # Copy alignment for header
                    target_para.alignment = source_para.alignment

                # Copy font formatting (but not bold for body cells)
                if source_para.runs and target_para.runs:
                    source_run = source_para.runs[0]
                    target_run = target_para.runs[0]

                    if source_run.font.name:
                        target_run.font.name = source_run.font.name

                    # Set font size: 10pt for body cells, keep header size for header cells
                    if is_header:
                        if source_run.font.size:
                            target_run.font.size = source_run.font.size
                    else:
                        # Body cells always use 10pt
                        target_run.font.size = Pt(10)

                    # Only copy bold if it's a header cell
                    if is_header and source_run.font.bold:
                        target_run.font.bold = True
                    else:
                        target_run.font.bold = False

        except Exception as e:
            logger.debug("[OpenXML] Error copying cell style: %s", str(e))

    def _generate_and_insert_pie_chart(self, doc: Document, by_the_numbers: List) -> None:
        """Generate pie chart using matplotlib and insert into Word document.

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
                logger.warning("[OpenXML] No data for pie chart")
                return

            # Extract counts
            counts = by_the_numbers[0] if by_the_numbers else None
            if not counts:
                return

            positive = getattr(counts, 'positive_count', 0) or 0
            neutral = getattr(counts, 'neutral_count', 0) or 0
            reconsider = getattr(counts, 'reconsider_count', 0) or 0

            logger.info("[OpenXML] Pie chart data: Positive=%d, Neutral=%d, Reconsider=%d",
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
            self.temp_dir = Path(tempfile.mkdtemp())
            chart_path = self.temp_dir / "pie_chart.png"
            fig.savefig(str(chart_path), dpi=150, bbox_inches='tight', pad_inches=0.1, facecolor='white')
            plt.close(fig)

            logger.info("[OpenXML] Pie chart generated successfully")

            # Find "PieChart" text or bookmark location and insert image
            chart_inserted = False

            # Search in paragraphs for "PieChart" placeholder
            for paragraph in doc.paragraphs:
                if "PieChart" in paragraph.text:
                    # Clear the paragraph text
                    paragraph.clear()
                    # Add spacing before the chart to push it down
                    paragraph_format = paragraph.paragraph_format
                    paragraph_format.space_before = Pt(12)  # Add 12pt space before

                    # Insert chart image (reduced width to fit page)
                    run = paragraph.add_run()
                    run.add_picture(str(chart_path), width=Inches(3.8))
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    chart_inserted = True
                    logger.info("[OpenXML] Pie chart inserted in paragraph")
                    break

            # If not found in paragraphs, search in tables
            if not chart_inserted:
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                if "PieChart" in paragraph.text:
                                    # Clear the paragraph text
                                    paragraph.clear()
                                    # Add spacing before the chart to push it down
                                    paragraph_format = paragraph.paragraph_format
                                    paragraph_format.space_before = Pt(12)  # Add 12pt space before

                                    # Insert chart image (reduced width to fit page)
                                    run = paragraph.add_run()
                                    run.add_picture(str(chart_path), width=Inches(3.5))
                                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    chart_inserted = True
                                    logger.info("[OpenXML] Pie chart inserted in table cell")
                                    break
                            if chart_inserted:
                                break
                        if chart_inserted:
                            break
                    if chart_inserted:
                        break

            if not chart_inserted:
                logger.warning("[OpenXML] PieChart placeholder not found in document")

        except Exception as e:
            logger.error("[OpenXML] Error generating pie chart: %s", str(e), exc_info=True)


# Global service instance
feedback_template_generator_openxml = FeedbackTemplateGeneratorOpenXML()
