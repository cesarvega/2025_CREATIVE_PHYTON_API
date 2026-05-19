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
        self._feedback_names: List[str] = []  # Test names for sound file embedding

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

                # Open .doc file (with retry for COM busy errors)
                logger.debug("[OpenXML] Opening source .doc file: %s", doc_path)
                source_path = str(doc_path.absolute())
                import time as _time
                for _open_attempt in range(3):
                    try:
                        doc = word_app.Documents.Open(source_path)
                        break
                    except Exception as open_err:
                        if _open_attempt < 2:
                            logger.warning("[OpenXML] Documents.Open attempt %d failed: %s. Retrying...", _open_attempt + 1, str(open_err))
                            _time.sleep(2)
                        else:
                            raise

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

                import time as _time
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        doc.SaveAs2(output_path_str, WD_FORMAT_XML_DOCUMENT)
                        logger.info("[OpenXML] SaveAs2 completed successfully")
                        break
                    except Exception as save_error:
                        if attempt < max_retries - 1:
                            logger.warning("[OpenXML] SaveAs2 attempt %d failed: %s. Retrying...", attempt + 1, str(save_error))
                            _time.sleep(2)
                        else:
                            logger.error("[OpenXML] SaveAs2 failed after %d attempts: %s", max_retries, str(save_error))
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
        with_sound_files: bool = False,
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

            # Format header with correct font size and style (includes logo insertion)
            logger.info("[OpenXML] Formatting header and inserting logo")
            try:
                self._format_document_header_modern(doc)
                logger.info("[OpenXML] Modern header applied successfully")
            except Exception as header_error:
                logger.warning("[OpenXML] Modern header failed, falling back to legacy layout: %s", str(header_error))
                self._format_document_header_OLD(doc)

            # PHASE 2: Populate tables with results (70-80%)
            logger.info("[OpenXML] Phase 2: Populating tables")
            self._populate_feedback_tables(doc, presentation_id, presentation_type)

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

            # PHASE 5: Embed sound files via COM (95-100%)
            if progress_callback:
                progress_callback(95)
            if with_sound_files and self._feedback_names:
                try:
                    self._embed_sound_files_via_com(output_path, display_name)
                except Exception as sound_error:
                    logger.warning("[OpenXML] Sound file embedding failed (non-fatal): %s", str(sound_error))

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
            # Default to phonetics template which includes Pronunciation column
            template_name = TemplateFilename.WORD_PHONETICS

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

    def _get_section_headers(self, section) -> List:
        """Return all header parts for a section (default, first-page, even), de-duplicated."""
        headers = []
        seen = set()
        for hdr in [section.header, getattr(section, "first_page_header", None), getattr(section, "even_page_header", None)]:
            if hdr is None:
                continue
            # Use id to avoid duplicate processing of shared parts
            hdr_id = id(hdr)
            if hdr_id in seen:
                continue
            seen.add(hdr_id)
            headers.append(hdr)
        return headers

    def _format_document_header_with_logo(self, doc: Document) -> None:
        """Format document header with logo in a clean table layout.

        This method is called AFTER placeholders are replaced, so all text is final.
        Creates a 2-column borderless table: logo on left, all header content on right.

        Args:
            doc: python-docx Document object
        """
        try:
            logger.info("[OpenXML] === Starting header formatting with logo ===")
            logo_path = settings.app_dir / "templates" / "bi_logo.png"

            for section_idx, section in enumerate(doc.sections):
                logger.info(f"[OpenXML] Processing section {section_idx + 1}")
                for header in self._get_section_headers(section):
                    # Step 1: Collect ALL text from the header (from all sources)
                    all_text_parts = []

                    logger.info(f"[OpenXML] Header has {len(header.paragraphs)} paragraphs and {len(header.tables)} tables")

                    # Collect from paragraphs
                    for idx, para in enumerate(header.paragraphs):
                        text = para.text.strip()
                        logger.debug(f"[OpenXML] Paragraph {idx}: '{text}'")
                        if text:
                            all_text_parts.append(text)

                    # Collect from tables
                    for table_idx, table in enumerate(header.tables):
                        logger.debug(f"[OpenXML] Processing table {table_idx}")
                        for row_idx, row in enumerate(table.rows):
                            for cell_idx, cell in enumerate(row.cells):
                                for para_idx, para in enumerate(cell.paragraphs):
                                    text = para.text.strip()
                                    logger.debug(f"[OpenXML] Table {table_idx}, Row {row_idx}, Cell {cell_idx}, Para {para_idx}: '{text}'")
                                    if text:
                                        all_text_parts.append(text)

                    # Combine all text with spaces
                    combined_header_text = "     ".join(all_text_parts)
                    logger.info(f"[OpenXML] Collected {len(all_text_parts)} text parts")
                    logger.info(f"[OpenXML] Combined text (first 200 chars): '{combined_header_text[:200]}'")

                    if not combined_header_text.strip():
                        logger.warning("[OpenXML] No text found in header! Skipping header reconstruction.")
                        # Don't clear the header if there's no text
                        continue

                    # Step 2: Clear the entire header
                    logger.info("[OpenXML] Clearing existing header content...")

                    # Remove all paragraphs
                    paragraphs_to_remove = list(header.paragraphs)
                    logger.info(f"[OpenXML] Removing {len(paragraphs_to_remove)} paragraphs")
                    for para in paragraphs_to_remove:
                        p_element = para._element
                        p_element.getparent().remove(p_element)

                    # Remove all tables
                    tables_to_remove = list(header.tables)
                    logger.info(f"[OpenXML] Removing {len(tables_to_remove)} tables")
                    for table in tables_to_remove:
                        t_element = table._element
                        t_element.getparent().remove(t_element)

                    logger.info("[OpenXML] Header cleared successfully")

                    # Step 3: Create new clean table structure
                    logger.info("[OpenXML] Creating new table structure...")
                    new_table = header.add_table(rows=1, cols=2)
                    new_table.autofit = False
                    logger.info("[OpenXML] New table created")

                    # Get cells
                    logo_cell = new_table.rows[0].cells[0]
                    text_cell = new_table.rows[0].cells[1]

                    # Set column widths
                    logo_cell.width = Inches(0.9)
                    text_cell.width = Inches(5.6)
                    logger.info("[OpenXML] Cell widths set")

                    # Configure logo cell
                    logo_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    # Add logo if available
                    if logo_path.exists():
                        logger.info("[OpenXML] Adding logo to first cell...")
                        logo_para = logo_cell.paragraphs[0]
                        logo_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        logo_para.paragraph_format.space_before = Pt(0)
                        logo_para.paragraph_format.space_after = Pt(0)
                        logo_para.paragraph_format.left_indent = Pt(0)

                        logo_run = logo_para.add_run()
                        logo_run.add_picture(
                            str(logo_path),
                            width=Inches(0.67),
                            height=Inches(0.38)
                        )
                        logger.info("[OpenXML] Logo added successfully")
                    else:
                        logger.warning(f"[OpenXML] Logo file not found at: {logo_path}")

                    # Configure text cell
                    logger.info("[OpenXML] Adding text to second cell...")
                    text_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    text_para = text_cell.paragraphs[0]
                    text_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    text_para.paragraph_format.space_before = Pt(0)
                    text_para.paragraph_format.space_after = Pt(0)

                    # Add the collected text
                    text_run = text_para.add_run(combined_header_text)
                    text_run.font.size = Pt(8.5)
                    text_run.font.bold = True
                    text_run.font.name = "Open Sans"
                    logger.info(f"[OpenXML] Text added: '{combined_header_text[:100]}...'")

                    # Remove all borders to make table invisible
                    logger.info("[OpenXML] Removing table borders...")
                    self._remove_table_borders(new_table)

                    logger.info("[OpenXML] Header formatted successfully with borderless table layout")

        except Exception as e:
            logger.error(f"[OpenXML] ERROR in header formatting: {str(e)}", exc_info=True)
            logger.error(f"[OpenXML] Error type: {type(e).__name__}")

    def _format_document_header_modern(self, doc: Document) -> None:
        """Rebuild header with simple single-row layout: Logo | Project Info | Company/Date."""
        try:
            logo_path = settings.app_dir / "templates" / "bi_logo.png"

            for section_idx, section in enumerate(doc.sections):
                for header in self._get_section_headers(section):

                    # Collect existing header text - CAREFULLY from cell by cell in tables
                    text_parts = []
                    
                    # First, check if header already has a table structure
                    if header.tables:
                        # Extract text from each cell separately to preserve structure
                        for table in header.tables:
                            for row in table.rows:
                                for cell_idx, cell in enumerate(row.cells):
                                    cell_texts = []
                                    for para in cell.paragraphs:
                                        text = para.text.strip()
                                        if text:
                                            cell_texts.append(text)
                                    if cell_texts:
                                        # Join texts within same cell with newline
                                        text_parts.append('\n'.join(cell_texts))
                    
                    # Also collect from paragraphs (if no table)
                    for para in header.paragraphs:
                        text = para.text.strip()
                        if text and text not in text_parts:  # Avoid duplicates
                            text_parts.append(text)

                    if not text_parts:
                        # Trigger fallback so header is never left empty
                        raise RuntimeError(f"No header text found for section {section_idx + 1}")

                    logger.info(f"[OpenXML] DEBUG - All collected text_parts: {text_parts}")

                    # Split text parts that contain tabs (original header has tab-separated columns)
                    left_parts = []  # Project info, Proprietary & Confidential
                    right_parts = []  # Company name, Date

                    for part in text_parts:
                        # Check if part contains tab - this means left and right text are in same string
                        if '\t' in part:
                            # Split by tab to separate left and right content
                            split_parts = part.split('\t')
                            if len(split_parts) >= 2:
                                left_parts.append(split_parts[0].strip())
                                right_parts.append(split_parts[1].strip())
                                logger.info(f"[OpenXML] DEBUG - Split by TAB - LEFT: '{split_parts[0].strip()}', RIGHT: '{split_parts[1].strip()}'")
                            else:
                                # Only one part after split, classify it
                                left_parts.append(split_parts[0].strip())
                                logger.info(f"[OpenXML] DEBUG - Added to LEFT: '{split_parts[0].strip()}'")
                        else:
                            # No tab, classify by content
                            part_lower = part.lower()
                            
                            # Skip if empty
                            if not part.strip():
                                continue
                                
                            # RIGHT COLUMN: Company names (with "pharmaceutical") or dates (with month names)
                            if 'pharmaceutical' in part_lower or \
                               any(month in part_lower for month in ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']):
                                right_parts.append(part)
                                logger.info(f"[OpenXML] DEBUG - Added to RIGHT: '{part}'")
                            # LEFT COLUMN: Everything else (project name, proprietary, confidential, etc.)
                            else:
                                left_parts.append(part)
                                logger.info(f"[OpenXML] DEBUG - Added to LEFT: '{part}'")

                    # Build the strings
                    left_text = "\n".join(left_parts) if left_parts else ""
                    right_text = "\n".join(right_parts) if right_parts else ""

                    logger.info(f"[OpenXML] DEBUG - LEFT TEXT: '{left_text}'")
                    logger.info(f"[OpenXML] DEBUG - RIGHT TEXT: '{right_text}'")

                    # Check if we have valid content
                    if not left_text and not right_text:
                        # Trigger fallback so header is never left empty
                        raise RuntimeError(f"No valid header text after parsing for section {section_idx + 1}")

                    # Clear existing content
                    for para in list(header.paragraphs):
                        p_element = para._element
                        p_element.getparent().remove(p_element)

                    for table in list(header.tables):
                        t_element = table._element
                        t_element.getparent().remove(t_element)

                    # Build simple single-row header
                    col_widths = [Inches(1.1), Inches(3.5), Inches(2.0)]
                    total_width = sum(col_widths)

                    header_table = header.add_table(rows=1, cols=3, width=total_width)
                    header_table.autofit = False

                    # Set individual column widths
                    for col_idx, width in enumerate(col_widths):
                        header_table.rows[0].cells[col_idx].width = width

                    row = header_table.rows[0]

                    # Configure all cells
                    for cell in row.cells:
                        self._set_cell_vertical_center(cell)
                        if cell.paragraphs:
                            para = cell.paragraphs[0]
                            para.paragraph_format.space_before = Pt(2)
                            para.paragraph_format.space_after = Pt(2)

                    # Logo cell
                    logo_cell = row.cells[0]
                    logo_para = logo_cell.paragraphs[0]
                    logo_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    if logo_path.exists():
                        logo_run = logo_para.add_run()
                        logo_run.add_picture(str(logo_path), width=Inches(0.8))
                    else:
                        logger.warning(f"[OpenXML] Logo file not found at: {logo_path}")

                    # Left text cell (Project info + Confidential)
                    left_cell = row.cells[1]
                    left_para = left_cell.paragraphs[0]
                    left_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    left_run = left_para.add_run(left_text)
                    left_run.font.size = Pt(9)
                    left_run.font.bold = True
                    left_run.font.name = "Open Sans"

                    # Right text cell (Company + Date) - use separate paragraphs for more spacing
                    right_cell = row.cells[2]

                    # Split right_text into separate parts for better spacing
                    right_text_lines = right_text.split('\n') if right_text else []

                    if right_text_lines:
                        # First line (Company name)
                        right_para = right_cell.paragraphs[0]
                        right_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        right_para.paragraph_format.space_after = Pt(2)  # Minimal space after first line
                        right_para.paragraph_format.right_indent = Pt(6)  # Indent from right edge
                        right_run = right_para.add_run(right_text_lines[0])
                        right_run.font.size = Pt(9)
                        right_run.font.bold = True
                        right_run.font.name = "Open Sans"

                        # Additional lines (Date, etc.)
                        for line in right_text_lines[1:]:
                            right_para2 = right_cell.add_paragraph()
                            right_para2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            right_para2.paragraph_format.space_before = Pt(0)
                            right_para2.paragraph_format.space_after = Pt(0)  # No extra space
                            right_para2.paragraph_format.right_indent = Pt(6)  # Indent from right edge
                            right_run2 = right_para2.add_run(line)
                            right_run2.font.size = Pt(9)
                            right_run2.font.bold = True
                            right_run2.font.name = "Open Sans"
                    else:
                        # Fallback: no text
                        right_para = right_cell.paragraphs[0]
                        right_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        right_para.paragraph_format.right_indent = Pt(6)

                    # Remove all borders to make table invisible
                    self._remove_table_borders(header_table)

                    # Add spacing after the header table to separate it from body content
                    # Set space after on the table itself
                    tbl = header_table._element
                    tblPr = tbl.tblPr
                    if tblPr is None:
                        tblPr = OxmlElement('w:tblPr')
                        tbl.insert(0, tblPr)
                    
                    # Remove existing spacing if present
                    for existing_spacing in tblPr.findall(qn('w:tblCellSpacing')):
                        tblPr.remove(existing_spacing)
                    
                    # Add paragraph after table for spacing
                    spacing_para = header.add_paragraph()
                    spacing_para.paragraph_format.space_after = Pt(12)  # 12pt space after header

            logger.info("[OpenXML] Modern header formatting completed")

        except Exception as e:
            logger.error("[OpenXML] Error building modern header: %s", str(e), exc_info=True)
            # Bubble up so caller can trigger legacy fallback
            raise

    def _format_document_header_OLD(self, doc: Document) -> None:
        """Format document header with correct font size (8.5pt) and bold, and add logo if available.

        Args:
            doc: python-docx Document object
        """
        try:
            logger.info("[OpenXML] Formatting document header")

            for section in doc.sections:
                for header in self._get_section_headers(section):

                    # Insert logo at the beginning of the header if logo file exists
                    logo_path = settings.app_dir / "templates" / "bi_logo.png"
                    if logo_path.exists():
                        try:
                            # Check if header has existing table
                            if header.tables:
                                # Work with existing table - add logo to first cell
                                first_table = header.tables[0]
                                if first_table.rows:
                                    first_row = first_table.rows[0]

                                    # Check if we need to add a column or use existing structure
                                    if len(first_row.cells) >= 2:
                                        # Use existing 2-column structure
                                        logo_cell = first_row.cells[0]
                                        text_cell = first_row.cells[1]
                                    else:
                                        # Single column - add logo to same cell as text (inline)
                                        logo_cell = first_row.cells[0]
                                        text_cell = None

                                    # Clear first cell
                                    for para in logo_cell.paragraphs:
                                        para.clear()

                                    # Configure logo cell
                                    logo_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                                    logo_cell.width = Inches(0.8)

                                    # Add logo to first cell
                                    if logo_cell.paragraphs:
                                        logo_para = logo_cell.paragraphs[0]
                                    else:
                                        logo_para = logo_cell.add_paragraph()

                                    logo_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    logo_para.paragraph_format.space_before = Pt(0)
                                    logo_para.paragraph_format.space_after = Pt(0)

                                    logo_run = logo_para.add_run()
                                    logo_run.add_picture(
                                        str(logo_path),
                                        width=Inches(0.67),
                                        height=Inches(0.38)
                                    )

                                    # If there's a separate text cell, configure it
                                    if text_cell:
                                        text_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                                        # Keep existing text - don't modify it

                                    # Remove borders from the table to make it look seamless
                                    self._remove_table_borders(first_table)

                                    logger.info("[OpenXML] Logo inserted in existing table structure")
                            else:
                                # No table in header - add logo inline in first paragraph
                                if header.paragraphs:
                                    first_para = header.paragraphs[0]

                                    # Save existing text
                                    existing_text = first_para.text

                                    # Clear paragraph
                                    first_para.clear()

                                    # Configure paragraph
                                    first_para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                    first_para.paragraph_format.space_before = Pt(0)
                                    first_para.paragraph_format.space_after = Pt(0)

                                    # Add logo inline
                                    logo_run = first_para.add_run()
                                    logo_run.add_picture(
                                        str(logo_path),
                                        width=Inches(0.67),
                                        height=Inches(0.38)
                                    )

                                    # Add space
                                    space_run = first_para.add_run("  ")

                                    # Add text back
                                    if existing_text.strip():
                                        text_run = first_para.add_run(existing_text)
                                        text_run.font.size = Pt(8.5)
                                        text_run.font.bold = True

                                    logger.info("[OpenXML] Logo inserted inline in header paragraph")

                        except Exception as logo_error:
                            logger.warning("[OpenXML] Error inserting logo: %s", str(logo_error), exc_info=True)
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

            logger.info("[OpenXML] Header formatting completed")

        except Exception as e:
            logger.warning("[OpenXML] Error formatting header: %s", str(e))
        
    def _populate_feedback_tables(self, doc: Document, presentation_id: int, presentation_type: str = "Normal") -> None:
        """Populate the feedback table using the nw_wdToIndicateFeedback SP.

        Args:
            doc: python-docx Document object
            presentation_id: Presentation ID
            presentation_type: Presentation type (Normal, Katakana, etc.)
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
                self._populate_feedback_table_from_sp(main_table, columns, rows, presentation_type)

        except Exception as e:
            logger.error("[OpenXML] Error populating feedback tables: %s", str(e), exc_info=True)

    def _populate_feedback_table_from_sp(self, table, columns: List[str], rows: List, presentation_type: str = "Normal") -> None:
        """Populate feedback table directly from SP results.

        For Normal presentations, the SP returns 7 columns:
        ' ', 'Test Name', 'Rationale', 'Positive', 'Neutral', 'Negative', 'Comments/Suggestions'
        Template has 8 columns: #, Test Name, Pronunciation (empty), Rationale, Positive, Neutral, Negative, Comments

        For Katakana presentations, the SP returns 8 columns:
        ' ', 'Test Name', 'Katakana Names', 'Rationale', 'Positive', 'Neutral', 'Negative', 'Comments/Suggestions'
        Template has 7 columns: #, Test Name, Katakana Names, Positive, Neutral, Negative, Comments
        (Rationale is removed per legacy behavior)

        Args:
            table: python-docx Table object
            columns: List of column names from SP
            rows: List of row tuples from SP
            presentation_type: Presentation type (Normal, Katakana, etc.)
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
                """Expand SP row into multiple rows when Test Name/Rationale/Katakana Names contain '##'."""
                test_name_parts = split_grouped_text(row_dict.get("Test Name"))
                rationale_parts = split_grouped_text(row_dict.get("Rationale"))
                katakana_parts = split_grouped_text(row_dict.get("Katakana Names"))

                # No grouping: return as-is (but with basic unescape/strip)
                if len(test_name_parts) <= 1 and len(rationale_parts) <= 1 and len(katakana_parts) <= 1:
                    return [
                        {
                            "Test Name": (test_name_parts[0] if test_name_parts else "").strip(),
                            "Rationale": (rationale_parts[0] if rationale_parts else "").strip(),
                            "Katakana Names": (katakana_parts[0] if katakana_parts else "").strip(),
                        }
                    ]

                # Align list lengths
                count = max(len(test_name_parts), len(rationale_parts), len(katakana_parts), 1)
                if not test_name_parts:
                    test_name_parts = [""] * count
                if not rationale_parts:
                    rationale_parts = [""] * count
                if not katakana_parts:
                    katakana_parts = [""] * count
                if len(test_name_parts) == 1 and count > 1:
                    test_name_parts = test_name_parts * count
                if len(rationale_parts) == 1 and count > 1:
                    rationale_parts = rationale_parts * count
                if len(katakana_parts) == 1 and count > 1:
                    katakana_parts = katakana_parts * count
                while len(test_name_parts) < count:
                    test_name_parts.append(test_name_parts[-1] if test_name_parts else "")
                while len(rationale_parts) < count:
                    rationale_parts.append(rationale_parts[-1] if rationale_parts else "")
                while len(katakana_parts) < count:
                    katakana_parts.append(katakana_parts[-1] if katakana_parts else "")

                expanded = []
                for i in range(count):
                    expanded.append(
                        {
                            "Test Name": (test_name_parts[i] or "").strip(),
                            "Rationale": (rationale_parts[i] or "").strip(),
                            "Katakana Names": (katakana_parts[i] or "").strip(),
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

            # Store test names for sound file embedding later
            self._feedback_names = [
                str(row_dict.get('Test Name') or '').strip()
                for row_dict in expanded_rows
            ]

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
                        run.font.name = "Open Sans"
                    # Apply vertical center using XML
                    self._set_cell_vertical_center(cell)

                    # Column 2: Test Name (Candidate) - CENTER ALIGNED, no bold
                    cell = new_row.cells[1]
                    cell.text = str(row_dict.get('Test Name') or '')
                    para = cell.paragraphs[0]
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.bold = False
                        run.font.name = "Open Sans"
                    # Apply vertical center using XML
                    self._set_cell_vertical_center(cell)

                    # Determine column layout based on presentation type
                    is_katakana = presentation_type and presentation_type.upper().startswith("KATAKANA")
                    num_cells = len(new_row.cells)

                    if is_katakana:
                        # Katakana template: 7 columns
                        # Col 3: Katakana Names, NO Rationale, votes start at col 4
                        cell = new_row.cells[2]
                        cell.text = str(row_dict.get('Katakana Names') or '')
                        para = cell.paragraphs[0]
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in para.runs:
                            run.font.size = Pt(10)
                            run.font.bold = False
                            run.font.name = "Open Sans"
                        self._set_cell_vertical_center(cell)

                        # Columns 4-7: Positive, Neutral, Negative, Comments/Suggestions
                        for col_idx in range(3, min(7, num_cells)):
                            cell = new_row.cells[col_idx]
                            cell.text = ""
                            para = cell.paragraphs[0]
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            if para.runs:
                                for run in para.runs:
                                    run.font.size = Pt(10)
                                    run.font.bold = False
                                    run.font.name = "Open Sans"
                            else:
                                run = para.add_run()
                                run.font.size = Pt(10)
                                run.font.bold = False
                                run.font.name = "Open Sans"
                            self._set_cell_vertical_center(cell)
                    else:
                        # Normal/Phonetics template: 8 columns
                        # Col 3: Pronunciation (empty), Col 4: Rationale, votes start at col 5
                        cell = new_row.cells[2]
                        cell.text = ''
                        para = cell.paragraphs[0]
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in para.runs:
                            run.font.size = Pt(10)
                            run.font.bold = False
                            run.font.name = "Open Sans"
                        self._set_cell_vertical_center(cell)

                        # Column 4: Rationale
                        cell = new_row.cells[3]
                        cell.text = str(row_dict.get('Rationale') or '')
                        para = cell.paragraphs[0]
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in para.runs:
                            run.font.size = Pt(10)
                            run.font.bold = False
                            run.font.name = "Open Sans"
                        self._set_cell_vertical_center(cell)

                        # Columns 5-8: Positive, Neutral, Negative, Comments/Suggestions
                        for col_idx in range(4, min(8, num_cells)):
                            cell = new_row.cells[col_idx]
                            cell.text = ""
                            para = cell.paragraphs[0]
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            if para.runs:
                                for run in para.runs:
                                    run.font.size = Pt(10)
                                    run.font.bold = False
                                    run.font.name = "Open Sans"
                            else:
                                run = para.add_run()
                                run.font.size = Pt(10)
                                run.font.bold = False
                                run.font.name = "Open Sans"
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

    def _find_sound_file(self, test_name: str, display_name: str) -> Optional[Path]:
        """Find MP3 sound file for a test name.

        Search order:
        1. Project-specific folder: C:\\SoundFiles\\{display_name}_MR\\{NAME}.MP3
        2. Global folder: C:\\SoundFiles\\{NAME}.MP3

        Args:
            test_name: The test name (e.g. "SKELTEGRA")
            display_name: Project display name for project-specific folder lookup

        Returns:
            Path to MP3 file if found, None otherwise
        """
        if not test_name:
            return None

        sound_dir = settings.sound_files_dir
        name_upper = test_name.strip().upper()

        # Try project-specific folder first
        if display_name:
            project_folder = sound_dir / f"{display_name}_MR"
            project_path = project_folder / f"{name_upper}.MP3"
            if project_path.exists():
                return project_path

        # Try global folder
        global_path = sound_dir / f"{name_upper}.MP3"
        if global_path.exists():
            return global_path

        return None

    def _embed_sound_files_via_com(self, doc_path: Path, display_name: str) -> None:
        """Embed MP3 sound files into the Pronunciation column using Word COM.

        Opens the saved document with Word COM and inserts OLE objects (MP3 files)
        as icons in the Pronunciation column (column 3) of the feedback table.

        Args:
            doc_path: Path to the saved .docx file
            display_name: Project display name for sound file lookup
        """
        import time as _time

        if not self._feedback_names:
            return

        # Check how many sound files exist before opening COM
        sound_files = []
        for name in self._feedback_names:
            mp3_path = self._find_sound_file(name, display_name)
            sound_files.append(mp3_path)

        found_count = sum(1 for f in sound_files if f is not None)
        if found_count == 0:
            logger.info("[OpenXML] No sound files found for any test names, skipping COM embedding")
            return

        logger.info("[OpenXML] Found %d/%d sound files, embedding via COM...", found_count, len(self._feedback_names))

        com_initialized = False
        word_app = None

        try:
            pythoncom.CoInitialize()
            com_initialized = True
        except Exception:
            pass

        try:
            from app.utils.com_manager import com_manager
            with com_manager.acquire("Word.Application for sound embedding"):
                word_app = win32com.client.Dispatch("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0

                doc_path_str = str(doc_path.absolute())
                doc = None

                for _attempt in range(3):
                    try:
                        doc = word_app.Documents.Open(doc_path_str)
                        break
                    except Exception as open_err:
                        if _attempt < 2:
                            logger.warning("[OpenXML] COM Documents.Open attempt %d failed: %s", _attempt + 1, str(open_err))
                            _time.sleep(2)
                        else:
                            raise

                if not doc:
                    raise RuntimeError("Failed to open document via COM")

                try:
                    # Get the first table in the document (feedback table)
                    if doc.Tables.Count < 1:
                        logger.warning("[OpenXML] No tables found in document for sound embedding")
                        return

                    table = doc.Tables(1)
                    # Table has header row (row 1), data starts at row 2
                    # Pronunciation is column 3

                    for idx, (name, mp3_path) in enumerate(zip(self._feedback_names, sound_files)):
                        if mp3_path is None:
                            continue

                        row_num = idx + 2  # +2 because row 1 is header, COM is 1-based
                        try:
                            if row_num > table.Rows.Count:
                                logger.warning("[OpenXML] Row %d exceeds table rows (%d), stopping", row_num, table.Rows.Count)
                                break

                            # Get the Pronunciation cell (column 3)
                            cell = table.Cell(row_num, 3)
                            cell_range = cell.Range

                            # Collapse range to start to avoid replacing existing content
                            cell_range.Collapse(1)  # wdCollapseStart

                            # Embed MP3 as OLE object with WMP icon
                            wmp_exe = r"C:\Program Files\Windows Media Player\wmplayer.exe"
                            ole_kwargs = {
                                "FileName": str(mp3_path.absolute()),
                                "LinkToFile": False,
                                "DisplayAsIcon": True,
                                "IconLabel": f"{name.upper()}.MP3",
                            }
                            # Use WMP icon if available for a recognizable play button
                            if Path(wmp_exe).exists():
                                ole_kwargs["IconFileName"] = wmp_exe
                                ole_kwargs["IconIndex"] = 0

                            cell_range.InlineShapes.AddOLEObject(**ole_kwargs)

                            # Center the OLE icon in the cell
                            for para in cell.Range.Paragraphs:
                                para.Alignment = 1  # wdAlignParagraphCenter

                            logger.debug("[OpenXML] Embedded sound: %s -> row %d", mp3_path.name, row_num)

                        except Exception as cell_err:
                            logger.warning("[OpenXML] Error embedding sound for '%s' in row %d: %s", name, row_num, str(cell_err))

                    # Save the document
                    doc.Save()
                    logger.info("[OpenXML] Sound files embedded and document saved via COM")

                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=False)
                        except Exception:
                            pass

        except Exception as e:
            logger.error("[OpenXML] Error embedding sound files via COM: %s", str(e), exc_info=True)
            raise

        finally:
            if word_app:
                try:
                    word_app.Quit()
                except Exception:
                    pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _remove_table_borders(self, table) -> None:
        """Remove all borders from a table to make it invisible.

        Args:
            table: python-docx Table object
        """
        try:
            tbl = table._element
            tblPr = tbl.tblPr
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr')
                tbl.insert(0, tblPr)

            # Remove existing borders
            for borders in tblPr.findall(qn('w:tblBorders')):
                tblPr.remove(borders)

            # Create new borders element with all borders set to 'none'
            tblBorders = OxmlElement('w:tblBorders')

            border_types = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
            for border_type in border_types:
                border = OxmlElement(f'w:{border_type}')
                border.set(qn('w:val'), 'none')
                border.set(qn('w:sz'), '0')
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), 'auto')
                tblBorders.append(border)

            tblPr.append(tblBorders)

            # Also remove borders from each cell
            for row in table.rows:
                for cell in row.cells:
                    tcPr = cell._element.get_or_add_tcPr()

                    # Remove existing cell borders
                    for tcBorders in tcPr.findall(qn('w:tcBorders')):
                        tcPr.remove(tcBorders)

                    # Add cell borders set to 'none'
                    tcBorders = OxmlElement('w:tcBorders')
                    for border_type in ['top', 'left', 'bottom', 'right']:
                        border = OxmlElement(f'w:{border_type}')
                        border.set(qn('w:val'), 'none')
                        border.set(qn('w:sz'), '0')
                        border.set(qn('w:space'), '0')
                        border.set(qn('w:color'), 'auto')
                        tcBorders.append(border)
                    tcPr.append(tcBorders)

            logger.debug("[OpenXML] Removed all borders from table")

        except Exception as e:
            logger.warning("[OpenXML] Error removing table borders: %s", str(e))

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

    def _set_cell_shading(self, cell, color_hex: str) -> None:
        """Apply background shading to a cell."""
        try:
            tcPr = cell._element.get_or_add_tcPr()
            for shd in tcPr.findall(qn('w:shd')):
                tcPr.remove(shd)

            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), color_hex)
            tcPr.append(shading_elm)
        except Exception as e:
            logger.debug("[OpenXML] Error setting cell shading: %s", str(e))

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
