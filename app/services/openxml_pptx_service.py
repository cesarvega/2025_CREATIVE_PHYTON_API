"""
OpenXML-based PowerPoint generation service.

This service provides a faster, more versatile alternative to COM-based PowerPoint generation
using the python-pptx library with direct OpenXML manipulation.
"""

import math
import textwrap
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

from app.models.presentation_models import ProcessedExcelData
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

NAVY_BLUE = RGBColor(45, 77, 145)


class OpenXMLPPTXService:
    """Service for generating PowerPoint presentations using OpenXML (python-pptx)."""

    def __init__(self):
        """Initialize the OpenXML PPTX service."""
        pass

    def generate_backup_presentation(
        self,
        pptx_bytes: bytes,
        excel_data: ProcessedExcelData,
        insert_position: int = 5,
        is_big_japanese: bool = False,
    ) -> bytes:
        """
        Generate a backup presentation using OpenXML.

        Args:
            pptx_bytes: Original PowerPoint file as bytes
            excel_data: Processed Excel data from excel_service
            insert_position: Zero-based index where to insert generated slides (from NameCandidateStartingSlide - 1)
            is_big_japanese: Whether to use big Japanese mode (katakana emphasized)

        Returns:
            bytes: Generated PowerPoint file as bytes

        Raises:
            ValueError: If inputs are invalid or processing fails
        """
        try:
            logger.info("Starting OpenXML backup generation at position %d", insert_position)

            # Load the presentation
            presentation = Presentation(BytesIO(pptx_bytes))

            # Convert ProcessedExcelData to the format expected by slide builder
            excel_rows = self._convert_processed_excel_to_rows(excel_data)

            # Build slides from Excel data
            inserted = self._build_slides_from_excel(
                presentation, excel_rows, insert_position, is_big_japanese
            )

            logger.info("Inserted %d slides using OpenXML", inserted)

            # Save to bytes
            buffer = BytesIO()
            presentation.save(buffer)
            buffer.seek(0)

            return buffer.read()

        except Exception as e:
            logger.error("OpenXML backup generation failed: %s", str(e))
            raise

    def _convert_processed_excel_to_rows(
        self, excel_data: ProcessedExcelData
    ) -> List[List[Any]]:
        """
        Convert ProcessedExcelData to list of rows format expected by slide builder.

        Args:
            excel_data: Processed Excel data from excel_service

        Returns:
            List of rows where first row is headers, subsequent rows are data
        """
        # Create header row
        headers = [
            "SEQ",
            "Category",
            "Name",
            "Rationale",
            "Pronunciation",
            "Katakana",
            "Group1",
            "Group2",
        ]

        # Create data rows
        rows = [headers]

        # Use total_rows_processed to get ALL rows (including category headers)
        total_rows = getattr(excel_data, 'total_rows_processed', len(excel_data.lst_types))

        logger.info(f"Converting ProcessedExcelData: total_rows={total_rows}, types={len(excel_data.lst_types)}, names={len(excel_data.lst_names)}, categories={len(excel_data.lst_categories)}")

        for i in range(total_rows):
            row = [
                excel_data.lst_types[i] if i < len(excel_data.lst_types) else "",
                excel_data.lst_categories[i] if i < len(excel_data.lst_categories) else "",
                excel_data.lst_names[i] if i < len(excel_data.lst_names) else "",
                excel_data.lst_rationales[i] if i < len(excel_data.lst_rationales) else "",
                excel_data.lst_notations[i] if i < len(excel_data.lst_notations) else "",
                excel_data.lst_kana[i] if i < len(excel_data.lst_kana) else "",
                excel_data.lst_name_sub_groups[i] if i < len(excel_data.lst_name_sub_groups) else "",
                "",  # Group2 placeholder
            ]
            rows.append(row)

            # Log first few rows for debugging
            if i < 5:
                logger.info(f"Row {i}: SEQ={row[0]}, Category={row[1][:50] if row[1] else ''}, Name={row[2]}, Group={row[6]}")

        logger.info(f"Converted {len(rows)-1} data rows (plus 1 header row)")
        return rows

    def insert_slide_with_background(
        self, presentation: Presentation, insert_at: int = 5
    ):
        """Add a single slide at the desired position while keeping the original background."""
        if not presentation.slides:
            raise ValueError("PowerPoint has no slides to copy a background from.")

        if insert_at < 0:
            insert_at = 0

        reference_index = (
            insert_at if insert_at < len(presentation.slides) else len(presentation.slides) - 1
        )
        reference_slide = presentation.slides[reference_index]
        layout = reference_slide.slide_layout
        new_slide = presentation.slides.add_slide(layout)
        self._copy_background(reference_slide, new_slide)
        self._move_new_slide(presentation, insert_at)
        self._clear_placeholders(new_slide)
        return new_slide

    def _copy_background(self, source_slide, target_slide) -> None:
        """Copy the background XML from the source slide to the target slide."""
        source_bg_nodes = source_slide._element.xpath("./p:cSld/p:bg")
        if not source_bg_nodes:
            return

        target_c_sld = target_slide._element.xpath("./p:cSld")
        if not target_c_sld:
            return

        target_c_sld = target_c_sld[0]
        for existing in target_c_sld.xpath("./p:bg"):
            target_c_sld.remove(existing)

        target_c_sld.insert(0, deepcopy(source_bg_nodes[0]))

    def _move_new_slide(self, presentation: Presentation, insert_at: int) -> None:
        """Reorder slide ids so the newly added slide lands at the requested position."""
        if insert_at < 0:
            insert_at = 0

        slide_id_list = presentation.slides._sldIdLst  # type: ignore[attr-defined]
        new_id = slide_id_list[-1]
        slide_id_list.remove(new_id)

        target_index = insert_at if insert_at <= len(slide_id_list) else len(slide_id_list)
        slide_id_list.insert(target_index, new_id)

    def _clear_placeholders(self, slide) -> None:
        """Remove default placeholders like 'Double-click to edit' before adding our own content."""
        # Iterate in reverse to avoid index issues when removing
        for shape in list(slide.shapes)[::-1]:
            try:
                if getattr(shape, "is_placeholder", False):
                    slide.shapes._spTree.remove(shape._element)  # type: ignore[attr-defined]
            except Exception:
                continue

    def _build_slides_from_excel(
        self,
        presentation: Presentation,
        excel_data: List[List[Any]],
        start_position: int = 5,
        is_big_japanese: bool = True,
    ) -> int:
        """Insert slides based on Excel rows. Returns count of slides inserted."""
        if not excel_data:
            return 0

        headers = self._normalize_headers(excel_data[0])
        try:
            seq_idx = headers["seq"]
            category_idx = headers["category"]
            name_idx = headers["name"]
            rationale_idx = headers["rationale"]
        except KeyError as exc:
            missing = exc.args[0]
            raise ValueError(f"Missing expected column: {missing}") from exc

        pronunciation_idx = headers.get("pronunciation")
        katakana_idx = headers.get("katakana")
        group1_idx = headers.get("group1")
        group2_idx = headers.get("group2")

        slide_items: List[Dict[str, Any]] = []
        group_index: Dict[Tuple[str, str], Dict[str, Any]] = {}

        for row in excel_data[1:]:
            seq_raw = self._get_cell(row, seq_idx)
            seq_str = str(seq_raw).strip() if seq_raw is not None else ""

            # Check if this is a category row (SEQ is alphabetic like "A", "B", "C")
            if seq_str and seq_str.isalpha():
                category_text = str(self._get_cell(row, category_idx) or "").strip()
                if category_text:
                    slide_items.append({"type": "category", "category_text": category_text})
                continue

            # Get row data
            name_text = str(self._get_cell(row, name_idx) or "").strip()
            rationale_text = str(self._get_cell(row, rationale_idx) or "").strip()
            category_text = str(self._get_cell(row, category_idx) or "").strip()
            pronunciation_text = (
                str(self._get_cell(row, pronunciation_idx) or "").strip()
                if pronunciation_idx is not None
                else ""
            )
            katakana_text = (
                str(self._get_cell(row, katakana_idx) or "").strip()
                if katakana_idx is not None
                else ""
            )
            group_label = self._extract_group_label(row, group1_idx, group2_idx)

            # Skip empty rows
            if not name_text and not category_text:
                continue

            # Check if this is a group row (empty SEQ + group label + names with ## separator)
            if group_label and "##" in name_text:
                # Split names by ## separator
                names = [n.strip() for n in name_text.split("##") if n.strip()]
                rows_for_group = []
                for name in names:
                    rows_for_group.append({
                        "category_text": category_text,
                        "name_text": name,
                        "rationale_text": "",
                        "pronunciation_text": "",
                        "katakana_text": "",
                    })

                # Add group slide
                slide_items.append({
                    "type": "group",
                    "group_label": group_label,
                    "rows": rows_for_group,
                    "category_text": category_text,
                })
            elif name_text:
                # Individual slide
                row_payload = {
                    "category_text": category_text,
                    "name_text": name_text,
                    "rationale_text": rationale_text,
                    "pronunciation_text": pronunciation_text,
                    "katakana_text": katakana_text,
                }
                slide_items.append({"type": "single", "row": row_payload})

        current_position = start_position
        inserted = 0
        for item in slide_items:
            slide = self.insert_slide_with_background(presentation, current_position)
            if item["type"] == "category":
                self._render_category_slide(presentation, slide, item["category_text"])
            elif item["type"] == "group":
                self._render_group_slide(
                    presentation,
                    slide,
                    category_text=item.get("category_text") or "",
                    group_label=item["group_label"],
                    rows=item["rows"],
                )
            else:
                row = item["row"]
                self._render_name_rationale_slide(
                    presentation,
                    slide,
                    category_text=row["category_text"],
                    name_text=row["name_text"],
                    rationale_text=row["rationale_text"],
                    pronunciation_text=row["pronunciation_text"],
                    katakana_text=row["katakana_text"],
                    is_big_japanese=is_big_japanese,
                )
            current_position += 1
            inserted += 1

        # Insert Name Summary slide at the end
        logger.info(f"Adding Name Summary slide at position {current_position}")
        summary_slide = self.insert_slide_with_background(presentation, current_position)
        self._render_name_summary_slide(presentation, summary_slide)
        inserted += 1

        return inserted

    def _normalize_headers(self, header_row: List[Any]) -> Dict[str, int]:
        mapping: Dict[str, int] = {}
        for idx, cell in enumerate(header_row):
            if cell is None:
                continue
            key = str(cell).strip().lower()
            if key:
                mapping[key] = idx
        return mapping

    def _get_cell(self, row: List[Any], idx: int) -> Any:
        return row[idx] if idx < len(row) else None

    def _wrap_text(self, text: str, max_chars: int, max_lines: int = 2) -> List[str]:
        """Return up to max_lines of wrapped text without breaking words."""
        wrapped = textwrap.wrap(
            text, max_chars, break_long_words=False, break_on_hyphens=False
        )
        if not wrapped:
            return []
        if len(wrapped) > max_lines:
            wrapped = wrapped[: max_lines - 1] + [" ".join(wrapped[max_lines - 1 :])]
        return wrapped[:max_lines]

    def _fit_font_size(
        self, text: str, base_size: int, min_size: int, max_chars: int
    ) -> int:
        """Heuristic: shrink font if text is longer than max_chars to keep within slide."""
        length = len(text.strip())
        if length == 0:
            return base_size
        if length <= max_chars:
            return base_size
        scaled = int(base_size * max_chars / length)
        return max(min_size, min(base_size, scaled))

    def _extract_group_label(
        self, row: List[Any], group1_idx: Optional[int], group2_idx: Optional[int]
    ) -> str:
        """Return the first non-empty group label from Group1/Group2 columns."""
        for idx in (group1_idx, group2_idx):
            if idx is None:
                continue
            value = self._get_cell(row, idx)
            if value is None:
                continue
            label = str(value).strip()
            if label:
                return label
        return ""

    def _render_category_slide(
        self, presentation: Presentation, slide, category_text: str
    ) -> None:
        """Render category text centered. If text contains '$', split into two lines and double bottom font size."""
        parts = [part.strip() for part in category_text.split("$", 1)]
        top_text = parts[0] if parts else ""
        bottom_text = parts[1] if len(parts) > 1 else ""

        textbox_width = presentation.slide_width - Inches(1.5)
        left = Inches(0.75)
        top = Inches(2.2)
        height = Inches(3.2)
        box = slide.shapes.add_textbox(left, top, textbox_width, height)
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True

        top_lines = self._wrap_text(top_text, max_chars=28, max_lines=2)
        bottom_lines = (
            self._wrap_text(bottom_text, max_chars=30, max_lines=2) if bottom_text else []
        )

        top_font = Pt(28) if len(top_lines) == 1 else Pt(26)
        bottom_font = Pt(38) if len(bottom_lines) == 1 else Pt(32)
        bottom_font = (
            Pt(max(bottom_font.pt, top_font.pt + 6)) if bottom_lines else bottom_font
        )

        for idx, line in enumerate(top_lines):
            para = tf.paragraphs[idx] if idx == 0 else tf.add_paragraph()
            para.text = line
            para.alignment = PP_ALIGN.CENTER
            para.font.size = top_font
            para.font.bold = True
            para.font.color.rgb = RGBColor(0, 0, 0)

        for line in bottom_lines:
            para = tf.add_paragraph()
            para.text = line
            para.alignment = PP_ALIGN.CENTER
            para.font.size = bottom_font
            para.font.bold = True
            para.font.color.rgb = RGBColor(0, 0, 139)  # deep navy

    def _render_name_rationale_slide(
        self,
        presentation: Presentation,
        slide,
        category_text: str,
        name_text: str,
        rationale_text: str,
        pronunciation_text: str = "",
        katakana_text: str = "",
        is_big_japanese: bool = False,
    ) -> None:
        """Render name/rationale slide with a simple layout similar to the provided design."""
        slide_width = presentation.slide_width

        # Category heading (text only, no bar)
        header_top = Inches(0.35)
        header_box = slide.shapes.add_textbox(
            Inches(0.75), header_top, slide_width - Inches(1.5), Inches(0.5)
        )
        header_tf = header_box.text_frame
        header_tf.clear()
        header_para = header_tf.paragraphs[0]
        header_para.text = category_text or "Category"
        header_para.font.size = Pt(20)
        header_para.font.bold = True
        header_para.font.color.rgb = NAVY_BLUE
        header_para.alignment = PP_ALIGN.LEFT

        # Name and pronunciation block
        top_block = header_top + Inches(0.8)
        block_width = slide_width - Inches(1.5)
        left = Inches(0.75)
        block_height = Inches(2.2)
        name_box = slide.shapes.add_textbox(left, top_block, block_width, block_height)
        name_tf = name_box.text_frame
        name_tf.clear()
        name_tf.word_wrap = True
        display_name = (
            katakana_text if (is_big_japanese and katakana_text) else (name_text or "Name")
        )
        name_para = name_tf.paragraphs[0]
        name_para.text = display_name
        name_font_size = self._fit_font_size(
            display_name, base_size=80, min_size=36, max_chars=18
        )
        name_para.font.size = Pt(name_font_size)
        name_para.font.bold = True
        name_para.alignment = PP_ALIGN.LEFT
        name_para.font.color.rgb = RGBColor(0, 0, 0)

        small_parts: List[str] = []
        if is_big_japanese and katakana_text:
            if name_text:
                small_parts.append(name_text)
            if pronunciation_text:
                small_parts.append(pronunciation_text)
        else:
            if pronunciation_text:
                small_parts.append(pronunciation_text)
            if katakana_text:
                small_parts.append(katakana_text)
        if small_parts:
            sub_para = name_tf.add_paragraph()
            sub_para.text = " / ".join(small_parts)
            sub_para.font.size = Pt(22)
            sub_para.font.bold = False
            sub_para.alignment = PP_ALIGN.LEFT
            sub_para.font.color.rgb = RGBColor(0, 0, 0)

        # Rationale title and text
        rat_title_top = top_block + block_height + Inches(0.4)
        rat_title = slide.shapes.add_textbox(left, rat_title_top, block_width, Inches(1.0))
        rat_tf = rat_title.text_frame
        rat_tf.clear()
        rat_tf.word_wrap = True
        rat_lines = self._wrap_text(
            rationale_text or "Rationale goes here", max_chars=60, max_lines=2
        )
        for idx, line in enumerate(rat_lines):
            para = rat_tf.paragraphs[idx] if idx == 0 else rat_tf.add_paragraph()
            para.text = line
            para.font.size = Pt(22)
            para.font.bold = True
            para.alignment = PP_ALIGN.LEFT
            para.font.color.rgb = RGBColor(0, 0, 0)

        # Placeholder boxes for New Names and Comments
        box_top = rat_title_top + Inches(0.9)
        box_height = Inches(0.9)
        box_width = (block_width - Inches(0.5)) / 2
        new_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, box_top, box_width, box_height
        )
        new_box.fill.background()
        new_box.line.color.rgb = RGBColor(120, 120, 120)
        new_box_tf = new_box.text_frame
        new_box_tf.text = "[New Names]"
        new_box_tf.paragraphs[0].font.size = Pt(14)
        new_box_tf.paragraphs[0].font.color.rgb = RGBColor(150, 150, 150)
        new_box_tf.paragraphs[0].alignment = PP_ALIGN.CENTER

        comments_left = left + box_width + Inches(0.5)
        comments_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, comments_left, box_top, box_width, box_height
        )
        comments_box.fill.background()
        comments_box.line.color.rgb = RGBColor(120, 120, 120)
        comments_tf = comments_box.text_frame
        comments_tf.text = "[Comments]"
        comments_tf.paragraphs[0].font.size = Pt(14)
        comments_tf.paragraphs[0].font.color.rgb = RGBColor(150, 150, 150)
        comments_tf.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Sentiment labels
        sentiment_top = box_top + box_height + Inches(0.4)
        sentiment_text = "◯ Positive        ◯ Neutral        ◯ Negative"
        sentiment_box = slide.shapes.add_textbox(
            left, sentiment_top, block_width, Inches(0.6)
        )
        sentiment_para = sentiment_box.text_frame.paragraphs[0]
        sentiment_para.text = sentiment_text
        sentiment_para.font.size = Pt(14)
        sentiment_para.font.color.rgb = RGBColor(50, 50, 50)
        sentiment_para.alignment = PP_ALIGN.LEFT

    def _render_group_slide(
        self,
        presentation: Presentation,
        slide,
        category_text: str,
        group_label: str,
        rows: List[Dict[str, str]],
    ) -> None:
        """Render a grouped slide with a compact checkbox grid that scales to many names."""
        slide_width = presentation.slide_width
        slide_height = presentation.slide_height

        total_rows = len(rows)
        columns = 3 if total_rows <= 24 else 4 if total_rows <= 48 else 5

        margin_left = Inches(0.4)
        margin_right = Inches(0.4)
        available_width = slide_width - margin_left - margin_right

        # Category heading (text only, no bar)
        heading_top = Inches(0.25)
        heading_box = slide.shapes.add_textbox(
            margin_left, heading_top, available_width, Inches(0.5)
        )
        heading_tf = heading_box.text_frame
        heading_tf.clear()
        heading_para = heading_tf.paragraphs[0]
        heading_para.text = category_text or "Category"
        heading_para.font.size = Pt(20)
        heading_para.font.bold = True
        heading_para.font.color.rgb = NAVY_BLUE
        heading_para.alignment = PP_ALIGN.LEFT

        group_top = heading_top + Inches(0.45)
        group_box = slide.shapes.add_textbox(
            margin_left, group_top, available_width, Inches(0.3)
        )
        group_tf = group_box.text_frame
        group_tf.clear()
        group_para = group_tf.paragraphs[0]
        group_para.text = f"Group {group_label}" if group_label else ""
        group_para.font.size = Pt(12)
        group_para.font.color.rgb = NAVY_BLUE
        group_para.alignment = PP_ALIGN.LEFT

        content_top = group_top + Inches(0.35)
        footer_reserved = Inches(1.6)
        content_height = max(Inches(1.0), slide_height - content_top - footer_reserved)

        gap = Inches(0.18)
        col_width = (available_width - gap * (columns - 1)) / columns
        rows_per_col = max(1, math.ceil(total_rows / columns))
        row_height = max(Inches(0.25), content_height / rows_per_col)

        checkbox_size_map = {3: Inches(0.16), 4: Inches(0.15), 5: Inches(0.14)}
        font_map = {3: Pt(18), 4: Pt(15), 5: Pt(13)}
        checkbox_size = checkbox_size_map[columns]
        name_font = font_map[columns]
        text_offset = checkbox_size + Inches(0.06)

        for col_idx in range(columns):
            start = col_idx * rows_per_col
            end = start + rows_per_col
            entries = rows[start:end]
            if not entries:
                continue
            x_start = margin_left + col_idx * (col_width + gap)
            for row_idx, entry in enumerate(entries):
                y_pos = content_top + row_height * row_idx
                box_y = y_pos + (row_height - checkbox_size) / 2
                box = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, x_start, box_y, checkbox_size, checkbox_size
                )
                box.fill.background()
                box.line.color.rgb = RGBColor(0, 0, 0)
                box.line.width = Pt(1)

                text_box = slide.shapes.add_textbox(
                    x_start + text_offset, y_pos, col_width - text_offset, row_height
                )
                tf = text_box.text_frame
                tf.clear()
                tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                para = tf.paragraphs[0]
                name_line = entry.get("name_text") or "Name Candidate"
                pronunciation = entry.get("pronunciation_text")
                katakana = entry.get("katakana_text")
                extra_parts = [part for part in [pronunciation, katakana] if part]
                if extra_parts:
                    name_line = f"{name_line} ({' / '.join(extra_parts)})"
                para.text = name_line
                para.font.size = name_font
                para.font.bold = True
                para.alignment = PP_ALIGN.LEFT
                para.font.color.rgb = RGBColor(0, 0, 0)

        # Group slides intentionally omit footer/comment boxes and logos to keep them clean.

    def _render_name_summary_slide(self, presentation: Presentation, slide) -> None:
        """Render Name Summary slide with blue header bar."""
        slide_width = presentation.slide_width

        # Blue header bar (matching reference image style)
        header_height = Inches(1.0)
        header_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            0,
            0,
            slide_width,
            header_height
        )
        header_box.fill.solid()
        header_box.fill.fore_color.rgb = NAVY_BLUE  # Dark blue like reference image
        header_box.line.fill.background()

        # Header text "Name Candidates - Summary" in white
        header_text_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.25), slide_width - Inches(1.0), Inches(0.5)
        )
        header_tf = header_text_box.text_frame
        header_tf.clear()
        header_para = header_tf.paragraphs[0]
        header_para.text = "Name Candidates - Summary"
        header_para.font.size = Pt(36)
        header_para.font.bold = True
        header_para.font.color.rgb = RGBColor(255, 255, 255)  # White text
        header_para.alignment = PP_ALIGN.LEFT

        logger.info("Rendered Name Summary slide")


# Singleton instance
openxml_pptx_service = OpenXMLPPTXService()
