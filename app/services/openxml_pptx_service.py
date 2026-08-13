"""
OpenXML-based PowerPoint generation service.

This service provides a faster, more versatile alternative to COM-based PowerPoint generation
using the python-pptx library with direct OpenXML manipulation.
"""

import math
import re
import textwrap
import zipfile
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

from app.config.settings import settings
from app.models.presentation_models import ProcessedExcelData
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

NAVY_BLUE = RGBColor(45, 77, 145)

# Front-end fonts (nw2 web viewer ships Gotham Light/Medium, family name "Gotham").
# bold=False resolves to the Light face, bold=True to the Medium face when both
# weights are installed; machines without Gotham fall back to the theme font.
FONT_GOTHAM = "Gotham"
TEXT_BLACK = RGBColor(0, 0, 0)
PLACEHOLDER_GRAY = RGBColor(119, 119, 119)  # #777 (front-end input placeholders)
BORDER_GRAY = RGBColor(187, 187, 187)  # #bbb (front-end input borders)
TOGGLE_TRACK_GRAY = RGBColor(204, 204, 204)  # #ccc (front-end switch track)
CHECKBOX_BORDER = RGBColor(221, 221, 221)  # #ddd (front-end group checkbox border)
RECRAFT_AMBER = RGBColor(251, 192, 45)  # #fbc02d (front-end "R" recraft box)

# ---- Render v2 constants (settings.use_backup_render_v2) --------------------
# Values measured live on the web viewer (group component CSS + DOM metrics,
# 1382px stage = 13.333in slide, so 1px = 0.6947pt).
RECRAFT_AMBER_V2 = RGBColor(249, 168, 37)  # #f9a825 (measured on the live viewer)
# Standalone medium-weight family (v2), used WITH b=1: every renderer starts
# from the same Medium base and applies its synthetic bold, matching the
# original heavy look. Clients resolve it from the embedded regular fntdata;
# the server registers templates/fonts/GothamMediumFamily.ttf transiently
# while converting viewer JPGs (see _transient_gotham_bold).
FONT_GOTHAM_MEDIUM = "Gotham Medium"
TOGGLE_LABEL_GRAY = RGBColor(85, 85, 85)  # #555 (Retain/Recraft card labels)
SWITCH_FILL_GRAY = RGBColor(158, 158, 158)  # #9e9e9e (active NONE segment)
GROUP_ROW_PITCH_IN = 0.54  # 56px row pitch (50px min-height + 6px gap)
GROUP_ROWS_PER_COLUMN = 7  # chunkOptionsIntoColumns(options, 7, 3)
GROUP_NAME_PT = 17.4  # 25px label (fixed, never scales with row count)
GROUP_NAME_SMALL_PT = 13.9  # 20px when the longest rationale exceeds 25 chars


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
        background_images: Optional[List[str]] = None,
        render_v2: bool = False,
    ) -> bytes:
        """
        Generate a backup presentation using OpenXML.

        Args:
            pptx_bytes: Original PowerPoint file as bytes
            excel_data: Processed Excel data from excel_service
            insert_position: Zero-based index where to insert generated slides (from NameCandidateStartingSlide - 1)
            is_big_japanese: Whether to use big Japanese mode (katakana emphasized)
            background_images: Ordered list of absolute image paths for background
                rotation (user's selection). Rotation advances on every
                name-evaluation slide (single and group), but the image is only
                painted on individual slides — mirroring the web viewer.
            render_v2: Feature-flagged render fixes (settings.use_backup_render_v2):
                group slides use the web viewer's exact layout algorithm and text
                properties are promoted to run level for viewer compatibility.

        Returns:
            bytes: Generated PowerPoint file as bytes

        Raises:
            ValueError: If inputs are invalid or processing fails
        """
        try:
            logger.info(
                "Starting OpenXML backup generation at position %d (render_v2=%s)",
                insert_position, render_v2,
            )

            # Load the presentation
            presentation = Presentation(BytesIO(pptx_bytes))

            # Convert ProcessedExcelData to the format expected by slide builder
            excel_rows = self._convert_processed_excel_to_rows(excel_data)

            # Build slides from Excel data
            inserted = self._build_slides_from_excel(
                presentation, excel_rows, insert_position, is_big_japanese,
                background_images=background_images,
                render_v2=render_v2,
            )

            logger.info("Inserted %d slides using OpenXML", inserted)

            if render_v2:
                # Copy paragraph-level defRPr onto runs lacking rPr so viewers
                # that ignore defRPr (PowerPoint Online, Google Slides,
                # LibreOffice) still render Gotham at the intended size instead
                # of falling back to a different default per text box.
                # Only the slides we generated are touched (contiguous block +
                # the Name Summary slide appended at the end).
                generated = list(presentation.slides)[insert_position:insert_position + inserted]
                for gen_slide in generated:
                    self._promote_paragraph_props_to_runs(gen_slide)

            # Save to bytes
            buffer = BytesIO()
            presentation.save(buffer)
            buffer.seek(0)
            result_bytes = buffer.read()

            if render_v2:
                # Embed Gotham in-process (zip surgery, ~milliseconds). Replaces
                # the old PowerPoint COM embed, which cold-started PowerPoint and
                # subsetted every font family (~60-90s per backup).
                result_bytes = self._embed_gotham_fonts(result_bytes)

            return result_bytes

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
            "GroupType",
            "GroupLetter",
        ]

        # Create data rows
        rows = [headers]

        # Use total_rows_processed to get ALL rows (including category headers)
        total_rows = getattr(excel_data, 'total_rows_processed', len(excel_data.lst_types))

        logger.info(f"Converting ProcessedExcelData: total_rows={total_rows}, types={len(excel_data.lst_types)}, names={len(excel_data.lst_names)}, categories={len(excel_data.lst_categories)}")

        group_types = getattr(excel_data, "lst_group_types", []) or []
        group_letters = getattr(excel_data, "lst_group_letters", []) or []

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
                group_types[i] if i < len(group_types) else "",
                group_letters[i] if i < len(group_letters) else "",
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

    def _apply_background_image(self, presentation: Presentation, slide, image_path: str) -> None:
        """Insert a full-bleed background image at the bottom of the z-order."""
        try:
            pic = slide.shapes.add_picture(
                str(image_path), 0, 0, presentation.slide_width, presentation.slide_height
            )
            # add_picture appends on top; move it to the back so slide content stays visible
            sp_tree = slide.shapes._spTree  # type: ignore[attr-defined]
            sp_tree.remove(pic._element)
            sp_tree.insert(2, pic._element)
        except Exception as e:
            logger.warning("Could not apply background image %s: %s", image_path, str(e))

    def _apply_solid_white_background(self, presentation: Presentation, slide) -> None:
        """Cover the inherited layout/master art with a plain white full-bleed rectangle.

        Name-evaluation slides render on a clean white canvas in the web viewer;
        shells whose layout carries title bars, logos or footers would otherwise
        bleed into the generated slides.
        """
        try:
            rect = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, 0, 0, presentation.slide_width, presentation.slide_height
            )
            rect.fill.solid()
            rect.fill.fore_color.rgb = RGBColor(255, 255, 255)
            rect.line.fill.background()
            rect.shadow.inherit = False
            sp_tree = slide.shapes._spTree  # type: ignore[attr-defined]
            sp_tree.remove(rect._element)
            sp_tree.insert(2, rect._element)
        except Exception as e:
            logger.warning("Could not apply white background: %s", str(e))

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
        background_images: Optional[List[str]] = None,
        render_v2: bool = False,
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

        # Note: pronunciation_idx contains notation data (e.g., "C", "T", "CB")
        # We'll extract and display it separately in the bottom-right corner
        pronunciation_idx = headers.get("pronunciation")
        katakana_idx = headers.get("katakana")
        group1_idx = headers.get("group1")
        group2_idx = headers.get("group2")
        group_type_idx = headers.get("grouptype")
        group_letter_idx = headers.get("groupletter")

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

            # Extract notation (e.g., "C", "T", "CB") from pronunciation column
            notation_text = (
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
                # Names and rationales are parallel ##-joined lists (see excel_service):
                # split both WITHOUT filtering so indexes stay aligned, then drop
                # entries whose name is empty (e.g. the trailing delimiter).
                names = [n.strip() for n in name_text.split("##")]
                rationales = [r.strip() for r in rationale_text.split("##")]
                rows_for_group = []
                for idx, name in enumerate(names):
                    if not name:
                        continue
                    rows_for_group.append({
                        "category_text": category_text,
                        "name_text": name,
                        "rationale_text": rationales[idx] if idx < len(rationales) else "",
                        "pronunciation_text": "",
                        "katakana_text": "",
                    })

                # group1 hides rationales in the front-end (hover only); group2 shows them
                group_type = (
                    str(self._get_cell(row, group_type_idx) or "").strip().lower()
                    if group_type_idx is not None
                    else ""
                )

                # Group letter (A/AR/B/C) decides the vote-control variant
                group_letter = (
                    str(self._get_cell(row, group_letter_idx) or "").strip().upper()
                    if group_letter_idx is not None
                    else ""
                )

                # Add group slide
                slide_items.append({
                    "type": "group",
                    "group_label": group_label,
                    "rows": rows_for_group,
                    "category_text": category_text,
                    "show_rationales": group_type != "group1",
                    "group_letter": group_letter,
                })
            elif name_text:
                # Individual slide
                row_payload = {
                    "category_text": category_text,
                    "name_text": name_text,
                    "rationale_text": rationale_text,
                    "notation_text": notation_text,
                    "katakana_text": katakana_text,
                }
                slide_items.append({"type": "single", "row": row_payload})

        current_position = start_position
        inserted = 0
        bg_count = len(background_images) if background_images else 0
        bg_index = 0
        for item in slide_items:
            slide = self.insert_slide_with_background(presentation, current_position)

            # Name-evaluation slides render on a clean white canvas (web viewer
            # behavior). Background rotation (user's selection): the index advances
            # on every name-evaluation slide (single AND group) — same as
            # nw_Details — but the image is only painted on individual slides.
            if item["type"] in ("single", "group"):
                bg_path = None
                if bg_count:
                    bg_path = background_images[bg_index % bg_count]
                    bg_index += 1
                if item["type"] == "single" and bg_path:
                    self._apply_background_image(presentation, slide, bg_path)
                else:
                    self._apply_solid_white_background(presentation, slide)

            if item["type"] == "category":
                self._render_category_slide(
                    presentation, slide, item["category_text"], render_v2=render_v2
                )
            elif item["type"] == "group":
                self._render_group_slide(
                    presentation,
                    slide,
                    category_text=item.get("category_text") or "",
                    group_label=item["group_label"],
                    rows=item["rows"],
                    show_rationales=item.get("show_rationales", True),
                    group_letter=item.get("group_letter", ""),
                    render_v2=render_v2,
                )
            else:
                row = item["row"]
                self._render_name_rationale_slide(
                    presentation,
                    slide,
                    category_text=row["category_text"],
                    name_text=row["name_text"],
                    rationale_text=row["rationale_text"],
                    notation_text=row["notation_text"],
                    katakana_text=row["katakana_text"],
                    is_big_japanese=is_big_japanese,
                    render_v2=render_v2,
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
        self, presentation: Presentation, slide, category_text: str, render_v2: bool = False
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

        # v2: "Gotham Medium" + b=1 — both server (transient TTF) and client
        # (embedded regular fntdata) start from the same Medium base and both
        # apply synthetic bold, reproducing the original heavy look everywhere.
        cat_font = FONT_GOTHAM_MEDIUM if render_v2 else FONT_GOTHAM
        cat_bold = True

        for idx, line in enumerate(top_lines):
            para = tf.paragraphs[idx] if idx == 0 else tf.add_paragraph()
            para.text = line
            para.alignment = PP_ALIGN.CENTER
            para.font.name = cat_font
            para.font.size = top_font
            para.font.bold = cat_bold
            para.font.color.rgb = RGBColor(0, 0, 0)

        for line in bottom_lines:
            para = tf.add_paragraph()
            para.text = line
            para.alignment = PP_ALIGN.CENTER
            para.font.name = cat_font
            para.font.size = bottom_font
            para.font.bold = cat_bold
            para.font.color.rgb = RGBColor(0, 0, 139)  # deep navy

    def _render_name_rationale_slide(
        self,
        presentation: Presentation,
        slide,
        category_text: str,
        name_text: str,
        rationale_text: str,
        notation_text: str = "",
        katakana_text: str = "",
        is_big_japanese: bool = False,
        render_v2: bool = False,
    ) -> None:
        """Render name/rationale slide with a simple layout similar to the provided design."""
        slide_width = presentation.slide_width
        slide_height = presentation.slide_height

        # Horizontal scale factor for non-16:9 decks (see DESIGN_WIDTH_IN)
        sx = slide_width / Inches(self.DESIGN_WIDTH_IN)

        # Category heading — front-end: GothamLight 30px (#000), x=53px, y≈36px
        header_top = Inches(0.36)
        header_box = slide.shapes.add_textbox(
            self._x(slide_width, 0.56), header_top, slide_width - Inches(1.2), Inches(0.5)
        )
        header_tf = header_box.text_frame
        header_tf.clear()
        header_para = header_tf.paragraphs[0]
        header_para.text = category_text or "Category"
        header_para.font.name = FONT_GOTHAM
        header_para.font.size = Pt(21)
        header_para.font.bold = False
        header_para.font.color.rgb = TEXT_BLACK
        header_para.alignment = PP_ALIGN.LEFT

        # Split the name at the FIRST parenthesis (front-end behavior):
        # "Healthy Futures (C)" -> main name "Healthy Futures" + notation "(C)".
        # If there is no "(" (or it is the first character), the whole text stays
        # as the main name and no corner notation is shown (e.g. "TruPrevent").
        raw_name = (name_text or "").strip()
        paren_idx = raw_name.find("(")
        if paren_idx > 0:
            main_name = raw_name[:paren_idx].strip()
            split_notation = raw_name[paren_idx:].strip()
        else:
            main_name = raw_name
            split_notation = ""
        effective_notation = (notation_text or "").strip() or split_notation

        # Name and katakana block — front-end: GothamMedium. Render v2: measured
        # live on the viewer, 110px on a 1382px stage = 76pt, box top y≈2.22in.
        # v1 keeps the original 55pt (kept for the flag-off regression).
        left = self._x(slide_width, 0.62)
        top_block = Inches(2.22) if render_v2 else Inches(2.35)
        block_width = slide_width - left - Inches(0.75)
        block_height = Inches(1.5) if render_v2 else Inches(1.4)
        name_box = slide.shapes.add_textbox(left, top_block, block_width, block_height)
        name_tf = name_box.text_frame
        name_tf.clear()
        name_tf.word_wrap = True
        display_name = (
            katakana_text if (is_big_japanese and katakana_text) else (main_name or "Name")
        )
        name_para = name_tf.paragraphs[0]
        name_para.text = display_name
        if render_v2:
            name_font_size = self._fit_font_size(
                display_name, base_size=76, min_size=32, max_chars=max(12, int(20 * sx))
            )
        else:
            name_font_size = self._fit_font_size(
                display_name, base_size=55, min_size=28, max_chars=max(14, int(27 * sx))
            )
        if render_v2:
            # Pure Medium, NO bold: the viewer renders the big name with the
            # GothamMedium webfont at w=500 (no synthesis). Clients resolve the
            # family from the embedded regular fntdata; the server from the
            # transiently registered TTF. (Category slides differ: they DO use
            # b=1 to reproduce the viewer's synthetic-bold JPG look.)
            name_para.font.name = FONT_GOTHAM_MEDIUM
            name_para.font.bold = False
        else:
            name_para.font.name = FONT_GOTHAM
            name_para.font.bold = True  # v1: resolves to Medium+synthetic on the server
        name_para.font.size = Pt(name_font_size)
        name_para.alignment = PP_ALIGN.LEFT
        name_para.font.color.rgb = TEXT_BLACK

        # Show name/katakana text below the main display name
        small_parts: List[str] = []
        if is_big_japanese and katakana_text:
            if main_name:
                small_parts.append(main_name)
        else:
            if katakana_text:
                small_parts.append(katakana_text)
        if small_parts:
            sub_para = name_tf.add_paragraph()
            sub_para.text = " / ".join(small_parts)
            sub_para.font.name = FONT_GOTHAM
            sub_para.font.size = Pt(22)
            sub_para.font.bold = False
            sub_para.alignment = PP_ALIGN.LEFT
            sub_para.font.color.rgb = TEXT_BLACK

        # Rationale — front-end: GothamLight 30px (~21pt), y≈4.1in
        rat_title_top = Inches(4.05)
        rat_title = slide.shapes.add_textbox(left, rat_title_top, block_width, Inches(1.0))
        rat_tf = rat_title.text_frame
        rat_tf.clear()
        rat_tf.word_wrap = True
        rat_lines = self._wrap_text(
            rationale_text or "Rationale goes here", max_chars=max(35, int(70 * sx)), max_lines=2
        )
        for idx, line in enumerate(rat_lines):
            para = rat_tf.paragraphs[idx] if idx == 0 else rat_tf.add_paragraph()
            para.text = line
            para.font.name = FONT_GOTHAM
            para.font.size = Pt(21)
            para.font.bold = False
            para.alignment = PP_ALIGN.LEFT
            para.font.color.rgb = TEXT_BLACK

        # Sentiment toggle row — front-end: starts x≈2.28in, pitch 1.62in, y≈5.5in.
        # The pitch has a floor so track + label never collide on narrow decks.
        sentiment_top = Inches(5.5)
        toggle_start = self._x(slide_width, 2.28)
        toggle_spacing = max(Inches(1.45), self._x(slide_width, 1.62))
        for idx, toggle_label in enumerate(["Positive", "Neutral", "Negative", "Recraft"]):
            self._add_toggle_switch(
                slide, toggle_start + toggle_spacing * idx, sentiment_top, toggle_label,
                render_v2=render_v2,
            )

        # Notation on the toggle row (e.g., "(C)", "(T)", "(CB)") — front-end: x≈10.9in
        if effective_notation:
            formatted_notation = f"({effective_notation})" if not effective_notation.startswith("(") else effective_notation
            notation_box = slide.shapes.add_textbox(
                self._x(slide_width, 10.2), sentiment_top - Inches(0.06), self._x(slide_width, 1.0), Inches(0.3)
            )
            notation_tf = notation_box.text_frame
            notation_tf.clear()
            notation_para = notation_tf.paragraphs[0]
            notation_para.text = formatted_notation
            notation_para.font.name = FONT_GOTHAM
            notation_para.font.size = Pt(11)
            notation_para.font.bold = False
            notation_para.alignment = PP_ALIGN.RIGHT
            notation_para.font.color.rgb = TEXT_BLACK

        # New Names / New Comments boxes — front-end: 4.04in wide, 0.77in tall,
        # left box at x=0.77in, right box at x=8.51in, y=6.04in (16:9 reference)
        box_w = self._x(slide_width, 4.04)
        self._add_input_box(slide, self._x(slide_width, 0.77), Inches(6.04), box_w, Inches(0.77), "New Names")
        self._add_input_box(slide, self._x(slide_width, 8.51), Inches(6.04), box_w, Inches(0.77), "New Comments")

    def _add_input_box(self, slide, left, top, width, height, placeholder: str) -> None:
        """Draw an input-style box with a gray left-aligned placeholder (front-end look)."""
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        box.fill.background()
        box.line.color.rgb = BORDER_GRAY
        box.line.width = Pt(0.75)
        box.shadow.inherit = False
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.15)
        para = tf.paragraphs[0]
        para.text = placeholder
        para.font.name = FONT_GOTHAM
        para.font.size = Pt(11)
        para.font.bold = False
        para.font.color.rgb = PLACEHOLDER_GRAY
        para.alignment = PP_ALIGN.LEFT

    def _add_toggle_switch(self, slide, left, top, label: str, render_v2: bool = False) -> None:
        """Draw a small pill-style toggle switch with a label, mimicking the
        front-end vote controls (Positive / Neutral / Negative / Recraft).

        v2: the front end renders these labels in Roboto 14px REGULAR (w=400,
        no bold) — a bold label reads much heavier than the app, so v2 drops
        the bold flag and uses 10pt."""
        track_w = Inches(0.35)
        track_h = Inches(0.19)
        track = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, track_w, track_h
        )
        try:
            track.adjustments[0] = 0.5  # fully rounded ends (pill shape)
        except Exception:
            pass
        track.fill.solid()
        track.fill.fore_color.rgb = TOGGLE_TRACK_GRAY
        track.line.fill.background()
        track.shadow.inherit = False

        knob_d = Inches(0.15)
        knob = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            left + Inches(0.02),
            top + (track_h - knob_d) / 2,
            knob_d,
            knob_d,
        )
        knob.fill.solid()
        knob.fill.fore_color.rgb = RGBColor(255, 255, 255)
        knob.line.color.rgb = RGBColor(170, 170, 170)
        knob.line.width = Pt(0.5)
        knob.shadow.inherit = False

        label_box = slide.shapes.add_textbox(
            left + track_w + Inches(0.12), top - Inches(0.05), Inches(1.2), Inches(0.3)
        )
        tf = label_box.text_frame
        tf.clear()
        para = tf.paragraphs[0]
        para.text = label
        para.font.name = FONT_GOTHAM
        para.font.size = Pt(10) if render_v2 else Pt(11)
        para.font.bold = not render_v2  # front: Roboto 14px w=400 (regular)
        para.font.color.rgb = TEXT_BLACK
        para.alignment = PP_ALIGN.LEFT

    # Groups larger than this fall back to the compact grid: with more rows the
    # rationale column would no longer be legible on a fixed-height slide.
    MAX_LIST_ROWS = 14

    # Horizontal positions were measured on the 16:9 web viewer (13.333in wide).
    # 4:3 decks (10in wide) would push right-side elements off the canvas, so
    # x-coordinates scale by the actual slide width. Vertical positions need no
    # scaling: PowerPoint slides are 7.5in tall in both aspect ratios.
    DESIGN_WIDTH_IN = 13.333

    def _x(self, slide_width, design_inches: float) -> int:
        """Scale a horizontal design coordinate (in inches, 16:9 reference) to the deck's width."""
        return int(slide_width * design_inches / self.DESIGN_WIDTH_IN)

    def _render_group_slide(
        self,
        presentation: Presentation,
        slide,
        category_text: str,
        group_label: str,
        rows: List[Dict[str, str]],
        show_rationales: bool = True,
        group_letter: str = "",
        render_v2: bool = False,
    ) -> None:
        """Render a grouped slide mirroring the front-end layout.

        Default: a vertical list where each row is `[checkbox] [R] Name .... Rationale`,
        matching how the web viewer shows grouped candidates. Falls back to the
        compact multi-column checkbox grid (names only) when the group is too
        large for the list to stay legible.

        The group letter (prefix of NameGroup) decides the vote-control variant,
        same as the front-end GroupSlideComponent:
        - A / AR / empty (old projects): standard checkbox per name
        - AR and B: additional amber "R" (recraft) box per name
        - B: cyclic vote button — its default state looks like a checkbox
        - C: voting disabled — names only, no boxes, aligned to the left margin
        """
        if render_v2:
            self._render_group_slide_v2(
                presentation,
                slide,
                category_text=category_text,
                rows=rows,
                show_rationales=show_rationales,
                group_letter=group_letter,
            )
            return

        slide_width = presentation.slide_width
        slide_height = presentation.slide_height

        total_rows = len(rows)

        margin_left = Inches(0.25)
        margin_right = Inches(0.4)
        available_width = slide_width - margin_left - margin_right

        # Category heading — front-end: GothamLight 30px (#000), x≈0.22in, y≈0.34in
        heading_top = Inches(0.34)
        heading_box = slide.shapes.add_textbox(
            margin_left, heading_top, available_width, Inches(0.5)
        )
        heading_tf = heading_box.text_frame
        heading_tf.clear()
        heading_para = heading_tf.paragraphs[0]
        heading_para.text = category_text or "Category"
        heading_para.font.name = FONT_GOTHAM
        heading_para.font.size = Pt(21)
        heading_para.font.bold = False
        heading_para.font.color.rgb = TEXT_BLACK
        heading_para.alignment = PP_ALIGN.LEFT

        # Note: Group label (e.g., "Group a1") is intentionally not shown in backup presentations
        # to keep slides clean and avoid cluttering the view

        content_top = heading_top + Inches(0.8)
        footer_reserved = Inches(1.6)
        content_height = max(Inches(1.0), slide_height - content_top - footer_reserved)

        # Vote-control variant from the group letter (front-end GroupSlideComponent)
        letter = (group_letter or "").strip().upper()
        show_checkbox = not letter.startswith("C")
        show_recraft = letter == "AR" or letter.startswith("B")

        if total_rows <= self.MAX_LIST_ROWS:
            self._render_group_list(
                slide,
                rows,
                margin_left=margin_left,
                margin_right=margin_right,
                slide_width=slide_width,
                available_width=available_width,
                content_top=content_top,
                content_height=content_height,
                show_rationales=show_rationales,
                show_checkbox=show_checkbox,
                show_recraft=show_recraft,
            )
            # New Names / New Comments boxes at the bottom (front-end group layout,
            # 16:9 reference coordinates scaled to the deck width)
            group_box_w = self._x(slide_width, 4.04)
            self._add_input_box(slide, self._x(slide_width, 1.99), Inches(6.05), group_box_w, Inches(0.8), "New Names")
            self._add_input_box(slide, self._x(slide_width, 7.27), Inches(6.05), group_box_w, Inches(0.8), "New Comments")
        else:
            self._render_group_grid(
                slide,
                rows,
                margin_left=margin_left,
                available_width=available_width,
                content_top=content_top,
                content_height=content_height,
                show_checkbox=show_checkbox,
                show_recraft=show_recraft,
            )

        # Group slides intentionally omit footer/comment boxes and logos to keep them clean.

    def _render_group_list(
        self,
        slide,
        rows: List[Dict[str, str]],
        *,
        margin_left,
        margin_right,
        slide_width,
        available_width,
        content_top,
        content_height,
        show_rationales: bool = True,
        show_checkbox: bool = True,
        show_recraft: bool = False,
    ) -> None:
        """Vertical list layout: `[checkbox] [R] Name .... Rationale` per row (front-end style)."""
        total_rows = len(rows)

        # Front-end geometry: checkbox at x≈0.25in (0.21in square, #ddd border),
        # amber "R" box at x≈0.55in, name column at x=1.15in, rationale column at
        # x=4.91in, ~0.54in row pitch with rows packed at the top. Mode C
        # (no voting) shows the names directly at the left margin.
        checkbox_size = Inches(0.21)
        recraft_left = Inches(0.55)
        name_left = self._x(slide_width, 1.15) if show_checkbox else margin_left
        rationale_left = self._x(slide_width, 4.91)
        name_col_width = rationale_left - name_left - Inches(0.2)
        rationale_width = slide_width - margin_right - rationale_left

        row_height = min(Inches(0.55), int(content_height / max(1, total_rows)))
        if total_rows <= 8:
            name_font, rationale_font = Pt(17), Pt(16)
        elif total_rows <= 11:
            name_font, rationale_font = Pt(15), Pt(14)
        else:
            name_font, rationale_font = Pt(13), Pt(12)

        for row_idx, entry in enumerate(rows):
            y_pos = content_top + row_height * row_idx
            box_y = y_pos + (row_height - checkbox_size) / 2

            if show_checkbox:
                box = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, margin_left, box_y, checkbox_size, checkbox_size
                )
                box.fill.background()
                box.line.color.rgb = CHECKBOX_BORDER
                box.line.width = Pt(1)
                box.shadow.inherit = False

            if show_checkbox and show_recraft:
                self._add_recraft_box(slide, recraft_left, box_y, checkbox_size)

            name_line = self._compose_group_name_line(entry)
            name_box = slide.shapes.add_textbox(
                name_left, y_pos, name_col_width, row_height
            )
            tf = name_box.text_frame
            tf.clear()
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = tf.paragraphs[0]
            para.text = name_line
            para.font.name = FONT_GOTHAM
            para.font.size = name_font
            para.font.bold = False
            para.alignment = PP_ALIGN.LEFT
            para.font.color.rgb = TEXT_BLACK

            rationale_text = (entry.get("rationale_text") or "").strip()
            if show_rationales and rationale_text:
                rat_box = slide.shapes.add_textbox(
                    rationale_left, y_pos, rationale_width, row_height
                )
                rtf = rat_box.text_frame
                rtf.clear()
                rtf.word_wrap = True
                rtf.vertical_anchor = MSO_ANCHOR.MIDDLE
                rpara = rtf.paragraphs[0]
                rpara.text = rationale_text
                rpara.font.name = FONT_GOTHAM
                rpara.font.size = rationale_font
                rpara.font.bold = False
                rpara.alignment = PP_ALIGN.LEFT
                rpara.font.color.rgb = TEXT_BLACK

    def _add_recraft_box(self, slide, left, top, size, color: Optional[RGBColor] = None) -> None:
        """Amber-bordered 'R' recraft box, mirroring the front-end control."""
        amber = color if color is not None else RECRAFT_AMBER
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, size, size)
        box.fill.background()
        box.line.color.rgb = amber
        box.line.width = Pt(1.25)
        box.shadow.inherit = False
        tf = box.text_frame
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.text = "R"
        para.font.name = FONT_GOTHAM
        para.font.size = Pt(10)
        para.font.bold = True
        para.font.color.rgb = amber
        para.alignment = PP_ALIGN.CENTER

    def _render_group_grid(
        self,
        slide,
        rows: List[Dict[str, str]],
        *,
        margin_left,
        available_width,
        content_top,
        content_height,
        show_checkbox: bool = True,
        show_recraft: bool = False,
    ) -> None:
        """Compact multi-column checkbox grid (names only) for large groups."""
        total_rows = len(rows)
        columns = 3 if total_rows <= 24 else 4 if total_rows <= 48 else 5

        gap = Inches(0.18)
        col_width = (available_width - gap * (columns - 1)) / columns
        rows_per_col = max(1, math.ceil(total_rows / columns))
        row_height = max(Inches(0.25), content_height / rows_per_col)

        checkbox_size_map = {3: Inches(0.16), 4: Inches(0.15), 5: Inches(0.14)}
        font_map = {3: Pt(18), 4: Pt(15), 5: Pt(13)}
        checkbox_size = checkbox_size_map[columns]
        name_font = font_map[columns]
        text_offset = Inches(0.02)
        if show_checkbox:
            text_offset += checkbox_size + Inches(0.06)
            if show_recraft:
                text_offset += checkbox_size + Inches(0.06)

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
                if show_checkbox:
                    box = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE, x_start, box_y, checkbox_size, checkbox_size
                    )
                    box.fill.background()
                    box.line.color.rgb = CHECKBOX_BORDER
                    box.line.width = Pt(1)
                    box.shadow.inherit = False
                    if show_recraft:
                        self._add_recraft_box(
                            slide, x_start + checkbox_size + Inches(0.06), box_y, checkbox_size
                        )

                text_box = slide.shapes.add_textbox(
                    x_start + text_offset, y_pos, col_width - text_offset, row_height
                )
                tf = text_box.text_frame
                tf.clear()
                tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                para = tf.paragraphs[0]
                para.text = self._compose_group_name_line(entry)
                para.font.name = FONT_GOTHAM
                para.font.size = name_font
                para.font.bold = False
                para.alignment = PP_ALIGN.LEFT
                para.font.color.rgb = TEXT_BLACK

    # ------------------------------------------------------------------
    # Render v2 (settings.use_backup_render_v2) — exact web-viewer algorithm
    # ------------------------------------------------------------------

    def _render_group_slide_v2(
        self,
        presentation: Presentation,
        slide,
        *,
        category_text: str,
        rows: List[Dict[str, str]],
        show_rationales: bool = True,
        group_letter: str = "",
    ) -> None:
        """Group slide replicating the web viewer's GroupSlideComponent exactly.

        Front-end algorithm (extracted from the nw2 Angular bundle):
        - Layout chosen by rationale VISIBILITY, not by row count: visible
          rationales -> single-column list (name | rationale); hidden/empty
          rationales -> columns of max 7 rows (chunkOptionsIntoColumns(_, 7, 3)),
          packed at the top with a fixed 0.54in pitch.
        - Name font is FIXED at 25px (17.4pt) — 20px (13.9pt) only when the
          longest rationale exceeds 25 characters. It never scales with count.
        - Every group slide shows the Retain/Recraft NONE-ALL card (when voting
          is enabled) and the New Names / New Comments boxes.

        Deliberate deviation: the front end silently truncates groups at 21
        names (3 columns x 7). The backup keeps ALL names by adding columns
        (and shrinking the font at 4-5 columns) — a paper form must be complete.
        """
        slide_width = presentation.slide_width
        total_rows = len(rows)

        margin_left = Inches(0.25)
        margin_right = Inches(0.4)
        available_width = slide_width - margin_left - margin_right

        # Category heading — same as v1 (matches front: GothamLight 30px ~21pt)
        heading_top = Inches(0.34)
        heading_box = slide.shapes.add_textbox(
            margin_left, heading_top, available_width, Inches(0.5)
        )
        heading_para = heading_box.text_frame.paragraphs[0]
        heading_para.text = category_text or "Category"
        heading_para.font.name = FONT_GOTHAM
        heading_para.font.size = Pt(21)
        heading_para.font.bold = False
        heading_para.font.color.rgb = TEXT_BLACK
        heading_para.alignment = PP_ALIGN.LEFT

        letter = (group_letter or "").strip().upper()
        show_checkbox = not letter.startswith("C")
        show_recraft = letter == "AR" or letter.startswith("B")

        rationale_values = [(entry.get("rationale_text") or "").strip() for entry in rows]
        rationales_visible = show_rationales and any(rationale_values)

        content_top = Inches(1.14)
        card_top = Inches(5.22)
        content_height = card_top - content_top - Inches(0.05)

        # Front: getDynamicFontSize() — 25px, or 20px when the longest visible
        # rationale exceeds 25 characters.
        if rationales_visible and max(len(r) for r in rationale_values) > 25:
            name_pt = GROUP_NAME_SMALL_PT
        else:
            name_pt = GROUP_NAME_PT

        checkbox_size = Inches(0.24)  # front checkbox input: 25px square

        if rationales_visible:
            # Single-column list: [checkbox] [R] Name .... Rationale
            row_pitch = min(Inches(GROUP_ROW_PITCH_IN), int(content_height / max(1, total_rows)))
            name_left = self._x(slide_width, 1.15) if show_checkbox else margin_left
            rationale_left = self._x(slide_width, 4.91)
            name_width = rationale_left - name_left - Inches(0.2)
            rationale_width = slide_width - margin_right - rationale_left

            for row_idx, entry in enumerate(rows):
                y_pos = content_top + row_pitch * row_idx
                box_y = y_pos + (row_pitch - checkbox_size) / 2
                if show_checkbox:
                    self._add_group_checkbox_v2(slide, margin_left, box_y, checkbox_size)
                    if show_recraft:
                        self._add_recraft_box(
                            slide, self._x(slide_width, 0.55), box_y, checkbox_size,
                            color=RECRAFT_AMBER_V2,
                        )
                self._add_group_text_v2(
                    slide, name_left, y_pos, name_width, row_pitch,
                    self._compose_group_name_line(entry), name_pt,
                )
                rationale_text = rationale_values[row_idx]
                if rationale_text:
                    self._add_group_text_v2(
                        slide, rationale_left, y_pos, rationale_width, row_pitch,
                        rationale_text, name_pt,
                    )
        else:
            # Columns of max 7 rows, packed at the top (front: chunkOptionsIntoColumns).
            # The front caps at 3 columns and DROPS the rest; the backup instead
            # adds columns (4-5) with a smaller font so every name is present.
            ncols = max(1, math.ceil(total_rows / GROUP_ROWS_PER_COLUMN))
            if ncols > 5:
                ncols = 5
            rows_per_col = max(GROUP_ROWS_PER_COLUMN, math.ceil(total_rows / ncols))
            row_pitch = min(Inches(GROUP_ROW_PITCH_IN), int(content_height / rows_per_col))
            # Front: column pitch is FIXED at available/3 regardless of how many
            # columns are populated (measured: col2 at the same x with 2 or 3
            # columns). Only our 4-5 column extension packs them tighter.
            col_pitch = int(available_width / max(3, ncols))
            if ncols == 4:
                name_pt = 14.0
            elif ncols >= 5:
                name_pt = 12.0

            # Front offsets within a column: checkbox at the column start, "R"
            # box at +0.30in, name at +0.89in (all 16:9 reference, x-scaled).
            recraft_offset = self._x(slide_width, 0.30)
            text_offset = self._x(slide_width, 0.89) if show_checkbox else 0

            for col_idx in range(ncols):
                entries = rows[col_idx * rows_per_col:(col_idx + 1) * rows_per_col]
                col_x = margin_left + col_pitch * col_idx
                for row_idx, entry in enumerate(entries):
                    y_pos = content_top + row_pitch * row_idx
                    box_y = y_pos + (row_pitch - checkbox_size) / 2
                    if show_checkbox:
                        self._add_group_checkbox_v2(slide, col_x, box_y, checkbox_size)
                        if show_recraft:
                            self._add_recraft_box(
                                slide, col_x + recraft_offset, box_y, checkbox_size,
                                color=RECRAFT_AMBER_V2,
                            )
                    self._add_group_text_v2(
                        slide, col_x + text_offset, y_pos,
                        col_pitch - text_offset - Inches(0.1), row_pitch,
                        self._compose_group_name_line(entry), name_pt,
                    )

        # Retain / Recraft NONE-ALL card (front: controls-card, shown whenever
        # voting is enabled; letter C disables voting entirely).
        if show_checkbox:
            self._add_retain_recraft_card_v2(
                slide, slide_width, card_top, include_recraft=show_recraft
            )

        # New Names / New Comments boxes — on EVERY group layout (front behavior)
        group_box_w = self._x(slide_width, 4.04)
        self._add_input_box(slide, self._x(slide_width, 1.99), Inches(6.05), group_box_w, Inches(0.8), "New Names")
        self._add_input_box(slide, self._x(slide_width, 7.27), Inches(6.05), group_box_w, Inches(0.8), "New Comments")

    def _add_group_checkbox_v2(self, slide, left, top, size) -> None:
        """Empty vote checkbox (#ddd border), front-end style."""
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, size, size)
        box.fill.background()
        box.line.color.rgb = CHECKBOX_BORDER
        box.line.width = Pt(1)
        box.shadow.inherit = False

    def _add_group_text_v2(self, slide, left, top, width, height, text: str, size_pt: float) -> None:
        """Vertically-centered Gotham Light black text (group rows)."""
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        para.text = text
        para.font.name = FONT_GOTHAM
        para.font.size = Pt(size_pt)
        para.font.bold = False
        para.font.color.rgb = TEXT_BLACK
        para.alignment = PP_ALIGN.LEFT

    def _add_retain_recraft_card_v2(
        self, slide, slide_width, card_top, *, include_recraft: bool
    ) -> None:
        """Bottom-left Retain/Recraft card with NONE-ALL switches (front controls-card)."""
        card_left = self._x(slide_width, 0.19)
        card_width = self._x(slide_width, 3.60 if include_recraft else 1.90)
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, card_left, card_top, card_width, Inches(0.78)
        )
        try:
            card.adjustments[0] = 0.08
        except Exception:
            pass
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = CHECKBOX_BORDER
        card.line.width = Pt(0.75)
        card.shadow.inherit = False

        toggles = [("Retain", 0.36)]
        if include_recraft:
            toggles.append(("Recraft", 2.06))
        for label_text, x_ref in toggles:
            x_pos = self._x(slide_width, x_ref)
            label_box = slide.shapes.add_textbox(
                x_pos, card_top + Inches(0.05), self._x(slide_width, 1.5), Inches(0.25)
            )
            lp = label_box.text_frame.paragraphs[0]
            lp.text = label_text
            lp.font.name = FONT_GOTHAM
            lp.font.size = Pt(11)
            lp.font.bold = False
            lp.font.color.rgb = TOGGLE_LABEL_GRAY
            lp.alignment = PP_ALIGN.LEFT
            self._add_none_all_switch_v2(
                slide, x_pos, card_top + Inches(0.37), self._x(slide_width, 1.50), Inches(0.28)
            )

    def _add_none_all_switch_v2(self, slide, left, top, width, height) -> None:
        """Two-segment NONE/ALL switch, NONE active (default state on the viewer)."""
        half = int(width / 2)
        track = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        track.fill.solid()
        track.fill.fore_color.rgb = RGBColor(255, 255, 255)
        track.line.color.rgb = CHECKBOX_BORDER
        track.line.width = Pt(0.75)
        track.shadow.inherit = False

        none_seg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, half, height)
        none_seg.fill.solid()
        none_seg.fill.fore_color.rgb = SWITCH_FILL_GRAY
        none_seg.line.fill.background()
        none_seg.shadow.inherit = False
        tf = none_seg.text_frame
        tf.margin_left = 0
        tf.margin_right = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        np_ = tf.paragraphs[0]
        np_.text = "NONE"
        np_.font.name = FONT_GOTHAM
        np_.font.size = Pt(8)
        np_.font.bold = True
        np_.font.color.rgb = RGBColor(255, 255, 255)
        np_.alignment = PP_ALIGN.CENTER

        all_box = slide.shapes.add_textbox(left + half, top, width - half, height)
        atf = all_box.text_frame
        atf.margin_left = 0
        atf.margin_right = 0
        atf.vertical_anchor = MSO_ANCHOR.MIDDLE
        ap = atf.paragraphs[0]
        ap.text = "ALL"
        ap.font.name = FONT_GOTHAM
        ap.font.size = Pt(8)
        ap.font.bold = True
        ap.font.color.rgb = PLACEHOLDER_GRAY
        ap.alignment = PP_ALIGN.CENTER

    def _embed_gotham_fonts(self, pptx_bytes: bytes) -> bytes:
        """Embed Gotham into the PPTX package (pure OpenXML zip surgery, no COM).

        The payload is NOT a raw TTF: PowerPoint only accepts fntdata parts in
        its own TTEmbed compressed format — raw TTFs trigger the "Install
        Embedded Fonts ... General Failure" dialog on machines without the
        font. The assets in templates/fonts/Gotham_*.fntdata were produced ONCE
        by PowerPoint itself (scratchpad make_fntdata.py: a deck whose Gotham
        runs contain the full printable Latin-1 charset, saved via COM with
        EmbedTrueTypeFonts, parts extracted). PowerPoint embeds only a
        <p:regular> face for this family — bold runs use synthetic bold, same
        as decks PowerPoint embeds natively.

        Non-fatal: returns the input unchanged if assets are missing or the
        package already embeds fonts.
        """
        try:
            fonts_dir = settings.app_dir / "templates" / "fonts"
            # (typeface, panose, asset prefix). Regular = Light, bold = true
            # Medium: on machines without Gotham, bold runs render the actual
            # Medium face instead of synthetic bold over Light. The viewer JPG
            # conversion mirrors this by transiently registering the
            # bold-flagged Medium (see presentation_service transient font).
            families = [
                ("Gotham", "02000504020000020004", "Gotham"),
                ("Gotham Medium", "02000604040000020004", "GothamMedium"),
            ]
            family_faces = []
            for typeface, panose, prefix in families:
                faces = [
                    (face, fonts_dir / f"{prefix}_{face}.fntdata")
                    for face in ("regular", "bold", "italic", "boldItalic")
                    if (fonts_dir / f"{prefix}_{face}.fntdata").exists()
                ]
                if faces:
                    family_faces.append((typeface, panose, faces))
            if not family_faces:
                logger.warning(
                    "No Gotham fntdata assets in %s, skipping font embedding", fonts_dir
                )
                return pptx_bytes

            src = zipfile.ZipFile(BytesIO(pptx_bytes))
            pres_xml = src.read("ppt/presentation.xml").decode("utf-8")
            if "p:embeddedFontLst" in pres_xml:
                logger.info("Presentation already embeds fonts, skipping")
                return pptx_bytes

            rels_xml = src.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
            content_types = src.read("[Content_Types].xml").decode("utf-8")

            font_rel_type = (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
            )
            max_rid = max(int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml))
            new_rels = []
            embedded_fonts = []
            parts = {}
            part_idx = 0
            for typeface, panose, faces in family_faces:
                face_tags = []
                for face, asset_path in faces:
                    part_idx += 1
                    rid = f"rId{max_rid + part_idx}"
                    part_name = f"fonts/font{part_idx}.fntdata"
                    new_rels.append(
                        f'<Relationship Id="{rid}" Type="{font_rel_type}" Target="{part_name}"/>'
                    )
                    face_tags.append(f'<p:{face} r:id="{rid}"/>')
                    parts[f"ppt/{part_name}"] = asset_path.read_bytes()
                embedded_fonts.append(
                    "<p:embeddedFont>"
                    f'<p:font typeface="{typeface}" panose="{panose}" pitchFamily="2" charset="0"/>'
                    + "".join(face_tags)
                    + "</p:embeddedFont>"
                )

            rels_xml = rels_xml.replace(
                "</Relationships>", "".join(new_rels) + "</Relationships>"
            )

            if 'Extension="fntdata"' not in content_types:
                content_types = content_types.replace(
                    "</Types>",
                    '<Default Extension="fntdata" ContentType="application/x-fontdata"/></Types>',
                )

            # Root attribute + embeddedFontLst (schema position: after notesSz)
            if "embedTrueTypeFonts" not in pres_xml:
                pres_xml = pres_xml.replace(
                    "<p:presentation ", '<p:presentation embedTrueTypeFonts="1" ', 1
                )
            font_lst = (
                "<p:embeddedFontLst>" + "".join(embedded_fonts) + "</p:embeddedFontLst>"
            )
            anchor = re.search(r"<p:notesSz[^>]*/>", pres_xml) or re.search(
                r"<p:sldSz[^>]*/>", pres_xml
            )
            if anchor:
                pres_xml = (
                    pres_xml[: anchor.end()] + font_lst + pres_xml[anchor.end():]
                )
            else:
                pres_xml = pres_xml.replace(
                    "</p:presentation>", font_lst + "</p:presentation>", 1
                )

            replaced = {
                "ppt/presentation.xml": pres_xml.encode("utf-8"),
                "ppt/_rels/presentation.xml.rels": rels_xml.encode("utf-8"),
                "[Content_Types].xml": content_types.encode("utf-8"),
            }

            out_buffer = BytesIO()
            with zipfile.ZipFile(out_buffer, "w", zipfile.ZIP_DEFLATED) as dst:
                for item in src.infolist():
                    data = replaced.get(item.filename) or src.read(item.filename)
                    dst.writestr(item, data)
                for part_name, data in parts.items():
                    dst.writestr(part_name, data)

            logger.info(
                "Embedded fonts via OpenXML (PowerPoint-native fntdata): %s",
                [(tf, [f for f, _ in fc]) for tf, _, fc in family_faces],
            )
            return out_buffer.getvalue()

        except Exception as embed_err:
            logger.warning(
                "OpenXML font embedding failed (non-fatal): %s", str(embed_err)
            )
            return pptx_bytes

    def _promote_paragraph_props_to_runs(self, slide) -> None:
        """Copy each paragraph's defRPr onto runs that have no rPr.

        python-pptx's `paragraph.font` writes to pPr/defRPr. PowerPoint desktop
        honors it, but PowerPoint Online, Google Slides and LibreOffice largely
        ignore defRPr — every text box falls back to a different default font,
        which reads as "the font changes on every slide". Promoting the same
        properties to run level is semantically identical per ECMA-376 and
        honored by every viewer.
        """
        from pptx.oxml.ns import qn

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame._txBody.findall(qn("a:p")):
                ppr = p.find(qn("a:pPr"))
                if ppr is None:
                    continue
                defrpr = ppr.find(qn("a:defRPr"))
                if defrpr is None:
                    continue
                for r in p.findall(qn("a:r")):
                    if r.find(qn("a:rPr")) is None:
                        rpr = deepcopy(defrpr)
                        rpr.tag = qn("a:rPr")
                        r.insert(0, rpr)

    @staticmethod
    def _compose_group_name_line(entry: Dict[str, str]) -> str:
        """Name plus optional pronunciation/katakana parenthetical."""
        name_line = entry.get("name_text") or "Name Candidate"
        extra_parts = [
            part
            for part in [entry.get("pronunciation_text"), entry.get("katakana_text")]
            if part
        ]
        if extra_parts:
            name_line = f"{name_line} ({' / '.join(extra_parts)})"
        return name_line

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
        header_para.font.name = FONT_GOTHAM
        header_para.font.size = Pt(36)
        header_para.font.bold = True
        header_para.font.color.rgb = RGBColor(255, 255, 255)  # White text
        header_para.alignment = PP_ALIGN.LEFT

        logger.info("Rendered Name Summary slide")


# Singleton instance
openxml_pptx_service = OpenXMLPPTXService()
