import re
from pathlib import Path
from typing import Optional

import pythoncom
from pptx import Presentation
from pptx.util import Inches, Pt
from win32com import client

from app.models.presentation_models import DetailItem
from app.utils.logging_utils import get_logger
from app.utils.path_utils import (
    build_relative_slide_path,
    resolve_project_output,
)

logger = get_logger(__name__)

class PPTXBuilderService:
    """Service for generating a PowerPoint file from PresentationData."""

    def __init__(
        self,
        templates_dir: Path | str = Path("templates"),
        base_template_path: Optional[Path | str] = None,
    ):
        self.templates_dir = Path(templates_dir)
        self.base_template_path = Path(base_template_path) if base_template_path else self.templates_dir / "Default.pptx"

        if not self.base_template_path.exists():
            logger.warning("Base PowerPoint template not found at %s", self.base_template_path)

    
    def generate_category_slide_image(
        self,
        *,
        category: str,
        display_name: str,
        slide_number: int,
        project_type: Optional[str] = None,
    ) -> Optional[str]:
        """Generate a category header slide image and return its relative path."""

        if not category:
            logger.warning("Cannot generate category slide without category name")
            return None

        fallback_subdir = "temp_category_slides"
        output_base, used_fallback = resolve_project_output(
            display_name,
            project_type,
            fallback_subdir=fallback_subdir,
        )
        if used_fallback:
            logger.warning(
                "Using fallback directory for category slides: %s",
                output_base.parent,
            )
        output_base.mkdir(parents=True, exist_ok=True)

        safe_category = self._sanitize_filename(category)
        ppt_filename = f"{slide_number:03d}_{safe_category}.pptx"
        image_filename = f"{slide_number:03d}_{safe_category}.jpg"

        ppt_path = output_base / ppt_filename
        image_path = output_base / image_filename

        try:
            prs, slide = self._prepare_category_slide_template()
            self._populate_category_slide(slide, category)
            prs.save(ppt_path)
            self._export_slide_to_image(ppt_path, image_path)
            if ppt_path.exists():
                ppt_path.unlink()

            relative_path = build_relative_slide_path(
                project_type,
                display_name,
                image_filename,
                fallback_root=fallback_subdir if used_fallback else None,
            )
            logger.info("Category slide generated at %s", relative_path)
            return relative_path
        except Exception as exc:  # pylint: disable=broad-except
            logger.error("Error generating category slide for '%s': %s", category, exc)
            return None

    def _prepare_category_slide_template(self):
        """Load a category template if available; otherwise create a blank slide."""
        template_path = self.base_template_path

        if not template_path.exists():
            fallback = Path("templates/BackgroundDefaultTemplate/template_default_seperator_2019.pptx")
            if fallback.exists():
                logger.warning(
                    "Base template missing at %s; using fallback template %s",
                    template_path,
                    fallback,
                )
                template_path = fallback
            else:
                logger.warning(
                    "Category template not found at %s, creating blank presentation",
                    template_path,
                )
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                return prs, slide

        logger.info("Loading category template from: %s", template_path)
        prs = Presentation(str(template_path))
        slide = prs.slides[0] if prs.slides else prs.slides.add_slide(prs.slide_layouts[5])
        return prs, slide

    def _populate_category_slide(self, slide, category: str) -> None:
        """Populate the category slide with header text by replacing Category1 and Category2."""

        # Split the category using the $ symbol
        parts = category.split('$')
        category1_text = parts[0].strip() if len(parts) > 0 else ""
        category2_text = parts[1].strip() if len(parts) > 1 else category.strip()
        
        logger.info("Splitting category: '%s' -> Category1: '%s', Category2: '%s'", 
                    category, category1_text, category2_text)
        
        # Search for and replace the placeholder text in the template
        replaced_count = 0
        for shape in slide.shapes:
            if shape.has_text_frame:
                # Get the current text
                current_text = shape.text_frame.text
                
                # Helper function to replace text while preserving formatting
                def replace_text(target_text, replacement_text):
                    nonlocal replaced_count
                    if target_text in current_text:
                        # Store formatting properties before clearing the text frame
                        font_size = None
                        font_bold = None
                        font_color = None
                        
                        # Try to capture formatting from the first existing run
                        if shape.text_frame.paragraphs and shape.text_frame.paragraphs[0].runs:
                            original_run = shape.text_frame.paragraphs[0].runs[0]
                            try:
                                if original_run.font.size:
                                    font_size = original_run.font.size
                            except Exception:
                                pass
                            try:
                                if original_run.font.bold is not None:
                                    font_bold = original_run.font.bold
                            except Exception:
                                pass
                            try:
                                # Check whether color.rgb exists and has a value
                                if hasattr(original_run.font.color, 'rgb') and original_run.font.color.rgb:
                                    font_color = original_run.font.color.rgb
                            except Exception:
                                pass
                        
                        # Clear the text frame and add the new text
                        shape.text_frame.clear()
                        p = shape.text_frame.paragraphs[0]
                        run = p.add_run()
                        run.text = replacement_text
                        
                        # Apply the stored formatting
                        try:
                            if font_size:
                                run.font.size = font_size
                            if font_bold is not None:
                                run.font.bold = font_bold
                            if font_color:
                                run.font.color.rgb = font_color
                        except Exception as e:
                            logger.warning("Could not apply formatting: %s", e)
                        
                        replaced_count += 1
                        logger.info("Replaced %s with: '%s'", target_text, replacement_text)
                        return True
                    return False
                
                # Replace Category1 or Category2
                if not replace_text("Category1", category1_text):
                    replace_text("Category2", category2_text)
        
        if replaced_count == 0:
            logger.warning("No Category1 or Category2 placeholders found in template")
        
    def _export_slide_to_image(self, ppt_path: Path, image_path: Path) -> None:
        """Export the first slide of a PPTX file to an image using PowerPoint automation."""
        pythoncom.CoInitialize()
        powerpoint = None
        presentation = None

        try:
            powerpoint = client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1
            presentation = powerpoint.Presentations.Open(str(ppt_path.resolve()), WithWindow=False)
            if presentation.Slides.Count == 0:
                raise ValueError("Presentation contains no slides to export")

            slide = presentation.Slides(1)
            # Export as JPG with high quality
            slide.Export(str(image_path.resolve()), "JPG")
            logger.info("Slide exported to JPG: %s", image_path)
            
        finally:
            if presentation:
                presentation.Close()
            if powerpoint:
                powerpoint.Quit()
            pythoncom.CoUninitialize()

    @staticmethod
    def _sanitize_filename(value: str) -> str:
        """Sanitize a string to be used as a filename."""
        # Remove the $ symbol and other special characters
        cleaned = re.sub(r"[^a-zA-Z0-9_-]+", "_", value.strip())
        # Limit length
        cleaned = cleaned[:100]
        return cleaned or "category"

    def _populate_slide(self, slide, detail: DetailItem):
        """Add dynamic text to slide placeholders."""
        # Name
        if detail.name:
            self._add_textbox(slide, detail.name, Inches(1), Inches(1), Inches(6), Inches(1), Pt(28), bold=True)

        # Category
        if detail.category:
            self._add_textbox(slide, f"Category: {detail.category}", Inches(1), Inches(2), Inches(6), Inches(0.8), Pt(18))

        # Rationale
        if detail.rationale:
            self._add_textbox(slide, detail.rationale, Inches(1), Inches(3), Inches(7), Inches(2), Pt(16))

        # Notation
        if detail.notation:
            self._add_textbox(slide, f"Notation: {detail.notation}", Inches(1), Inches(5.5), Inches(6), Inches(0.6), Pt(14))

        # Kana (phonetics)
        if detail.kana:
            self._add_textbox(slide, detail.kana, Inches(1), Inches(6.2), Inches(6), Inches(0.6), Pt(14))

    def _add_textbox(self, slide, text, left, top, width, height, font_size=Pt(18), bold=False):
        textbox = slide.shapes.add_textbox(left, top, width, height)
        tf = textbox.text_frame
        run = tf.paragraphs[0].add_run()
        run.text = text
        font = run.font
        font.size = font_size
        font.bold = bold


pptx_builder_service = PPTXBuilderService(
    base_template_path=Path("templates/BackgroundDefaultTemplate/template_default_seperator_2019.pptx")
)