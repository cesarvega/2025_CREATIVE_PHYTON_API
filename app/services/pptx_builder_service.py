import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Sequence, Tuple

import pythoncom
from pptx import Presentation
from pptx.util import Inches, Pt
from win32com import client

from app.config.settings import settings
from app.models.excel_models import ProcessedExcelData
from app.models.presentation_models import (
    CreatePresentationRequest,
    DetailItem,
    PresentationBuildOptions,
)
from app.utils.logging_utils import get_logger
from app.utils.path_utils import (
    build_relative_slide_path,
    resolve_project_output,
)

logger = get_logger(__name__)


PP_SAVE_AS_PPTX = 24
PP_SAVE_AS_PPTM = 25
FALLBACK_PRESENTATIONS_DIR = "generated_presentations"
PRESENTATIONS_SUBDIR = "Presentations"

_DELIMITER_PATTERN = re.compile(r"(?:##|\$\$)+")


def tokenize_delimited_block(value: Optional[str]) -> List[str]:
    """Split the legacy delimiter markers (##/$$) into individual entries."""

    if not value:
        return []

    tokens = [segment.strip() for segment in _DELIMITER_PATTERN.split(value) if segment and segment.strip()]
    return tokens


def normalize_delimited_block(value: Optional[str], *, joiner: str = "\n") -> str:
    """Return a human-friendly text block from legacy delimiter markers."""

    return joiner.join(tokenize_delimited_block(value))


def split_category_labels(value: Optional[str]) -> Tuple[str, str]:
    """Split category text into primary/secondary labels using the `$` separator."""

    if not value:
        return "", ""

    parts = [part.strip() for part in value.split("$") if part and part.strip()]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[1]


@dataclass
class PresentationBuildArtifacts:
    """Artifacts produced when composing final PowerPoint deliverables."""

    printable_path: Optional[Path]
    macro_path: Optional[Path]
    total_slides: int
    slide_start: int
    slide_end: Optional[int]
    download_urls: Dict[str, str] = field(default_factory=dict)
    summary: Dict[str, List[str]] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)

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

    def compose_presentation(
        self,
        *,
        request: CreatePresentationRequest,
        options: PresentationBuildOptions,
        details: List[Dict[str, Any]],
        excel_data: ProcessedExcelData,
    ) -> PresentationBuildArtifacts:
        """Compose final PPT deliverables using existing templates and Excel data."""

        template_paths = self._resolve_template_paths(options)
        _ = excel_data  # Reserved for future enhancements (e.g., aggregated metrics)
        output_base, used_fallback = resolve_project_output(
            request.display_name,
            request.project_type,
            fallback_subdir=FALLBACK_PRESENTATIONS_DIR,
        )
        presentation_root = output_base / PRESENTATIONS_SUBDIR
        presentation_root.mkdir(parents=True, exist_ok=True)

        base_name = options.output_basename(request.display_name)
        if options.slide_start > 1 or options.slide_end:
            range_suffix = f"_s{options.slide_start}"
            if options.slide_end:
                range_suffix += f"-{options.slide_end}"
            base_name = f"{base_name}{range_suffix}"
        timestamp = datetime.utcnow().strftime("%Y%m%d%H%M%S")
        base_name = f"{base_name}_{timestamp}"

        printable_path = presentation_root / f"{base_name}.pptx"
        macro_path = presentation_root / f"{base_name}.pptm"

        filtered_details = self._filter_details(details, options.slide_start, options.slide_end)
        summary_tracker: Dict[str, List[str]] = {}
        warnings: List[str] = []

        pythoncom.CoInitialize()
        app = None
        output_presentation = None
        template_handles: Dict[str, Any] = {}
        created_slides = 0

        try:
            try:
                app = client.Dispatch("PowerPoint.Application")
            except Exception as dispatch_error:  # pylint: disable=broad-except
                logger.error(
                    "Unable to start PowerPoint automation: %s",
                    dispatch_error,
                )
                raise RuntimeError(
                    "PowerPoint automation is unavailable. Ensure Microsoft PowerPoint is installed and accessible."
                ) from dispatch_error
            try:
                app.Visible = 1
            except Exception as visibility_error:  # pylint: disable=broad-except
                logger.debug(
                    "PowerPoint visibility flag could not be set: %s",
                    visibility_error,
                )

            template_handles = self._open_template_presentations(app, template_paths)
            base_template = template_handles["base"]

            output_presentation = app.Presentations.Add()
            output_presentation.PageSetup.SlideWidth = base_template.PageSetup.SlideWidth
            output_presentation.PageSetup.SlideHeight = base_template.PageSetup.SlideHeight
            output_presentation.ApplyTemplate(str(template_paths["base"].resolve()))

            for detail in filtered_details:
                template_key = self._select_template_kind(detail)
                if template_key == "skip":
                    continue

                source_presentation = template_handles.get(template_key)
                if source_presentation is None:
                    source_presentation = app.Presentations.Open(
                        str(template_paths[template_key].resolve()),
                        WithWindow=False,
                    )
                    template_handles[template_key] = source_presentation

                new_slide = self._copy_slide_from_template(
                    source_presentation,
                    output_presentation,
                    template_paths[template_key],
                )
                created_slides += 1

                if template_key == "separator":
                    self._populate_separator_slide(new_slide, detail, warnings)
                elif template_key == "summary":
                    self._populate_summary_slide(new_slide, summary_tracker, warnings)
                elif template_key == "group":
                    self._populate_group_slide(new_slide, detail, warnings)
                elif template_key == "multi":
                    self._populate_multi_slide(new_slide, detail, summary_tracker, warnings)
                else:
                    self._populate_base_slide(new_slide, detail, summary_tracker, warnings)

            printable_path_out: Optional[Path] = None
            macro_path_out: Optional[Path] = None

            if created_slides == 0:
                warnings.append("No slides generated from the provided Excel data.")
            else:
                if options.include_print_ready_version:
                    output_presentation.SaveAs(str(printable_path), PP_SAVE_AS_PPTX)
                    printable_path_out = printable_path

                if options.include_macro_version:
                    if options.include_print_ready_version:
                        output_presentation.SaveCopyAs(str(macro_path), PP_SAVE_AS_PPTM)
                    else:
                        output_presentation.SaveAs(str(macro_path), PP_SAVE_AS_PPTM)
                    macro_path_out = macro_path

                if not options.include_print_ready_version and not options.include_macro_version:
                    output_presentation.SaveAs(str(printable_path), PP_SAVE_AS_PPTX)
                    printable_path_out = printable_path

            download_urls: Dict[str, str] = {}
            if options.return_urls:
                if printable_path_out:
                    download_urls["printable"] = self._relative_download_path(printable_path_out)
                if macro_path_out:
                    download_urls["macro"] = self._relative_download_path(macro_path_out)

            artifacts = PresentationBuildArtifacts(
                printable_path=printable_path_out,
                macro_path=macro_path_out,
                total_slides=created_slides,
                slide_start=options.slide_start,
                slide_end=options.slide_end,
                download_urls=download_urls,
                summary={key: list(values) for key, values in summary_tracker.items()},
                warnings=warnings,
            )

        except Exception as exc:  # pylint: disable=broad-except
            logger.error("Error composing presentation: %s", exc)
            raise
        finally:
            if output_presentation is not None:
                try:
                    output_presentation.Close()
                except Exception:  # pragma: no cover - best effort cleanup
                    pass

            for handle in template_handles.values():
                try:
                    handle.Close()
                except Exception:  # pragma: no cover - best effort cleanup
                    pass

            if app is not None:
                try:
                    app.Quit()
                except Exception:  # pragma: no cover
                    pass

            pythoncom.CoUninitialize()

        return artifacts

    def _relative_download_path(self, artifact_path: Path) -> str:
        """Return a relative URL-like path when possible."""

        candidates = [settings.nw_files_dir, settings.base_dir]
        for base in candidates:
            try:
                return artifact_path.relative_to(base).as_posix()
            except ValueError:
                continue
        return artifact_path.as_posix()

    def _resolve_template_paths(self, options: PresentationBuildOptions) -> Dict[str, Path]:
        root = Path(options.templates_root) if options.templates_root else (self.templates_dir / options.template_pack)
        root.mkdir(parents=True, exist_ok=True)

        return {
            "base": self._resolve_template_path(root, options.base_template),
            "multi": self._resolve_template_path(root, options.multi_template),
            "group": self._resolve_template_path(root, options.group_template),
            "separator": self._resolve_template_path(root, options.separator_template),
            "summary": self._resolve_template_path(root, options.summary_template),
        }

    @staticmethod
    def _resolve_template_path(root: Path, template_name: str) -> Path:
        candidate = Path(template_name)
        if not candidate.is_absolute():
            candidate = root / template_name
        if not candidate.exists():
            raise FileNotFoundError(f"Template not found: {candidate}")
        return candidate

    @staticmethod
    def _open_template_presentations(app, template_paths: Dict[str, Path]) -> Dict[str, Any]:
        handles: Dict[str, Any] = {}
        for key, path in template_paths.items():
            handles[key] = app.Presentations.Open(str(path.resolve()), WithWindow=False)
        return handles

    @staticmethod
    def _copy_slide_from_template(template_presentation, target_presentation, template_path: Path):
        try:
            template_presentation.Slides(1).Copy()
            slide_range = target_presentation.Slides.Paste()
            return slide_range.Item(1)
        except Exception as exc:  # pylint: disable=broad-except
            logger.warning(
                "Clipboard-based slide copy failed (%s); falling back to InsertFromFile for template %s",
                exc,
                template_path,
            )
            return PPTXBuilderService._insert_slide_from_file(
                target_presentation,
                template_path,
                original_error=exc,
            )

    @staticmethod
    def _insert_slide_from_file(target_presentation, template_path: Path, original_error: Exception):
        try:
            slides = target_presentation.Slides
            dest_count = getattr(slides, "Count", 0)
            insert_index = dest_count if dest_count > 0 else 0
            slide_range = slides.InsertFromFile(
                str(template_path.resolve()),
                insert_index,
                1,
                1,
            )
            return slide_range.Item(1)
        except Exception as fallback_error:  # pylint: disable=broad-except
            logger.error(
                "InsertFromFile fallback failed after copy error %s: %s",
                original_error,
                fallback_error,
            )
            raise

    @staticmethod
    def _filter_details(details: List[Dict[str, Any]], start: int, end: Optional[int]) -> List[Dict[str, Any]]:
        if start <= 1 and end is None:
            return list(details)

        filtered: List[Dict[str, Any]] = []
        for item in details:
            slide_number = item.get("slide_number")
            if slide_number is None:
                filtered.append(item)
                continue

            if slide_number < start:
                continue
            if end is not None and slide_number > end:
                continue
            filtered.append(item)

        return filtered

    @staticmethod
    def _detail_context(detail: Dict[str, Any], default: str) -> str:
        tokens: List[str] = []
        slide_number = detail.get("slide_number")
        if slide_number is not None:
            tokens.append(f"slide #{slide_number}")

        name = detail.get("name") or detail.get("group_name")
        if name:
            tokens.append(f'"{name}"')

        category = detail.get("category")
        if category:
            tokens.append(f'category "{category}"')

        if not tokens:
            return default

        joined = ", ".join(tokens)
        return f"{default} ({joined})"

    @staticmethod
    def _iter_text_shapes(slide) -> Iterator[Any]:
        count = getattr(slide.Shapes, "Count", 0)
        for idx in range(1, count + 1):
            shape = slide.Shapes(idx)
            try:
                has_text_frame = bool(getattr(shape, "HasTextFrame", 0))
            except Exception:
                has_text_frame = False
            if not has_text_frame:
                continue
            try:
                _ = shape.TextFrame.TextRange
            except Exception:
                continue
            yield shape

    def _collect_placeholder_shapes(self, slide, placeholder: str) -> List[Any]:
        placeholder_lower = placeholder.lower()
        matches: List[Any] = []
        for shape in self._iter_text_shapes(slide):
            try:
                text = shape.TextFrame.TextRange.Text
            except Exception:
                continue
            if placeholder_lower in text.lower():
                matches.append(shape)
        return matches

    def _replace_placeholder_text(
        self,
        slide,
        placeholder: str,
        replacement: str,
        *,
        replace_all: bool = False,
        warnings: Optional[List[str]] = None,
        required: bool = False,
        context: Optional[str] = None,
    ) -> int:
        placeholder_lower = placeholder.lower()
        replaced = 0
        for shape in self._iter_text_shapes(slide):
            try:
                text = shape.TextFrame.TextRange.Text
            except Exception:
                continue
            if placeholder_lower in text.lower():
                self._set_shape_text(shape, replacement)
                replaced += 1
                if not replace_all:
                    break
        if replaced == 0 and required and warnings is not None:
            slide_index = getattr(slide, "SlideIndex", None)
            location = context or (f"slide #{slide_index}" if slide_index else "slide")
            warnings.append(
                f"Placeholder '{placeholder}' not found in {location}."
            )
        return replaced

    @staticmethod
    def _set_shape_text(shape, text: str) -> None:
        try:
            shape.TextFrame.TextRange.Text = text or ""
        except Exception:
            logger.debug("Could not assign text to shape: %s", getattr(shape, "Name", "<unnamed>"))

    def _populate_separator_slide(self, slide, detail: Dict[str, Any], warnings: List[str]) -> None:
        category = detail.get("category") or detail.get("slide_description") or ""
        primary, secondary = split_category_labels(category)
        if not primary and secondary:
            primary, secondary = secondary, ""
        context = self._detail_context(detail, "separator slide")
        self._replace_placeholder_text(
            slide,
            "Category1",
            primary,
            warnings=warnings,
            required=bool(primary),
            context=context,
        )
        self._replace_placeholder_text(
            slide,
            "Category2",
            secondary,
            replace_all=True,
            warnings=warnings,
            required=bool(secondary),
            context=context,
        )

    def _populate_group_slide(self, slide, detail: Dict[str, Any], warnings: List[str]) -> None:
        group_name = detail.get("group_name") or detail.get("name") or "Group"
        category = detail.get("category") or ""
        context = self._detail_context(detail, "group slide")
        self._replace_placeholder_text(
            slide,
            "[New Names]",
            group_name,
            warnings=warnings,
            required=bool(group_name),
            context=context,
        )
        self._replace_placeholder_text(
            slide,
            "Category",
            category,
            warnings=warnings,
            required=bool(category),
            context=context,
        )

    def _populate_multi_slide(
        self,
        slide,
        detail: Dict[str, Any],
        summary_tracker: Dict[str, List[str]],
        warnings: List[str],
    ) -> None:
        category = detail.get("category") or detail.get("group_name") or "General"
        names = tokenize_delimited_block(detail.get("name")) or [detail.get("name", "")]
        summary_tracker.setdefault(category, []).extend([name for name in names if name])

        context = self._detail_context(detail, "multi slide")
        name_shapes = self._collect_placeholder_shapes(slide, "Name Candidate")
        if not name_shapes and any(name.strip() for name in names):
            warnings.append(
                f"No 'Name Candidate' placeholders found in {context}; candidate names will not render."
            )
        for idx, value in enumerate(names):
            if idx < len(name_shapes):
                self._set_shape_text(name_shapes[idx], value)
            else:
                warnings.append(
                    f"Slide '{category}' has more names than placeholders; extra name '{value}' was omitted."
                )
                break

        for idx in range(len(names), len(name_shapes)):
            self._set_shape_text(name_shapes[idx], "")

        self._replace_placeholder_text(
            slide,
            "Category",
            category,
            warnings=warnings,
            required=bool(category),
            context=context,
        )
        rationale = normalize_delimited_block(detail.get("rationale"))
        self._replace_placeholder_text(
            slide,
            "[Comments]",
            rationale,
            warnings=warnings,
            required=False,
            context=context,
        )

    def _populate_base_slide(
        self,
        slide,
        detail: Dict[str, Any],
        summary_tracker: Dict[str, List[str]],
        warnings: List[str],
    ) -> None:
        category = detail.get("category") or detail.get("group_name") or "General"
        names = tokenize_delimited_block(detail.get("name")) or [detail.get("name", "")]
        summary_tracker.setdefault(category, []).extend([name for name in names if name])

        primary_name = names[0] if names else ""
        context = self._detail_context(detail, "base slide")
        self._replace_placeholder_text(
            slide,
            "Name Candidate",
            primary_name,
            warnings=warnings,
            required=bool(primary_name),
            context=context,
        )
        self._replace_placeholder_text(
            slide,
            "Category",
            category,
            warnings=warnings,
            required=bool(category),
            context=context,
        )

        rationale_text = normalize_delimited_block(detail.get("rationale"))
        self._replace_placeholder_text(
            slide,
            "Rationale Goes Here",
            rationale_text,
            warnings=warnings,
            required=bool(rationale_text),
            context=context,
        )

        notation_or_comments = detail.get("notation") or detail.get("name_sub_group") or ""
        additional_comments = normalize_delimited_block(notation_or_comments)
        self._replace_placeholder_text(
            slide,
            "[Comments]",
            additional_comments,
            warnings=warnings,
            required=False,
            context=context,
        )

        kana = detail.get("kana") or ""
        self._replace_placeholder_text(
            slide,
            "Pronunciation/Katakana",
            kana,
            warnings=warnings,
            required=False,
            context=context,
        )

        for placeholder in ("Neutral", "Negative", "Positive"):
            self._replace_placeholder_text(
                slide,
                placeholder,
                "",
                warnings=warnings,
                required=False,
                context=context,
            )

    def _populate_summary_slide(
        self,
        slide,
        summary_tracker: Dict[str, List[str]],
        warnings: Optional[List[str]] = None,
    ) -> None:
        lines: List[str] = []
        for category, names in summary_tracker.items():
            if not names:
                continue
            unique_names: List[str] = list(dict.fromkeys(name for name in names if name))
            if not unique_names:
                continue
            lines.append(f"{category}: {', '.join(unique_names)}")

        body = "Name Candidates - Summary"
        if lines:
            body = f"{body}\n\n" + "\n".join(lines)
        self._replace_placeholder_text(
            slide,
            "Name Candidates - Summary",
            body,
            replace_all=True,
            warnings=warnings,
            required=True,
            context="summary slide",
        )

    def _select_template_kind(self, detail: Dict[str, Any]) -> str:
        slide_type = (detail.get("slide_type") or "").lower()
        name = detail.get("name") or ""
        group_name = detail.get("group_name") or ""

        if slide_type == "image":
            if detail.get("category") or detail.get("slide_description"):
                return "separator"
            return "skip"

        if slide_type == "namesummary":
            return "summary"

        if slide_type in {"bsr_group", "nsr_group"}:
            return "group"

        if group_name and group_name.strip() and group_name.strip() == name.strip() and not detail.get("rationale"):
            return "group"

        if detail.get("name_sub_group"):
            return "multi"

        tokenized_name = tokenize_delimited_block(name)
        tokenized_rationale = tokenize_delimited_block(detail.get("rationale"))
        if len(tokenized_name) > 1 or len(tokenized_rationale) > 1:
            return "multi"

        return "base"

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