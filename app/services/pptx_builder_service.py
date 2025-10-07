import re
import math
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

    def compose_presentation_with_original_slides(
        self,
        *,
        request: CreatePresentationRequest,
        options: PresentationBuildOptions,
        details: List[Dict[str, Any]],
        excel_data: ProcessedExcelData,
        original_pptx_path: str,
        page_number_insert: int,
    ) -> PresentationBuildArtifacts:
        """
        Compose complete PowerPoint by combining original PPTX slides with template-generated slides.
        
        Process:
        1. Insert original slides from user's PPTX (before page_number_insert)
        2. Insert slides generated from Excel data
        3. Insert remaining original slides (after page_number_insert)
        
        Args:
            request: Presentation creation request
            options: Build options for presentation
            details: Slide details from Excel processing
            excel_data: Processed Excel data
            original_pptx_path: Path to original PPTX file sent by user
            page_number_insert: Position where Excel-generated slides should be inserted
            
        Returns:
            PresentationBuildArtifacts with paths and metadata
        """
        _ = excel_data  # Reserved for future enhancements
        template_paths = self._resolve_template_paths(options)
        output_base, _ = resolve_project_output(
            request.display_name,
            request.project_type,
            fallback_subdir=FALLBACK_PRESENTATIONS_DIR,
        )
        presentation_root = output_base / PRESENTATIONS_SUBDIR
        presentation_root.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.utcnow().strftime("%Y%m%d%H%M%S")
        base_name = f"{request.display_name}_{timestamp}"
        printable_path = presentation_root / f"{base_name}.pptx"
        macro_path = presentation_root / f"{base_name}.pptm"

        summary_tracker: Dict[str, List[str]] = {}
        warnings: List[str] = []
        total_slides_created = 0

        pythoncom.CoInitialize()
        app = None
        output_presentation = None
        original_pptx = None
        template_handles: Dict[str, Any] = {}

        try:
            # Initialize PowerPoint
            try:
                app = client.Dispatch("PowerPoint.Application")
                app.Visible = 1
            except Exception as dispatch_error:
                logger.error("Unable to start PowerPoint automation: %s", dispatch_error)
                raise RuntimeError(
                    "PowerPoint automation is unavailable. Ensure Microsoft PowerPoint is installed."
                ) from dispatch_error

            # Open template presentations
            template_handles = self._open_template_presentations(app, template_paths)
            base_template = template_handles["base"]

            # Create new presentation with base template configuration
            output_presentation = app.Presentations.Add()
            output_presentation.PageSetup.SlideWidth = base_template.PageSetup.SlideWidth
            output_presentation.PageSetup.SlideHeight = base_template.PageSetup.SlideHeight
            output_presentation.ApplyTemplate(str(template_paths["base"].resolve()))

            # Open original PPTX sent by user
            original_pptx = app.Presentations.Open(str(original_pptx_path), WithWindow=False)
            total_original_slides = original_pptx.Slides.Count

            logger.info(
                "Composing presentation: %d original slides, insert at position %d, %d generated slides",
                total_original_slides,
                page_number_insert,
                len(details),
            )

            # SECTION A: Insert original slides before page_number_insert
            if page_number_insert > 1:
                slides_to_copy = min(page_number_insert - 1, total_original_slides)
                for i in range(1, slides_to_copy + 1):
                    try:
                        slide = original_pptx.Slides(i)
                        slide.Copy()
                        output_presentation.Slides.Paste(output_presentation.Slides.Count + 1)
                        total_slides_created += 1
                        logger.debug("Copied original slide %d", i)
                    except Exception as copy_error:
                        logger.warning("Failed to copy original slide %d: %s", i, copy_error)
                        warnings.append(f"Failed to copy original slide {i}")

            logger.info("Inserted %d prefix slides from original PPTX", total_slides_created)

            # SECTION B: Insert template-generated slides from Excel data
            generated_count = 0
            for detail in details:
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

                try:
                    new_slide = self._copy_slide_from_template(
                        source_presentation,
                        output_presentation,
                        template_paths[template_key],
                    )
                    total_slides_created += 1
                    generated_count += 1

                    # Populate slide with data based on template type
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
                        
                except Exception as slide_error:
                    logger.error("Failed to create template slide: %s", slide_error)
                    warnings.append(f"Failed to create template slide: {str(slide_error)}")

            logger.info("Inserted %d template-generated slides", generated_count)

            # SECTION C: Insert remaining original slides after generated slides
            if page_number_insert <= total_original_slides:
                suffix_count = 0
                for i in range(page_number_insert, total_original_slides + 1):
                    try:
                        slide = original_pptx.Slides(i)
                        slide.Copy()
                        output_presentation.Slides.Paste(output_presentation.Slides.Count + 1)
                        total_slides_created += 1
                        suffix_count += 1
                        logger.debug("Copied original slide %d", i)
                    except Exception as copy_error:
                        logger.warning("Failed to copy original slide %d: %s", i, copy_error)
                        warnings.append(f"Failed to copy original slide {i}")
                
                logger.info("Inserted %d suffix slides from original PPTX", suffix_count)

            # Save output files
            printable_path_out: Optional[Path] = None
            macro_path_out: Optional[Path] = None

            if total_slides_created == 0:
                warnings.append("No slides were generated in the presentation.")
            else:
                if options.include_print_ready_version:
                    output_presentation.SaveAs(str(printable_path.resolve()), PP_SAVE_AS_PPTX)
                    printable_path_out = printable_path
                    logger.info("Saved printable presentation: %s", printable_path)

                if options.include_macro_version:
                    if options.include_print_ready_version:
                        output_presentation.SaveCopyAs(str(macro_path.resolve()), PP_SAVE_AS_PPTM)
                    else:
                        output_presentation.SaveAs(str(macro_path.resolve()), PP_SAVE_AS_PPTM)
                    macro_path_out = macro_path
                    logger.info("Saved macro presentation: %s", macro_path)

                if not options.include_print_ready_version and not options.include_macro_version:
                    output_presentation.SaveAs(str(printable_path.resolve()), PP_SAVE_AS_PPTX)
                    printable_path_out = printable_path

            # Generate download URLs if requested
            download_urls: Dict[str, str] = {}
            if options.return_urls:
                if printable_path_out:
                    download_urls["printable"] = self._relative_download_path(printable_path_out)
                if macro_path_out:
                    download_urls["macro"] = self._relative_download_path(macro_path_out)

            artifacts = PresentationBuildArtifacts(
                printable_path=printable_path_out,
                macro_path=macro_path_out,
                total_slides=total_slides_created,
                slide_start=1,
                slide_end=total_slides_created,
                download_urls=download_urls,
                summary={key: list(values) for key, values in summary_tracker.items()},
                warnings=warnings,
            )

            logger.info(
                "Presentation composition complete: %d total slides, %d warnings",
                total_slides_created,
                len(warnings),
            )

            return artifacts

        except Exception as exc:
            logger.error("Error composing presentation with original slides: %s", exc)
            raise
            
        finally:
            # Cleanup COM objects
            if original_pptx is not None:
                try:
                    original_pptx.Close()
                except Exception:
                    pass

            if output_presentation is not None:
                try:
                    output_presentation.Close()
                except Exception:
                    pass

            for handle in template_handles.values():
                try:
                    handle.Close()
                except Exception:
                    pass

            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass

            pythoncom.CoUninitialize()

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


    def _find_checkbox_for_text(self, slide, text_shape) -> Any:
        """Finds the checkbox (shape) that is near the given text shape."""
        try:
            text_left = getattr(text_shape, "Left", 0)
            text_top = getattr(text_shape, "Top", 0)

            # Find shapes that are to the left of the text and at a similar height
            horizontal_tolerance = 800000  # ~0.83 inches
            vertical_tolerance = 200000    # ~0.21 inches

            count = getattr(slide.Shapes, "Count", 0)
            closest_checkbox = None
            min_distance = float('inf')

            for idx in range(1, count + 1):
                shape = slide.Shapes(idx)
                shape_left = getattr(shape, "Left", 0)
                shape_top = getattr(shape, "Top", 0)

                # The checkbox must be to the left of the text
                if shape_left < text_left and abs(shape_left - text_left) <= horizontal_tolerance:
                    # And at a similar height
                    if abs(shape_top - text_top) <= vertical_tolerance:
                        # Ensure it is not a text frame (checkboxes do not have text)
                        try:
                            has_text_frame = bool(getattr(shape, "HasTextFrame", 0))
                            if has_text_frame:
                                continue
                        except Exception:
                            pass

                        # Compute distance
                        distance = abs(shape_left - text_left) + abs(shape_top - text_top)
                        if distance < min_distance:
                            min_distance = distance
                            closest_checkbox = shape

            return closest_checkbox
        except Exception as exc:
            logger.debug("Error finding checkbox for text shape: %s", exc)
            return None

    def _group_name_placeholder_columns(self, shapes: Sequence[Any]) -> List[Dict[str, Any]]:
        """Group placeholder text shapes into columns based on their horizontal position."""
        tolerance = 120000  # ~0.13 inches
        columns: List[Dict[str, Any]] = []
        for shape in sorted(shapes, key=lambda s: (getattr(s, "Left", 0), getattr(s, "Top", 0))):
            left = getattr(shape, "Left", None)
            top = getattr(shape, "Top", None)
            if left is None or top is None:
                continue
            for column in columns:
                if abs(column["left"] - left) <= tolerance:
                    column["shapes"].append(shape)
                    break
            else:
                columns.append({"left": left, "shapes": [shape]})
        for column in columns:
            column["shapes"].sort(key=lambda s: getattr(s, "Top", 0))
        columns.sort(key=lambda c: c["left"])
        return columns

    def _create_checkbox_shape(self, slide, left: int, top: int, size: int = 152400) -> Any:
        """Creates a new checkbox (square) at the specified position."""
        try:
            # msoShapeRectangle = 1
            checkbox = slide.Shapes.AddShape(1, left, top, size, size)

            # Configure the checkbox as a border-only square with no fill
            checkbox.Fill.Visible = 0  # No fill
            checkbox.Line.Visible = 1  # With border
            checkbox.Line.ForeColor.RGB = 0  # Black
            checkbox.Line.Weight = 1.5  # Border thickness

            return checkbox
        except Exception as e:
            logger.warning("Error creating checkbox: %s", e)
            return None

    def _find_group_for_text_shape(self, slide, text_shape) -> Any:
        """Finds the group that contains a text shape."""
        try:
            text_shape_id = getattr(text_shape, "Id", None)
            if not text_shape_id:
                return None

            count = getattr(slide.Shapes, "Count", 0)
            for idx in range(1, count + 1):
                try:
                    shape = slide.Shapes(idx)
                    # Type 6 = msoGroup
                    if getattr(shape, "Type", None) == 6:
                        group_items = getattr(shape, "GroupItems", None)
                        if group_items:
                            for i in range(1, group_items.Count + 1):
                                item = group_items.Item(i)
                                if getattr(item, "Id", None) == text_shape_id:
                                    return shape
                except:
                    continue

            return None
        except Exception as e:
            logger.debug("Error finding group: %s", e)
            return None

    def _calculate_layout_requirements(self, names_count: int, max_per_column: int = 18) -> Tuple[int, int]:
        """Calculates how many columns and rows are needed."""
        columns_needed = math.ceil(names_count / max_per_column)
        names_per_column = math.ceil(names_count / columns_needed)
        return columns_needed, names_per_column

    def _duplicate_group_for_new_column(self, slide, ref_group, new_left: int) -> Any:
        """Duplicates a complete group (checkbox + text) for a new column."""
        try:
            group_range = ref_group.Duplicate()
            new_group = group_range.Item(1)
            new_group.Left = int(new_left)
            new_group.Top = ref_group.Top

            # Clear text in the duplicated group
            group_items = getattr(new_group, "GroupItems", None)
            if group_items:
                for i in range(1, group_items.Count + 1):
                    item = group_items.Item(i)
                    try:
                        has_text = bool(getattr(item, "HasTextFrame", 0))
                        if has_text:
                            self._set_shape_text(item, "")
                    except Exception:
                        pass

            new_group.Visible = False
            return new_group
        except Exception as e:
            logger.error("Error duplicating group: %s", e)
            return None

    def _duplicate_group_for_new_row(self, slide, ref_group, vertical_spacing: int) -> Any:
        """Duplicates a group to create a new row in the same column."""
        try:
            group_range = ref_group.Duplicate()
            new_group = group_range.Item(1)
            new_group.Top = int(ref_group.Top + vertical_spacing)
            new_group.Left = ref_group.Left

            # Clear text
            group_items = getattr(new_group, "GroupItems", None)
            if group_items:
                for i in range(1, group_items.Count + 1):
                    item = group_items.Item(i)
                    try:
                        has_text = bool(getattr(item, "HasTextFrame", 0))
                        if has_text:
                            self._set_shape_text(item, "")
                    except Exception:
                        pass

            new_group.Visible = False
            return new_group
        except Exception as e:
            logger.error("Error duplicating row group: %s", e)
            return None

    def _calculate_column_spacing(self, columns: List[Dict[str, Any]]) -> int:
        """Calculates the spacing between existing columns."""
        if len(columns) > 1:
            spacing = columns[1]["left"] - columns[0]["left"]
            logger.debug("Column spacing calculated: %d EMUs", spacing)
            return spacing
        # Default spacing: ~2.4 inches
        default_spacing = 2300000
        logger.debug("Using default column spacing: %d EMUs", default_spacing)
        return default_spacing

    def _create_additional_columns(
        self,
        slide,
        columns: List[Dict[str, Any]],
        additional_needed: int,
        slide_width: int,
    ) -> int:
        """Creates additional columns by duplicating the first column."""
        if not columns or not columns[0].get("groups"):
            logger.warning("No reference column available for duplication")
            return 0

        ref_column = columns[0]
        col_spacing = self._calculate_column_spacing(columns)
        ref_group = ref_column["groups"][0]
        group_width = getattr(ref_group, "Width", 914400)

        created = 0
        for _ in range(additional_needed):
            last_column = columns[-1]
            new_column_left = last_column["left"] + col_spacing

            # Check if it fits horizontally
            if new_column_left + group_width > slide_width * 0.95:
                # Try with reduced spacing
                reduced_spacing = int(col_spacing * 0.75)
                new_column_left = last_column["left"] + reduced_spacing

                if new_column_left + group_width > slide_width * 0.95:
                    logger.warning("No more horizontal space for additional columns")
                    break

            new_column = {"left": new_column_left, "groups": []}

            # Duplicate all groups from the reference column
            for ref_group in ref_column["groups"]:
                new_group = self._duplicate_group_for_new_column(slide, ref_group, new_column_left)
                if new_group:
                    new_column["groups"].append(new_group)

            if new_column["groups"]:
                columns.append(new_column)
                created += 1
                logger.info("Created additional column at position %d", new_column_left)

        return created

    def _extend_column_rows(
        self,
        slide,
        column: Dict[str, Any],
        target_rows: int,
        slide_height: int,
    ) -> int:
        """Extends a column with additional rows if needed."""
        groups_list = column["groups"]
        current_rows = len(groups_list)

        if current_rows >= target_rows or not groups_list:
            return 0

        # Calculate vertical spacing
        if len(groups_list) > 1:
            vertical_spacing = groups_list[-1].Top - groups_list[-2].Top
        else:
            vertical_spacing = 500000  # ~0.52 inches by default

        template_group = groups_list[-1]
        created = 0

        for _ in range(target_rows - current_rows):
            new_top = int(template_group.Top + vertical_spacing)

            # Check vertical limit
            if new_top > slide_height * 0.85:
                logger.debug("Vertical space limit reached in column")
                break

            new_group = self._duplicate_group_for_new_row(slide, template_group, vertical_spacing)
            if new_group:
                groups_list.append(new_group)
                template_group = new_group
                created += 1

        column["groups"] = groups_list
        return created

    def _assign_name_to_group(self, group, name: str, font_size: int = 12) -> bool:
        """Assigns a name to the text shape within a group."""
        try:
            group_items = getattr(group, "GroupItems", None)
            if not group_items:
                return False

            for i in range(1, group_items.Count + 1):
                item = group_items.Item(i)
                try:
                    has_text = bool(getattr(item, "HasTextFrame", 0))
                    if has_text:
                        self._set_shape_text(item, name)
                        # Adjust font size based on name length
                        if len(name) > 20:
                            item.TextFrame.TextRange.Font.Size = font_size - 2
                        elif len(name) > 15:
                            item.TextFrame.TextRange.Font.Size = font_size - 1
                        else:
                            item.TextFrame.TextRange.Font.Size = font_size
                        return True
                except Exception as e:
                    logger.debug("Error setting text in group item: %s", e)
                    continue

            return False
        except Exception as e:
            logger.error("Error assigning name to group: %s", e)
            return False

    def _group_shapes_into_columns(self, shapes: List[Any]) -> List[Dict[str, Any]]:
        """Groups shapes into columns by horizontal position."""
        tolerance = 150000  # ~0.16 inches
        columns = []

        for shape in sorted(shapes, key=lambda s: (getattr(s, "Left", 0), getattr(s, "Top", 0))):
            left = getattr(shape, "Left", 0)

            found = False
            for column in columns:
                if abs(column["left"] - left) <= tolerance:
                    column["groups"].append(shape)
                    found = True
                    break

            if not found:
                columns.append({"left": left, "groups": [shape]})

        # Sort groups within each column by Top
        for column in columns:
            column["groups"].sort(key=lambda s: getattr(s, "Top", 0))

        # Sort columns by Left
        columns.sort(key=lambda c: c["left"])

        return columns

    def _layout_multi_candidate_names(
        self,
        *,
        slide,
        placeholder_shapes: List[Any],
        names: List[str],
        context: str,
        warnings: List[str],
    ) -> bool:
        """
        Distributes names in a multi-column layout, creating additional columns and rows as needed.
        """
        if not placeholder_shapes:
            logger.warning("No placeholder shapes provided for layout in %s", context)
            return False
        if not names:
            logger.debug("No names to layout in %s", context)
            return True

        try:
            # Find groups that contain the placeholders
            groups = []
            groups_seen = set()

            for text_shape in placeholder_shapes:
                group = self._find_group_for_text_shape(slide, text_shape)
                if group:
                    group_id = getattr(group, "Id", id(group))
                    if group_id not in groups_seen:
                        groups.append(group)
                        groups_seen.add(group_id)

            # If there are no groups, work with individual shapes
            if not groups:
                logger.info("No groups found, working with individual shapes")
                columns = self._group_name_placeholder_columns(placeholder_shapes)
                return self._layout_without_groups(slide, columns, names, context)

            # Clear text from existing groups
            for group in groups:
                try:
                    group_items = getattr(group, "GroupItems", None)
                    if group_items:
                        for i in range(1, group_items.Count + 1):
                            item = group_items.Item(i)
                            if getattr(item, "HasTextFrame", 0):
                                self._set_shape_text(item, "")
                    group.Visible = False
                except Exception as e:
                    logger.warning("Error clearing group: %s", e)

            # Organize into columns
            columns = self._group_shapes_into_columns(groups)
            if not columns:
                logger.warning("Could not group into columns")
                return False

            # Get slide dimensions
            try:
                slide_width = slide.Parent.PageSetup.SlideWidth
                slide_height = slide.Parent.PageSetup.SlideHeight
            except Exception:
                slide_width = 9144000
                slide_height = 6858000

            # Calculate layout requirements
            max_per_column = max(len(col["groups"]) for col in columns) if columns else 18
            columns_needed, names_per_column = self._calculate_layout_requirements(len(names), max_per_column)

            logger.info("Layout requirements: %d columns needed, %d names per column", columns_needed, names_per_column)

            # Create additional columns if necessary
            current_columns = len(columns)
            if columns_needed > current_columns:
                additional_needed = columns_needed - current_columns
                created = self._create_additional_columns(slide, columns, additional_needed, slide_width)
                logger.info("Created %d additional columns", created)

            # Extend rows in each column if needed
            for col_idx, column in enumerate(columns):
                if len(column["groups"]) < names_per_column:
                    self._extend_column_rows(slide, column, names_per_column, slide_height)

            # Distribute names across the groups
            successful_placements = 0
            name_idx = 0

            for col_idx, column in enumerate(columns):
                for row_idx, group in enumerate(column["groups"]):
                    if name_idx >= len(names):
                        group.Visible = False
                    else:
                        if self._assign_name_to_group(group, names[name_idx]):
                            group.Visible = True
                            successful_placements += 1
                            logger.debug("Assigned '%s' to col %d, row %d", names[name_idx][:20], col_idx, row_idx)
                        name_idx += 1

            logger.info("Layout complete: %d/%d names placed", successful_placements, len(names))

            if successful_placements < len(names):
                logger.warning("Only placed %d of %d names", successful_placements, len(names))

            return successful_placements >= len(names) * 0.8

        except Exception as exc:
            logger.error("Error in multi-candidate layout: %s", exc, exc_info=True)
            return False

    def _group_shapes_into_columns(self, shapes: List[Any]) -> List[Dict[str, Any]]:
        """Groups shapes/groups into columns based on their horizontal position."""
        tolerance = 120000
        columns = []

        for shape in sorted(shapes, key=lambda s: (getattr(s, "Left", 0), getattr(s, "Top", 0))):
            left = getattr(shape, "Left", 0)

            # Find existing column
            found = False
            for column in columns:
                if abs(column["left"] - left) <= tolerance:
                    column["groups"].append(shape)
                    found = True
                    break

            if not found:
                columns.append({"left": left, "groups": [shape]})

        # Sort shapes within each column by Top
        for column in columns:
            column["groups"].sort(key=lambda s: getattr(s, "Top", 0))

        # Sort columns by Left
        columns.sort(key=lambda c: c["left"])

        return columns

    def _find_checkbox_near_text(self, slide, text_shape) -> Any:
        """Finds the checkbox (oval/rectangle) near the text shape."""
        try:
            text_left = getattr(text_shape, "Left", 0)
            text_top = getattr(text_shape, "Top", 0)

            # Find shapes near the text (to the left)
            horizontal_tolerance = 400000  # ~0.4 inches
            vertical_tolerance = 150000    # ~0.15 inches

            count = getattr(slide.Shapes, "Count", 0)
            for idx in range(1, count + 1):
                shape = slide.Shapes(idx)

                # Ignore the same shape
                if getattr(shape, "Id", None) == getattr(text_shape, "Id", None):
                    continue

                shape_type = getattr(shape, "Type", None)
                # Type 1 = msoAutoShape (ovals, rectangles)
                if shape_type == 1:
                    shape_left = getattr(shape, "Left", 0)
                    shape_top = getattr(shape, "Top", 0)

                    # The checkbox must be to the left of the text
                    if shape_left < text_left:
                        h_dist = abs(text_left - shape_left)
                        v_dist = abs(text_top - shape_top)

                        if h_dist <= horizontal_tolerance and v_dist <= vertical_tolerance:
                            return shape

            return None
        except Exception as e:
            logger.debug("Error finding checkbox: %s", e)
            return None

    def _layout_without_groups(self, slide, columns, names, context) -> bool:
        """Layout with individual shapes - duplicates text and checkboxes separately."""
        try:
            logger.info("Layout with individual shapes: %d names, %d base columns",
                       len(names), len(columns))

            # Get slide dimensions
            try:
                slide_width = slide.Parent.PageSetup.SlideWidth
                slide_height = slide.Parent.PageSetup.SlideHeight
            except:
                slide_width = 9144000
                slide_height = 6858000

            # PRE-MAP all checkboxes to their texts to avoid repeated lookups
            checkbox_map = {}  # {text_shape_id: checkbox_shape}
            logger.info("Pre-mapping checkboxes...")

            for column in columns:
                for text_shape in column["shapes"]:
                    text_id = getattr(text_shape, "Id", None)
                    if text_id:
                        checkbox = self._find_checkbox_near_text(slide, text_shape)
                        if checkbox:
                            checkbox_map[text_id] = checkbox
                            logger.debug("Checkbox mapped for text Id=%s", text_id)

            logger.info("Checkboxes mapped: %d", len(checkbox_map))

            # Calculate current capacity
            num_base_columns = len(columns)
            max_rows = max(len(col["shapes"]) for col in columns)
            total_capacity = num_base_columns * max_rows

            logger.info("Current capacity: %d (%d cols x %d rows)", total_capacity, num_base_columns, max_rows)

            # Create additional columns if necessary
            if len(names) > total_capacity:
                cols_needed = math.ceil((len(names) - total_capacity) / max_rows)
                logger.info("Creating %d additional columns", cols_needed)

                if num_base_columns > 1:
                    col_spacing = columns[1]["left"] - columns[0]["left"]
                else:
                    col_spacing = 2300000

                for new_col_idx in range(cols_needed):
                    last_column = columns[-1]
                    new_column_left = last_column["left"] + col_spacing

                    # Validate if it fits
                    ref_shape = columns[0]["shapes"][0]
                    shape_width = getattr(ref_shape, "Width", 914400)

                    if new_column_left + shape_width > slide_width * 0.9:
                        logger.warning("No more space for additional columns")
                        break

                    new_column = {"left": new_column_left, "shapes": []}

                    # Duplicate text shapes from the first column
                    for ref_text in columns[0]["shapes"]:
                        try:
                            # Duplicate text
                            text_range = ref_text.Duplicate()
                            new_text = text_range.Item(1)
                            new_text.Left = int(new_column_left)
                            new_text.Top = ref_text.Top
                            self._set_shape_text(new_text, "")
                            new_text.Visible = False

                            # Duplicate associated checkbox using the map
                            ref_text_id = getattr(ref_text, "Id", None)
                            ref_checkbox = checkbox_map.get(ref_text_id)

                            if ref_checkbox:
                                try:
                                    cb_range = ref_checkbox.Duplicate()
                                    new_cb = cb_range.Item(1)
                                    cb_offset = ref_checkbox.Left - ref_text.Left
                                    new_cb.Left = int(new_column_left + cb_offset)
                                    new_cb.Top = ref_text.Top
                                    new_cb.Visible = False

                                    # Map the new checkbox to the new text
                                    new_text_id = getattr(new_text, "Id", None)
                                    if new_text_id:
                                        checkbox_map[new_text_id] = new_cb
                                        logger.debug("New checkbox mapped: text Id=%s", new_text_id)
                                except Exception as e:
                                    logger.warning("Error duplicating checkbox: %s", e)

                            new_column["shapes"].append(new_text)
                        except Exception as e:
                            logger.warning("Error duplicating shape: %s", e)

                    if new_column["shapes"]:
                        columns.append(new_column)
                        logger.info("Column %d created with %d shapes", len(columns), len(new_column["shapes"]))

            # Create additional rows if necessary
            num_columns = len(columns)
            names_per_column = math.ceil(len(names) / num_columns)

            for col_idx, column in enumerate(columns):
                while len(column["shapes"]) < names_per_column:
                    if not column["shapes"]:
                        break

                    template_text = column["shapes"][-1]
                    try:
                        # Duplicate text
                        text_range = template_text.Duplicate()
                        new_text = text_range.Item(1)

                        if len(column["shapes"]) > 1:
                            v_spacing = column["shapes"][-1].Top - column["shapes"][-2].Top
                        else:
                            v_spacing = 500000

                        new_top = int(template_text.Top + v_spacing)

                        if new_top > slide_height * 0.85:
                            break

                        new_text.Top = new_top
                        new_text.Left = template_text.Left
                        self._set_shape_text(new_text, "")
                        new_text.Visible = False

                        # Duplicate checkbox using the map
                        template_text_id = getattr(template_text, "Id", None)
                        template_cb = checkbox_map.get(template_text_id)

                        if template_cb:
                            try:
                                cb_range = template_cb.Duplicate()
                                new_cb = cb_range.Item(1)
                                new_cb.Top = new_top
                                new_cb.Left = template_cb.Left
                                new_cb.Visible = False

                                # Map the new checkbox
                                new_text_id = getattr(new_text, "Id", None)
                                if new_text_id:
                                    checkbox_map[new_text_id] = new_cb
                            except Exception as e:
                                logger.warning("Error duplicating vertical checkbox: %s", e)

                        column["shapes"].append(new_text)
                    except Exception as e:
                        logger.warning("Error duplicating row: %s", e)
                        break

            # Assign names using the checkbox map
            successful_placements = 0
            name_idx = 0

            for col_idx, column in enumerate(columns):
                for row_idx, text_shape in enumerate(column["shapes"]):
                    text_id = getattr(text_shape, "Id", None)
                    checkbox = checkbox_map.get(text_id) if text_id else None

                    if name_idx >= len(names):
                        # No more names, hide
                        text_shape.Visible = False
                        if checkbox:
                            checkbox.Visible = False
                    else:
                        # Assign name
                        try:
                            self._set_shape_text(text_shape, names[name_idx])
                            text_shape.TextFrame.TextRange.Font.Size = 12
                            text_shape.Visible = True

                            if checkbox:
                                checkbox.Visible = True

                            logger.debug("Assigned '%s' to col %d, row %d (text Id=%s)",
                                       names[name_idx][:20], col_idx, row_idx, text_id)

                            successful_placements += 1
                            name_idx += 1
                        except Exception as e:
                            logger.warning("Error assigning name to col %d, row %d: %s",
                                         col_idx, row_idx, e)
                            name_idx += 1

            logger.info("Layout complete: %d/%d names placed in %d columns",
                       successful_placements, len(names), num_columns)

            return successful_placements >= len(names) * 0.5

        except Exception as exc:
            logger.error("Error in layout: %s", exc, exc_info=True)
            return False
    
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
        """Populate group slide with adaptive layout for multiple names."""
        group_name = detail.get("group_name") or detail.get("name") or "Group"
        category = detail.get("category") or ""
        context = self._detail_context(detail, "group slide")
        
        # Tokenize names in case they contain delimiters
        names = tokenize_delimited_block(group_name) or [group_name]
        cleaned_names = [name.strip() for name in names if name and name.strip()]
        
        logger.debug("Group slide: %d names found in '%s'", len(cleaned_names), group_name[:50])
        
        # First try to find and use adaptive layout for multiple names
        placeholder_shapes = self._collect_placeholder_shapes(slide, "[New Names]")
        
        if cleaned_names and placeholder_shapes:
            if len(cleaned_names) > 1:
                logger.info("Attempting adaptive layout for %d names in group slide", len(cleaned_names))
                # Use adaptive layout for multiple names
                layout_ok = self._layout_multi_candidate_names(
                    slide=slide,
                    placeholder_shapes=placeholder_shapes,
                    names=cleaned_names,
                    context=context,
                    warnings=warnings,
                )
                
                if not layout_ok:
                    logger.warning("Adaptive layout failed, using fallback for group slide")
                    # Fallback: distribute names across available placeholders
                    for idx, value in enumerate(cleaned_names):
                        if idx < len(placeholder_shapes):
                            self._set_shape_text(placeholder_shapes[idx], value)
                        else:
                            warnings.append(
                                f"Group slide '{category}' has more names ({len(cleaned_names)}) than "
                                f"placeholders ({len(placeholder_shapes)}); extra name '{value}' was omitted."
                            )
                            break
                    # Clear remaining placeholders
                    for idx in range(len(cleaned_names), len(placeholder_shapes)):
                        self._set_shape_text(placeholder_shapes[idx], "")
            else:
                # Single name - use simple replacement
                self._replace_placeholder_text(
                    slide,
                    "[New Names]",
                    cleaned_names[0] if cleaned_names else group_name,
                    warnings=warnings,
                    required=bool(cleaned_names),
                    context=context,
                )
        else:
            # No placeholder shapes found - use text replacement
            combined_names = ", ".join(cleaned_names) if len(cleaned_names) > 1 else (cleaned_names[0] if cleaned_names else group_name)
            self._replace_placeholder_text(
                slide,
                "[New Names]",
                combined_names,
                warnings=warnings,
                required=bool(group_name),
                context=context,
            )
        
        # Populate category
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
        """Populate multi-candidate slide with adaptive layout."""
        category = detail.get("category") or detail.get("group_name") or "General"
        names = tokenize_delimited_block(detail.get("name")) or [detail.get("name", "")]
        summary_tracker.setdefault(category, []).extend([name for name in names if name])

        context = self._detail_context(detail, "multi slide")
        placeholder_shapes = self._collect_placeholder_shapes(slide, "Name Candidate")
        
        cleaned_names = [name.strip() for name in names if name and name.strip()]
        
        logger.debug("Multi slide: %d names found", len(cleaned_names))
        
        if not placeholder_shapes:
            if cleaned_names:
                warnings.append(
                    f"No 'Name Candidate' placeholders found in {context}; "
                    f"candidate names will not render."
                )
            return
        
        if not cleaned_names:
            # Clear all placeholders if no names
            for shape in placeholder_shapes:
                self._set_shape_text(shape, "")
            return

        logger.info("Attempting adaptive layout for %d names in multi slide", len(cleaned_names))
        
        # Always try adaptive layout first
        layout_ok = self._layout_multi_candidate_names(
            slide=slide,
            placeholder_shapes=placeholder_shapes,
            names=cleaned_names,
            context=context,
            warnings=warnings,
        )

        if not layout_ok:
            logger.warning("Adaptive layout failed, using fallback for multi slide")
            # Fallback: simple distribution across existing placeholders
            for idx, value in enumerate(cleaned_names):
                if idx < len(placeholder_shapes):
                    self._set_shape_text(placeholder_shapes[idx], value)
                else:
                    warnings.append(
                        f"Slide '{category}' has more names ({len(cleaned_names)}) than "
                        f"placeholders ({len(placeholder_shapes)}); extra name '{value}' was omitted."
                    )
                    break
            # Clear remaining placeholders
            for idx in range(len(cleaned_names), len(placeholder_shapes)):
                self._set_shape_text(placeholder_shapes[idx], "")

        # Populate category
        self._replace_placeholder_text(
            slide,
            "Category",
            category,
            warnings=warnings,
            required=bool(category),
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

        # Leave the comments placeholder untouched so the template formatting remains.

        kana = detail.get("kana") or ""
        self._replace_placeholder_text(
            slide,
            "Pronunciation/Katakana",
            kana,
            warnings=warnings,
            required=False,
            context=context,
        )

        # Preserve vote placeholders (Positive/Neutral/Negative) without overwriting template text.

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
            if detail.get("category"):
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

    def _insert_image_slide(
        self,
        presentation,
        image_path: str,
        base_template,
        warnings: Optional[List[str]] = None,
    ):
        """
        Insert a slide with an image from the original PPTX conversion.
        
        This method creates a new slide and inserts a JPG image that was generated
        from the original PowerPoint file, filling the entire slide area.
        
        Args:
            presentation: PowerPoint presentation COM object
            image_path: Path to the JPG image file
            base_template: Base template presentation for layout reference
            warnings: Optional list to collect warning messages
            
        Returns:
            The newly created slide object
        """
        if warnings is None:
            warnings = []
        
        try:
            # Resolve image path
            img_path = Path(image_path)
            if not img_path.exists():
                warning_msg = f"Image file not found: {image_path}"
                logger.warning(warning_msg)
                warnings.append(warning_msg)
                return None
            
            # Add blank slide using base template layout
            slide_index = presentation.Slides.Count + 1
            slide = presentation.Slides.Add(
                slide_index,
                base_template.Slides(1).Layout
            )
            
            # Get slide dimensions
            slide_width = presentation.PageSetup.SlideWidth
            slide_height = presentation.PageSetup.SlideHeight
            
            # Insert image to fill entire slide
            slide.Shapes.AddPicture(
                FileName=str(img_path.resolve()),
                LinkToFile=False,  # Embed image in presentation
                SaveWithDocument=True,
                Left=0,
                Top=0,
                Width=slide_width,
                Height=slide_height
            )
            
            logger.debug("Inserted image slide from: %s", image_path)
            return slide
            
        except Exception as exc:  # pylint: disable=broad-except
            error_msg = f"Failed to insert image slide from {image_path}: {exc}"
            logger.error(error_msg)
            warnings.append(error_msg)
            return None


pptx_builder_service = PPTXBuilderService(
    base_template_path=Path("templates/BackgroundDefaultTemplate/template_default_seperator_2019.pptx")
)
