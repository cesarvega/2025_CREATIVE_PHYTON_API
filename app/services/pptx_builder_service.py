import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Tuple

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
        progress_callback=None,
    ) -> PresentationBuildArtifacts:
        """Compose final PPT deliverables using existing templates and Excel data.
        
        Args:
            request: Presentation creation request
            options: Build options
            details: List of detail items
            excel_data: Processed Excel data
            progress_callback: Optional callback function to report progress (int 0-100)
        """

        if progress_callback:
            progress_callback(72)

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
            # Protect entire PowerPoint build operation with semaphore
            from app.utils.com_manager import com_manager
            with com_manager.acquire("PowerPoint.Application - Build Presentation"):
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

                    # Use cached template handle, open lazily only if needed
                    source_presentation = template_handles.get(template_key)
                    if source_presentation is None:
                        source_presentation = app.Presentations.Open(
                            str(template_paths[template_key].resolve()),
                            ReadOnly=True,
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

                if progress_callback:
                    progress_callback(88)

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
            # Protect entire PowerPoint composition operation with semaphore
            from app.utils.com_manager import com_manager
            with com_manager.acquire("PowerPoint.Application - Compose with Original"):
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

                    # Use cached template handle, open lazily only if needed
                    source_presentation = template_handles.get(template_key)
                    if source_presentation is None:
                        source_presentation = app.Presentations.Open(
                            str(template_paths[template_key].resolve()),
                            ReadOnly=True,
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
        """Open template presentations with caching optimization."""
        handles: Dict[str, Any] = {}
        for key, path in template_paths.items():
            # Open with ReadOnly=True and WithWindow=False for better performance
            handles[key] = app.Presentations.Open(
                str(path.resolve()),
                ReadOnly=True,
                WithWindow=False
            )
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
            result = slides.InsertFromFile(
                str(template_path.resolve()),
                insert_index,
                1,
                1,
            )

            # InsertFromFile can return either a SlideRange object or an integer
            # If it's an integer, it's the index of the inserted slide
            if isinstance(result, int):
                # Get the slide by index
                return slides.Item(result)
            else:
                # It's a SlideRange, get the first item
                return result.Item(1)
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
        """Iterate over text shapes in a slide.

        Optimized with early exit and minimal COM calls.
        """
        count = getattr(slide.Shapes, "Count", 0)
        if count == 0:
            return

        for idx in range(1, count + 1):
            shape = slide.Shapes(idx)
            try:
                has_text_frame = bool(getattr(shape, "HasTextFrame", 0))
            except Exception:
                continue  # Skip shapes that don't support HasTextFrame

            if not has_text_frame:
                continue

            try:
                _ = shape.TextFrame.TextRange
            except Exception:
                continue  # Skip shapes without valid TextRange

            yield shape

    def _collect_placeholder_shapes(self, slide, placeholder: str) -> List[Any]:
        """Collect all shapes containing the placeholder text, sorted by position.

        Optimized with early exit and position caching.
        """
        placeholder_lower = placeholder.lower()
        matches: List[Tuple[Any, float, float]] = []  # (shape, top, left)

        for shape in self._iter_text_shapes(slide):
            try:
                text = shape.TextFrame.TextRange.Text
                text_lower = text.lower().strip()
            except Exception:
                continue

            # Check if it's a name candidate placeholder (could be empty or contain the placeholder text)
            if placeholder_lower in text_lower or "name candidate" in text_lower:
                # Cache position values to avoid multiple COM calls
                top = getattr(shape, "Top", 0)
                left = getattr(shape, "Left", 0)
                matches.append((shape, top, left))
                logger.debug(
                    "Found placeholder: text='%s' (len=%d), Left=%s, Top=%s",
                    text_lower[:30] if text_lower else "[empty]",
                    len(text),
                    left,
                    top
                )

        # Sort by Top (row) first, then by Left (column) within each row
        # This ensures placeholders are processed row-by-row, left-to-right
        matches.sort(key=lambda item: (item[1], item[2]))

        logger.info("Collected %d placeholders, sorted by position (row-by-row)", len(matches))
        # Return only shapes, without cached positions
        return [item[0] for item in matches]

    def _find_checkbox_near_shape(self, slide, text_shape) -> Any:
        """Find the checkbox (square) near a text shape (left or right side).

        Optimized with early exit and spatial indexing.
        """
        try:
            text_left = getattr(text_shape, "Left", 0)
            text_top = getattr(text_shape, "Top", 0)
            text_id = getattr(text_shape, "Id", None)

            # Tolerance for finding nearby shapes
            horizontal_tolerance = 400000  # ~0.4 inches
            vertical_tolerance = 150000    # ~0.15 inches

            count = getattr(slide.Shapes, "Count", 0)
            closest_checkbox = None
            min_distance = float('inf')

            # Early exit if no shapes
            if count == 0:
                return None

            for idx in range(1, count + 1):
                shape = slide.Shapes(idx)

                # Skip the text shape itself (cached ID comparison)
                if text_id is not None and getattr(shape, "Id", None) == text_id:
                    continue

                shape_type = getattr(shape, "Type", None)
                # Type 1 = msoAutoShape (rectangles, ovals, etc.)
                if shape_type != 1:
                    continue

                shape_left = getattr(shape, "Left", 0)
                shape_top = getattr(shape, "Top", 0)

                # Check both left and right side of the text
                h_dist = abs(text_left - shape_left)
                v_dist = abs(text_top - shape_top)

                # Early exit if outside tolerance
                if h_dist > horizontal_tolerance or v_dist > vertical_tolerance:
                    continue

                # Calculate total distance
                distance = h_dist + v_dist
                if distance < min_distance:
                    min_distance = distance
                    closest_checkbox = shape

            return closest_checkbox
        except Exception as exc:
            logger.debug("Error finding checkbox for shape: %s", exc)
            return None

    def _assign_names_to_placeholders(
        self,
        placeholder_shapes: List[Any],
        names: List[str],
        context: str,
    ) -> None:
        """Assign names to existing placeholder shapes (max 18) and delete unused ones."""
        max_candidates = len(placeholder_shapes)

        # Get the slide from the first shape
        slide = None
        if placeholder_shapes:
            try:
                slide = placeholder_shapes[0].Parent
            except Exception:
                pass

        # FIRST: Map all checkboxes to their placeholders BEFORE any deletion
        checkbox_map = {}  # {placeholder_id: checkbox_shape}
        if slide:
            for shape in placeholder_shapes:
                shape_id = getattr(shape, "Id", None)
                if shape_id:
                    checkbox = self._find_checkbox_near_shape(slide, shape)
                    if checkbox:
                        checkbox_map[shape_id] = checkbox
                        logger.debug("Mapped checkbox to placeholder Id=%s in %s", shape_id, context)

        logger.info("Mapped %d checkboxes to placeholders in %s", len(checkbox_map), context)

        # SECOND: Assign names and collect unused shapes
        used_shapes = []
        unused_shapes = []

        for idx, shape in enumerate(placeholder_shapes):
            try:
                # Get current text and position for debugging
                try:
                    current_text = shape.TextFrame.TextRange.Text.strip()[:30]
                    shape_left = getattr(shape, "Left", 0)
                    shape_top = getattr(shape, "Top", 0)
                except:
                    current_text = "[cannot read]"
                    shape_left = 0
                    shape_top = 0

                if idx < len(names):
                    # Assign name to placeholder
                    self._set_shape_text(shape, names[idx])
                    shape.Visible = True
                    used_shapes.append(shape)
                    logger.info(
                        "→ Assigned '%s' to placeholder #%d (was: '%s', Left=%d, Top=%d) in %s",
                        names[idx][:30], idx + 1, current_text, shape_left, shape_top, context
                    )
                else:
                    # Mark for deletion
                    unused_shapes.append(shape)
                    logger.info(
                        "✗ Marked placeholder #%d for deletion (text: '%s', Left=%d, Top=%d) in %s",
                        idx + 1, current_text, shape_left, shape_top, context
                    )
            except Exception as exc:
                logger.warning("Error processing placeholder %d in %s: %s", idx + 1, context, exc)

        logger.info("Assignment complete: %d used, %d unused placeholders in %s", len(used_shapes), len(unused_shapes), context)

        # THIRD: Delete unused placeholders and their pre-mapped checkboxes
        deleted_count = 0
        hidden_count = 0

        for idx, shape in enumerate(unused_shapes):
            try:
                shape_id = getattr(shape, "Id", None)

                # Try to get current text for debugging
                try:
                    current_text = shape.TextFrame.TextRange.Text[:30]
                except:
                    current_text = "[cannot read]"

                # Delete associated checkbox using the pre-mapped association
                if shape_id and shape_id in checkbox_map:
                    checkbox = checkbox_map[shape_id]
                    try:
                        checkbox.Delete()
                        logger.debug("✓ Deleted checkbox for unused placeholder Id=%s (text='%s') in %s", shape_id, current_text, context)
                    except Exception as cb_exc:
                        logger.warning("✗ Could not delete checkbox for Id=%s: %s", shape_id, cb_exc)
                        try:
                            checkbox.Visible = False
                            hidden_count += 1
                        except Exception:
                            pass

                # Delete the text placeholder
                shape.Delete()
                deleted_count += 1
                logger.debug("✓ Deleted unused placeholder %d/%d (Id=%s, text='%s') in %s",
                           idx + 1, len(unused_shapes), shape_id, current_text, context)

            except Exception as exc:
                logger.warning("✗ Error deleting unused placeholder %d/%d (Id=%s): %s",
                             idx + 1, len(unused_shapes), shape_id, exc)
                # Fallback: try to hide it
                try:
                    shape.Visible = False
                    hidden_count += 1
                    logger.debug("↓ Hidden placeholder %d/%d instead of deleting in %s", idx + 1, len(unused_shapes), context)
                except Exception:
                    logger.error("✗✗ Could not delete OR hide placeholder %d/%d in %s", idx + 1, len(unused_shapes), context)

        logger.info("Cleanup complete in %s: %d deleted, %d hidden", context, deleted_count, hidden_count)

        if len(names) > max_candidates:
            logger.warning(
                "Only %d of %d names were assigned in %s (template limit: %d)",
                max_candidates, len(names), context, max_candidates
            )

    
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
        """Set text on a shape with error handling.

        Optimized to avoid unnecessary operations.
        """
        if text is None:
            text = ""

        try:
            shape.TextFrame.TextRange.Text = text
        except Exception as exc:
            logger.debug(
                "Could not assign text to shape: %s (error: %s)",
                getattr(shape, "Name", "<unnamed>"),
                str(exc)[:50]
            )

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
        """Populate group slide using existing placeholders (max 18)."""
        group_name = detail.get("group_name") or detail.get("name") or "Group"
        category = detail.get("category") or ""
        context = self._detail_context(detail, "group slide")

        # Tokenize names in case they contain delimiters
        names = tokenize_delimited_block(group_name) or [group_name]
        cleaned_names = [name.strip() for name in names if name and name.strip()]

        logger.debug("Group slide: %d names found in '%s'", len(cleaned_names), group_name[:50])

        # Collect placeholder shapes
        placeholder_shapes = self._collect_placeholder_shapes(slide, "Name Candidate")
        logger.info("Found %d placeholder shapes for 'Name Candidate' in group slide", len(placeholder_shapes))

        if cleaned_names and placeholder_shapes:
            # Assign names to existing placeholders
            self._assign_names_to_placeholders(placeholder_shapes, cleaned_names, context)
        elif not cleaned_names and placeholder_shapes:
            # No names provided: hide all placeholder shapes
            for shape in placeholder_shapes:
                try:
                    shape.Visible = False
                except Exception:
                    pass
        else:
            # No placeholder shapes found - use text replacement fallback
            combined_names = ", ".join(cleaned_names) if len(cleaned_names) > 1 else (cleaned_names[0] if cleaned_names else group_name)
            self._replace_placeholder_text(
                slide,
                "Name Candidate",
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
        """Populate multi-candidate slide using existing placeholders (max 18)."""
        category = detail.get("category") or detail.get("group_name") or "General"
        names = tokenize_delimited_block(detail.get("name")) or [detail.get("name", "")]
        summary_tracker.setdefault(category, []).extend([name for name in names if name])

        context = self._detail_context(detail, "multi slide")
        cleaned_names = [name.strip() for name in names if name and name.strip()]

        logger.debug("Multi slide: %d names found", len(cleaned_names))

        # Collect placeholder shapes
        placeholder_shapes = self._collect_placeholder_shapes(slide, "Name Candidate")

        if not placeholder_shapes:
            if cleaned_names:
                warnings.append(
                    f"No 'Name Candidate' placeholders found in {context}; "
                    f"candidate names will not render."
                )
            return

        if not cleaned_names:
            # Hide all placeholders if no names
            for shape in placeholder_shapes:
                try:
                    shape.Visible = False
                except Exception:
                    pass
            return

        # Assign names to existing placeholders
        self._assign_names_to_placeholders(placeholder_shapes, cleaned_names, context)

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
        """Populate summary slide - leaves content empty, only keeps the title.

        The summary slide template is used as-is without filling in candidate names.
        This allows users to manually add content if needed.
        """
        # Don't populate the summary slide with candidate names
        # The template slide is used as-is
        logger.debug("Summary slide created from template (content not auto-populated)")

        # Optionally, you can still replace the title placeholder if needed
        # but leave the body empty for manual editing
        # self._replace_placeholder_text(
        #     slide,
        #     "Name Candidates - Summary",
        #     "",  # Empty content
        #     replace_all=False,
        #     warnings=warnings,
        #     required=False,
        #     context="summary slide",
        # )

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
            # Protect PowerPoint export operation with semaphore
            from app.utils.com_manager import com_manager
            with com_manager.acquire("PowerPoint.Application - Export Slide to Image"):
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
            try:
                if presentation:
                    presentation.Close()
            except Exception as e:
                logger.debug("Presentation already closed or error closing: %s", str(e))

            try:
                if powerpoint:
                    powerpoint.Quit()
            except Exception as e:
                logger.debug("PowerPoint already closed or error quitting: %s", str(e))

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
