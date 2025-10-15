"""
Presentation creation orchestration service.
"""

import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

from fastapi import HTTPException

from app.config.db import DatabaseConnectionError, create_connection
from app.models.excel_models import (
    ProcessedExcelData,
)
from app.models.presentation_models import (
    CreatePresentationRequest,
    CreatePresentationResponse,
    DetailItem,
    PresentationBuildOptions,
    PresentationData,
)
from app.models.response_models import PPTXConversionResponse
from app.services.excel_service import GROUP_MARKERS, process_excel_file
from app.services.pptx_service import pptx_service
from app.services.pptx_builder_service import pptx_builder_service
from app.services.word_service import generate_feedback_document
from app.services.email_service import send_presentation_emails
from app.utils.logging_utils import get_logger
from app.utils.path_utils import resolve_project_output


logger = get_logger(__name__)


# Template metadata cache to avoid repeated DB queries
_template_metadata_cache: Dict[str, Dict[str, Any]] = {}
_cache_timestamp: Optional[float] = None
_CACHE_TTL_SECONDS = 300  # 5 minutes


class PresentationService:
    """Service for orchestrating complete presentation creation.
    
    This service coordinates between multiple services (Excel, PPTX, Word, Email)
    to create complete presentations with all associated artifacts.
    """

    def create_presentation(
        self, request: CreatePresentationRequest
    ) -> CreatePresentationResponse:
        """
        Orchestrate the complete presentation creation process.

        Args:
            request: Complete presentation creation request

        Returns:
            CreatePresentationResponse: Result of the presentation creation
        """
        start_time = time.time()

        try:
            # 1. Process Excel file
            logger.info("Processing Excel file: %s", request.excel_filename)
            excel_data = self._process_excel_file(request)

            # 2. Convert PPTX file
            logger.info("Converting PPTX file: %s", request.pptx_filename)
            pptx_data = self._convert_pptx_file(request)

            # 3. Generate slides from Excel arrays
            logger.info("Generating slides from Excel data")
            slides_data = self._generate_slides_from_excel(
                excel_data=excel_data,
                pptx_data=pptx_data,
                request=request,
            )

            # 4. Template rotation no longer used - backgrounds handled differently

            logger.info(
                "Slides data generated: %d detail items, background=%s",
                len(slides_data.get("details", [])),
                slides_data.get("background_name"),
            )

            # 5. Generate physical PowerPoint file (if enabled)
            # generated_files = None
            # if request.generate_physical_pptx:
            #     logger.info("Generating physical PowerPoint file")
            #     try:
            #         ppt_files = self._generate_physical_powerpoint(
            #             slides_data=slides_data,
            #             excel_data=excel_data,
            #             request=request,
            #             pptx_data=pptx_data,
            #         )
            #         generated_files = ppt_files
                    
            #         # Update slides_data with generated file path for DB storage
            #         if ppt_files.get("printable_path"):
            #             slides_data["powerpoint_file"] = ppt_files["printable_path"]
                    
            #         logger.info(
            #             "Physical PowerPoint generated: %s (total slides: %s)",
            #             ppt_files.get("printable_path", "N/A"),
            #             ppt_files.get("total_slides", "0"),
            #         )
            #     except Exception as ppt_error:
            #         logger.error("Failed to generate physical PowerPoint: %s", str(ppt_error))
            #         # Continue execution - physical file generation is optional
            #         generated_files = {
            #             "error": str(ppt_error),
            #             "printable_path": "",
            #             "macro_path": "",
            #             "total_slides": "0",
            #             "warnings": "Failed to generate",
            #         }
            # else:
            #     logger.info("Physical PowerPoint generation skipped (generate_physical_pptx=False)")

            # 6. Create presentation in DB
            logger.info("Creating presentation in database")
            presentation_result = self._create_presentation_in_db(slides_data, request)

            # presentation_id = presentation_result.get("presentation_id")
            # if presentation_id:
            #     feedback_path = self._generate_feedback_document(
            #         presentation_id, request, excel_data
            #     )
            #     if feedback_path:
            #         logger.info("Feedback document generated: %s", feedback_path)

            #     self._send_notification_emails(presentation_id, request)

            processing_time = time.time() - start_time

            return CreatePresentationResponse(
                message="Presentation created successfully",
                presentation_id=presentation_result.get("presentation_id"),
                total_slides=len(slides_data["details"]),
                excel_data=excel_data,
                pptx_data=pptx_data,
                processing_time_seconds=processing_time,
            )

        except Exception as e:
            logger.error("Error creating presentation: %s", str(e))
            raise

    def build_presentation_files(
        self,
        *,
        build_request: CreatePresentationRequest,
        options: PresentationBuildOptions,
    ):
        """Generate PowerPoint deliverables directly from Excel data."""

        logger.info(
            "Assembling PPT files for project=%s display=%s slide_range=%s-%s",
            build_request.project,
            build_request.display_name,
            options.slide_start,
            options.slide_end,
        )

        excel_data = self._process_excel_file(build_request)

        stub_conversion = PPTXConversionResponse(
            message="Generated for PPT assembly",
            conversion_id=build_request.display_name,
            project_type=build_request.project_type,
            total_images=0,
            images=[],
            thumbnails=[],
            titles=[],
            pptx_file=options.base_template,
        )

        slides_dict = self._generate_slides_from_excel(
            excel_data=excel_data,
            pptx_data=stub_conversion,
            request=build_request,
        )

        details = slides_dict.get("details", [])

        return pptx_builder_service.compose_presentation(
            request=build_request,
            options=options,
            details=details,
            excel_data=excel_data,
        )

    def presentation_exists(
        self,
        project_name: str,
        display_name: str,
        exclude_id: Optional[int] = None,
    ) -> bool:
        """
        Check if a presentation with the given project and display name already exists.

        Args:
            project_name: The name of the project.
            display_name: The display name of the presentation.
            exclude_id: Optional presentation ID to exclude from the check.

        Returns:
            True if a matching presentation exists, False otherwise.
        """
        sql = "SELECT TOP 1 1 FROM [BI_GUIDELINES].[dbo].[nw_Master] WHERE Project = ? AND DisplayName = ?"
        params = [project_name, display_name]

        if exclude_id is not None:
            sql += " AND PresentationId != ?"
            params.append(exclude_id)

        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(sql, tuple(params))
                    return cursor.fetchone() is not None
        except Exception as e:
            logger.error("Error checking if presentation exists: %s", e, exc_info=True)
            raise HTTPException(status_code=500, detail="Database error while checking for presentation.") from e

    def _process_excel_file(
        self, request: CreatePresentationRequest
    ) -> ProcessedExcelData:
        """Process Excel file using the existing Excel service."""
        try:
            # Derive is_phonetics from presentation_type
            # If presentation_type is "Phonetics", process phonetic columns
            is_phonetics = request.presentation_type.lower() == "phonetics"

            # Process the Excel file
            result = process_excel_file(
                file_content=request.excel_file,
                is_phonetics=is_phonetics,
                has_groups=request.has_groups,
                test_name_order=request.test_name_order,
            )

            logger.info(
                "Excel processing completed: %d rows processed, test_name_order=%s, is_phonetics=%s (from presentation_type='%s')",
                result.total_rows_processed,
                request.test_name_order,
                is_phonetics,
                request.presentation_type,
            )
            return result

        except Exception as e:
            logger.error("Error processing Excel file: %s", str(e))
            raise

    def _convert_pptx_file(
        self, request: CreatePresentationRequest
    ) -> PPTXConversionResponse:
        """Convert PPTX file using the existing PPTX service."""
        try:
            # Convert PPTX to images
            result_dict = pptx_service.convert_pptx_to_images(
                file_content=request.pptx_file,
                filename=request.pptx_filename,
                display_name=request.display_name,
                project_type=request.project_type,
            )

            # Convert dict to PPTXConversionResponse
            result = PPTXConversionResponse(**result_dict)

            logger.info(
                "PPTX conversion completed: %d images generated", result.total_images
            )
            return result

        except Exception as e:
            logger.error("Error converting PPTX file: %s", str(e))
            raise

    def _generate_physical_powerpoint(
        self,
        *,
        slides_data: Dict[str, Any],
        excel_data: ProcessedExcelData,
        request: CreatePresentationRequest,
        pptx_data: PPTXConversionResponse,
    ) -> Dict[str, str]:
        """
        Generate physical PowerPoint file combining original PPTX slides with template-generated slides.
        
        This method orchestrates the creation of a complete PowerPoint presentation by:
        1. Taking slides from the original uploaded PPTX (converted to images)
        2. Generating slides from Excel data using templates
        3. Combining them in the correct order based on page_number setting
        
        Args:
            slides_data: Dictionary containing slide details and metadata
            excel_data: Processed Excel data with candidate information
            request: Original presentation creation request
            pptx_data: Converted PPTX data with image paths
            
        Returns:
            Dictionary with paths to generated files:
            - printable_path: Path to .pptx file
            - macro_path: Path to .pptm file (if generated)
        """
        try:
            # Build presentation build options from request
            build_options = PresentationBuildOptions(
                template_pack=request.template_pack or "BackgroundDefaultTemplate",
                base_template=request.base_template or "template_default_2019.pptx",
                multi_template=request.multi_template or "template_default_withgroups2019.pptx",
                group_template=request.group_template or "template_default_withgroup_2019.pptx",
                separator_template=request.separator_template or "template_default_seperator_2019.pptx",
                summary_template=request.summary_template or "template_default_summary2019.pptx",
                slide_start=1,
                slide_end=None,
                include_print_ready_version=True,
                include_macro_version=request.include_macro_version,
                return_urls=False,  # Not needed for internal generation
            )
            
            logger.info(
                "Generating physical PowerPoint: %d original images, %d template slides, insert at position %d",
                len(pptx_data.images) if pptx_data.images else 0,
                len(slides_data.get("details", [])),
                request.page_number,
            )
            # Save the original PPTX file to disk so COM can open it
            output_base, _ = resolve_project_output(
                request.display_name,
                request.project_type,
                fallback_subdir="generated_presentations",
            )
            project_folder = output_base
            pptx_original_path = project_folder / request.pptx_filename
            if not pptx_original_path.exists():
                with open(pptx_original_path, "wb") as f:
                    f.write(request.pptx_file)

            artifacts = pptx_builder_service.compose_presentation_with_original_slides(
                request=request,
                options=build_options,
                details=slides_data["details"],
                excel_data=excel_data,
                original_pptx_path=str(pptx_original_path.resolve()),
                page_number_insert=request.page_number,
            )
            
            # Convert PresentationBuildArtifacts to Dict[str, str] for response model
            result = {
                "printable_path": str(artifacts.printable_path) if artifacts.printable_path else "",
                "macro_path": str(artifacts.macro_path) if artifacts.macro_path else "",
                "total_slides": str(artifacts.total_slides),
                "warnings": ", ".join(artifacts.warnings) if artifacts.warnings else "None",
            }
            
            logger.info(
                "Physical PowerPoint generated successfully: %d total slides, %d warnings",
                artifacts.total_slides,
                len(artifacts.warnings),
            )
            
            if artifacts.warnings:
                for warning in artifacts.warnings:
                    logger.warning("PowerPoint generation warning: %s", warning)
            
            return result
            
        except Exception as e:
            logger.error("Error generating physical PowerPoint: %s", str(e))
            raise

    def _generate_slides_from_excel(
        self,
        *,
        excel_data: ProcessedExcelData,
        pptx_data: PPTXConversionResponse,
        request: CreatePresentationRequest,
    ) -> PresentationData:
        details: List[DetailItem] = []
        last_group_name = ""
        default_template = self._get_template_metadata("Default")
        default_template_id = default_template["template_id"]

    # 1) Initial PPTX slides (before the Excel-generated slides)
        ppt_images = pptx_data.images or []
        prefix_count = 0
        if request.page_number > 1 and ppt_images:
            prefix_count = min(request.page_number - 1, len(ppt_images))
            for idx in range(prefix_count):
                image_path = ppt_images[idx]
                description = Path(image_path).stem.replace("_", " ") if image_path else f"Slide {idx+1}"
                details.append(DetailItem(
                    slide_number=idx + 1,
                    slide_type="Image",
                    slide_bg_file_name=Path(image_path).as_posix() if image_path else "",
                    slide_description=description,
                    group_name="",
                    category="",
                    name="",
                    rationale="",
                    notation="",
                    kana="",
                    logo_filename="",
                    template_id=0,
                    name_sub_group=""
                ))

    # 2) Slides generated from the Excel data
        slide_number = request.page_number
        total_rows = excel_data.total_rows_processed

        for index in range(total_rows):
            marker = (excel_data.lst_types[index] or "").strip().upper()
            category = (excel_data.lst_categories[index] or "").strip()
            name = (excel_data.lst_names[index] or "").strip()
            rationale = (excel_data.lst_rationales[index] or "").strip()
            notation = (excel_data.lst_notations[index] or "").strip()
            kana = (excel_data.lst_kana[index] or "").strip()
            logo = (excel_data.lst_logos[index] or "").strip()
            name_sub_group = (excel_data.lst_name_sub_groups[index] or "").strip()

            is_group_marker = marker in GROUP_MARKERS
            has_delimiter = "##" in name or "$$" in name
            is_grouped_slide = bool(name_sub_group)
            
            if category and (not name and not rationale and not notation and not kana and not logo):
                logger.info(
                    "🖼️ Adding category header image slide: Index=%s, Category='%s', Name=%s, Rationale=%s, Notation=%s, Kana=%s",
                    index,
                    category,
                    name,
                    rationale,
                    notation,
                    kana
                )
                image_path = pptx_builder_service.generate_category_slide_image(
                    category=category,
                    display_name=request.display_name,
                    slide_number=slide_number,
                    project_type=request.project_type,
                )
                details.append(DetailItem(
                    slide_number=slide_number,
                    slide_type="Image",
                    slide_bg_file_name=image_path or "",
                    slide_description=category,
                    group_name=last_group_name,
                    category=category,
                    name=name,
                    rationale=excel_data.lst_rationales[index],
                    notation=excel_data.lst_notations[index],
                    kana=excel_data.lst_kana[index],
                    logo_filename=excel_data.lst_logos[index],
                    template_id=0,
                    name_sub_group=name_sub_group,
                ))
                slide_number += 1

            # 🟢 Group header slide ("A", "B", "C", etc.)
            elif is_group_marker:
                logger.info(f"🟢 Adding group header slide: Index={index}, Marker='{marker}'")
                detail = self._create_group_slide(
                    excel_data=excel_data,
                    index=index,
                    slide_number=slide_number,
                    request=request,
                    default_template_id=default_template_id,
                )
                last_group_name = detail.group_name or last_group_name
                details.append(detail)
                slide_number += 1

            # 🟡 Grouped slide (has a subgroup or delimiter)
            elif is_grouped_slide or has_delimiter:
                logger.info(f"🟡 Adding grouped names slide: Index={index}, SubGroup='{name_sub_group}'")
                detail = self._create_individual_slide(
                    excel_data=excel_data,
                    index=index,
                    slide_number=slide_number,
                    request=request,
                    current_group=last_group_name,
                    default_template_id=default_template_id,
                )
                details.append(detail)
                slide_number += 1

            # 🔵 Individual slide
            else:
                logger.info(f"🔵 Adding individual name slide: Index={index}, Name='{name}'")
                detail = self._create_individual_slide(
                    excel_data=excel_data,
                    index=index,
                    slide_number=slide_number,
                    request=request,
                    current_group=last_group_name,
                    default_template_id=default_template_id,
                )
                details.append(detail)
                slide_number += 1

    # 3) Summary slide for NW/DW
        if request.project_type.lower() in {"nw", "dw"}:
            summary_slide = self._create_summary_slide(
                slide_number=slide_number,
                last_group=last_group_name,
                request=request,
                default_template_id=default_template_id,
            )
            details.append(summary_slide)
            slide_number += 1

    # 4) Remaining PPTX slides (after the Excel slides)
        if ppt_images and prefix_count < len(ppt_images):
            for tail_idx, image_path in enumerate(ppt_images[prefix_count:], start=1):
                description = Path(image_path).stem.replace("_", " ") if image_path else f"Slide {prefix_count + tail_idx}"
                details.append(DetailItem(
                    slide_number=slide_number,
                    slide_type="Image",
                    slide_bg_file_name=Path(image_path).as_posix() if image_path else "",
                    slide_description=description,
                    group_name="",
                    category="",
                    name="",
                    rationale="",
                    notation="",
                    kana="",
                    logo_filename="",
                    template_id=0,
                    name_sub_group=""
                ))
                slide_number += 1

    # Background and template info
        background_type = request.background_type or "Default"
        background_name = request.background_name or "Default"

        powerpoint_file = pptx_data.pptx_file or ""
        excel_file = ""  # Will be populated if needed in the future

        return PresentationData(
            project=request.project,
            display_name=request.display_name,
            powerpoint_file=powerpoint_file,
            excel_file=excel_file,
            background_type=background_type,
            background_name=background_name,
            page_number=request.page_number,
            presentation_type=request.presentation_type,
            user_name=request.user_name,
            mobile_link_bsr=request.mobile_link_bsr,
            participant_vote=request.participant_vote,
            is_wide_ppt=request.is_wide_ppt,
            is_aws_email=request.is_aws_email,
            details=details,
        ).model_dump()

    
    def _create_group_slide(
        self,
        excel_data: ProcessedExcelData,
        index: int,
        slide_number: int,
        request: CreatePresentationRequest,
        default_template_id: int,
    ) -> DetailItem:
        """Create a group slide (A-Z slide types) with improved metadata handling."""
        slide_type = self._get_slide_type_for_group(
            request.project_type, request.presentation_type
        )
        group_name = excel_data.lst_categories[index] or excel_data.lst_names[index]

        return DetailItem(
            slide_number=slide_number,
            slide_type=slide_type,
            slide_bg_file_name=self._get_default_group_background(
                request, slide_number
            ),
            slide_description=group_name,
            group_name=group_name,
            category=excel_data.lst_categories[index],
            name=group_name,
            rationale="",
            notation="",
            kana="",
            logo_filename="",
            template_id=default_template_id,
            name_sub_group="",
        )
        
    def _create_individual_slide(
        self,
        excel_data: ProcessedExcelData,
        index: int,
        slide_number: int,
        request: CreatePresentationRequest,
        current_group: str,
        default_template_id: int,
    ) -> DetailItem:
        """Create an individual name evaluation slide while tracking the current group."""
        slide_type = self._get_slide_type_for_individual(
            request.project_type, request.presentation_type
        )
        group_name = current_group or excel_data.lst_categories[index]

        return DetailItem(
            slide_number=slide_number,
            slide_type=slide_type,
            slide_bg_file_name="",  # Will be assigned with template rotation
            slide_description=f"{excel_data.lst_names[index]} - Name Evaluation",
            group_name=group_name,
            category=excel_data.lst_categories[index],
            name=excel_data.lst_names[index],
            rationale=excel_data.lst_rationales[index],
            notation=excel_data.lst_notations[index],
            kana=excel_data.lst_kana[index],
            logo_filename=excel_data.lst_logos[index],
            template_id=default_template_id,  # May be updated with template rotation
            name_sub_group=excel_data.lst_name_sub_groups[index],
        )

    def _create_summary_slide(
        self,
        slide_number: int,
        last_group: str,
        request: CreatePresentationRequest,
        default_template_id: int,
    ) -> DetailItem:
        """Create a summary slide for NW/DW projects."""
        return DetailItem(
            slide_number=slide_number,
            slide_type="NameSummary",
            slide_bg_file_name=self._get_default_group_background(
                request, slide_number
            ),
            slide_description="Brainstorm",
            group_name=last_group,
            category="",
            name="",
            rationale="",
            notation="",
            kana="",
            logo_filename="",
            template_id=default_template_id,
            name_sub_group="",
        )

    def _create_existing_ppt_slides(
        self,
        pptx_data: PPTXConversionResponse,
        *,
        max_slides: Optional[int] = None,
        start_slide_number: int = 0,
    ) -> Tuple[List[Dict[str, Any]], int]:
        """Create slides from existing PPT images.

        Returns a tuple of (slides, last_slide_number).
        """
        slides: List[Dict[str, Any]] = []
        slide_number = start_slide_number

        for index, image_path in enumerate(pptx_data.images, start=1):
            if max_slides is not None and len(slides) >= max_slides:
                break

            slide_number += 1

            slide_description = f"Slide {index}"
            slide_source = ""

            if image_path:
                slide_source = image_path

                if "://" in image_path:
                    parsed = urlparse(image_path)
                    filename = Path(parsed.path).name
                    if filename:
                        slide_description = Path(filename).stem.replace("_", " ")
                else:
                    slide_description = Path(image_path).stem.replace("_", " ")
                    slide_source = Path(image_path).as_posix()

            slides.append(
                {
                    "slide_number": slide_number,
                    "slide_type": "Image",
                    "slide_bg_file_name": slide_source,
                    "slide_description": slide_description,
                    "group_name": "",
                    "category": "",
                    "name": "",
                    "rationale": "",
                    "notation": "",
                    "kana": "",
                    "logo_filename": "",
                    "template_id": 0,
                    "name_sub_group": "",
                }
            )

        return slides, slide_number

    def _get_default_group_background(
        self, request: CreatePresentationRequest, slide_number: int
    ) -> str:
        """Get the default background image path for group slides."""
        project_folder = request.display_name if request.display_name else request.project
        base_folder = (
            "BRS_slides"
            if request.project_type.lower() in {"bsr", "nsr"}
            else "nw_slides"
        )
        return f"{base_folder}/{project_folder}/{slide_number:03d}.jpg"

    def _get_slide_type_for_group(self, project_type: str, presentation_type: str) -> str:
        """Get the slide type for group slides based on the project type."""
        if presentation_type == "Normal-NoNeutral":
            return "NameEvaluation_noNeutral"
        if project_type.lower() == "bsr":
            return "BSR_Group"
        if project_type.lower() == "nsr":
            return "NSR_Group"
        return "NameEvaluation"

    def _get_slide_type_for_individual(
        self, project_type: str, presentation_type: str
    ) -> str:
        """Get the slide type for individual slides based on the project type."""
        if presentation_type == "Normal-NoNeutral":
            return "NameEvaluation_noNeutral"
        if project_type.lower() == "bsr":
            return "BSR-Japan" if "Japan" in presentation_type else "BSR"
        if project_type.lower() == "nsr":
            return "NSR-Japan" if "Japan" in presentation_type else "NSR"
        return "NameEvaluation"

    def _get_template_metadata(self, template_name: str) -> Dict[str, Any]:
        """Get template metadata from the database with robust error handling and caching.

        Optimized with in-memory cache (5 min TTL) to reduce DB queries.
        """
        global _template_metadata_cache, _cache_timestamp

        # Check cache validity
        current_time = time.time()
        if _cache_timestamp is None or (current_time - _cache_timestamp) > _CACHE_TTL_SECONDS:
            # Cache expired, clear it
            _template_metadata_cache.clear()
            _cache_timestamp = current_time
            logger.debug("Template metadata cache expired and cleared")

        # Check if template metadata is in cache
        if template_name in _template_metadata_cache:
            logger.debug("Template metadata cache hit for '%s'", template_name)
            return _template_metadata_cache[template_name].copy()

        # Cache miss - fetch from database
        metadata = {"template_id": 5, "background": template_name}

        conn = None
        cursor = None
        try:
            conn = create_connection()
            cursor = conn.cursor()

            sql = (
                "SELECT TOP 1 templateid "
                "FROM [BI_GUIDELINES].dbo.nw_Templates "
                "WHERE TemplateName = ?"
            )

            cursor.execute(sql, (template_name,))
            row = cursor.fetchone()

            if row and row[0]:
                metadata["template_id"] = int(row[0])
            else:
                logger.warning(
                    "Template '%s' not found in database, using default template",
                    template_name,
                )

        except DatabaseConnectionError as exc:
            logger.warning(
                "Could not connect to database for template lookup (%s), using default template",
                exc,
            )
        except Exception as exc:  # pylint: disable=broad-except
            logger.error(
                "Error getting template metadata for '%s': %s", template_name, str(exc)
            )
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

        # Store in cache
        _template_metadata_cache[template_name] = metadata.copy()
        logger.debug("Template metadata cached for '%s'", template_name)

        return metadata

    def _generate_feedback_document(
        self, presentation_id: int, request: CreatePresentationRequest, excel_data: ProcessedExcelData
    ) -> str:
        """Generate a feedback template document when the Word service is available."""
        try:
            return generate_feedback_document(
                presentation_id=presentation_id,
                project_type=request.project_type,
                presentation_type=request.presentation_type,
                excel_data=excel_data,
                display_name=request.display_name,
                user_name=request.user_name,
            )
        except Exception as e:
            logger.error("Error generating feedback document: %s", str(e))
            return ""

    def _send_notification_emails(self, presentation_id: int, request: CreatePresentationRequest):
        """Send notification emails when the Email service is available."""
        try:
            send_presentation_emails(
                presentation_id=presentation_id,
                presentation_type=request.presentation_type,
                project_type=request.project_type,
                display_name=request.display_name,
                user_name=request.user_name,
                participant_vote=bool(request.participant_vote),
                is_wide_ppt=bool(request.is_wide_ppt),
            )
        except Exception as e:
            logger.error("Error sending notification emails: %s", str(e))
            # Don't raise - emails are not critical for the main flow
    def _create_presentation_in_db(
        self, slides_data: Dict[str, Any], request: CreatePresentationRequest
    ) -> Dict[str, Any]:
        """Create presentation in database using BI Guidelines functionality directly."""
        # Validate supported project types (placeholder for future stricter rules)
        if request.project_type.lower() not in {"nw", "dw"}:
            logger.warning(
                "Project type '%s' may not be fully supported", request.project_type
            )

        conn = None
        cursor = None
        try:
            try:
                conn = create_connection()
            except DatabaseConnectionError as exc:
                logger.error("Could not connect to database: %s", exc)
                raise HTTPException(
                    status_code=500,
                    detail="Could not connect to database.",
                ) from exc

            cursor = conn.cursor()
            cursor.execute("SET NOCOUNT ON;")

            # Delete existing presentation before inserting new records
            self._delete_existing_presentation(cursor, slides_data)

            # Insert master record
            presentation_id = self._insert_master_record(cursor, slides_data)

            # Insert detail records
            self._insert_detail_records(cursor, presentation_id, slides_data)

            conn.commit()
            logger.info("Presentation created successfully in database: %s", presentation_id)

            return {
                "message": "Presentation created successfully.",
                "presentation_id": presentation_id,
            }

        except Exception as exc:
            if conn:
                conn.rollback()
            error_msg = f"Error creating presentation in database: {exc}"
            logger.error(error_msg)
            raise HTTPException(status_code=500, detail=error_msg) from exc
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def _delete_existing_presentation(self, cursor, slides_data: Dict[str, Any]) -> None:
        """Delete existing presentation if it exists."""
        try:
            logger.info(
                "Deleting existing presentation if exists: %s / %s",
                slides_data["project"],
                slides_data["display_name"],
            )
            cursor.execute(
                "EXEC [dbo].[nw_DeletePresentation] ?, ?",
                (slides_data["project"], slides_data["display_name"]),
            )
        except Exception as exc:
            logger.warning("Error deleting existing presentation: %s", exc)
            # Continue - this is not critical

    def _insert_master_record(self, cursor, slides_data: Dict[str, Any]) -> int:
        """Insert the master presentation record with improved handling."""
        master_sql = """
            EXEC [dbo].[nw_InsertPresentationMaster_sep2025]
                @Project=?, @DisplayName=?, @MainPptFileName=?, @NameCandidateFileName=?,
                @NameCandidateBGType=?, @NameCandidateBGName=?, @NameCandidateStartingSlide=?,
                @PresentationType=?, @UploadedBy=?, @BSRDisplayName=?,
                @isParticipantsVote=?, @isWideScreenPPT=?, @isAWSLinkReq=?;
        """
        master_params = (
            slides_data["project"],
            slides_data["display_name"],
            slides_data["powerpoint_file"],
            slides_data["excel_file"],
            slides_data["background_type"],
            slides_data["background_name"],
            slides_data["page_number"],
            slides_data["presentation_type"],
            slides_data["user_name"],
            slides_data["mobile_link_bsr"],
            slides_data["participant_vote"],
            slides_data["is_wide_ppt"],
            slides_data["is_aws_email"],
        )

        logger.info("Executing master stored procedure...")
        cursor.execute(master_sql, master_params)

        presentation_id = None
        while True:
            try:
                row = cursor.fetchone()
                if row:
                    presentation_id = int(row[0])
                    logger.info("PresentationId created: %s", presentation_id)
                    break
            except Exception as exc:  # pylint: disable=broad-except
                message = str(exc)
                if "No results" in message and "Previous SQL" in message:
                    logger.debug(
                        "Stored procedure returned no result set for presentation ID; applying fallback lookup."
                    )
                else:
                    logger.error("Error fetching presentation ID: %s", message)
            if not cursor.nextset():
                break

        if not presentation_id:
            presentation_id = self._lookup_presentation_id(cursor, slides_data)

        if not presentation_id:
            raise HTTPException(
                status_code=500,
                detail="Could not get PresentationId from the master record.",
            )

        return presentation_id

    def _lookup_presentation_id(
        self, cursor, slides_data: Dict[str, Any]
    ) -> Optional[int]:
        """Fallback lookup for presentation ID when stored procedure returns no result set."""
        try:
            lookup_sql = (
                "SELECT TOP 1 PresentationId "
                "FROM [BI_GUIDELINES].[dbo].[nw_Master] "
                "WHERE Project = ? AND DisplayName = ? "
                "ORDER BY PresentationId DESC"
            )
            cursor.execute(lookup_sql, (slides_data["project"], slides_data["display_name"]))
            row = cursor.fetchone()
            if row and row[0]:
                presentation_id = int(row[0])
                logger.info(
                    "PresentationId resolved via fallback lookup: %s", presentation_id
                )
                return presentation_id
        except Exception as exc:  # pylint: disable=broad-except
            logger.warning(
                "Fallback presentation ID lookup failed for %s/%s: %s",
                slides_data.get("project"),
                slides_data.get("display_name"),
                exc,
            )
        return None

    def _insert_detail_records(
        self, cursor, presentation_id: int, slides_data: Dict[str, Any]
    ) -> None:
        """Insert detail records for each slide while handling null values."""
        detail_sql = """
            EXEC [dbo].[nw_InsertPresentationDetail_copy]
                @PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?,
                @SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?,
                @NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?,
                @TemplateId=?, @NameSubGroup=?;
        """
        
        for slide in slides_data["details"]:
            detail_params = (
                presentation_id,
                slide["slide_number"],
                slide["slide_type"],
                slide.get("slide_bg_file_name") or "",
                slide.get("slide_description") or "",
                slide.get("group_name") or "",
                slide.get("category") or "",
                slide.get("name") or "",
                slide.get("rationale") or "",
                slide.get("notation") or "",
                slide.get("kana") or "",
                slide.get("logo_filename") or "",
                slide.get("template_id", 0),
                slide.get("name_sub_group") or "",
            )
            cursor.execute(detail_sql, detail_params)


# Global presentation service instance
presentation_service = PresentationService()