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

    def create_bsr_presentation(
        self,
        *,
        project_name: str,
        display_name: str,
        slide_number: int,
        presentation_type: str,
        user_name: str,
        is_wide_ppt: int,
        pptx_content: bytes,
        pptx_filename: str,
        categories: Optional[Dict[str, Any]] = None,
        progress_callback=None,
    ) -> Dict[str, Any]:
        """
        Create a BSR (Board Sales Request) presentation.

        Args:
            project_name: Name of the project
            display_name: Display name for the presentation
            slide_number: Slide number where summary slide will be inserted
            presentation_type: Type (BSR or BSR-Japan)
            user_name: User creating the presentation
            is_wide_ppt: 1 for 16:9, 0 for 4:3
            pptx_content: PowerPoint file content
            pptx_filename: PowerPoint filename
            categories: Optional categories configuration with add_categories flag
            progress_callback: Optional callback function to report progress (int 0-100)

        Returns:
            Dict with presentation_id, total_slides, and categories_added

        Raises:
            HTTPException: If validation fails or creation errors occur
        """
        try:
            # Extract just the folder name if display_name contains path separators
            # This prevents path duplication issues when constructing full paths
            from pathlib import Path as PathLib
            clean_display_name = PathLib(display_name).name if "/" in display_name or "\\" in display_name else display_name

            logger.info(
                "Creating BSR presentation: project=%s, display=%s, slide_number=%d",
                project_name,
                clean_display_name,
                slide_number,
            )

            # 1. Validate presentation doesn't exist (30%)
            if progress_callback:
                progress_callback(30)
            self._validate_bsr_presentation_not_exists(project_name, clean_display_name)

            # 2. Validate display name hasn't been used (35%)
            if progress_callback:
                progress_callback(35)
            self._validate_bsr_display_name_not_used(clean_display_name)

            # 3. Convert PowerPoint to images (40-55%)
            if progress_callback:
                progress_callback(40)
            logger.info("Converting PowerPoint to images")
            # For BSR we store assets under the bipresents (bsr_slides) root
            pptx_data = pptx_service.convert_pptx_to_images(
                file_content=pptx_content,
                filename=pptx_filename,
                display_name=clean_display_name,  # Use display_name for BSR folder structure
                project_type="bipresents",  # Ensure URLs resolve to bsr_slides
            )
            logger.info("PPTX Data - Images: %s", pptx_data.get("images"))
            logger.info("PPTX Data - Thumbnails: %s", pptx_data.get("thumbnails"))

            # 4. Extract slide titles from PowerPoint (55-60%)
            if progress_callback:
                progress_callback(55)
            logger.info("Extracting slide titles")
            slide_titles = self._extract_slide_titles(pptx_content, pptx_filename)

            # 5. Insert presentation master record (60-70%)
            if progress_callback:
                progress_callback(60)
            logger.info("Inserting presentation master record")
            presentation_id = self._insert_bsr_master_record(
                project_name=project_name,
                display_name=clean_display_name,
                pptx_filename=pptx_filename,
                slide_number=slide_number,
                presentation_type=presentation_type,
                user_name=user_name,
                is_wide_ppt=is_wide_ppt,
            )

            # 6. Insert presentation detail records (70-80%)
            if progress_callback:
                progress_callback(70)
            logger.info("Inserting presentation detail records")
            total_slides = self._insert_bsr_detail_records(
                presentation_id=presentation_id,
                display_name=clean_display_name,  # Use display_name for image paths
                slide_number=slide_number,
                slide_titles=slide_titles,
                pptx_data=pptx_data,
            )

            # 7. Process categories if provided (80-90%)
            if progress_callback:
                progress_callback(80)
            categories_created = []
            if categories and categories.get("add_categories", False):
                try:
                    logger.info("Processing BSR categories for presentation ID=%d", presentation_id)
                    categories_created = self._process_bsr_categories(
                        presentation_id=presentation_id,
                        categories_data=categories
                    )
                    logger.info("Successfully added %d categories", len(categories_created))
                except Exception as e:
                    # Log error but don't fail the entire process since presentation is already created
                    logger.error(
                        "Error processing categories for presentation %d: %s",
                        presentation_id,
                        str(e),
                        exc_info=True
                    )
                    # Optionally re-raise if you want categories to be mandatory
                    # raise HTTPException(status_code=500, detail=f"Categories error: {str(e)}")

            if progress_callback:
                progress_callback(90)

            logger.info(
                "BSR presentation created successfully: ID=%d, Total Slides=%d, Categories=%d",
                presentation_id,
                total_slides,
                len(categories_created),
            )

            return {
                "presentation_id": presentation_id,
                "total_slides": total_slides,
                "categories_added": len(categories_created),
            }

        except HTTPException:
            raise
        except Exception as e:
            logger.error("Error creating BSR presentation: %s", str(e), exc_info=True)
            raise HTTPException(
                status_code=500,
                detail=f"Failed to create BSR presentation: {str(e)}"
            ) from e

    def create_presentation(
        self, request: CreatePresentationRequest, progress_callback=None
    ) -> CreatePresentationResponse:
        """
        Orchestrate the complete presentation creation process.

        Args:
            request: Complete presentation creation request
            progress_callback: Optional callback function to report progress (int 0-100)

        Returns:
            CreatePresentationResponse: Result of the presentation creation
        """
        start_time = time.time()

        try:
            # Clean display_name if it contains path separators
            # This prevents path duplication when the frontend sends full paths
            from pathlib import Path as PathLib
            if "/" in request.display_name or "\\" in request.display_name:
                original_display_name = request.display_name
                request.display_name = PathLib(request.display_name).name
                logger.info(
                    "Cleaned display_name: '%s' -> '%s'",
                    original_display_name,
                    request.display_name
                )

            # 1. Process Excel file (30-40%)
            if progress_callback:
                progress_callback(30)
            logger.info("Processing Excel file: %s", request.excel_filename)
            excel_data = self._process_excel_file(request)

            # 2. Convert PPTX file (40-50%)
            if progress_callback:
                progress_callback(40)
            logger.info("Converting PPTX file: %s", request.pptx_filename)
            pptx_data = self._convert_pptx_file(request)

            # 3. Save original Excel file (50-55%)
            if progress_callback:
                progress_callback(50)
            logger.info("Saving original Excel file")
            excel_relative_path = self._save_original_excel(request)

            # 4. Generate slides from Excel arrays (55-70%)
            if progress_callback:
                progress_callback(55)
            logger.info("Generating slides from Excel data")
            slides_data = self._generate_slides_from_excel(
                excel_data=excel_data,
                pptx_data=pptx_data,
                request=request,
            )

            # Store Excel path in slides_data for database persistence
            slides_data["excel_file"] = excel_relative_path

            logger.info(
                "Slides data generated: %d detail items, background=%s, excel_file=%s",
                len(slides_data.get("details", [])),
                slides_data.get("background_name"),
                slides_data.get("excel_file"),
            )
            
            # 5. Generate physical PowerPoint file (70-85%)
            if progress_callback:
                progress_callback(70)
            generated_files = None
            if request.create_backup == 1:
                logger.info("Generating physical PowerPoint backup file")
                try:
                    ppt_files = self._generate_physical_powerpoint(
                        slides_data=slides_data,
                        excel_data=excel_data,
                        request=request,
                        pptx_data=pptx_data,
                    )
                    generated_files = ppt_files

                    # Update slides_data with generated file path for DB storage
                    if ppt_files.get("printable_path"):
                        slides_data["powerpoint_file"] = ppt_files["printable_path"]

                    logger.info(
                        "Physical PowerPoint backup generated: %s (total slides: %s)",
                        ppt_files.get("printable_path", "N/A"),
                        ppt_files.get("total_slides", "0"),
                    )

                except Exception as ppt_error:
                    logger.error("Failed to generate physical PowerPoint backup: %s", str(ppt_error))
                    # Continue execution - physical file generation is optional
                    generated_files = {
                        "error": str(ppt_error),
                        "printable_path": "",
                        "macro_path": "",
                        "total_slides": "0",
                        "warnings": "Failed to generate",
                    }
            else:
                logger.info("Physical PowerPoint backup generation skipped (create_backup=0)")

            # 6. Create presentation in DB (85-95%)
            if progress_callback:
                progress_callback(85)
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

            if progress_callback:
                progress_callback(95)

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
        progress_callback=None,
    ):
        """Generate PowerPoint deliverables directly from Excel data.
        
        Args:
            build_request: Complete presentation creation request
            options: Build options for the presentation
            progress_callback: Optional callback function to report progress (int 0-100)
        """

        logger.info(
            "Assembling PPT files for project=%s display=%s slide_range=%s-%s",
            build_request.project,
            build_request.display_name,
            options.slide_start,
            options.slide_end,
        )

        # Process Excel (30-50%)
        if progress_callback:
            progress_callback(30)
        excel_data = self._process_excel_file(build_request)

        if progress_callback:
            progress_callback(50)

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

        # Generate slides (50-70%)
        if progress_callback:
            progress_callback(60)
        slides_dict = self._generate_slides_from_excel(
            excel_data=excel_data,
            pptx_data=stub_conversion,
            request=build_request,
        )

        details = slides_dict.get("details", [])

        if progress_callback:
            progress_callback(70)

        # Compose presentation (70-90%)
        return pptx_builder_service.compose_presentation(
            request=build_request,
            options=options,
            details=details,
            excel_data=excel_data,
            progress_callback=progress_callback,
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

    def generate_backup_presentation(
        self, presentation_id: int, progress_callback=None
    ) -> Dict[str, Any]:
        """
        Generate a backup PowerPoint presentation from saved Excel and PPTX files.

        This method retrieves a previously created presentation from the database,
        reads the saved original Excel and PowerPoint files, and regenerates the
        complete PowerPoint backup file.

        Args:
            presentation_id: The ID of the presentation to generate backup for
            progress_callback: Optional callback function to report progress (int 0-100)

        Returns:
            Dict containing:
                - presentation_id: ID of the presentation
                - printable_path: Path to generated .pptx file
                - macro_path: Path to generated .pptm file (if applicable)
                - total_slides: Total number of slides in the presentation
                - warnings: List of any warnings during generation

        Raises:
            HTTPException: If presentation not found, files missing, or generation fails
        """
        try:
            logger.info("Generating backup for presentation ID: %d", presentation_id)

            # 1. Retrieve presentation information from database (30-40%)
            if progress_callback:
                progress_callback(30)
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(
                        """
                        SELECT
                            Project,
                            DisplayName,
                            MainPptFileName,
                            NameCandidateFileName,
                            NameCandidateBGType,
                            NameCandidateBGName,
                            NameCandidateStartingSlide,
                            PresentationType,
                            UploadedBy,
                            BSRDisplayName,
                            isParticipantsVote,
                            isWideScreenPPT,
                            isAWSLinkReq
                        FROM [BI_GUIDELINES].[dbo].[nw_Master]
                        WHERE PresentationId = ?
                        """,
                        (presentation_id,)
                    )

                    row = cursor.fetchone()

                    if not row:
                        raise HTTPException(
                            status_code=404,
                            detail=f"Presentation with ID {presentation_id} not found"
                        )

                    # Extract presentation metadata
                    project = row[0]
                    display_name = row[1]
                    main_ppt_filename = row[2]
                    excel_filename = row[3]
                    background_type = row[4] or "Default"
                    background_name = row[5] or "Default"
                    page_number = row[6] or 1
                    presentation_type = row[7] or "Normal"
                    user_name = row[8] or ""
                    mobile_link_bsr = row[9] or ""
                    participant_vote = row[10] or 0
                    is_wide_ppt = row[11] or 0
                    is_aws_email = row[12] or 0

                    logger.info(
                        "Retrieved presentation info: project=%s, display=%s, type=%s, excel_file=%s, ppt_file=%s",
                        project,
                        display_name,
                        presentation_type,
                        excel_filename,
                        main_ppt_filename
                    )

            # 2. Resolve file paths
            # Extract just the folder name if display_name contains path separators
            # This prevents path duplication issues when the display_name was stored with a path
            from pathlib import Path as PathLib
            clean_display_name = PathLib(display_name).name if "/" in display_name or "\\" in display_name else display_name

            output_base, _ = resolve_project_output(
                clean_display_name,
                "NW",  # Default to NW project type
                fallback_subdir="generated_presentations",
            )

            # Load saved Excel file
            # Use the filename from database, or fallback to original_data.xlsx for backwards compatibility
            excel_filename_to_use = excel_filename if excel_filename else "original_data.xlsx"
            excel_path = output_base / excel_filename_to_use

            logger.info(
                "Looking for Excel file: excel_filename_from_db=%s, path=%s",
                excel_filename,
                excel_path
            )

            if not excel_path.exists():
                # Try fallback for old presentations
                excel_path_fallback = output_base / "original_data.xlsx"
                if excel_path_fallback.exists():
                    excel_path = excel_path_fallback
                    logger.info("Using fallback Excel filename: original_data.xlsx")
                else:
                    # List all Excel files in the directory to help diagnose the issue
                    excel_files = list(output_base.glob("*.xlsx"))
                    logger.error(
                        "Excel file not found. Expected: %s, Available files: %s",
                        excel_path,
                        [f.name for f in excel_files]
                    )
                    raise HTTPException(
                        status_code=404,
                        detail=f"Original Excel file not found: {excel_path}. Available files: {[f.name for f in excel_files]}"
                    )

            # Load saved PowerPoint template file
            pptx_path = output_base / main_ppt_filename if main_ppt_filename else None
            if not pptx_path or not pptx_path.exists():
                # Try alternate naming
                pptx_files = list(output_base.glob("*.pptx"))
                if pptx_files:
                    pptx_path = pptx_files[0]
                else:
                    raise HTTPException(
                        status_code=404,
                        detail=f"Original PowerPoint file not found in: {output_base}"
                    )

            logger.info("Found Excel file: %s", excel_path)
            logger.info("Found PPTX file: %s", pptx_path)

            # 3. Read file contents (40-50%)
            if progress_callback:
                progress_callback(40)
            with open(excel_path, "rb") as f:
                excel_content = f.read()

            with open(pptx_path, "rb") as f:
                pptx_content = f.read()

            # 4. Create request object from database metadata (50-55%)
            if progress_callback:
                progress_callback(50)
            request = CreatePresentationRequest(
                project=project,
                display_name=display_name,
                background_type=background_type,
                background_name=background_name,
                page_number=page_number,
                presentation_type=presentation_type,
                user_name=user_name,
                mobile_link_bsr=mobile_link_bsr,
                participant_vote=participant_vote,
                is_wide_ppt=is_wide_ppt,
                is_aws_email=is_aws_email,
                excel_file=excel_content,
                excel_filename=excel_path.name,
                pptx_file=pptx_content,
                pptx_filename=pptx_path.name,
                create_backup=1,  # Always generate backup for this endpoint
                has_groups=True,  # Default assumption
                test_name_order="Default",  # Use Default for backup generation
                project_type="NW",  # Default to NW
            )

            # 5. Process Excel file (55-65%)
            if progress_callback:
                progress_callback(55)
            logger.info("Processing Excel file for backup generation")
            excel_data = self._process_excel_file(request)

            # 6. Load existing PPTX image data (65-70%)
            if progress_callback:
                progress_callback(65)
            # (reconverting would delete the entire folder including the Excel file)
            logger.info("Loading existing PPTX images for backup generation")
            pptx_data = self._load_existing_pptx_images(output_base, pptx_path.name)

            # 7. Generate slides data (70-75%)
            if progress_callback:
                progress_callback(70)
            logger.info("Generating slides data for backup")
            slides_data = self._generate_slides_from_excel(
                excel_data=excel_data,
                pptx_data=pptx_data,
                request=request,
            )

            # 8. Generate physical PowerPoint backup (75-90%)
            if progress_callback:
                progress_callback(75)
            logger.info("Generating physical PowerPoint backup")
            ppt_files = self._generate_physical_powerpoint(
                slides_data=slides_data,
                excel_data=excel_data,
                request=request,
                pptx_data=pptx_data,
            )

            logger.info(
                "Backup generation completed: %s (total slides: %s)",
                ppt_files.get("printable_path", "N/A"),
                ppt_files.get("total_slides", "0"),
            )

            # 9. Update the database with the generated PowerPoint filename (90-95%)
            if progress_callback:
                progress_callback(90)
            printable_path = ppt_files.get("printable_path", "")
            if printable_path:
                from pathlib import Path as PathLib
                generated_filename = PathLib(printable_path).name

                logger.info(
                    "Updating database with generated PowerPoint filename: %s",
                    generated_filename
                )

                with create_connection() as conn:
                    with conn.cursor() as cursor:
                        cursor.execute(
                            """
                            UPDATE [BI_GUIDELINES].[dbo].[nw_Master]
                            SET MainPptFileName = ?
                            WHERE PresentationId = ?
                            """,
                            (generated_filename, presentation_id)
                        )
                        conn.commit()
                        logger.info("Database updated successfully with new PowerPoint filename")

            return {
                "presentation_id": presentation_id,
                "printable_path": ppt_files.get("printable_path", ""),
                "macro_path": ppt_files.get("macro_path", ""),
                "total_slides": ppt_files.get("total_slides", "0"),
                "warnings": ppt_files.get("warnings", "None"),
            }

        except HTTPException:
            raise
        except Exception as e:
            logger.error("Error generating backup for presentation %d: %s", presentation_id, str(e), exc_info=True)
            raise HTTPException(
                status_code=500,
                detail=f"Failed to generate backup: {str(e)}"
            ) from e

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

    def _save_original_excel(
        self, request: CreatePresentationRequest
    ) -> str:
        """Save the original Excel file to the project folder for future backup generation.

        Args:
            request: The presentation creation request containing the Excel file

        Returns:
            str: Relative path to the saved Excel file
        """
        try:
            # Resolve project output folder
            output_base, _ = resolve_project_output(
                request.display_name,
                request.project_type,
                fallback_subdir="generated_presentations",
            )

            # Ensure the directory exists
            output_base.mkdir(parents=True, exist_ok=True)

            # Save Excel file with original filename (for backup generation)
            # Use sanitized filename to prevent path traversal
            safe_filename = Path(request.excel_filename).name  # Get just the filename, no path
            excel_path = output_base / safe_filename

            logger.info(
                "🔵 Saving Excel file: display_name=%s, output_base=%s, filename=%s, full_path=%s",
                request.display_name,
                output_base,
                safe_filename,
                excel_path
            )
            logger.info("🔵 Excel file size: %d bytes", len(request.excel_file))
            logger.info("🔵 Directory exists: %s", output_base.exists())
            logger.info("🔵 Directory is writable: %s", output_base.is_dir())

            with open(excel_path, "wb") as f:
                bytes_written = f.write(request.excel_file)
                logger.info("🔵 Bytes written to file: %d", bytes_written)

            logger.info("✅ Original Excel file saved successfully to: %s", excel_path)
            logger.info("✅ File exists after save: %s", excel_path.exists())
            if excel_path.exists():
                logger.info("✅ File size on disk: %d bytes", excel_path.stat().st_size)

            # Return relative path for database storage (just the filename)
            return safe_filename

        except Exception as e:
            logger.error("❌ Error saving original Excel file: %s", str(e))
            import traceback
            logger.error("❌ Traceback: %s", traceback.format_exc())
            raise

    def _load_existing_pptx_images(
        self, output_base: Path, pptx_filename: str
    ) -> PPTXConversionResponse:
        """Load existing PPTX image data from disk without reconverting.

        This is used for backup generation to avoid deleting the folder.

        Args:
            output_base: Base directory where images are stored
            pptx_filename: Name of the original PPTX file

        Returns:
            PPTXConversionResponse with existing image paths
        """
        from app.utils.path_utils import get_relative_slide_root

        try:
            # Get list of existing image files
            image_files = sorted(output_base.glob("*.jpg"))

            # Get relative URL root for proper URL construction
            url_root = get_relative_slide_root("NW")

            # Build image URLs
            images = []
            for img_file in image_files:
                if url_root:
                    # Create web-accessible URL
                    image_url = f"{url_root}/{output_base.name}/{img_file.name}"
                else:
                    # Fallback to file path
                    image_url = str(img_file.as_posix())
                images.append(image_url)

            logger.info("Loaded %d existing PPTX images from %s", len(images), output_base)

            return PPTXConversionResponse(
                message="Existing images loaded successfully",
                conversion_id=output_base.name,  # Use folder name as conversion ID
                project_type="NW",
                images=images,
                thumbnails=[],  # Not needed for backup
                titles=[""] * len(images),  # Empty titles for existing images
                total_images=len(images),
                pptx_file=pptx_filename
            )

        except Exception as e:
            logger.error("Error loading existing PPTX images: %s", str(e))
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
            # Build presentation build options with default templates
            build_options = PresentationBuildOptions(
                template_pack="BackgroundDefaultTemplate",
                base_template="template_default_2019.pptx",
                multi_template="template_default_withgroups2019.pptx",
                group_template="template_default_withgroup_2019.pptx",
                separator_template="template_default_seperator_2019.pptx",
                summary_template="template_default_summary2019.pptx",
                slide_start=1,
                slide_end=None,
                include_print_ready_version=True,
                include_macro_version=False,  # Default: no macro version
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
            slide_description=excel_data.lst_names[index],
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
        logger.info(
            "Inserting master record with excel_file=%s, powerpoint_file=%s",
            slides_data.get("excel_file"),
            slides_data.get("powerpoint_file")
        )

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
        """Insert detail records for each slide with background rotation for NameEvaluation slides.

        This method applies background rotation ONLY for NameEvaluation type presentations.
        Background paths are fetched from nw_Templates table and rotated through slides.
        """
        from app.services.template_selector import background_image_selector

        # 1. Determine if we should apply background rotation
        background_type = slides_data.get("background_type", "Default")
        background_name = slides_data.get("background_name", "")

        use_background_rotation = (
            background_type.lower() == "rotate"
            and background_name
            and background_name.lower() != "default"
        )

        # 2. Fetch background paths from database if rotating
        background_paths = {}
        selected_bg_names = []

        if use_background_rotation:
            selected_bg_names = background_image_selector.parse_background_names(background_name)
            names_to_fetch = set(selected_bg_names)
            names_to_fetch.add("Default")

            # Query nw_Templates table for template IDs and paths
            from app.services.bi_guidelines_service import bi_guidelines_service
            background_paths = bi_guidelines_service.get_background_templates_by_names(
                list(names_to_fetch)
            )

            logger.info(
                "Background rotation enabled: %d templates will be rotated across NameEvaluation slides",
                len(selected_bg_names)
            )

        # 3. Insert slides with background rotation
        detail_sql = """
            EXEC [dbo].[nw_InsertPresentationDetail_copy]
                @PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?,
                @SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?,
                @NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?,
                @TemplateId=?, @NameSubGroup=?;
        """

        bg_index = 0
        for slide in slides_data["details"]:
            slide_type = slide.get("slide_type", "")

            # Determine the background path and template ID for this slide
            path_to_save = slide.get("slide_bg_file_name") or ""
            template_id_to_save = slide.get("template_id", 0)

            # Apply background rotation ONLY for NameEvaluation slides
            if use_background_rotation and slide_type == "NameEvaluation":
                # Rotate through selected backgrounds
                template_name_for_slide = selected_bg_names[bg_index % len(selected_bg_names)]
                template_info = background_paths.get(template_name_for_slide)

                if template_info:
                    path_to_save = template_info.get('template_file_name', "")
                    template_id_to_save = template_info.get('template_id', 0)
                else:
                    # Fallback to Default if template not found
                    default_info = background_paths.get("Default", {})
                    path_to_save = default_info.get('template_file_name', "")
                    template_id_to_save = default_info.get('template_id', 0)
                    logger.warning(
                        "Template '%s' not found, using Default background for slide %d",
                        template_name_for_slide,
                        slide["slide_number"]
                    )

                bg_index += 1

            # Insert the slide with the determined background path and template ID
            detail_params = (
                presentation_id,
                slide["slide_number"],
                slide_type,
                path_to_save,  # Full background image path
                slide.get("slide_description") or "",
                slide.get("group_name") or "",
                slide.get("category") or "",
                slide.get("name") or "",
                slide.get("rationale") or "",
                slide.get("notation") or "",
                slide.get("kana") or "",
                slide.get("logo_filename") or "",
                template_id_to_save,  # Correct template ID for the applied background
                slide.get("name_sub_group") or "",
            )
            cursor.execute(detail_sql, detail_params)

    def _validate_bsr_presentation_not_exists(
        self, project_name: str, display_name: str
    ) -> None:
        """Validate that BSR presentation doesn't already exist."""
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(
                        "EXEC [BI_GUIDELINES].[dbo].[BSR_CheckIfPresentationExists] ?, ?",
                        (project_name, display_name)
                    )
                    result = cursor.fetchone()
                    if result and result[0] > 0:
                        raise HTTPException(
                            status_code=400,
                            detail=f"BSR presentation already exists: {project_name}/{display_name}"
                        )
        except HTTPException:
            raise
        except Exception as e:
            logger.warning("Error checking BSR presentation existence: %s", str(e))

    def _validate_bsr_display_name_not_used(self, display_name: str) -> None:
        """Validate that BSR display name hasn't been used."""
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(
                        "EXEC [BI_GUIDELINES].[dbo].[BSR_CheckIfDisplayNameHasBeenUsed] ?",
                        (display_name,)
                    )
                    result = cursor.fetchone()
                    if result and result[0]:
                        raise HTTPException(
                            status_code=400,
                            detail=f"Display name already used by another project: {result[0]}"
                        )
        except HTTPException:
            raise
        except Exception as e:
            logger.warning("Error checking BSR display name: %s", str(e))

    def _extract_slide_titles(self, pptx_content: bytes, pptx_filename: str) -> List[str]:
        """Extract titles from PowerPoint slides."""
        from io import BytesIO
        from pptx import Presentation

        titles = []
        try:
            prs = Presentation(BytesIO(pptx_content))
            for slide in prs.slides:
                title = ""
                if slide.shapes.title:
                    title = slide.shapes.title.text
                titles.append(title)
            logger.info("Extracted %d slide titles", len(titles))
        except Exception as e:
            logger.error("Error extracting slide titles: %s", str(e))

        return titles

    def _insert_bsr_master_record(
        self,
        *,
        project_name: str,
        display_name: str,
        pptx_filename: str,
        slide_number: int,
        presentation_type: str,
        user_name: str,
        is_wide_ppt: int,
    ) -> int:
        """Insert BSR master presentation record."""
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    # Use positional parameters
                    master_sql = "EXEC [BI_GUIDELINES].[dbo].[bsr_InsertPresentationMaster] ?, ?, ?, ?, ?, ?, ?, ?, ?, ?;"

                    logger.info("Executing BSR master stored procedure...")
                    cursor.execute(
                        master_sql,
                        (
                            project_name,
                            display_name,
                            pptx_filename,
                            "",  # OriginalPPTPath
                            "",  # PPTPath
                            "",  # XLPath
                            slide_number,
                            presentation_type,
                            user_name,
                            is_wide_ppt,
                        )
                    )

                    # Try to get presentation ID from result set
                    presentation_id = None
                    while True:
                        try:
                            row = cursor.fetchone()
                            if row:
                                presentation_id = int(row[0])
                                logger.info("BSR PresentationId created: %s", presentation_id)
                                break
                        except Exception as exc:
                            message = str(exc)
                            if "No results" in message and "Previous SQL" in message:
                                logger.debug(
                                    "BSR stored procedure returned no result set for presentation ID; applying fallback lookup."
                                )
                            else:
                                logger.error("Error fetching BSR presentation ID: %s", message)
                        if not cursor.nextset():
                            break

                    # Fallback: Query for the presentation ID if SP didn't return it
                    if not presentation_id:
                        logger.info("Attempting fallback lookup for BSR presentation ID...")
                        lookup_sql = (
                            "SELECT TOP 1 PresentationId "
                            "FROM [BI_GUIDELINES].[dbo].[bsr_Master] "
                            "WHERE Project = ? AND DisplayName = ? "
                            "ORDER BY PresentationId DESC"
                        )
                        cursor.execute(lookup_sql, (project_name, display_name))
                        row = cursor.fetchone()
                        if row and row[0]:
                            presentation_id = int(row[0])
                            logger.info("BSR PresentationId resolved via fallback lookup: %s", presentation_id)

                    if not presentation_id:
                        raise HTTPException(
                            status_code=500,
                            detail="Could not get PresentationId from BSR master record"
                        )

                    conn.commit()
                    logger.info("BSR master record created with ID: %d", presentation_id)
                    return presentation_id

        except HTTPException:
            raise
        except Exception as e:
            logger.error("Error inserting BSR master record: %s", str(e))
            raise HTTPException(
                status_code=500,
                detail=f"Failed to insert BSR master record: {str(e)}"
            ) from e

    def _insert_bsr_detail_records(
        self,
        *,
        presentation_id: int,
        display_name: str,
        slide_number: int,
        slide_titles: List[str],
        pptx_data: Dict[str, Any],
    ) -> int:
        """Insert BSR detail records for each slide."""
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    print(pptx_data)
                    images = pptx_data.get("images") or []
                    total_slides = len(images)

                    slide_idx = 1

                    # Insert slides BEFORE summary position
                    for idx in range(min(slide_number - 1, len(images))):
                        # Use the actual image path from pptx_data, not thumbnails
                        image_path = images[idx]
                        print(image_path)
                        title = slide_titles[idx] if idx < len(slide_titles) else f"Slide {idx + 1}"

                        cursor.execute(
                            "EXEC [BI_GUIDELINES].[dbo].[bsr_InsertPresentationDetail] ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?;",
                            (
                                presentation_id,
                                slide_idx,
                                "Image",
                                image_path,
                                title,
                                "",  # @Param6
                                "",  # @Param7
                                "",  # @Param8
                                "",  # @Param9
                                "",  # @Param10
                                "",  # @Param11
                                "",  # @Param12
                                0,   # @BGTemplateId
                            )
                        )
                        slide_idx += 1

                    # Insert SUMMARY slide at specified position (slide_number - 1 because 0-indexed)
                    if slide_number - 1 < len(images):
                        summary_image_path = images[slide_number - 1]
                        cursor.execute(
                            "EXEC [BI_GUIDELINES].[dbo].[bsr_InsertPresentationDetail] ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?;",
                            (
                                presentation_id,
                                slide_idx,
                                "NameSummary",
                                summary_image_path,
                                "Brainstorm",
                                "",  # @Param6
                                "",  # @Param7
                                "",  # @Param8
                                "",  # @Param9
                                "",  # @Param10
                                "",  # @Param11
                                "",  # @Param12
                                0,   # @BGTemplateId
                            )
                        )
                        slide_idx += 1

                    # Insert remaining slides AFTER summary
                    for idx in range(slide_number, len(images)):
                        # Use the actual image path from pptx_data, not thumbnails
                        image_path = images[idx]
                        title = slide_titles[idx] if idx < len(slide_titles) else f"Slide {idx + 1}"

                        cursor.execute(
                            "EXEC [BI_GUIDELINES].[dbo].[bsr_InsertPresentationDetail] ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?;",
                            (
                                presentation_id,
                                slide_idx,
                                "Image",
                                image_path,
                                title,
                                "",  # @Param6
                                "",  # @Param7
                                "",  # @Param8
                                "",  # @Param9
                                "",  # @Param10
                                "",  # @Param11
                                "",  # @Param12
                                0,   # @BGTemplateId
                            )
                        )
                        slide_idx += 1

                    # Normalize any 'Thumbnails' paths to full-size image paths for this presentation
                    try:
                        # Attempt normalization on common BSR detail tables/views
                        for table_name in (
                            "[BI_GUIDELINES].[dbo].[BSR_Details]",
                            "[BI_GUIDELINES].[dbo].[bsr_Details]",
                            "[BI_GUIDELINES].[dbo].[bsr_Detail]",
                        ):
                            try:
                                cursor.execute(
                                    (
                                        f"UPDATE {table_name} "
                                        "SET SlideBGFileName = REPLACE(REPLACE(SlideBGFileName, '/Thumbnails/', '/'), '/thumbnails/', '/') "
                                        "WHERE PresentationId = ?"
                                    ),
                                    (presentation_id,),
                                )
                            except Exception:
                                # Ignore if table doesn't exist in this environment
                                pass

                        # Force-set each slide's path to the exact full-size image we generated
                        try:
                            for i, img in enumerate(images, start=1):
                                for tname in (
                                    "[BI_GUIDELINES].[dbo].[BSR_Details]",
                                    "[BI_GUIDELINES].[dbo].[bsr_Details]",
                                    "[BI_GUIDELINES].[dbo].[bsr_Detail]",
                                ):
                                    try:
                                        cursor.execute(
                                            (
                                                f"UPDATE {tname} SET SlideBGFileName = ? "
                                                "WHERE PresentationId = ? AND SlideNumber = ?"
                                            ),
                                            (img, presentation_id, i),
                                        )
                                    except Exception:
                                        # ignore if table not present
                                        pass
                        except Exception as set_exc:
                            logger.warning("BSR force-set image paths skipped: %s", str(set_exc))

                        # Log a few rows after normalization to verify
                        for table_name in (
                            "[BI_GUIDELINES].[dbo].[BSR_Details]",
                            "[BI_GUIDELINES].[dbo].[bsr_Details]",
                        ):
                            try:
                                cursor.execute(
                                    (
                                        f"SELECT TOP 3 SlideNumber, SlideBGFileName FROM {table_name} "
                                        "WHERE PresentationId = ? ORDER BY SlideNumber"
                                    ),
                                    (presentation_id,),
                                )
                                rows = cursor.fetchall() or []
                                if rows:
                                    samples = ", ".join(
                                        [f"{r[0]}=>{r[1]}" for r in rows if len(r) >= 2]
                                    )
                                    logger.info(
                                        "BSR detail path samples after normalization from %s: %s",
                                        table_name,
                                        samples,
                                    )
                            except Exception:
                                pass
                    except Exception as norm_exc:
                        logger.warning("BSR detail path normalization skipped: %s", str(norm_exc))

                    conn.commit()
                    logger.info("Inserted %d BSR detail records", slide_idx - 1)
                    return slide_idx - 1

        except Exception as e:
            logger.error("Error inserting BSR detail records: %s", str(e))
            raise HTTPException(
                status_code=500,
                detail=f"Failed to insert BSR detail records: {str(e)}"
            ) from e

    def _process_bsr_categories(
        self,
        presentation_id: int,
        categories_data: Dict[str, Any]
    ) -> List[Dict[str, Any]]:
        """
        Process and save additional categories for a BSR project.

        Args:
            presentation_id: ID of the newly created BSR project
            categories_data: Object with categories structure

        Returns:
            List of created categories with their IDs

        Process:
            1. Validate no duplicate categories exist
            2. Insert category1 (always)
            3. Insert category2 (only if mode='both')
            4. For each category:
               a. Insert into BSR_CATEGORY
               b. Get generated category_id
               c. Insert each element into BSR_CATEGORY_ELEMENTS
        """
        created_categories = []
        mode = categories_data.get("mode")

        # Validate that no categories exist for this project
        existing_count = self._check_existing_categories(presentation_id)
        if existing_count > 0:
            raise ValueError(
                f"Project {presentation_id} already has {existing_count} categories"
            )

        # Process Category 1 (always present)
        cat1_data = categories_data.get("category1")
        if cat1_data:
            # Validate that category name doesn't exist
            if self._check_category_name_exists(presentation_id, cat1_data["name"]):
                raise ValueError(f"Category '{cat1_data['name']}' already exists")

            cat1_result = self._insert_category_with_elements(
                presentation_id=presentation_id,
                category_name=cat1_data["name"],
                elements=cat1_data["elements"]
            )
            created_categories.append(cat1_result)

        # Process Category 2 (only if mode='both')
        if mode == "both":
            cat2_data = categories_data.get("category2")
            if cat2_data:
                # Validate that category name doesn't exist
                if self._check_category_name_exists(presentation_id, cat2_data["name"]):
                    raise ValueError(f"Category '{cat2_data['name']}' already exists")

                cat2_result = self._insert_category_with_elements(
                    presentation_id=presentation_id,
                    category_name=cat2_data["name"],
                    elements=cat2_data["elements"]
                )
                created_categories.append(cat2_result)

        return created_categories

    def _check_existing_categories(self, presentation_id: int) -> int:
        """
        Count how many categories exist for a project.

        Returns:
            Number of existing categories
        """
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    query = """
                        SELECT COUNT(*) as count
                        FROM [BI_GUIDELINES].[dbo].[BSR_CATEGORY]
                        WHERE BSRPROJECTID = ?
                    """
                    cursor.execute(query, (presentation_id,))
                    result = cursor.fetchone()
                    return result[0] if result else 0
        except Exception as e:
            logger.error("Error checking existing categories: %s", str(e))
            return 0

    def _check_category_name_exists(
        self,
        presentation_id: int,
        category_name: str
    ) -> bool:
        """
        Check if a category with that name already exists in the project.

        Returns:
            True if exists, False if not
        """
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    query = """
                        SELECT COUNT(*) as count
                        FROM [BI_GUIDELINES].[dbo].[BSR_CATEGORY]
                        WHERE BSRPROJECTID = ? AND CATEGORY = ?
                    """
                    cursor.execute(query, (presentation_id, category_name))
                    result = cursor.fetchone()
                    return (result[0] if result else 0) > 0
        except Exception as e:
            logger.error("Error checking category name exists: %s", str(e))
            return False

    def _insert_category_with_elements(
        self,
        presentation_id: int,
        category_name: str,
        elements: List[str]
    ) -> Dict[str, Any]:
        """
        Insert a category and its elements into the database.

        Args:
            presentation_id: BSR project ID
            category_name: Category name (e.g., "Region")
            elements: List of elements (e.g., ["North", "South", "East"])

        Returns:
            Dict with category_id and elements count

        Process:
            1. INSERT into BSR_CATEGORY
            2. Get generated category_id
            3. Multiple INSERTs into BSR_CATEGORY_ELEMENTS
        """
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    # Step 1: Insert category
                    insert_category_query = """
                        INSERT INTO [BI_GUIDELINES].[dbo].[BSR_CATEGORY]
                        (BSRPROJECTID, CATEGORY)
                        VALUES (?, ?)
                    """
                    cursor.execute(insert_category_query, (presentation_id, category_name))

                    # Step 2: Get generated ID
                    get_id_query = """
                        SELECT id
                        FROM [BI_GUIDELINES].[dbo].[BSR_CATEGORY]
                        WHERE BSRPROJECTID = ? AND CATEGORY = ?
                    """
                    cursor.execute(get_id_query, (presentation_id, category_name))
                    result = cursor.fetchone()

                    if not result:
                        raise ValueError(f"Failed to retrieve category ID for '{category_name}'")

                    category_id = result[0]

                    # Step 3: Insert elements one by one
                    insert_element_query = """
                        INSERT INTO [BI_GUIDELINES].[dbo].[BSR_CATEGORY_ELEMENTS]
                        (CATEGORY_ID, CATEGORY_MEMBERS)
                        VALUES (?, ?)
                    """

                    elements_inserted = 0
                    for element in elements:
                        element_clean = element.strip()
                        if element_clean:  # Only insert if not empty
                            cursor.execute(insert_element_query, (category_id, element_clean))
                            elements_inserted += 1

                    conn.commit()

                    logger.info(
                        "Inserted category '%s' (ID=%d) with %d elements for project %d",
                        category_name,
                        category_id,
                        elements_inserted,
                        presentation_id
                    )

                    return {
                        "category_id": category_id,
                        "name": category_name,
                        "elements_count": elements_inserted
                    }
        except Exception as e:
            logger.error("Error inserting category with elements: %s", str(e))
            raise HTTPException(
                status_code=500,
                detail=f"Failed to insert category '{category_name}': {str(e)}"
            ) from e

    def update_project_categories(
        self,
        presentation_id: int,
        categories_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Update categories for an existing BSR project.

        Equivalente VB.NET: FnInsertCategory_FromUpdate (línea 3419)

        Process:
            1. Delete existing categories (CASCADE will delete elements)
            2. If mode="none", only delete and return
            3. If mode="single", insert only category1
            4. If mode="both", insert category1 and category2

        Args:
            presentation_id: BSR project ID
            categories_data: Dict with mode, category1, category2

        Returns:
            Dict with message, categories_updated, deleted_count
        """
        try:
            # STEP 1: Delete existing categories
            deleted_count = self._delete_project_categories(presentation_id)

            # If mode is "none", only delete
            if categories_data.get("mode") == "none":
                return {
                    "message": "Categories deleted successfully",
                    "categories_updated": 0,
                    "deleted_count": deleted_count
                }

            # STEP 2: Insert new categories
            created_categories = []

            # Insert Category 1
            cat1_data = categories_data.get("category1")
            if cat1_data:
                # Validate that category name doesn't exist
                if self._check_category_name_exists(presentation_id, cat1_data["name"]):
                    raise ValueError(f"Category '{cat1_data['name']}' already exists")

                cat1_result = self._insert_category_with_elements(
                    presentation_id=presentation_id,
                    category_name=cat1_data["name"],
                    elements=cat1_data["elements"]
                )
                created_categories.append(cat1_result)

            # Insert Category 2 if mode is "both"
            if categories_data.get("mode") == "both":
                cat2_data = categories_data.get("category2")
                if cat2_data:
                    # Validate that category name doesn't exist
                    if self._check_category_name_exists(presentation_id, cat2_data["name"]):
                        raise ValueError(f"Category '{cat2_data['name']}' already exists")

                    cat2_result = self._insert_category_with_elements(
                        presentation_id=presentation_id,
                        category_name=cat2_data["name"],
                        elements=cat2_data["elements"]
                    )
                    created_categories.append(cat2_result)

            logger.info(
                "Updated categories for project %d: deleted=%d, created=%d",
                presentation_id,
                deleted_count,
                len(created_categories)
            )

            return {
                "message": "Categories updated successfully",
                "categories_updated": len(created_categories),
                "deleted_count": deleted_count
            }

        except ValueError as e:
            logger.error("Validation error updating categories: %s", str(e))
            raise HTTPException(status_code=400, detail=str(e)) from e
        except Exception as e:
            logger.error("Error updating categories for project %d: %s", presentation_id, str(e))
            raise HTTPException(
                status_code=500,
                detail=f"Failed to update categories: {str(e)}"
            ) from e

    def _delete_project_categories(self, presentation_id: int) -> int:
        """
        Delete all categories and their elements for a project.

        Equivalente VB.NET:
        - fnDeleteCategoryfromUpdate() (línea 3018)
        - fnDeleteCategoryElements_fromUpdate() (línea 3131)

        Process:
            1. Get category IDs to delete
            2. Delete elements from BSR_CATEGORY_ELEMENTS
            3. Delete categories from BSR_CATEGORY

        Returns:
            Number of categories deleted
        """
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    # Step 1: Get category IDs to delete
                    get_category_ids_query = """
                        SELECT id
                        FROM [BI_GUIDELINES].[dbo].[BSR_CATEGORY]
                        WHERE BSRPROJECTID = ?
                    """
                    cursor.execute(get_category_ids_query, (presentation_id,))
                    category_ids = cursor.fetchall()

                    if not category_ids or len(category_ids) == 0:
                        logger.info("No categories to delete for project %d", presentation_id)
                        return 0

                    # Step 2: Delete elements for each category
                    delete_elements_query = """
                        DELETE FROM [BI_GUIDELINES].[dbo].[BSR_CATEGORY_ELEMENTS]
                        WHERE category_id = ?
                    """

                    for cat_id_row in category_ids:
                        cat_id = cat_id_row[0]
                        cursor.execute(delete_elements_query, (cat_id,))

                    # Step 3: Delete categories
                    delete_categories_query = """
                        DELETE FROM [BI_GUIDELINES].[dbo].[BSR_CATEGORY]
                        WHERE BSRPROJECTID = ?
                    """
                    cursor.execute(delete_categories_query, (presentation_id,))

                    conn.commit()

                    logger.info(
                        "Deleted %d categories for project %d",
                        len(category_ids),
                        presentation_id
                    )

                    return len(category_ids)

        except Exception as e:
            logger.error("Error deleting categories for project %d: %s", presentation_id, str(e))
            raise HTTPException(
                status_code=500,
                detail=f"Failed to delete categories: {str(e)}"
            ) from e

    def count_project_categories(self, presentation_id: int) -> int:
        """
        Count how many categories a project has.

        Equivalente VB.NET: fnCheckCategory_fromUpdate() (línea 3195)

        Returns:
            Number of categories
        """
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    query = """
                        SELECT COUNT(*) as count
                        FROM [BI_GUIDELINES].[dbo].[BSR_CATEGORY]
                        WHERE BSRPROJECTID = ?
                    """
                    cursor.execute(query, (presentation_id,))
                    result = cursor.fetchone()
                    return result[0] if result else 0
        except Exception as e:
            logger.error("Error counting categories for project %d: %s", presentation_id, str(e))
            return 0

    def check_bsr_project_exists(self, presentation_id: int) -> bool:
        """
        Verify if a BSR project exists.

        Returns:
            True if project exists, False otherwise
        """
        try:
            with create_connection() as conn:
                with conn.cursor() as cursor:
                    query = """
                        SELECT COUNT(*) as count
                        FROM [BI_GUIDELINES].[dbo].[bsr_Master]
                        WHERE PresentationId = ?
                    """
                    cursor.execute(query, (presentation_id,))
                    result = cursor.fetchone()
                    return (result[0] if result else 0) > 0
        except Exception as e:
            logger.error("Error checking if BSR project %d exists: %s", presentation_id, str(e))
            return False


# Global presentation service instance
presentation_service = PresentationService()
