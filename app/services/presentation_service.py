"""
Presentation creation orchestration service.
"""

import time
from typing import Any, Dict, List

from fastapi import HTTPException

from app.api.routes.bi_guidelines import get_db_connection
from app.models.excel_models import (
    CreatePresentationRequest,
    CreatePresentationResponse,
    PPTXConversionResponse,
    ProcessedExcelData,
)
from app.models.nw_master_request import DetailItem, PresentationData
from app.services.excel_service import excel_processing_service
from app.services.pptx_service import pptx_service
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class PresentationService:
    """Service for orchestrating complete presentation creation."""

    def __init__(self):
        self.excel_service = excel_processing_service
        self.pptx_service = pptx_service

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
                excel_data, pptx_data, request
            )

            # 4. Apply templates if specified
            if request.template_rotation:
                slides_data = self._apply_template_rotation(
                    slides_data, request.template_rotation
                )

            # 5. Create presentation in DB
            logger.info("Creating presentation in database")
            presentation_result = self._create_presentation_in_db(slides_data, request)

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

    def _process_excel_file(
        self, request: CreatePresentationRequest
    ) -> ProcessedExcelData:
        """Process Excel file using the existing Excel service."""
        try:
            # Process the Excel file
            result = self.excel_service.process_excel_file(
                file_content=request.excel_file,
                is_phonetics=request.is_phonetics,
                has_groups=request.has_groups,
            )

            logger.info(
                "Excel processing completed: %d rows processed",
                result.total_rows_processed,
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
            result_dict = self.pptx_service.convert_pptx_to_images(
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

    def _generate_slides_from_excel(
        self,
        excel_data: ProcessedExcelData,
        pptx_data: PPTXConversionResponse,
        request: CreatePresentationRequest,
    ) -> Dict[str, Any]:
        """
        Generate slides data from Excel arrays - matching VB.NET logic.
        """
        slides = []
        slide_number = request.page_number

        # Determine slide type based on lst_types (A-Z = groups, numbers = individual)
        for j in range(excel_data.lst_max_item_number + 1):
            if (
                excel_data.lst_types[j]
                and excel_data.lst_types[j].upper() in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            ):
                # Group slide
                slide = self._create_group_slide(
                    excel_data, pptx_data, j, slide_number, request
                )
            else:
                # Individual slide
                slide = self._create_individual_slide(
                    excel_data, pptx_data, j, slide_number, request
                )

            slides.append(slide)
            slide_number += 1

        # Create complete structure for DB
        presentation_data = {
            "project": request.project,
            "display_name": request.display_name,
            "powerpoint_file": f"{request.project}_{request.display_name}.pptx",
            "excel_file": f"{request.project}_candidates.xlsx",
            "background_type": request.background_type,
            "background_name": request.background_name,
            "page_number": request.page_number,
            "presentation_type": request.presentation_type,
            "user_name": request.user_name,
            "bsr_display_name": request.bsr_display_name or request.display_name,
            "mobile_link_bsr": request.mobile_link_bsr,
            "participant_vote": request.participant_vote,
            "is_wide_ppt": request.is_wide_ppt,
            "is_aws_email": request.is_aws_email,
            "details": slides,
        }

        return presentation_data

    def _create_group_slide(
        self,
        excel_data: ProcessedExcelData,
        _pptx_data: PPTXConversionResponse,
        index: int,
        slide_number: int,
        request: CreatePresentationRequest,
    ) -> Dict[str, Any]:
        """Create a group slide (A-Z types)."""
        return {
            "slide_number": slide_number,
            "slide_type": "Image",
            "slide_bg_file_name": f"nw_slides/{request.project}/{slide_number:03d}.jpg",
            "slide_description": excel_data.lst_categories[index],
            "group_name": excel_data.lst_categories[index],
            "category": excel_data.lst_categories[index],
            "name": excel_data.lst_categories[index],  # Group name
            "rationale": "",
            "notation": "",
            "kana": "",
            "logo_filename": "",
            "template_id": 0,
            "name_sub_group": "",
        }

    def _create_individual_slide(
        self,
        excel_data: ProcessedExcelData,
        _pptx_data: PPTXConversionResponse,
        index: int,
        slide_number: int,
        _request: CreatePresentationRequest,
    ) -> Dict[str, Any]:
        """Create an individual name evaluation slide."""
        return {
            "slide_number": slide_number,
            "slide_type": "NameEvaluation",
            "slide_bg_file_name": "",  # Will be assigned with template rotation
            "slide_description": f"{excel_data.lst_names[index]} - Name Evaluation",
            "group_name": excel_data.lst_categories[index] or "Default",
            "category": excel_data.lst_categories[index],
            "name": excel_data.lst_names[index],
            "rationale": excel_data.lst_rationales[index],
            "notation": excel_data.lst_notations[index],
            "kana": excel_data.lst_kana[index],
            "logo_filename": excel_data.lst_logos[index],
            "template_id": 0,  # Will be updated with template
            "name_sub_group": excel_data.lst_name_sub_groups[index],
        }

    def _apply_template_rotation(
        self, slides_data: Dict[str, Any], template_rotation: List[str]
    ) -> Dict[str, Any]:
        """Apply template rotation to slides."""
        if not template_rotation:
            return slides_data

        for i, slide in enumerate(slides_data["details"]):
            template_index = i % len(template_rotation)
            slide["slide_bg_file_name"] = template_rotation[template_index]
            slide["template_id"] = self._get_template_id(
                template_rotation[template_index]
            )

        return slides_data

    def _get_template_id(self, template_name: str) -> int:
        """Get template ID from template name by querying the database."""
        try:
            conn = get_db_connection()
            if not conn:
                logger.warning(
                    "Could not connect to database for template lookup, using default template"
                )
                return 5  # Default template

            cursor = conn.cursor()

            # Query the nw_Templates table for the template ID
            sql = (
                "SELECT TOP 1 templateid "
                "FROM [BI_GUIDELINES].dbo.nw_Templates "
                "WHERE TemplateName = ?"
            )

            cursor.execute(sql, (template_name,))

            row = cursor.fetchone()
            cursor.close()
            conn.close()

            if row:
                return row[0]
            else:
                logger.warning(
                    "Template '%s' not found in database, using default template",
                    template_name,
                )
                return 5  # Default template

        except Exception as e:  # pylint: disable=broad-except
            logger.error(
                "Error getting template ID for '%s': %s", template_name, str(e)
            )
            return 5  # Default template on error

    def _create_presentation_in_db(
        self, slides_data: Dict[str, Any], _request: CreatePresentationRequest
    ) -> Dict[str, Any]:
        """Create presentation in database using BI Guidelines functionality directly."""
        try:
            # Prepare the data for the BI Guidelines functionality
            presentation_payload = PresentationData(
                project=slides_data["project"],
                display_name=slides_data["display_name"],
                powerpoint_file=slides_data["powerpoint_file"],
                excel_file=slides_data["excel_file"],
                background_type=slides_data["background_type"],
                background_name=slides_data["background_name"],
                page_number=slides_data["page_number"],
                presentation_type=slides_data["presentation_type"],
                user_name=slides_data["user_name"],
                bsr_display_name=slides_data["bsr_display_name"],
                mobile_link_bsr=slides_data.get("mobile_link_bsr"),
                participant_vote=slides_data["participant_vote"],
                is_wide_ppt=slides_data["is_wide_ppt"],
                is_aws_email=slides_data["is_aws_email"],
                details=[
                    DetailItem(
                        slide_number=slide["slide_number"],
                        slide_type=slide["slide_type"],
                        slide_bg_file_name=slide.get("slide_bg_file_name", ""),
                        slide_description=slide["slide_description"],
                        group_name=slide.get("group_name"),
                        category=slide.get("category"),
                        name=slide.get("name"),
                        rationale=slide.get("rationale"),
                        notation=slide.get("notation"),
                        kana=slide.get("kana"),
                        logo_filename=slide.get("logo_filename"),
                        template_id=slide.get("template_id", 0),
                        name_sub_group=slide.get("name_sub_group"),
                    )
                    for slide in slides_data.get("details", [])
                ],
            )

            # Use BI Guidelines functionality directly instead of HTTP call
            logger.info(
                "Creating presentation in database using BI Guidelines functionality"
            )

            conn = get_db_connection()
            if not conn:
                raise HTTPException(
                    status_code=500, detail="Could not connect to database."
                )

            cursor = conn.cursor()

            # Start transaction
            cursor.execute("SET NOCOUNT ON;")

            # 1. Insert into the master table
            master_sql = """
                EXEC [dbo].[nw_InsertPresentationMaster_sep2025]
                    @Project=?, @DisplayName=?, @MainPptFileName=?, @NameCandidateFileName=?,
                    @NameCandidateBGType=?, @NameCandidateBGName=?, @NameCandidateStartingSlide=?,
                    @PresentationType=?, @UploadedBy=?, @BSRDisplayName=?,
                    @isParticipantsVote=?, @isWideScreenPPT=?, @isAWSLinkReq=?;
            """
            master_params = (
                presentation_payload.project,
                presentation_payload.display_name,
                presentation_payload.powerpoint_file,
                presentation_payload.excel_file,
                presentation_payload.background_type,
                presentation_payload.background_name,
                presentation_payload.page_number,
                presentation_payload.presentation_type,
                presentation_payload.user_name,
                presentation_payload.bsr_display_name,
                presentation_payload.participant_vote,
                presentation_payload.is_wide_ppt,
                presentation_payload.is_aws_email,
            )

            logger.info("Executing master stored procedure...")
            cursor.execute(master_sql, master_params)

            # Get the presentation_id
            presentation_id = None
            while True:
                try:
                    row = cursor.fetchone()
                    if row:
                        presentation_id = row[0]
                        logger.info("PresentationId created: %s", presentation_id)
                        break
                except Exception as e:  # pylint: disable=broad-except
                    logger.error("Error fetching presentation ID: %s", e)
                if not cursor.nextset():
                    break

            if not presentation_id:
                raise HTTPException(
                    status_code=500,
                    detail="Could not get PresentationId from the master record.",
                )

            # 2. Insert into the details table
            detail_sql = """
                EXEC [dbo].[nw_InsertPresentationDetail_copy]
                    @PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?,
                    @SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?,
                    @NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?,
                    @TemplateId=?, @NameSubGroup=?;
            """
            for item in presentation_payload.details:
                detail_params = (
                    presentation_id,
                    item.slide_number,
                    item.slide_type,
                    item.slide_bg_file_name,
                    item.slide_description,
                    item.group_name,
                    item.category,
                    item.name,
                    item.rationale,
                    item.notation,
                    item.kana,
                    item.logo_filename,
                    item.template_id,
                    item.name_sub_group,
                )
                cursor.execute(detail_sql, detail_params)

            # Commit the transaction
            conn.commit()
            logger.info("Presentation created successfully in database")

            cursor.close()
            conn.close()

            return {
                "message": "Presentation created successfully.",
                "presentation_id": presentation_id,
            }

        except Exception as e:
            error_msg = f"Error creating presentation in database: {str(e)}"
            logger.error(error_msg)
            raise HTTPException(status_code=500, detail=error_msg) from e


# Global presentation service instance
presentation_service = PresentationService()
