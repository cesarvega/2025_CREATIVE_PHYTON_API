"""
Presentation creation and management models.
"""

from typing import List, Optional

from pydantic import BaseModel

from app.models.response_models import PPTXConversionResponse
from app.models.excel_models import ProcessedExcelData


class CreatePresentationMetadata(BaseModel):
    """Metadata payload supplied by clients alongside the uploaded files."""

    project: str
    display_name: str
    presentation_type: str = "Nonproprietary"
    user_name: str
    bsr_display_name: Optional[str] = None
    mobile_link_bsr: Optional[str] = None
    participant_vote: int = 1
    is_wide_ppt: int = 0
    is_aws_email: int = 0
    is_design_mode: bool = False
    background_type: str = "Image"
    background_name: str = ""
    page_number: int = 1
    project_type: str = "bipresents"
    is_phonetics: bool = False
    has_groups: bool = False
    template_rotation: Optional[List[str]] = None

    def to_service_request(
        self,
        *,
        excel_content: bytes,
        excel_filename: str,
        pptx_content: bytes,
        pptx_filename: str,
    ) -> "CreatePresentationRequest":
        """Convert metadata into a CreatePresentationRequest with uploaded files."""

        return CreatePresentationRequest(
            excel_file=excel_content,
            pptx_file=pptx_content,
            excel_filename=excel_filename,
            pptx_filename=pptx_filename,
            is_phonetics=self.is_phonetics,
            has_groups=self.has_groups,
            project=self.project,
            display_name=self.display_name,
            presentation_type=self.presentation_type,
            user_name=self.user_name,
            bsr_display_name=self.bsr_display_name,
            mobile_link_bsr=self.mobile_link_bsr,
            participant_vote=self.participant_vote,
            is_wide_ppt=self.is_wide_ppt,
            is_aws_email=self.is_aws_email,
            is_design_mode=self.is_design_mode,
            background_type=self.background_type,
            background_name=self.background_name,
            page_number=self.page_number,
            project_type=self.project_type,
            template_rotation=self.template_rotation,
        )


class CreatePresentationRequest(BaseModel):
    """Request model for creating a complete presentation."""

    # Input files (processed internally by the service)
    excel_file: bytes  # Excel file content
    pptx_file: bytes  # PPTX file content

    # Processing parameters
    excel_filename: str = "data.xlsx"
    pptx_filename: str = "presentation.pptx"
    is_phonetics: bool = False
    has_groups: bool = False

    # Presentation data
    project: str
    display_name: str
    presentation_type: str = "Nonproprietary"
    user_name: str
    bsr_display_name: Optional[str] = None
    mobile_link_bsr: Optional[str] = None
    participant_vote: int = 1
    is_wide_ppt: int = 0
    is_aws_email: int = 0
    is_design_mode: bool = False

    # Additional configuration
    background_type: str = "Image"
    background_name: str = ""
    page_number: int = 1
    project_type: str = "bipresents"  # "bipresents" or "nw"
    template_rotation: Optional[List[str]] = None  # List of templates for rotation


class CreatePresentationResponse(BaseModel):
    """Response model for presentation creation."""

    message: str
    presentation_id: Optional[int] = None
    total_slides: int = 0
    excel_data: ProcessedExcelData
    pptx_data: PPTXConversionResponse
    processing_time_seconds: float


class DetailItem(BaseModel):
    """Represents a single record in the nw_Details table."""

    slide_number: int
    slide_type: str
    slide_bg_file_name: Optional[str] = None
    slide_description: Optional[str] = None
    group_name: Optional[str] = None
    category: Optional[str] = None
    name: Optional[str] = None
    rationale: Optional[str] = None
    notation: Optional[str] = None
    kana: Optional[str] = None
    logo_filename: Optional[str] = None
    template_id: int
    name_sub_group: Optional[str] = None


class PresentationData(BaseModel):
    """Represents the complete payload for creating a new presentation."""

    project: str
    display_name: str
    powerpoint_file: str
    excel_file: str
    background_type: str
    background_name: str
    page_number: int
    presentation_type: str
    user_name: str
    bsr_display_name: str
    mobile_link_bsr: Optional[str] = None
    participant_vote: int
    is_wide_ppt: int
    is_aws_email: int
    is_design_mode: bool = False
    is_printed: bool = False
    details: List[DetailItem]