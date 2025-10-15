"""
Base models for all presentation project types.

These models contain common fields and logic shared across NW, DW, BSR, and NSR projects.
Project-specific models should inherit from these base models and extend as needed.
"""

from typing import Dict, List, Optional
from enum import Enum
from abc import ABC

from pydantic import BaseModel, ConfigDict, Field

from app.models.response_models import PPTXConversionResponse
from app.models.excel_models import ProcessedExcelData


class ProjectType(str, Enum):
    """Supported project types."""
    NW = "NW"
    DW = "DW"
    BSR = "BSR"
    NSR = "NSR"


class TestNameOrder(str, Enum):
    """Test name ordering options."""
    DEFAULT = "Default"
    RANDOMIZE = "Randomize"
    RANDOMIZE_TOP_5 = "Randomize_top_5"


class BackgroundType(str, Enum):
    """Background type options."""
    DEFAULT = "Default"
    ROTATE = "Rotate"


class BaseCreatePresentationMetadata(BaseModel, ABC):
    """Base metadata payload for presentation creation.

    All project-specific metadata models should inherit from this base.
    Contains common fields required by all project types.
    """

    model_config = ConfigDict(use_enum_values=True)

    # ===== PRINCIPAL DATA FIELDS (COMMON TO ALL PROJECTS) =====
    project: str = Field(
        ...,
        description="Project identifier"
    )
    display_name: str = Field(
        ...,
        description="Presentation name"
    )
    presentation_type: str = Field(
        "Normal",
        description="Presentation type. Default: 'Normal'",
    )
    project_type: ProjectType = Field(
        ...,
        description="Project kind: 'NW', 'DW', 'BSR', or 'NSR'",
    )
    user_name: str = Field(
        ...,
        description="User creating the presentation"
    )
    page_number: int = Field(
        1,
        ge=1,
        description="Starting page number. Default: 1",
    )

    # ===== CONFIGURATION FLAGS (COMMON) =====
    participant_vote: int = Field(
        1,
        ge=0,
        le=1,
        description="Participant vote flag (0 or 1). Default: 1",
    )
    is_wide_ppt: int = Field(
        0,
        ge=0,
        le=1,
        description="Aspect ratio: 0=4:3, 1=16:9. Default: 0",
    )
    is_aws_email: int = Field(
        0,
        ge=0,
        le=1,
        description="AWS email flag (0 or 1). Default: 0",
    )

    # ===== CONTROL FLAGS (COMMON) =====
    has_groups: bool = Field(
        True,
        description="Whether the Excel sheet contains groups. Default: true",
    )

    # ===== RANDOMIZATION OPTIONS (COMMON) =====
    test_name_order: TestNameOrder = Field(
        TestNameOrder.DEFAULT,
        description="Ordering strategy: 'Default', 'Randomize', or 'Randomize_top_5'. Default: 'Default'",
    )

    # ===== BACKGROUND OPTIONS (COMMON) =====
    background_type: BackgroundType = Field(
        BackgroundType.DEFAULT,
        description="Background type: 'Default' or 'Rotate'. Default: 'Default'",
    )
    background_name: Optional[str] = Field(
        default="",
        description="Background template names separated by '|' (e.g., 'BMW_1|BrandDNA'). Default: ''",
    )


class BaseCreatePresentationRequest(BaseModel, ABC):
    """Base request model for creating a complete presentation.

    Contains common fields required for all project types.
    Project-specific requests should inherit and extend as needed.
    """

    excel_file: bytes = Field(
        ..., description="Raw bytes for the uploaded Excel workbook."
    )
    pptx_file: bytes = Field(
        ..., description="Raw bytes for the uploaded PowerPoint template."
    )

    excel_filename: str = Field(
        "data.xlsx",
        description="Original filename of the Excel workbook.",
    )
    pptx_filename: str = Field(
        "presentation.pptx",
        description="Original filename of the PPTX template.",
    )
    has_groups: bool = Field(
        False, description="Flag indicating whether the Excel file contains group/sub-group rows."
    )
    test_name_order: TestNameOrder = Field(
        TestNameOrder.DEFAULT,
        description="Ordering strategy for test names: Default, Randomize, or Randomize_top_5."
    )

    project: str = Field(
        ..., description="Project identifier."
    )
    display_name: str = Field(
        ..., description="Human-readable display name."
    )
    presentation_type: str = Field(
        "Normal",
        description="Presentation classification.",
    )
    user_name: str = Field(..., description="Username that initiated the creation process.")
    participant_vote: int = Field(1, description="Participant vote flag.")
    is_wide_ppt: int = Field(0, ge=0, le=1, description="Widescreen flag (0/1).")
    is_aws_email: int = Field(0, ge=0, le=1, description="AWS email flag (0/1).")

    background_type: str = Field("Default", description="Background selection type.")
    background_name: str = Field("", description="Background asset name.")
    page_number: int = Field(1, ge=1, description="Page offset for numbering.")
    project_type: str = Field("NW", description="Project type identifier.")


class BaseCreatePresentationResponse(BaseModel):
    """Base response model for presentation creation.

    Contains common response fields for all project types.
    """

    model_config = ConfigDict(use_enum_values=True)

    message: str = Field(
        ..., description="Human-readable status message."
    )
    presentation_id: Optional[int] = Field(
        default=None,
        description="Database presentation record identifier.",
    )
    total_slides: int = Field(
        0, description="Total number of slides generated."
    )
    excel_data: ProcessedExcelData = Field(
        ..., description="Processed Excel content."
    )
    pptx_data: PPTXConversionResponse = Field(
        ..., description="PPTX conversion metadata."
    )
    processing_time_seconds: float = Field(
        ..., description="Total processing time."
    )


class BaseDetailItem(BaseModel):
    """Base model for a single slide detail record.

    Represents a single record in the presentation details table.
    Common to all project types.
    """

    slide_number: int = Field(
        ..., ge=1, description="Slide number within the presentation sequence."
    )
    slide_type: str = Field(
        ..., description="Slide classification."
    )
    slide_bg_file_name: Optional[str] = Field(
        default=None,
        description="Background file applied to the slide.",
    )
    slide_description: Optional[str] = Field(
        default=None, description="Narrative description."
    )
    group_name: Optional[str] = Field(
        default=None, description="Group name when applicable."
    )
    category: Optional[str] = Field(
        default=None, description="Category of the candidate or slide grouping."
    )
    name: Optional[str] = Field(
        default=None, description="Primary candidate or item name."
    )
    rationale: Optional[str] = Field(
        default=None, description="Supporting rationale."
    )
    notation: Optional[str] = Field(
        default=None, description="Phonetic notation."
    )
    kana: Optional[str] = Field(
        default=None, description="Japanese kana rendering."
    )
    logo_filename: Optional[str] = Field(
        default=None,
        description="Optional logo filename.",
    )
    template_id: int = Field(
        ..., description="Template identifier."
    )
    name_sub_group: Optional[str] = Field(
        default=None, description="Sub-group label."
    )


class BasePresentationData(BaseModel):
    """Base model for complete presentation payload.

    Represents the complete data structure for creating a new presentation.
    Common to all project types.
    """

    project: str = Field(..., description="Project identifier.")
    display_name: str = Field(..., description="User-friendly name.")
    powerpoint_file: str = Field(..., description="Generated PowerPoint filename.")
    excel_file: str = Field(..., description="Source Excel filename.")
    background_type: str = Field(..., description="Background type.")
    background_name: str = Field(..., description="Background asset name.")
    page_number: int = Field(..., description="Page numbering offset.")
    presentation_type: str = Field(..., description="Presentation classification.")
    user_name: str = Field(..., description="User responsible.")
    participant_vote: int = Field(..., description="Participant vote flag.")
    is_wide_ppt: int = Field(..., description="Widescreen flag.")
    is_aws_email: int = Field(..., description="AWS email flag.")
    details: List[BaseDetailItem] = Field(
        ..., description="Collection of slide detail records."
    )
