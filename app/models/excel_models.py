"""
Excel processing models for data transformation.
"""

from typing import List, Optional

from pydantic import BaseModel, Field

from app.models.response_models import PPTXConversionResponse


class ExcelRowData(BaseModel):
    """Represents a single row of data from the Excel file."""

    # Column A: Type (A-Z for groups, numbers for individual items)
    Type: Optional[str] = None

    # Column B: Category
    Category: Optional[str] = None

    # Column C: Name + Notation combined (needs getTestName processing)
    NameWithNotation: Optional[str] = None

    # Column D: Rationale
    Rationale: Optional[str] = None

    # Column E: Kana (Japanese reading)
    Kana: Optional[str] = None

    # Column F: Logo filename
    Logo: Optional[str] = None

    # Column G/H: NameSubGroup (depending on configuration)
    NameSubGroup: Optional[str] = None


class ProcessedExcelData(BaseModel):
    """Data structure containing processed Excel arrays."""

    # Array data structure for Excel processing
    lst_categories: List[str] = Field(default_factory=list)
    lst_names: List[str] = Field(default_factory=list)
    lst_rationales: List[str] = Field(default_factory=list)
    lst_notations: List[str] = Field(default_factory=list)
    lst_types: List[str] = Field(default_factory=list)
    lst_kana: List[str] = Field(default_factory=list)
    lst_logos: List[str] = Field(default_factory=list)
    lst_name_sub_groups: List[str] = Field(default_factory=list)

    # Metadata
    lst_max_item_number: int = 0
    total_rows_processed: int = 0


class ExcelProcessingRequest(BaseModel):
    """Request parameters for Excel processing."""

    presentation_id: Optional[str] = ""
    display_name: Optional[str] = ""
    start_index: Optional[int] = 0
    is_phonetics: Optional[bool] = False
    has_groups: Optional[bool] = False


class ExcelProcessingResponse(BaseModel):
    """Response model for Excel processing - returns processed data arrays."""

    message: str
    data: ProcessedExcelData
    processing_id: Optional[str] = None


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
