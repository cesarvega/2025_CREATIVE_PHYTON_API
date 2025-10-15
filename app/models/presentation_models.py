"""
Presentation creation and management models.
"""

from typing import Dict, List, Optional
from enum import Enum

from pydantic import BaseModel, ConfigDict, Field, model_validator

from app.models.response_models import PPTXConversionResponse
from app.models.excel_models import ProcessedExcelData
from app.utils.path_utils import sanitize_folder_name


class ProjectKind(str, Enum):
    """Project type enumeration."""
    NW = "NW"
    BSR = "BSR"
    NSR = "NSR"
    DW = "DW"


class TestNameOrder(str, Enum):
    """Test name ordering options."""
    DEFAULT = "Default"
    RANDOMIZE = "Randomize"
    RANDOMIZE_TOP_5 = "Randomize_top_5"


class BackgroundType(str, Enum):
    """Background type options."""
    DEFAULT = "Default"
    ROTATE = "Rotate"


class CreatePresentationMetadata(BaseModel):
    """Metadata payload supplied by clients alongside the uploaded files."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "project": "NW_PROJECT_2025",
                "display_name": "Name Evaluation",
                "presentation_type": "Nonproprietary",
                "user_name": "analyst",
                "bsr_display_name": "Brand Study Results",
                "mobile_link_bsr": "https://example.com/bsr",
                "participant_vote": 1,
                "is_wide_ppt": 0,
                "is_aws_email": 0,
                "is_design_mode": False,
                "background_type": "Image",
                "background_name": "blue_gradient",
                "page_number": 1,
                "project_type": "nw",
                "is_phonetics": True,
                "has_groups": True,
                "template_rotation": [
                    "template_default_2019.pptx",
                    "template_default_withgroups2019.pptx",
                ],
            }
        }
    )

    project: str = Field(
        ..., description="Internal project identifier used for folder paths and database records."
    )
    display_name: str = Field(
        ..., description="Human-friendly name for generated assets (folders, output files, database records)."
    )
    presentation_type: str = Field(
        "Nonproprietary",
        description="Presentation classification stored in the database (for example 'Nonproprietary').",
    )
    user_name: str = Field(
        ..., description="Username recorded for auditing when the presentation is created."
    )
    bsr_display_name: Optional[str] = Field(
        default=None,
        description="Optional Brand Study Results label displayed in generated slides.",
    )
    mobile_link_bsr: Optional[str] = Field(
        default=None,
        description="Optional URL that is embedded in the presentation for mobile BSR access.",
    )
    participant_vote: int = Field(
        1,
        description="Numeric flag persisted in the master record to indicate participant vote status (legacy compatibility).",
    )
    is_wide_ppt: int = Field(
        0,
        ge=0,
        le=1,
        description="Whether widescreen (16:9) templates should be used (0 = standard, 1 = widescreen).",
    )
    is_aws_email: int = Field(
        0,
        ge=0,
        le=1,
        description="Flag that instructs downstream automation to send AWS email notifications (0/1).",
    )
    is_design_mode: bool = Field(
        False,
        description="Enable design-mode processing to keep placeholders for manual adjustments.",
    )
    background_type: str = Field(
        "Image",
        description="Background selection mode, typically 'Image' or 'Color'.",
    )
    background_name: str = Field(
        "",
        description="Specific background asset name to apply when generating slides.",
    )
    page_number: int = Field(
        1,
        ge=1,
        description="Starting page number offset written into generated slides.",
    )
    project_type: str = Field(
        "NW",
        description="Presentation family that determines directory routing (NW, BSR, NSR, DW).",
    )
    is_phonetics: bool = Field(
        False,
        description="Indicates that phonetic columns (kana/notation) should be processed from the Excel sheet.",
    )
    has_groups: bool = Field(
        True,
        description="Whether the Excel sheet contains subgroup rows that must be converted into group slides.",
    )
    test_name_order: TestNameOrder = Field(
        TestNameOrder.DEFAULT,
        description="Ordering strategy for test names: Default (keep original), Randomize (shuffle all), or Randomize_top_5 (shuffle only first 5).",
    )
    template_rotation: Optional[List[str]] = Field(
        default=None,
        description="Optional list of PPTX filenames that should be rotated when building slides.",
    )
    
    # Physical PowerPoint generation options (optional, will use defaults if not provided)
    template_pack: Optional[str] = Field(
        default=None,
        description="Template pack folder to use for generating physical PowerPoint slides.",
    )
    base_template: Optional[str] = Field(
        default=None,
        description="Base template file for individual candidate slides.",
    )
    multi_template: Optional[str] = Field(
        default=None,
        description="Template file for slides containing multiple grouped candidates.",
    )
    group_template: Optional[str] = Field(
        default=None,
        description="Template file for group header slides.",
    )
    separator_template: Optional[str] = Field(
        default=None,
        description="Template file for separator slides between sections.",
    )
    summary_template: Optional[str] = Field(
        default=None,
        description="Template file for summary slides.",
    )
    include_macro_version: bool = Field(
        default=False,
        description="Generate a macro-enabled (.pptm) version of the PowerPoint file.",
    )
    generate_physical_pptx: bool = Field(
        default=True,
        description="Generate physical PowerPoint file combining original and template slides.",
    )

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
            test_name_order=self.test_name_order,
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
            # Physical PowerPoint generation options
            template_pack=self.template_pack,
            base_template=self.base_template,
            multi_template=self.multi_template,
            group_template=self.group_template,
            separator_template=self.separator_template,
            summary_template=self.summary_template,
            include_macro_version=self.include_macro_version,
            generate_physical_pptx=self.generate_physical_pptx,
        )


class CreatePresentationRequest(BaseModel):
    """Request model for creating a complete presentation."""

    excel_file: bytes = Field(
        ..., description="Raw bytes for the uploaded Excel workbook."
    )
    pptx_file: bytes = Field(
        ..., description="Raw bytes for the uploaded PowerPoint template."
    )

    excel_filename: str = Field(
        "data.xlsx",
        description="Original filename of the Excel workbook (used for diagnostics and storage).",
    )
    pptx_filename: str = Field(
        "presentation.pptx",
        description="Original filename of the PPTX template (used for diagnostics and storage).",
    )
    is_phonetics: bool = Field(
        False, description="Process phonetic columns (kana/notation) from the Excel data."
    )
    has_groups: bool = Field(
        False, description="Flag indicating whether the Excel file contains group/sub-group rows."
    )
    test_name_order: TestNameOrder = Field(
        TestNameOrder.DEFAULT,
        description="Ordering strategy for test names: Default, Randomize, or Randomize_top_5."
    )

    project: str = Field(
        ..., description="Project identifier persisted in the master presentation record."
    )
    display_name: str = Field(
        ..., description="Human-readable display name for generated assets and database entries."
    )
    presentation_type: str = Field(
        "Nonproprietary",
        description="Presentation classification stored in the database.",
    )
    user_name: str = Field(..., description="Username that initiated the creation process.")
    bsr_display_name: Optional[str] = Field(
        default=None,
        description="Optional Brand Study Results label to include in slides.",
    )
    mobile_link_bsr: Optional[str] = Field(
        default=None, description="Optional URL pointing to the BSR mobile experience."
    )
    participant_vote: int = Field(
        1,
        description="Numeric flag persisted with the master record for reporting compatibility.",
    )
    is_wide_ppt: int = Field(
        0,
        ge=0,
        le=1,
        description="Indicates if widescreen templates should be applied (0 = standard, 1 = widescreen).",
    )
    is_aws_email: int = Field(
        0,
        ge=0,
        le=1,
        description="Flag enabling downstream AWS email distribution workflows (0/1).",
    )
    is_design_mode: bool = Field(
        False,
        description="Run in design mode to keep placeholders for manual editing instead of final assets.",
    )

    background_type: str = Field(
        "Image",
        description="Background selection type, e.g. 'Image', 'Color', or 'None'.",
    )
    background_name: str = Field(
        "",
        description="Specific background asset name applied to slides when background_type='Image'.",
    )
    page_number: int = Field(
        1,
        ge=1,
        description="Page offset for numbering generated slides.",
    )
    project_type: str = Field(
        "bipresents",
        description="Directory and business unit routing (typically 'bipresents' or 'nw').",
    )
    template_rotation: Optional[List[str]] = Field(
        default=None,
        description="Optional ordered list of template filenames to rotate per generated slide.",
    )
    
    # Physical PowerPoint generation options
    template_pack: Optional[str] = Field(
        default="BackgroundDefaultTemplate",
        description="Template pack folder to use for generating physical PowerPoint slides.",
    )
    base_template: Optional[str] = Field(
        default="template_default_2019.pptx",
        description="Base template file for individual candidate slides.",
    )
    multi_template: Optional[str] = Field(
        default="template_default_withgroups2019.pptx",
        description="Template file for slides containing multiple grouped candidates.",
    )
    group_template: Optional[str] = Field(
        default="template_default_withgroup_2019.pptx",
        description="Template file for group header slides (A, B, C, etc.).",
    )
    separator_template: Optional[str] = Field(
        default="template_default_seperator_2019.pptx",
        description="Template file for separator slides between sections.",
    )
    summary_template: Optional[str] = Field(
        default="template_default_summary2019.pptx",
        description="Template file for summary slides at the end of presentation.",
    )
    include_macro_version: bool = Field(
        default=False,
        description="Generate a macro-enabled (.pptm) version of the PowerPoint file.",
    )
    generate_physical_pptx: bool = Field(
        default=True,
        description="Generate physical PowerPoint file combining original slides and template-generated slides.",
    )


class CreatePresentationResponse(BaseModel):
    """Response model for presentation creation."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "message": "Presentation created successfully",
                "presentation_id": 12345,
                "total_slides": 24,
                "excel_data": {"slides": [], "total_slides": 24},
                "pptx_data": {
                    "message": "Conversion completed successfully",
                    "conversion_id": "NW_PROJECT_2025",
                    "total_images": 24,
                },
                "processing_time_seconds": 42.18,
            }
        }
    )

    message: str = Field(
        ..., description="Human-readable status message describing the outcome."
    )
    presentation_id: Optional[int] = Field(
        default=None,
        description="Identifier of the newly created presentation record in the database, when available.",
    )
    total_slides: int = Field(
        0, description="Total number of slides generated for the presentation."
    )
    excel_data: ProcessedExcelData = Field(
        ..., description="Detailed breakdown of processed Excel content used for slide generation."
    )
    pptx_data: PPTXConversionResponse = Field(
        ..., description="Metadata about the PPTX conversion process and generated images."
    )
    processing_time_seconds: float = Field(
        ..., description="Total processing time spent inside the orchestration pipeline."
    )


class DetailItem(BaseModel):
    """Represents a single record in the nw_Details table."""

    slide_number: int = Field(
        ..., ge=1, description="Slide number within the presentation sequence."
    )
    slide_type: str = Field(
        ..., description="Slide classification (for example 'Group', 'Name', 'Summary')."
    )
    slide_bg_file_name: Optional[str] = Field(
        default=None,
        description="Background file applied to the slide, if any.",
    )
    slide_description: Optional[str] = Field(
        default=None, description="Narrative description associated with the slide."
    )
    group_name: Optional[str] = Field(
        default=None, description="Name of the group when the slide represents a group summary."
    )
    category: Optional[str] = Field(
        default=None, description="Category of the candidate or slide grouping."
    )
    name: Optional[str] = Field(
        default=None, description="Primary candidate or item name displayed on the slide."
    )
    rationale: Optional[str] = Field(
        default=None, description="Supporting rationale or commentary."
    )
    notation: Optional[str] = Field(
        default=None, description="Phonetic notation or transliteration of the candidate."
    )
    kana: Optional[str] = Field(
        default=None, description="Japanese kana rendering when phonetics are enabled."
    )
    logo_filename: Optional[str] = Field(
        default=None,
        description="Optional logo filename associated with the slide content.",
    )
    template_id: int = Field(
        ..., description="Template identifier used when generating this slide."
    )
    name_sub_group: Optional[str] = Field(
        default=None, description="Sub-group label for nested grouping scenarios."
    )


class PresentationData(BaseModel):
    """Represents the complete payload for creating a new presentation."""

    project: str = Field(
        ..., description="Project identifier referencing the master record."
    )
    display_name: str = Field(
        ..., description="User-friendly name assigned to the presentation."
    )
    powerpoint_file: str = Field(
        ..., description="Filename of the generated PowerPoint asset."
    )
    excel_file: str = Field(
        ..., description="Filename of the source Excel workbook."
    )
    background_type: str = Field(
        ..., description="Background selection type stored with the presentation."
    )
    background_name: str = Field(
        ..., description="Background asset name stored with the presentation."
    )
    page_number: int = Field(
        ..., description="Page numbering offset saved to the master record."
    )
    presentation_type: str = Field(
        ..., description="Presentation classification stored with the master record."
    )
    user_name: str = Field(
        ..., description="User responsible for creating the presentation."
    )
    bsr_display_name: str = Field(
        ..., description="Brand Study Results label persisted for reporting."
    )
    mobile_link_bsr: Optional[str] = Field(
        default=None, description="Optional BSR hyperlink for mobile experiences."
    )
    participant_vote: int = Field(
        ..., description="Participant vote flag stored with the presentation."
    )
    is_wide_ppt: int = Field(
        ..., description="Flag indicating whether widescreen templates were used."
    )
    is_aws_email: int = Field(
        ..., description="Flag indicating whether AWS email processes should run."
    )
    is_design_mode: bool = Field(
        False,
        description="Whether the presentation was generated with design-mode placeholders.",
    )
    is_printed: bool = Field(
        False, description="Mark presentations generated for printing workflows"
    )
    details: List[DetailItem] = Field(
        ..., description="Collection of slide detail records persisted in nw_Details."
    )


class PresentationBuildOptions(BaseModel):
    """Options controlling how PPT files are assembled from Excel data."""

    template_pack: str = Field(
        "BackgroundDefaultTemplate",
        description="Directory containing the set of templates to use during assembly.",
    )
    base_template: str = Field(
        "template_default_2019.pptx",
        description="Default template applied to candidate slides when no rotation is provided.",
    )
    multi_template: str = Field(
        "template_default_withgroups2019.pptx",
        description="Template used for slides that contain multiple candidates grouped together.",
    )
    group_template: str = Field(
        "template_default_withgroup_2019.pptx",
        description="Template applied to group summary slides.",
    )
    separator_template: str = Field(
        "template_default_seperator_2019.pptx",
        description="Template dropped between sections when separator slides are requested.",
    )
    summary_template: str = Field(
        "template_default_summary2019.pptx",
        description="Template applied to summary slides at the end of the deck.",
    )
    templates_root: Optional[str] = Field(
        default=None,
        description="Override for the absolute path containing template packs (defaults to configuration).",
    )

    slide_start: int = Field(
        1,
        ge=1,
        description="First slide number to generate from the Excel candidates (1-based).",
    )
    slide_end: Optional[int] = Field(
        default=None,
        ge=1,
        description="Optional last slide number to generate; if omitted all candidates are processed.",
    )

    output_filename: Optional[str] = Field(
        default=None,
        description="Custom filename prefix for generated files; defaults to metadata.display_name.",
    )
    include_macro_version: bool = Field(
        False,
        description="Generate a macro-enabled (.pptm) version when templates include macros.",
    )
    include_print_ready_version: bool = Field(
        True,
        description="Produce a ready-to-print PPTX version with hidden helper slides removed.",
    )
    return_bytes: bool = Field(
        False,
        description="Inline the generated files as base64 strings in the API response.",
    )
    return_urls: bool = Field(
        True,
        description="Include downloadable URLs for generated files in the API response.",
    )

    def output_basename(self, display_name: str) -> str:
        """Return the sanitized filename prefix for generated files."""

        candidate = self.output_filename or display_name
        return sanitize_folder_name(candidate)


class PresentationBuildMetadata(CreatePresentationMetadata):
    """Metadata supplied for the on-demand PPT assembly endpoint."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                **CreatePresentationMetadata.model_config["json_schema_extra"][
                    "example"
                ],
                "template_pack": "TemplateWithGroups",
                "base_template": "template_default_withgroups2019.pptx",
                "slide_start": 1,
                "slide_end": 10,
                "output_filename": "NW_Project_2025_Presentation",
                "include_macro_version": True,
                "include_print_ready_version": True,
                "return_bytes": True,
                "return_urls": True,
            }
        }
    )

    template_pack: str = Field(
        "BackgroundDefaultTemplate",
        description="Template pack folder to load relative to the configured templates root.",
    )
    base_template: str = Field(
        "template_default_2019.pptx",
        description="Default PPTX applied to slides when no group-specific template is required.",
    )
    multi_template: str = Field(
        "template_default_withgroups2019.pptx",
        description="PPTX applied when multiple candidates share the same slide.",
    )
    group_template: str = Field(
        "template_default_withgroup_2019.pptx",
        description="PPTX for group overview slides.",
    )
    separator_template: str = Field(
        "template_default_seperator_2019.pptx",
        description="PPTX used to separate sections in the deck.",
    )
    summary_template: str = Field(
        "template_default_summary2019.pptx",
        description="Template for summary slides appended at the end of the deck.",
    )
    templates_root: Optional[str] = Field(
        default=None,
        description="Optional override for the base directory containing template packs.",
    )

    slide_start: int = Field(
        1,
        ge=1,
        description="First candidate row (1-based) from the Excel sheet to include in the generated deck.",
    )
    slide_end: Optional[int] = Field(
        default=None,
        ge=1,
        description="Last candidate row (1-based) to include; if omitted, all remaining rows are processed.",
    )

    output_filename: Optional[str] = Field(
        default=None,
        description="Custom filename prefix for generated files; defaults to the display_name value.",
    )
    include_macro_version: bool = Field(
        False,
        description="Generate a macro-enabled version of the deck when macros are available.",
    )
    include_print_ready_version: bool = Field(
        True,
        description="Produce a print-optimized PPTX alongside the standard output.",
    )
    return_bytes: bool = Field(
        False,
        description="Return base64-encoded file contents inline in the API response.",
    )
    return_urls: bool = Field(
        True,
        description="Return URLs pointing to generated files stored on disk.",
    )

    @model_validator(mode="after")
    def _validate_slide_range(self) -> "PresentationBuildMetadata":
        if self.slide_start < 1:
            raise ValueError("slide_start must be greater than or equal to 1")
        if self.slide_end is not None and self.slide_end < self.slide_start:
            raise ValueError("slide_end must be greater than or equal to slide_start")
        return self

    def to_build_options(self) -> PresentationBuildOptions:
        """Convert metadata into internal builder options."""

        return PresentationBuildOptions(
            template_pack=self.template_pack,
            base_template=self.base_template,
            multi_template=self.multi_template,
            group_template=self.group_template,
            separator_template=self.separator_template,
            summary_template=self.summary_template,
            templates_root=self.templates_root,
            slide_start=self.slide_start,
            slide_end=self.slide_end,
            output_filename=self.output_filename,
            include_macro_version=self.include_macro_version,
            include_print_ready_version=self.include_print_ready_version,
            return_bytes=self.return_bytes,
            return_urls=self.return_urls,
        )

    def to_creation_request(
        self,
        *,
        excel_content: bytes,
        excel_filename: str,
        pptx_filename: Optional[str] = None,
    ) -> CreatePresentationRequest:
        """Generate a CreatePresentationRequest without requiring an uploaded PPTX."""

        return super().to_service_request(
            excel_content=excel_content,
            excel_filename=excel_filename,
            pptx_content=b"",
            pptx_filename=pptx_filename or self.base_template,
        )


class PresentationBuildResponse(BaseModel):
    """Response payload for the PPT assembly endpoint."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "message": "Presentation assembled in 12.45s",
                "total_slides": 18,
                "slide_start": 1,
                "slide_end": 18,
                "printable_pptx": "C:/inetpub/wwwroot/nw_slides/NW_Project_2025/printable.pptx",
                "macro_pptx": "C:/inetpub/wwwroot/nw_slides/NW_Project_2025/macro.pptm",
                "printable_base64": "UEsDBBQAAAAIA...",
                "macro_base64": None,
                "download_urls": {
                    "printable": "/static/nw/NW_Project_2025/printable.pptx",
                    "macro": "/static/nw/NW_Project_2025/macro.pptm",
                },
                "api_downloads": {
                    "printable": "/api/presentations/files/eyJwYXRoIjogIkM6XFxcinetpubXFxcXG..."
                },
                "summary": {"groups": ["Group A", "Group B"]},
                "warnings": ["Slide 5 skipped: missing logo asset"],
            }
        }
    )

    message: str = Field(
        ..., description="Human-readable summary of the assembly operation outcome."
    )
    total_slides: int = Field(
        ..., ge=0, description="Total number of slides generated during assembly."
    )
    slide_start: int = Field(
        ..., ge=1, description="First slide number included in the generated deck."
    )
    slide_end: Optional[int] = Field(
        default=None, ge=1, description="Last slide number included in the generated deck."
    )

    printable_pptx: Optional[str] = Field(
        default=None,
        description="Absolute path to the print-ready PPTX when persisted to disk.",
    )
    macro_pptx: Optional[str] = Field(
        default=None,
        description="Absolute path to the macro-enabled PPTM when generated.",
    )

    printable_base64: Optional[str] = Field(
        default=None,
        description="Base64-encoded printable PPTX when return_bytes is enabled.",
    )
    macro_base64: Optional[str] = Field(
        default=None,
        description="Base64-encoded macro PPTM when return_bytes is enabled.",
    )

    download_urls: Dict[str, str] = Field(
        default_factory=dict,
        description="Mapping of artifact type to relative file-system paths for troubleshooting.",
    )
    api_downloads: Dict[str, str] = Field(
        default_factory=dict,
        description="Mapping of artifact type to authenticated API download URLs.",
    )
    summary: Dict[str, List[str]] = Field(
        default_factory=dict,
        description="Aggregated summary information (such as group listings or warnings).",
    )
    warnings: List[str] = Field(
        default_factory=list,
        description="Collection of non-fatal warnings encountered during assembly.",
    )