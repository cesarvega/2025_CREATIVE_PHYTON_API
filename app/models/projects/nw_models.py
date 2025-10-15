"""
NW project-specific models.

Models for Name Writing (NW) presentations that inherit from base models
and add NW-specific fields and validation.
"""

from typing import Optional
from pydantic import Field

from app.models.projects.base_models import (
    BaseCreatePresentationMetadata,
    BaseCreatePresentationRequest,
    BaseCreatePresentationResponse,
    BaseDetailItem,
    BasePresentationData,
)


class NWCreatePresentationMetadata(BaseCreatePresentationMetadata):
    """NW-specific metadata for presentation creation.

    Extends base metadata with NW-specific fields like mobile_link_bsr.
    """

    mobile_link_bsr: str = Field(
        "",
        description="Mobile link for BSR. Default: ''",
    )

    # Physical PowerPoint generation options (NW-specific)
    template_pack: Optional[str] = Field(
        default="BackgroundDefaultTemplate",
        description="Template pack folder for NW presentations.",
    )
    base_template: Optional[str] = Field(
        default="template_default_2019.pptx",
        description="Base template file for NW individual slides.",
    )
    multi_template: Optional[str] = Field(
        default="template_default_withgroups2019.pptx",
        description="Template file for NW grouped slides.",
    )
    group_template: Optional[str] = Field(
        default="template_default_withgroup_2019.pptx",
        description="Template file for NW group header slides.",
    )
    separator_template: Optional[str] = Field(
        default="template_default_seperator_2019.pptx",
        description="Template file for NW separator slides.",
    )
    summary_template: Optional[str] = Field(
        default="template_default_summary2019.pptx",
        description="Template file for NW summary slides.",
    )
    include_macro_version: bool = Field(
        default=False,
        description="Generate macro-enabled (.pptm) version for NW.",
    )
    generate_physical_pptx: bool = Field(
        default=True,
        description="Generate physical PowerPoint file for NW presentations.",
    )

    def to_service_request(
        self,
        *,
        excel_content: bytes,
        excel_filename: str,
        pptx_content: bytes,
        pptx_filename: str,
    ) -> "NWCreatePresentationRequest":
        """Convert metadata into an NWCreatePresentationRequest."""

        return NWCreatePresentationRequest(
            excel_file=excel_content,
            pptx_file=pptx_content,
            excel_filename=excel_filename,
            pptx_filename=pptx_filename,
            has_groups=self.has_groups,
            test_name_order=self.test_name_order,
            project=self.project,
            display_name=self.display_name,
            presentation_type=self.presentation_type,
            user_name=self.user_name,
            mobile_link_bsr=self.mobile_link_bsr,
            participant_vote=self.participant_vote,
            is_wide_ppt=self.is_wide_ppt,
            is_aws_email=self.is_aws_email,
            background_type=self.background_type,
            background_name=self.background_name or "",
            page_number=self.page_number,
            project_type=str(self.project_type),
            template_pack=self.template_pack,
            base_template=self.base_template,
            multi_template=self.multi_template,
            group_template=self.group_template,
            separator_template=self.separator_template,
            summary_template=self.summary_template,
            include_macro_version=self.include_macro_version,
            generate_physical_pptx=self.generate_physical_pptx,
        )


class NWCreatePresentationRequest(BaseCreatePresentationRequest):
    """NW-specific request model for creating presentations.

    Extends base request with NW-specific fields.
    """

    mobile_link_bsr: Optional[str] = Field(
        default=None, description="Mobile link for BSR in NW projects."
    )

    # Physical PowerPoint generation options for NW
    template_pack: Optional[str] = Field(
        default="BackgroundDefaultTemplate",
        description="Template pack folder for NW.",
    )
    base_template: Optional[str] = Field(
        default="template_default_2019.pptx",
        description="Base template for NW slides.",
    )
    multi_template: Optional[str] = Field(
        default="template_default_withgroups2019.pptx",
        description="Multi-candidate template for NW.",
    )
    group_template: Optional[str] = Field(
        default="template_default_withgroup_2019.pptx",
        description="Group template for NW.",
    )
    separator_template: Optional[str] = Field(
        default="template_default_seperator_2019.pptx",
        description="Separator template for NW.",
    )
    summary_template: Optional[str] = Field(
        default="template_default_summary2019.pptx",
        description="Summary template for NW.",
    )
    include_macro_version: bool = Field(
        default=False,
        description="Generate .pptm version for NW.",
    )
    generate_physical_pptx: bool = Field(
        default=True,
        description="Generate physical PPTX for NW.",
    )


class NWCreatePresentationResponse(BaseCreatePresentationResponse):
    """NW-specific response model for presentation creation.

    Extends base response with NW-specific response fields if needed.
    Currently uses base response structure.
    """
    pass


class NWDetailItem(BaseDetailItem):
    """NW-specific detail item model.

    Extends base detail item with NW-specific fields if needed.
    Currently uses base detail structure.
    """
    pass


class NWPresentationData(BasePresentationData):
    """NW-specific presentation data model.

    Extends base presentation data with NW-specific fields.
    """

    mobile_link_bsr: Optional[str] = Field(
        default=None, description="Mobile link for BSR in NW presentations."
    )
