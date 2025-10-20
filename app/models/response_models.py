"""
Response models for API endpoints.
"""

from datetime import datetime
from typing import List, Optional

from pydantic import BaseModel, Field


class PPTXConversionResponse(BaseModel):
    """Response model for PPTX conversion."""

    message: str
    conversion_id: str
    project_type: str
    total_images: int
    images: List[str]
    thumbnails: List[str] = []
    titles: List[str]
    pptx_file: str = ""


class ProjectsResponse(BaseModel):
    """Response model for paginated projects list."""

    projects: List[str]
    page: int
    limit: int
    total: int


class ActivePresentation(BaseModel):
    """Model for a single active presentation."""

    presentation_id: int
    project: str
    display_name: str
    uploaded_by: Optional[str] = None
    uploaded_date: Optional[datetime] = None
    link: str
    presentation_status: str
    last_update_date: Optional[datetime] = None


class ActivePresentationsResponse(BaseModel):
    """Response model for active presentations list with pagination."""

    presentations: List[ActivePresentation]
    page: int
    limit: int
    total: int


class DisplayNamesResponse(BaseModel):
    """Response model for BSR display names list with pagination."""

    display_names: List[str]
    page: int
    limit: int
    total: int


class TemplateGroup(BaseModel):
    """Model for a single template group (background template).

    Represents a background image template from the nw_Templates table.
    The template_file_name contains the relative path to the background image
    (e.g., 'images/BackGrounds/Backgrounds2019/BrandDNA.jpg').
    """

    template_group_id: int
    template_name: str
    category: Optional[str] = None
    template_file_name: Optional[str] = Field(
        default=None,
        description="Relative path to the background image file"
    )


class TemplateGroupsResponse(BaseModel):
    """Response model for template groups list."""

    template_groups: List[TemplateGroup]
    total: int


class ReplaceProjectImagesResponse(BaseModel):
    """Response model for replacing slide images of an existing project."""

    message: str
    project_name: str
    project_type: str
    total_images: int
    images: List[str]
    thumbnails: List[str] = []


class CreateSimpleDWResponse(BaseModel):
    """Response model for simplified DW PPTX-to-images + DB detail insertion."""

    message: str
    project_name: str
    display_name: str
    project_type: str  # expected 'dw'
    slide_type: str  # e.g., 'Design'
    user_name: str
    total_images: int
    images: List[str]
    thumbnails: List[str] = []
    presentation_id: Optional[int] = None
