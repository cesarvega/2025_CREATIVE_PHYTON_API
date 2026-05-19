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
    project_type: Optional[str] = None
    presentation_type: Optional[str] = None


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
    
    The thumbnail_url provides a lightweight version (300x169px) for fast loading
    in list views, while preview_url provides the full-resolution original.
    """

    template_group_id: int
    template_name: str
    category: Optional[str] = None
    template_file_name: Optional[str] = Field(
        default=None,
        description="Relative path to the background image file"
    )
    thumbnail_url: Optional[str] = Field(
        None, 
        description="URL to lightweight thumbnail (300x169px) for fast loading in lists"
    )
    preview_url: Optional[str] = Field(
        None, 
        description="URL to full-resolution background image"
    )


class TemplateGroupsResponse(BaseModel):
    """Response model for template groups list."""

    template_groups: List[TemplateGroup]
    total: int
    custom_count: int = Field(default=0, description="Number of custom themes")
    system_count: int = Field(default=0, description="Number of system themes")


class CustomThemesListResponse(BaseModel):
    """Response model for background templates list (simplified)."""
    
    success: bool = True
    templates: List["TemplateGroup"]  # Forward reference
    total: int
    page: int = Field(default=1, ge=1)
    limit: int = Field(default=50, ge=1, le=100)


class CustomThemeCreateResponse(BaseModel):
    """Response model for background template creation."""
    
    success: bool = True
    message: str = "Background template created successfully"
    template: "TemplateGroup"  # Forward reference


class CustomThemeDeleteResponse(BaseModel):
    """Response model for background template deletion."""
    
    success: bool = True
    message: str = "Background template deleted successfully"
    deleted_template_id: int



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
    overwritten: bool = False
