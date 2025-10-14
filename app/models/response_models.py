"""
Response models for API endpoints.
"""

from datetime import datetime
from typing import List, Optional

from pydantic import BaseModel


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
    """Model for a single template group."""

    template_group_id: int
    template_name: str
    category: Optional[str] = None


class TemplateGroupsResponse(BaseModel):
    """Response model for template groups list."""

    template_groups: List[TemplateGroup]
    total: int
