"""
PowerPoint conversion models.
"""

from typing import List, Optional

from pydantic import BaseModel


class PPTXProcessingRequest(BaseModel):
    """Request model for PowerPoint processing."""

    presentation_id: Optional[str] = None
    project_type: str = "bipresents"  # "bipresents" or "nw"


class PPTXProcessingResponse(BaseModel):
    """Response model for PowerPoint processing."""

    message: str
    conversion_id: str
    total_slides: int
    image_urls: List[str]
    thumbnail_urls: List[str]
    slide_titles: List[str]