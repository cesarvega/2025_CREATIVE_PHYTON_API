"""
Response models for API endpoints.
"""

from typing import List

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
