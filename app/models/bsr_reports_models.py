"""Models for BSR Reports and Download Results functionality."""

from __future__ import annotations

from datetime import datetime
from typing import List, Optional

from pydantic import BaseModel, Field, ConfigDict


class BSRGenerateReportRequest(BaseModel):
    """Request model for generating BSR reports."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "presentation_id": 12345,
                "display_name": "Test_BSR_Presentation"
            }
        }
    )

    presentation_id: int = Field(
        ...,
        gt=0,
        description="BSR Presentation ID to generate reports for"
    )
    display_name: Optional[str] = Field(
        None,
        description="Display name of the presentation (optional, will be fetched if not provided)"
    )


class BSRGenerateReportResponse(BaseModel):
    """Response model for BSR report generation."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "message": "BSR reports generated successfully",
                "excel_file": "C:\\BRS_files\\downloads\\Test_Presentation.xls",
                "word_file": "C:\\BRS_files\\downloads\\Test_Presentation.doc",
                "processing_time_seconds": 12.5
            }
        }
    )

    message: str = Field(..., description="Status message")
    excel_file: Optional[str] = Field(None, description="Path to generated Excel file")
    word_file: Optional[str] = Field(None, description="Path to generated Word file")
    processing_time_seconds: float = Field(..., description="Total processing time")
    warnings: List[str] = Field(default_factory=list, description="Any warnings during generation")
    excel_download_token: Optional[str] = Field(None, description="Download token for Excel file")
    word_download_token: Optional[str] = Field(None, description="Download token for Word file")
    zip_file: Optional[str] = Field(None, description="Path to ZIP containing both files")
    zip_download_token: Optional[str] = Field(None, description="Download token for ZIP file")


# Models for BSR Excel Report Data


class BSRExcelReportRow(BaseModel):
    """Row data from bsr_GetExcelReport stored procedure."""

    id: Optional[int] = None
    name: str = Field(..., description="Candidate name")
    source: str = Field(..., description="Source of the name (PC/Mobile/Moderator)")
    is_mobile: str = Field(..., description="Mobile flag: '0' or '1'")
    concept: Optional[str] = Field(None, description="Associated concept")
    created_date: datetime = Field(..., description="Creation date")
    categories_elements: Optional[str] = Field(
        None, description="Comma-separated category values"
    )
    categories: Optional[str] = Field(
        None, description="Comma-separated category names"
    )


# Models for BSR Word Report Data


class BSRSlideNote(BaseModel):
    """Slide note/comment from BSR_PageComments table."""

    slide_description: str = Field(..., description="Slide description")
    comments: Optional[str] = Field(None, description="Slide comments")


class BSRProjectConcept(BaseModel):
    """Project concept from bsr_ProjectConcepts table."""

    concept: Optional[str] = Field(None, description="Concept name")
    html: str = Field(..., description="HTML content for the concept")


class BSRProjectClient(BaseModel):
    """Client information from DayMaster.Projects table."""

    client: str = Field(..., description="Client name")
