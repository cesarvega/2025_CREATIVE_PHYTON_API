"""Models for NW Reports and Download Results functionality."""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import List, Optional

from pydantic import BaseModel, Field


class ReportType(str, Enum):
    """Types of reports that can be generated."""
    EXCEL = "excel"
    WORD = "word"
    ANALYTICS = "analytics"
    BSR = "bsr"


class SummaryType(str, Enum):
    """Summary types for Word report phonetics."""
    BY_THE_NUMBERS = "ByTheNumbers"
    POSITIVE = "Positive"
    NEUTRAL = "Neutral"
    RECONSIDER = "Reconsider"
    NEW_NAMES = "NewNames"
    EXPLORE = "Explore"
    AVOID = "Avoid"
    NOTES = "Notes"
    STATISTICS = "Statistics"
    POSITIVE_PHONETICS = "Positive_Phonetics"
    NEGATIVE_PHONETICS = "Negative_Phonetics"
    RECONSIDER_PHONETICS = "Reconsider_Phonetics"


class VoteValue(str, Enum):
    """Vote values transformation."""
    NEGATIVE = "Negative"
    NEUTRAL = "Neutral"
    POSITIVE = "Positive"


# Excel Report Models

class RetainedName(BaseModel):
    """Retained name from nw_dlRetainedNames_withRecraft SP."""
    name: str
    category: Optional[str] = None
    rationale: Optional[str] = None
    vote: Optional[str] = None
    recraft_flag: Optional[bool] = False


class NewlyCreatedName(BaseModel):
    """Newly created name from nw_CombineNewNames SP."""
    name: str
    category: Optional[str] = None
    rationale: Optional[str] = None


class RootConcept(BaseModel):
    """Root or concept to explore/avoid."""
    concept: str
    description: Optional[str] = None


class OpenNote(BaseModel):
    """Open note from nw_dlOpenNotes SP."""
    note_id: int
    note_text: str
    created_by: Optional[str] = None
    created_date: Optional[datetime] = None


class VoteByGroup(BaseModel):
    """Vote results by group from nw_Votesbygroups SP."""
    group_name: str
    positive_votes: int = 0
    neutral_votes: int = 0
    negative_votes: int = 0
    total_votes: int = 0


class VotedParticipant(BaseModel):
    """Participant who voted from nw_VotedParticipants SP."""
    participant_id: int
    participant_name: str
    email: Optional[str] = None
    voted_date: Optional[datetime] = None


# Word Report Models

class WordReportReplacement(BaseModel):
    """Value to replace in Word template from nw_wdValuesToReplace SP."""
    placeholder: str
    value: str


class WordReportResult(BaseModel):
    """Result for Word report from nw_wdGetResults or nw_wdGetResults_Phonetics SP."""
    name: str
    pronunciation: Optional[str] = None  # For phonetics reports
    category: Optional[str] = None
    rationale: Optional[str] = None
    vote: Optional[str] = None
    name_rationale_part1: Optional[str] = None
    name_rationale_part2: Optional[str] = None

    # For ByTheNumbers summary type
    positive_count: Optional[int] = None
    neutral_count: Optional[int] = None
    reconsider_count: Optional[int] = None
    new_names_count: Optional[int] = None
    total_count: Optional[int] = None


# Analytics Report Models

class ProjectAnalytics(BaseModel):
    """Project analytics data from NW_ProjectAnalytics SP."""
    metric_name: str
    metric_value: str


class RegionSpecificAnalytics(BaseModel):
    """Region-specific analytics from NW_RegionSpecificAnalytics SP."""
    region: str
    metric_name: str
    metric_value: str


# BSR Report Models

class BSRProjectConcept(BaseModel):
    """BSR project concept."""
    concept_id: int
    concept_name: str
    description: Optional[str] = None


class BSRPageComment(BaseModel):
    """BSR page comment."""
    comment_id: int
    page_number: int
    comment_text: str
    created_by: Optional[str] = None


class BSRExcelReportData(BaseModel):
    """BSR Excel report data from bsr_GetExcelReport SP."""
    project_id: int
    project_name: str
    data: dict


# Request Models

class DownloadResultsRequest(BaseModel):
    """Request for downloading NW results.

    Only ``presentation_id`` is required. The API always generates both
    the Excel results workbook and the Word report document. Any logic for
    including votes/participants sheets is handled internally.
    """
    presentation_id: int = Field(
        ..., description="Presentation ID to generate results for"
    )


# Feedback Template Models

class CreateFeedbackTemplateRequest(BaseModel):
    """Request for creating a Feedback Template document."""
    presentation_id: int = Field(..., description="Presentation ID to generate feedback template for")


class CreateFeedbackTemplateResponse(BaseModel):
    """Response for create feedback template endpoint."""
    success: bool = True
    message: str = "Feedback template generated successfully"
    file_path: Optional[str] = None
    file_name: Optional[str] = None
    download_token: Optional[str] = None
    presentation_id: int
    presentation_type: Optional[str] = None
    generated_at: datetime = Field(default_factory=datetime.utcnow)
    warnings: List[str] = Field(default_factory=list)


# Response Models

class DownloadResultsResponse(BaseModel):
    """Response for download results endpoint."""
    success: bool = True
    message: str = "Report generated successfully"
    file_path: Optional[str] = None
    file_name: Optional[str] = None
    download_token: Optional[str] = None
    # Optional explicit tokens and filenames for each artifact
    excel_download_token: Optional[str] = None
    word_download_token: Optional[str] = None
    excel_file_name: Optional[str] = None
    word_file_name: Optional[str] = None
    report_type: ReportType
    presentation_id: int
    generated_at: datetime = Field(default_factory=datetime.utcnow)
    warnings: List[str] = Field(default_factory=list)
