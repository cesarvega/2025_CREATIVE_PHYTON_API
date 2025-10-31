"""Models for Global Analytics Statistics Reports.

This module contains Pydantic models for the Statistics tab functionality,
which generates Excel reports with data from ALL projects (not individual projects).

IMPORTANT: These are DIFFERENT from the individual project reports in nw_reports_models.py
"""

from __future__ import annotations

from typing import List, Literal

from pydantic import BaseModel, Field


# NW Analytics Models

class NWProjectAnalytics(BaseModel):
    """Analytics data for a single NW project (from global statistics).

    This represents one row in the "Project-Specific" sheet of the NW Analytics Excel report.
    Data is aggregated from ALL NW projects, not a single project.

    Column order matches the Excel template exactly:
    A: Region, B: Active Lead, C: Project Name, D: DisplayName,
    E: ClientName, F: Percent Retained, G: Percent Positive,
    H: Percent Neutral, I: Percent Negative, J: Names Created
    """
    region: str = Field(default="", description="Region of the project (US, EU, APAC, etc.)")
    active_lead: str = Field(default="", description="Active lead/director of the project")
    project_name: str = Field(default="", description="Project name")
    display_name: str = Field(default="", description="Display name of the presentation")
    client_name: str = Field(default="", description="Client name")
    percent_retained: str = Field(default="0%", description="Percentage retained (Positive + Neutral)")
    percent_positive: str = Field(default="0%", description="Percentage of positive votes")
    percent_neutral: str = Field(default="0%", description="Percentage of neutral votes")
    percent_negative: str = Field(default="0%", description="Percentage of negative votes")
    names_created: int = Field(default=0, description="Number of newly created names")


class NWRegionAnalytics(BaseModel):
    """Aggregated analytics data by region for NW projects.

    This represents one row in the "Region-Specific" sheet of the NW Analytics Excel report.
    Data is aggregated across all projects in a region.
    """
    region: str = Field(..., description="Region name")
    active_lead: str = Field(default="", description="Active lead for the region")
    projects_to_date: int = Field(default=0, description="Total number of projects in the region")
    percent_retained: str = Field(default="0%", description="Aggregated percentage retained")
    percent_positive: str = Field(default="0%", description="Aggregated percentage positive")
    percent_neutral: str = Field(default="0%", description="Aggregated percentage neutral")
    percent_negative: str = Field(default="0%", description="Aggregated percentage negative")
    avg_names_per_project: float = Field(default=0.0, description="Average names per project")


# BSR Analytics Models

class BSRProjectAnalytics(BaseModel):
    """Analytics data for a single BSR project (from global statistics).

    This represents one row in the "Project-Specific" sheet of the BSR Analytics Excel report.

    Column order matches the Excel template exactly:
    A: Region, B: Active Lead, C: Project Name, D: DisplayName,
    E: ClientName, F: Names Created (Web), G: Names Created (Mobile),
    H: Total Names Created
    """
    region: str = Field(default="", description="Region of the project")
    active_lead: str = Field(default="", description="Active lead/director")
    project_name: str = Field(default="", description="Project name")
    display_name: str = Field(default="", description="Display name")
    client_name: str = Field(default="", description="Client name")
    names_created_web: int = Field(default=0, description="Number of names created via Web (PC)")
    names_created_mobile: int = Field(default=0, description="Number of names created via Mobile")
    total_names_created: int = Field(default=0, description="Total names created (Web + Mobile)")


class BSRRegionAnalytics(BaseModel):
    """Aggregated analytics data by region for BSR projects.

    This represents one row in the "Region-Specific" sheet of the BSR Analytics Excel report.
    """
    region: str = Field(..., description="Region name")
    active_lead: str = Field(default="", description="Active lead for the region")
    projects_to_date: int = Field(default=0, description="Total number of projects")
    pc_count: int = Field(default=0, description="Aggregated PC access count")
    mobile_count: int = Field(default=0, description="Aggregated mobile access count")
    total: int = Field(default=0, description="Total accesses (PC + Mobile)")


# Request/Response Models

class AnalyticsDownloadRequest(BaseModel):
    """Request body for downloading analytics Excel report.

    The client must first call the GET endpoints to retrieve projects and regions data,
    then pass that data to this endpoint for Excel generation.
    """
    project_type: Literal["NW", "BSR"] = Field(
        ...,
        description="Type of project report to generate (NW or BSR)"
    )
    projects: List[dict] = Field(
        default_factory=list,
        description="List of project analytics data (from /api/analytics/{type}/projects)"
    )
    regions: List[dict] = Field(
        default_factory=list,
        description="List of region analytics data (from /api/analytics/{type}/regions)"
    )


class AnalyticsResponse(BaseModel):
    """Response model for analytics GET endpoints."""
    success: bool = True
    message: str = "Analytics data retrieved successfully"
    data: List[dict] = Field(default_factory=list, description="Analytics data")
    total: int = Field(default=0, description="Total number of records")


class NWProjectAnalyticsResponse(BaseModel):
    """Response for NW project analytics endpoint."""
    success: bool = True
    message: str = "NW project analytics retrieved successfully"
    data: List[NWProjectAnalytics] = Field(default_factory=list)
    total: int = 0


class NWRegionAnalyticsResponse(BaseModel):
    """Response for NW region analytics endpoint."""
    success: bool = True
    message: str = "NW region analytics retrieved successfully"
    data: List[NWRegionAnalytics] = Field(default_factory=list)
    total: int = 0


class BSRProjectAnalyticsResponse(BaseModel):
    """Response for BSR project analytics endpoint."""
    success: bool = True
    message: str = "BSR project analytics retrieved successfully"
    data: List[BSRProjectAnalytics] = Field(default_factory=list)
    total: int = 0


class BSRRegionAnalyticsResponse(BaseModel):
    """Response for BSR region analytics endpoint."""
    success: bool = True
    message: str = "BSR region analytics retrieved successfully"
    data: List[BSRRegionAnalytics] = Field(default_factory=list)
    total: int = 0
