"""
BI Guidelines specific models for database operations.
"""

from pydantic import BaseModel, Field
from typing import List, Optional


class NWMasterRequest(BaseModel):
    project: str = Field(..., min_length=1, max_length=100)
    uploaded_by: str = Field(..., min_length=1, max_length=50)
    presentation_status: str = Field(..., min_length=1, max_length=50)


class ReloadProjectSoundsRequest(BaseModel):
    """Request model for reloading project sounds."""
    display_name: str = Field(
        ...,
        min_length=1,
        max_length=255,
        description="Display name of the project to reload sounds for"
    )

    class Config:
        json_schema_extra = {
            "example": {
                "display_name": "SOLE_19Nov2024"
            }
        }


class ProjectUpdateRequest(BaseModel):
    """Request model for updating NW project presentation master details.

    Maps to the nw_UpdatePresentationMaster stored procedure which updates:
    - DisplayName
    - PresentationStatus
    - BSRDisplayName
    - LastUpdateDate (automatic)
    """
    presentation_id: int = Field(
        ...,
        gt=0,
        description="ID of the presentation to update"
    )
    display_name: str = Field(
        ...,
        min_length=1,
        max_length=100,
        description="The new display name for the NW project"
    )
    presentation_status: str = Field(
        ...,
        min_length=1,
        max_length=10,
        description="The new presentation status (e.g., 'OPEN', 'CLOSED')"
    )
    bsr_display_name: Optional[str] = Field(
        None,
        max_length=100,
        description="The BSR display name associated with this presentation (optional)"
    )

    class Config:
        json_schema_extra = {
            "example": {
                "presentation_id": 8286,
                "display_name": "SOLE_19Nov2024",
                "presentation_status": "OPEN",
                "bsr_display_name": "SOLE_BSR_2024"
            }
        }


class BSRProjectUpdateRequest(BaseModel):
    """Request model for updating BSR project presentation master details.

    Maps to the BSR_UpdatePresentationMaster stored procedure which updates:
    - DisplayName
    - PresentationStatus
    """
    presentation_id: int = Field(
        ...,
        gt=0,
        description="ID of the BSR presentation to update"
    )
    display_name: str = Field(
        ...,
        min_length=1,
        max_length=100,
        description="The new display name for the BSR presentation"
    )
    status: str = Field(
        ...,
        min_length=1,
        max_length=20,
        description="The new presentation status (e.g., 'OPEN', 'CLOSED')"
    )

    class Config:
        json_schema_extra = {
            "example": {
                "presentation_id": 12345,
                "display_name": "Test_Presentation_Updated",
                "status": "CLOSED"
            }
        }