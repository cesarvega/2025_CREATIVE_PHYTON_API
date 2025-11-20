"""
BI Guidelines specific models for database operations.
"""

from datetime import datetime
from pydantic import BaseModel, Field, field_validator
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


# --- Custom Background Templates Models ---
# These models work with the existing nw_Templates table


class BackgroundTemplateCreate(BaseModel):
    """Request model for creating a new background template.
    
    Works with existing [BI_GUIDELINES].[dbo].[nw_Templates] table.
    Only supports JPG/PNG image uploads.
    """
    template_name: str = Field(
        ...,
        min_length=3,
        max_length=100,
        description="Name of the background template (unique)"
    )
    template_group: str = Field(
        ...,
        min_length=3,
        max_length=100,
        description="Group/category for organizing templates (e.g., 'BMW Theme', 'Corporate')"
    )

    @field_validator("template_name", "template_group")
    @classmethod
    def validate_no_special_chars(cls, v: str) -> str:
        """Validate that name/group contain only alphanumeric, spaces, dashes, underscores."""
        import re
        if not re.match(r'^[a-zA-Z0-9 _-]+$', v):
            raise ValueError(
                f"'{v}' contains invalid characters. Only alphanumeric, spaces, dashes, and underscores allowed."
            )
        return v.strip()

    class Config:
        json_schema_extra = {
            "example": {
                "template_name": "BMW Background 2025",
                "template_group": "BMW Theme"
            }
        }


class BackgroundTemplateUpdate(BaseModel):
    """Request model for updating a background template.
    
    Only template_name and template_group can be updated.
    To change the image, delete and recreate.
    """
    template_name: Optional[str] = Field(
        None,
        min_length=3,
        max_length=100,
        description="New name for the template"
    )
    template_group: Optional[str] = Field(
        None,
        min_length=3,
        max_length=100,
        description="New group/category"
    )

    @field_validator("template_name", "template_group")
    @classmethod
    def validate_no_special_chars(cls, v: Optional[str]) -> Optional[str]:
        """Validate that name/group contain only alphanumeric, spaces, dashes, underscores."""
        if v is None:
            return v
        import re
        if not re.match(r'^[a-zA-Z0-9 _-]+$', v):
            raise ValueError(
                f"'{v}' contains invalid characters. Only alphanumeric, spaces, dashes, and underscores allowed."
            )
        return v.strip()

    class Config:
        json_schema_extra = {
            "example": {
                "template_name": "Updated BMW Background",
                "template_group": "BMW Theme 2025"
            }
        }

