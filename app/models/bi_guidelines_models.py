"""
BI Guidelines specific models for database operations.
"""

from pydantic import BaseModel, Field
from typing import List, Optional


class NWMasterRequest(BaseModel):
    project: str = Field(..., min_length=1, max_length=100)
    uploaded_by: str = Field(..., min_length=1, max_length=50)
    presentation_status: str = Field(..., min_length=1, max_length=50)