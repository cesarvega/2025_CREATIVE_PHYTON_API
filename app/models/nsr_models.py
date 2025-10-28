"""Models for NSR (Name Selection Research) configuration."""

from __future__ import annotations

from typing import List, Optional

from pydantic import BaseModel, Field, ConfigDict


class NSRRuleConfig(BaseModel):
    """NSR rule configuration."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "rule_id": 101,
                "is_on": 1
            }
        }
    )

    rule_id: int = Field(..., ge=101, le=117, description="Rule ID (101-117)")
    is_on: int = Field(..., ge=0, le=1, description="Rule state: 0=OFF, 1=ON")


class NSRProjectConfigResponse(BaseModel):
    """Response model for NSR project configuration."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "project_name": "ATEST01",
                "rules": [
                    {"rule_id": 101, "is_on": 1},
                    {"rule_id": 102, "is_on": 0}
                ]
            }
        }
    )

    project_name: str = Field(..., description="Project display name")
    rules: List[NSRRuleConfig] = Field(default_factory=list, description="List of rule configurations")


class NSRUpdateRuleRequest(BaseModel):
    """Request model for updating NSR rule."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "project_name": "ATEST01",
                "rule_id": 101,
                "is_on": 1
            }
        }
    )

    project_name: str = Field(..., min_length=1, description="Project display name")
    rule_id: int = Field(..., ge=101, le=117, description="Rule ID (101-117)")
    is_on: int = Field(..., ge=0, le=1, description="Rule state: 0=OFF, 1=ON")


class NSRUpdateRuleResponse(BaseModel):
    """Response model for NSR rule update."""

    message: str = Field(..., description="Status message")
    project_name: str = Field(..., description="Project display name")
    rule_id: int = Field(..., description="Rule ID that was updated")
    is_on: int = Field(..., description="New state of the rule")
    operation: str = Field(..., description="Operation performed: 'insert' or 'update'")


class NSRProjectsListResponse(BaseModel):
    """Response model for NSR projects list."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "success": True,
                "data": ["Project_Alpha_NSR", "Project_Beta_NSR", "Project_Gamma_NSR"]
            }
        }
    )

    success: bool = Field(default=True, description="Request success status")
    data: List[str] = Field(..., description="List of NSR project display names")


class NSRRuleExistsResponse(BaseModel):
    """Response model for checking if NSR rule exists."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "success": True,
                "exists": True
            }
        }
    )

    success: bool = Field(default=True, description="Request success status")
    exists: bool = Field(..., description="Whether the rule exists for the project")


class NSRInitializeResponse(BaseModel):
    """Response model for initializing NSR rules."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "success": True,
                "message": "All rules initialized successfully",
                "displayName": "NewProject_NSR",
                "rulesCreated": 16
            }
        }
    )

    success: bool = Field(..., description="Operation success status")
    message: str = Field(..., description="Status message")
    displayName: str = Field(..., description="Project display name")
    rulesCreated: int = Field(..., description="Number of rules created")


# NSR Rule Definitions for documentation
NSR_RULES = {
    # Panel2 - Contains
    101: {"name": "Contains Y", "type": "contains", "value": "Y"},
    102: {"name": "Contains H", "type": "contains", "value": "H"},
    103: {"name": "Contains W", "type": "contains", "value": "W"},
    104: {"name": "Contains J", "type": "contains", "value": "J"},
    105: {"name": "Contains K", "type": "contains", "value": "K"},

    # Panel5 - USAN Checks
    106: {"name": "Check USAN Violation", "type": "usan", "value": "violation"},
    108: {"name": "Check USAN Nomenclature", "type": "usan", "value": "nomenclature"},
    116: {"name": "Check USAN Search", "type": "usan", "value": "search"},
    117: {"name": "Check USAN MedNet", "type": "usan", "value": "mednet"},

    # Panel3 - Prefix
    109: {"name": "Prefix AR", "type": "prefix", "value": "AR"},
    110: {"name": "Prefix DEX", "type": "prefix", "value": "DEX"},
    111: {"name": "Prefix ES", "type": "prefix", "value": "ES"},
    112: {"name": "Prefix STR", "type": "prefix", "value": "STR"},
    113: {"name": "Prefix LEV", "type": "prefix", "value": "LEV"},
    114: {"name": "Prefix X", "type": "prefix", "value": "X"},
    115: {"name": "Prefix RAC", "type": "prefix", "value": "RAC"},
}
