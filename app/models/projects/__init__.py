"""
Project-specific models package.

This package contains models organized by project type (NW, DW, BSR, NSR).
Each project type has its own models that inherit from base models.
"""

from app.models.projects.base_models import (
    BaseCreatePresentationMetadata,
    BaseCreatePresentationRequest,
    BaseCreatePresentationResponse,
    BaseDetailItem,
    BasePresentationData,
)

__all__ = [
    "BaseCreatePresentationMetadata",
    "BaseCreatePresentationRequest",
    "BaseCreatePresentationResponse",
    "BaseDetailItem",
    "BasePresentationData",
]
