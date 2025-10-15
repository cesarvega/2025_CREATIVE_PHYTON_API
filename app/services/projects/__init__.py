"""
Project-specific services package.

This package contains services organized by project type (NW, DW, BSR, NSR).
Each project type has its own service that inherits from BasePresentationService.
"""

from app.services.projects.base_presentation_service import BasePresentationService

__all__ = [
    "BasePresentationService",
]
