"""
Application-wide constants and enumerations.

This module centralizes all magic strings and constant values used throughout
the application to improve maintainability and reduce duplication.
"""

from typing import Final


# Presentation Types
class PresentationType:
    """Valid presentation type identifiers."""
    
    NORMAL: Final[str] = "Normal"
    NORMAL_NO_NEUTRAL: Final[str] = "Normal-NoNeutral"
    PHONETICS: Final[str] = "Phonetics"
    KATAKANA: Final[str] = "Katakana"
    KATAKANA_BIG_JAP: Final[str] = "Katakana_BigJap"
    NONPROPRIETARY: Final[str] = "Nonproprietary"
    TAGLINE: Final[str] = "Tagline"
    DESIGN: Final[str] = "Design"
    
    @classmethod
    def all(cls) -> list[str]:
        """Get all valid presentation types."""
        return [
            cls.NORMAL,
            cls.NORMAL_NO_NEUTRAL,
            cls.PHONETICS,
            cls.KATAKANA,
            cls.KATAKANA_BIG_JAP,
            cls.NONPROPRIETARY,
            cls.TAGLINE,
            cls.DESIGN,
        ]


# Project Types
class ProjectType:
    """Valid project type identifiers."""
    
    BIPRESENTS: Final[str] = "bipresents"
    BSR: Final[str] = "bsr"
    NSR: Final[str] = "nsr"
    NW: Final[str] = "nw"
    
    @classmethod
    def all(cls) -> list[str]:
        """Get all valid project types."""
        return [cls.BIPRESENTS, cls.BSR, cls.NSR, cls.NW]


# Slide Types
class SlideType:
    """Slide type identifiers used in presentations."""
    
    NAME_EVALUATION: Final[str] = "NameEvaluation"
    NAME_EVALUATION_NO_NEUTRAL: Final[str] = "NameEvaluation_noNeutral"
    BSR_GROUP: Final[str] = "BSR_Group"
    NSR_GROUP: Final[str] = "NSR_Group"
    
    @classmethod
    def all(cls) -> list[str]:
        """Get all valid slide types."""
        return [
            cls.NAME_EVALUATION,
            cls.NAME_EVALUATION_NO_NEUTRAL,
            cls.BSR_GROUP,
            cls.NSR_GROUP,
        ]


# Word Template Types (lowercase for backward compatibility)
class WordTemplateType:
    """Word document template type identifiers."""
    
    NORMAL: Final[str] = "normal"
    NORMAL_NO_NEUTRAL: Final[str] = "normal-noneutral"
    KATAKANA: Final[str] = "katakana"
    KATAKANA_BIG_JAP: Final[str] = "katakana_bigjap"
    PHONETICS: Final[str] = "phonetics"
    NONPROPRIETARY: Final[str] = "nonproprietary"
    TAGLINE: Final[str] = "tagline"


# Template Filenames
class TemplateFilename:
    """Default template filenames."""
    
    # Word templates
    WORD_RATIONALES: Final[str] = "InputDocumentRationales.doc"
    WORD_KATAKANA: Final[str] = "InputDocumentRationales_KataKana.doc"
    WORD_PHONETICS: Final[str] = "InputDocumentRationales_Phonetics.doc"
    
    # PowerPoint templates
    PPTX_DEFAULT_2019: Final[str] = "template_default_2019.pptx"
    PPTX_WITH_GROUPS_2019: Final[str] = "template_default_withgroups2019.pptx"
    PPTX_WITH_GROUP_2019: Final[str] = "template_default_withgroup_2019.pptx"
    PPTX_SEPARATOR_2019: Final[str] = "template_default_seperator_2019.pptx"
    PPTX_SUMMARY_2019: Final[str] = "template_default_summary2019.pptx"


# File Extensions
class FileExtension:
    """Valid file extensions."""
    
    XLSX: Final[str] = ".xlsx"
    XLS: Final[str] = ".xls"
    PPTX: Final[str] = ".pptx"
    PPTM: Final[str] = ".pptm"
    DOC: Final[str] = ".doc"
    DOCX: Final[str] = ".docx"
    JPG: Final[str] = ".jpg"
    JPEG: Final[str] = ".jpeg"
    PNG: Final[str] = ".png"


# HTTP Status Codes (for documentation)
class HttpStatus:
    """Common HTTP status codes."""
    
    OK: Final[int] = 200
    CREATED: Final[int] = 201
    BAD_REQUEST: Final[int] = 400
    UNAUTHORIZED: Final[int] = 401
    FORBIDDEN: Final[int] = 403
    NOT_FOUND: Final[int] = 404
    REQUEST_TIMEOUT: Final[int] = 408
    PAYLOAD_TOO_LARGE: Final[int] = 413
    UNPROCESSABLE_ENTITY: Final[int] = 422
    INTERNAL_SERVER_ERROR: Final[int] = 500
    SERVICE_UNAVAILABLE: Final[int] = 503


# Database Field Limits
class DbLimits:
    """Database field size limits."""
    
    MAX_PROJECT_NAME_LENGTH: Final[int] = 255
    MAX_DISPLAY_NAME_LENGTH: Final[int] = 255
    MAX_USER_NAME_LENGTH: Final[int] = 100
    MAX_FILE_PATH_LENGTH: Final[int] = 500


# Default Values
class Defaults:
    """Default configuration values."""
    
    CHART_WIDTH: Final[float] = 3.5
    CHART_HEIGHT: Final[float] = 3.5
    CHART_DPI: Final[int] = 100
    FONT_NAME: Final[str] = "Calibri"
    FONT_SIZE: Final[int] = 11
    MAX_EXCEL_ROWS: Final[int] = 2000
