"""
Configuration settings for the Report Generator API.
"""

import os
from pathlib import Path

from pydantic_settings import BaseSettings

from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class Settings(BaseSettings):
    """Application settings."""

    # App configuration
    app_name: str = "Report Generator API"
    app_version: str = "2.0.0"
    debug: bool = False

    # Server configuration
    host: str = "0.0.0.0"
    port: int = 50100
    
    # Base directories for different project types
    base_dir_bipresents: Path = Path("C:/inetpub/wwwroot/bipresents/bsr_slides")
    base_dir_nw: Path = Path("C:/inetpub/wwwroot/nw2/nw_slides")

    # File paths
    base_dir: Path = Path("C:/inetpub/wwwroot/CreativePythonAPI")
    nw_files_dir: Path = base_dir / "NW_Files"
    powerpoint_dir: Path = nw_files_dir / "PowerPoint_files"
    output_dir: Path = nw_files_dir / "output_images"

    # Application paths
    app_dir: Path = Path(__file__).parent.parent.parent.resolve()
    static_dir: Path = app_dir / "static"
    default_images_dir: Path = static_dir / "default_images"
    
    # SQL Server connection string for BI_GUIDELINES
    sql_connection_string: str = (
        "DRIVER={SQL Server};"
        "SERVER=192.168.0.85;"
        "DATABASE=BI_GUIDELINES;"
        "UID=sqlguide;"
        "PWD=sqlguidepwd;"
        "TrustServerCertificate=yes;"
        "Connection Timeout=30;"
        "Encrypt=no;"
    )

    # CORS settings
    cors_origins: list = ["*"]
    cors_methods: list = ["*"]
    cors_headers: list = ["*"]

    # Logging configuration
    log_level: str = "INFO"

    # Chart settings
    default_chart_width: float = 3.5
    default_chart_height: float = 3.5
    chart_dpi: int = 100

    # Document settings
    default_font_name: str = "Calibri"
    default_font_size: int = 11

    # PowerPoint conversion settings
    pptx_image_format: str = "PNG"
    pptx_image_quality: int = 90

    # File cleanup settings
    cleanup_older_than_hours: int = 24
    max_file_size_mb: int = 100
    
    # Projects types
    PROJECT_TYPE_BIPRESENTS: str = "bipresents"
    PROJECT_TYPE_NW: str = "nw"
    
    def get_base_dir_for_project_type(self, project_type: str) -> Path:
        """Get the base directory for a specific project type."""
        if project_type == self.PROJECT_TYPE_BIPRESENTS:
            return self.base_dir_bipresents
        elif project_type == self.PROJECT_TYPE_NW:
            return self.base_dir_nw
        else:
            raise ValueError(f"Unknown project type: {project_type}")

    def ensure_directories(self):
        """Create necessary directories if they don't exist"""
        directories = [
            self.base_dir_bipresents,
            self.base_dir_nw,
            self.default_images_dir,
            self.nw_files_dir,
            self.output_dir,
            self.powerpoint_dir,
        ]

        for directory in directories:
            try:
                directory.mkdir(parents=True, exist_ok=True)
                logger.info(f"Ensured directory exists: {directory}")
            except Exception as e:
                logger.error(
                    "Warning: Could not create directory %s: %s", directory, str(e)
                )
                raise


# Global settings instance
settings = Settings()

# Directory paths for easy access
BASE_DIR = settings.base_dir
BASE_DIR_BIPRESENTS = settings.base_dir_bipresents
BASE_DIR_NW = settings.base_dir_nw
NW_FILES_DIR = settings.nw_files_dir
POWERPOINT_DIR = settings.powerpoint_dir
OUTPUT_DIR = settings.output_dir
APP_DIR = settings.app_dir
DEFAULT_IMAGES_DIR = settings.default_images_dir