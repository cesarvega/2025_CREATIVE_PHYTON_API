"""
Configuration settings for the Report Generator API.
"""

from pathlib import Path

from pydantic_settings import BaseSettings

from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class Settings(BaseSettings):
    """Application settings with environment-based configuration."""

    model_config = {"env_file": ".env", "env_file_encoding": "utf-8"}

    # Environment configuration
    environment: str = "development"  # "development" or "production"

    # App configuration
    app_name: str = "Report Generator API"
    app_version: str = "2.0.0"
    debug: bool = False

    # Server configuration
    host: str = "0.0.0.0"
    port: int = 50100

    def _get_path_for_env(self, prod_path: str, dev_path: str) -> Path:
        """OPTIMIZATION: Helper method to reduce duplication in path properties.

        Args:
            prod_path: Path to use in production environment
            dev_path: Path to use in development environment

        Returns:
            Path object for the current environment
        """
        if self.is_production:
            return Path(prod_path)
        return Path(dev_path)

    @property
    def base_dir_bipresents(self) -> Path:
        """Get base directory for bipresents based on environment"""
        return self._get_path_for_env(
            "C:/inetpub/wwwroot/bipresents/bsr_slides",
            "NW_Files/bipresents/bsr_slides"
        )

    @property
    def base_dir_nw(self) -> Path:
        """Get base directory for nw based on environment"""
        return self._get_path_for_env(
            "C:/inetpub/wwwroot/nw2/nw_slides",
            "NW_Files/nw2/nw_slides"
        )

    @property
    def base_dir(self) -> Path:
        """Get base directory based on environment"""
        return self._get_path_for_env(
            "C:/inetpub/wwwroot/CreativePythonAPI",
            "NW_Files/CreativePythonAPI"
        )

    @property
    def nw_files_dir(self) -> Path:
        """Get NW files directory based on environment"""
        return self._get_path_for_env(
            "C:/inetpub/wwwroot/CreativePythonAPI/NW_Files",
            "NW_Files"
        )

    @property
    def backgrounds_dir(self) -> Path:
        """Get backgrounds directory for custom background templates.
        
        Production: C:/inetpub/wwwroot/nw2/assets/images/BackGrounds
        Development: NW_Files/backgrounds
        """
        return self._get_path_for_env(
            "C:/inetpub/wwwroot/nw2/assets/images/BackGrounds",
            "NW_Files/backgrounds"
        )

    @property
    def sound_files_dir(self) -> Path:
        """Get sound files directory for MP3 pronunciation files.

        Production: C:/SoundFiles
        Development: NW_Files/SoundFiles
        """
        return self._get_path_for_env(
            "C:/SoundFiles",
            "NW_Files/SoundFiles"
        )

    # Application paths
    app_dir: Path = Path(__file__).parent.parent.parent.resolve()

    # SQL Server connection string for BI_GUIDELINES
    # This value is loaded from the SQL_CONNECTION_STRING environment variable in .env file
    sql_connection_string: str = (
        "DRIVER={SQL Server};SERVER=localhost;DATABASE=master;Trusted_Connection=yes;"
    )

    # SQL Server connection string for DAYMASTER database
    # This value is loaded from the DAYMASTER_CONNECTION_STRING environment variable in .env file
    daymaster_connection_string: str = (
        "DRIVER={SQL Server};SERVER=sql08;DATABASE=DAYMASTER;UID=sqlguide;PWD=sqlguidepwd;TrustServerCertificate=yes;Connection Timeout=30;"
    )

    # CORS settings
    cors_origins: str = "*"  # Comma-separated list or "*" for all
    cors_methods: str = "*"  # Comma-separated list or "*" for all
    cors_headers: str = "*"  # Comma-separated list or "*" for all

    @staticmethod
    def _parse_csv_or_wildcard(value: str) -> list[str]:
        """
        Parse a comma-separated string or wildcard into a list.
        
        Args:
            value: A string that's either "*" or comma-separated values
            
        Returns:
            A list containing either ["*"] or parsed trimmed values
        """
        if value == "*":
            return ["*"]
        return [item.strip() for item in value.split(",")]

    @property
    def cors_origins_list(self) -> list[str]:
        """Get CORS origins as a list"""
        return self._parse_csv_or_wildcard(self.cors_origins)

    @property
    def cors_methods_list(self) -> list[str]:
        """Get CORS methods as a list"""
        return self._parse_csv_or_wildcard(self.cors_methods)

    @property
    def cors_headers_list(self) -> list[str]:
        """Get CORS headers as a list"""
        return self._parse_csv_or_wildcard(self.cors_headers)

    # Logging configuration
    log_level: str = "INFO"

    # Chart settings
    default_chart_width: float = 3.5
    default_chart_height: float = 3.5
    chart_dpi: int = 100

    @property
    def is_production(self) -> bool:
        """Check if running in production environment"""
        return self.environment.lower() == "production"

    @property
    def is_development(self) -> bool:
        """Check if running in development environment"""
        return self.environment.lower() == "development"

    # Document settings
    default_font_name: str = "Calibri"
    default_font_size: int = 11

    # PowerPoint conversion settings
    pptx_image_format: str = "JPG"
    pptx_image_quality: int = 90

    # File cleanup settings
    cleanup_older_than_hours: int = 24
    max_file_size_mb: int = 100

    # Projects types
    PROJECT_TYPE_BIPRESENTS: str = "bipresents"
    PROJECT_TYPE_NW: str = "nw"
    PROJECT_TYPE_DW: str = "dw"
    PROJECT_TYPE_BSR: str = "BSR"

    # Email notification settings
    email_creative: str = "creative@brandinstitute.com"
    email_nonproprietary: str = "Chicago-Nonproprietary@brandinstitute.com"
    email_enabled: bool = True  # Toggle email notifications

    # Concurrency settings
    concurrency_max_workers: int = 1  # PowerPoint COM: only 1 simultaneous
    concurrency_queue_size: int = 100
    task_timeout_seconds: int = 600  # 10 minutes maximum
    progress_update_interval: int = 5  # seconds

    # Socket.IO settings
    socketio_cors_origins: str = "http://localhost:4200,https://tools.brandinstitute.com,null"
    socketio_async_mode: str = "asgi"  # asgi mode for FastAPI integration

    # Backup generation settings
    use_openxml_backup: bool = True  # Feature flag: True = OpenXML (faster), False = COM (legacy)

    # Feedback template generation settings
    use_openxml_feedback: bool = True  # Feature flag: True = OpenXML (5-10x faster), False = COM (legacy)

    # Word report generation settings
    use_openxml_word_reports: bool = True  # Feature flag: True = OpenXML (5-10x faster), False = COM (legacy)

    @property
    def socketio_cors_origins_list(self) -> list[str]:
        """Get Socket.IO CORS origins as a list"""
        return [origin.strip() for origin in self.socketio_cors_origins.split(",")]

    def get_base_dir_for_project_type(self, project_type: str) -> Path:
        """Get the base directory for a specific project type."""
        if project_type == self.PROJECT_TYPE_BIPRESENTS:
            return self.base_dir_bipresents
        elif project_type == self.PROJECT_TYPE_NW:
            return self.base_dir_nw
        elif project_type == self.PROJECT_TYPE_DW:
            # DW shares the same base directory as NW (nw2/nw_slides)
            return self.base_dir_nw
        elif project_type == self.PROJECT_TYPE_BSR or project_type.upper() == "BSR":
            # BSR uses bipresents directory structure
            return self.base_dir_bipresents
        else:
            raise ValueError(f"Unknown project type: {project_type}")

    def get_environment_info(self) -> dict:
        """Get current environment configuration info"""
        return {
            "environment": self.environment,
            "is_production": self.is_production,
            "is_development": self.is_development,
            "base_dir_bipresents": str(self.base_dir_bipresents),
            "base_dir_nw": str(self.base_dir_nw),
            "nw_files_dir": str(self.nw_files_dir),
            "host": self.host,
            "port": self.port,
        }

    def ensure_directories(self):
        """Create necessary directories if they don't exist"""
        directories = [
            self.base_dir_bipresents,
            self.base_dir_nw,
            self.nw_files_dir,
        ]

        for directory in directories:
            try:
                directory.mkdir(parents=True, exist_ok=True)
                logger.info("Ensured directory exists: %s", directory)
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
APP_DIR = settings.app_dir
