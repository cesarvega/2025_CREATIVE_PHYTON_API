"""Unit tests for configuration settings."""

from pathlib import Path

import pytest

from app.config.settings import Settings


@pytest.mark.unit
class TestSettings:
    """Test suite for application settings."""

    def test_settings_initialization(self):
        """Test that settings can be initialized."""
        settings = Settings()
        assert settings is not None
        assert settings.app_name is not None
        assert settings.app_version is not None

    def test_cors_parsing_wildcard(self):
        """Test parsing CORS settings with wildcard."""
        settings = Settings(cors_origins="*")
        assert settings.cors_origins_list == ["*"]

    def test_cors_parsing_csv(self):
        """Test parsing CORS settings with CSV."""
        settings = Settings(cors_origins="http://localhost:3000,http://localhost:8080")
        cors_list = settings.cors_origins_list
        assert len(cors_list) == 2
        assert "http://localhost:3000" in cors_list
        assert "http://localhost:8080" in cors_list

    def test_cors_methods_parsing(self):
        """Test parsing CORS methods."""
        settings = Settings(cors_methods="GET,POST,PUT")
        methods = settings.cors_methods_list
        assert len(methods) == 3
        assert "GET" in methods
        assert "POST" in methods
        assert "PUT" in methods

    def test_cors_headers_parsing(self):
        """Test parsing CORS headers."""
        settings = Settings(cors_headers="Content-Type,Authorization")
        headers = settings.cors_headers_list
        assert len(headers) == 2
        assert "Content-Type" in headers
        assert "Authorization" in headers

    def test_is_production_flag(self):
        """Test production environment flag."""
        settings_dev = Settings(environment="development")
        assert settings_dev.is_production is False
        assert settings_dev.is_development is True

        settings_prod = Settings(environment="production")
        assert settings_prod.is_production is True
        assert settings_prod.is_development is False

    def test_base_dir_bipresents_development(self):
        """Test base directory for bipresents in development."""
        settings = Settings(environment="development")
        base_dir = settings.base_dir_bipresents
        assert isinstance(base_dir, Path)
        assert "NW_Files" in str(base_dir)

    def test_base_dir_nw_development(self):
        """Test base directory for NW in development."""
        settings = Settings(environment="development")
        base_dir = settings.base_dir_nw
        assert isinstance(base_dir, Path)
        assert "NW_Files" in str(base_dir)

    def test_get_base_dir_for_project_type_bipresents(self):
        """Test getting base directory for bipresents project type."""
        settings = Settings()
        base_dir = settings.get_base_dir_for_project_type("bipresents")
        assert isinstance(base_dir, Path)

    def test_get_base_dir_for_project_type_nw(self):
        """Test getting base directory for NW project type."""
        settings = Settings()
        base_dir = settings.get_base_dir_for_project_type("nw")
        assert isinstance(base_dir, Path)

    def test_get_base_dir_for_project_type_invalid(self):
        """Test getting base directory for invalid project type."""
        settings = Settings()
        with pytest.raises(ValueError):
            settings.get_base_dir_for_project_type("invalid_type")

    def test_get_environment_info(self):
        """Test getting environment information."""
        settings = Settings()
        info = settings.get_environment_info()
        
        assert "environment" in info
        assert "is_production" in info
        assert "is_development" in info
        assert "base_dir_bipresents" in info
        assert "base_dir_nw" in info
        assert "host" in info
        assert "port" in info

    def test_email_settings(self):
        """Test email configuration settings."""
        settings = Settings()
        assert settings.email_creative is not None
        assert "@brandinstitute.com" in settings.email_creative
        assert settings.email_nonproprietary is not None
        assert "@brandinstitute.com" in settings.email_nonproprietary

    def test_file_size_limit(self):
        """Test file size limit setting."""
        settings = Settings()
        assert settings.max_file_size_mb > 0
        assert isinstance(settings.max_file_size_mb, int)

    def test_project_type_constants(self):
        """Test project type constants."""
        settings = Settings()
        assert settings.PROJECT_TYPE_BIPRESENTS == "bipresents"
        assert settings.PROJECT_TYPE_NW == "nw"
