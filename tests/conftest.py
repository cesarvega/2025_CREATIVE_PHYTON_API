"""
Pytest configuration and shared fixtures for all tests.

This file provides common fixtures, test utilities, and configuration
that can be used across all test modules.
"""

import io
from pathlib import Path
from typing import Any, Dict

import pytest
from openpyxl import Workbook


# Test data directory
TEST_DATA_DIR = Path(__file__).parent / "test_data"


@pytest.fixture
def sample_excel_bytes() -> bytes:
    """Create a sample Excel file with test data.
    
    Returns:
        bytes: Excel file content as bytes.
    """
    workbook = Workbook()
    worksheet = workbook.active
    
    # Headers
    headers = [
        "Type", "Category", "Name", "Rationale",
        "Kana", "Logo", "NameSubGroup", "AltSubGroup"
    ]
    worksheet.append(headers)
    
    # Sample data rows
    data = [
        ["", "Category 1", "Alpha", "Reason A", "", "", "", ""],
        ["", "Category 1", "Beta", "Reason B", "", "logo.png", "", ""],
        ["", "Category 2", "Gamma", "Reason C", "カタカナ", "", "", ""],
        ["", "Category 2", "Delta", "Reason D", "", "", "GroupA", ""],
    ]
    
    for row in data:
        worksheet.append(row)
    
    # Save to bytes
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


@pytest.fixture
def sample_excel_with_groups() -> bytes:
    """Create an Excel file with group data.
    
    Returns:
        bytes: Excel file content with group information.
    """
    workbook = Workbook()
    worksheet = workbook.active
    
    headers = [
        "Type", "Category", "Name", "Rationale",
        "Kana", "Logo", "NameSubGroup", "AltSubGroup"
    ]
    worksheet.append(headers)
    
    data = [
        ["", "Category 1", "Name1", "Reason1", "", "", "GroupA", ""],
        ["", "Category 1", "Name2", "Reason2", "", "", "GroupA", ""],
        ["", "Category 1", "Name3", "Reason3", "", "", "GroupB", ""],
        ["", "Category 2", "Name4", "Reason4", "", "", "GroupB", ""],
        ["", "Category 2", "Name5", "Reason5", "", "", "", "GroupC"],
    ]
    
    for row in data:
        worksheet.append(row)
    
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


@pytest.fixture
def sample_presentation_metadata() -> Dict[str, Any]:
    """Sample presentation metadata for testing.
    
    Returns:
        dict: Valid presentation metadata.
    """
    return {
        "project": "TEST_PROJECT",
        "display_name": "Test Presentation",
        "presentation_type": "Normal",
        "user_name": "test_user",
        "project_type": "nw",
        "background_image": "default.jpg",
        "is_phonetics": False,
        "has_groups": False,
        "no_neutral": False,
        "participant_vote": 1,
    }


@pytest.fixture
def sample_build_metadata() -> Dict[str, Any]:
    """Sample build metadata for testing.
    
    Returns:
        dict: Valid build metadata.
    """
    return {
        "project": "TEST_BUILD",
        "display_name": "Test Build",
        "presentation_type": "Normal",
        "user_name": "test_user",
        "project_type": "nw",
        "background_image": "default.jpg",
        "template_pack": "BackgroundDefaultTemplate",
        "base_template": "template_default_2019.pptx",
        "slide_start": 1,
        "slide_end": 10,
    }


@pytest.fixture
def mock_settings(monkeypatch):
    """Mock settings for testing.
    
    This fixture allows tests to override settings without affecting
    the actual configuration.
    """
    from app.config.settings import settings
    
    # Store original values
    original_values = {}
    
    def _set_setting(key: str, value: Any):
        """Set a setting value for testing."""
        if key not in original_values:
            original_values[key] = getattr(settings, key)
        monkeypatch.setattr(settings, key, value)
    
    yield _set_setting
    
    # Cleanup is automatic with monkeypatch


@pytest.fixture
def temp_test_dir(tmp_path: Path) -> Path:
    """Create a temporary directory for test files.
    
    Args:
        tmp_path: pytest's built-in temporary path fixture.
        
    Returns:
        Path: Temporary directory for test files.
    """
    test_dir = tmp_path / "test_files"
    test_dir.mkdir(exist_ok=True)
    return test_dir


# Test markers for categorizing tests
def pytest_configure(config):
    """Register custom pytest markers."""
    config.addinivalue_line(
        "markers", "unit: Unit tests that don't require external dependencies"
    )
    config.addinivalue_line(
        "markers", "integration: Integration tests that may require database or files"
    )
    config.addinivalue_line(
        "markers", "slow: Tests that take a long time to run"
    )
    config.addinivalue_line(
        "markers", "api: API endpoint tests"
    )
