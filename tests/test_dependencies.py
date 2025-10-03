"""Unit tests for FastAPI dependencies."""

import io
import json

import pytest
from fastapi import HTTPException, UploadFile
from openpyxl import Workbook

from app.api.dependencies import (
    parse_build_metadata,
    parse_presentation_metadata,
    validate_display_name,
    validate_excel_file,
    validate_excel_file_with_size,
    validate_file_size,
    validate_project_type,
    validate_pptx_file,
    validate_pptx_file_with_size,
)


@pytest.mark.unit
class TestValidateDependencies:
    """Test suite for validation dependencies."""

    @pytest.mark.asyncio
    async def test_validate_excel_file_valid(self):
        """Test validating a valid Excel file."""
        file = UploadFile(filename="test.xlsx", file=io.BytesIO(b"test"))
        result = await validate_excel_file(file)
        assert result.filename == "test.xlsx"

    @pytest.mark.asyncio
    async def test_validate_excel_file_invalid_extension(self):
        """Test validating an Excel file with invalid extension."""
        file = UploadFile(filename="test.txt", file=io.BytesIO(b"test"))
        with pytest.raises(HTTPException) as exc_info:
            await validate_excel_file(file)
        assert exc_info.value.status_code == 400

    @pytest.mark.asyncio
    async def test_validate_pptx_file_valid(self):
        """Test validating a valid PPTX file."""
        file = UploadFile(filename="test.pptx", file=io.BytesIO(b"test"))
        result = await validate_pptx_file(file)
        assert result.filename == "test.pptx"

    @pytest.mark.asyncio
    async def test_validate_pptx_file_invalid_extension(self):
        """Test validating a PPTX file with invalid extension."""
        file = UploadFile(filename="test.doc", file=io.BytesIO(b"test"))
        with pytest.raises(HTTPException) as exc_info:
            await validate_pptx_file(file)
        assert exc_info.value.status_code == 400

    @pytest.mark.asyncio
    async def test_validate_file_size_valid(self):
        """Test validating file size within limit."""
        content = b"test content"
        file = UploadFile(filename="test.xlsx", file=io.BytesIO(content))
        result = await validate_file_size(file, max_size_mb=1)
        assert result == content

    @pytest.mark.asyncio
    async def test_validate_file_size_too_large(self):
        """Test validating file size exceeding limit."""
        content = b"x" * (11 * 1024 * 1024)  # 11 MB
        file = UploadFile(filename="test.xlsx", file=io.BytesIO(content))
        with pytest.raises(HTTPException) as exc_info:
            await validate_file_size(file, max_size_mb=10)
        assert exc_info.value.status_code == 413

    @pytest.mark.asyncio
    async def test_validate_file_size_empty(self):
        """Test validating empty file."""
        file = UploadFile(filename="test.xlsx", file=io.BytesIO(b""))
        with pytest.raises(HTTPException) as exc_info:
            await validate_file_size(file, max_size_mb=10)
        assert exc_info.value.status_code == 400

    @pytest.mark.asyncio
    async def test_validate_excel_file_with_size(self):
        """Test composite validation for Excel file with size."""
        content = b"test"
        file = UploadFile(filename="test.xlsx", file=io.BytesIO(content))
        result_file, result_content = await validate_excel_file_with_size(file)
        assert result_file.filename == "test.xlsx"
        assert result_content == content

    @pytest.mark.asyncio
    async def test_validate_pptx_file_with_size(self):
        """Test composite validation for PPTX file with size."""
        content = b"test"
        file = UploadFile(filename="test.pptx", file=io.BytesIO(content))
        result_file, result_content = await validate_pptx_file_with_size(file)
        assert result_file.filename == "test.pptx"
        assert result_content == content


@pytest.mark.unit
class TestParsingDependencies:
    """Test suite for parsing dependencies."""

    @pytest.mark.asyncio
    async def test_parse_presentation_metadata_valid(self, sample_presentation_metadata):
        """Test parsing valid presentation metadata."""
        metadata_json = json.dumps(sample_presentation_metadata)
        
        class MockForm:
            def __init__(self, value):
                self.value = value
        
        result = await parse_presentation_metadata(metadata_json)
        assert result.project == sample_presentation_metadata["project"]
        assert result.display_name == sample_presentation_metadata["display_name"]

    @pytest.mark.asyncio
    async def test_parse_presentation_metadata_invalid_json(self):
        """Test parsing invalid JSON."""
        with pytest.raises(HTTPException) as exc_info:
            await parse_presentation_metadata("not valid json")
        assert exc_info.value.status_code == 400

    @pytest.mark.asyncio
    async def test_validate_display_name_valid(self):
        """Test validating a valid display name."""
        result = await validate_display_name("Test Project")
        assert result == "Test Project"

    @pytest.mark.asyncio
    async def test_validate_display_name_empty(self):
        """Test validating an empty display name."""
        with pytest.raises(HTTPException) as exc_info:
            await validate_display_name("")
        assert exc_info.value.status_code == 400

    @pytest.mark.asyncio
    async def test_validate_project_type_valid(self):
        """Test validating a valid project type."""
        result = await validate_project_type("nw")
        assert result == "nw"

    @pytest.mark.asyncio
    async def test_validate_project_type_invalid(self):
        """Test validating an invalid project type."""
        with pytest.raises(HTTPException) as exc_info:
            await validate_project_type("invalid")
        assert exc_info.value.status_code == 400
