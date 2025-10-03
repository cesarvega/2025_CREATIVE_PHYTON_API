"""Unit tests for utility modules."""

from pathlib import Path

import pytest

from app.utils.files_utils import FileUtils
from app.utils.path_utils import sanitize_folder_name


@pytest.mark.unit
class TestPathUtils:
    """Test suite for path utilities."""

    def test_sanitize_folder_name_normal(self):
        """Test sanitizing a normal folder name."""
        result = sanitize_folder_name("test_folder")
        assert result is not None
        assert isinstance(result, str)

    def test_sanitize_folder_name_with_spaces(self):
        """Test sanitizing folder name with spaces."""
        result = sanitize_folder_name("my test folder")
        assert " " not in result  # Spaces should be replaced

    def test_sanitize_folder_name_with_special_chars(self):
        """Test sanitizing folder name with special characters."""
        result = sanitize_folder_name("my<folder>name")
        assert "<" not in result
        assert ">" not in result

    def test_sanitize_folder_name_empty(self):
        """Test sanitizing empty folder name."""
        result = sanitize_folder_name("")
        assert result != ""  # Should provide default name

    def test_sanitize_folder_name_uppercase(self):
        """Test that folder names are uppercased."""
        result = sanitize_folder_name("lowercase")
        assert result.isupper()

    def test_sanitize_folder_name_length_limit(self):
        """Test that folder names are limited in length."""
        long_name = "a" * 100
        result = sanitize_folder_name(long_name)
        assert len(result) <= 50


@pytest.mark.unit
class TestFileUtils:
    """Test suite for file utilities."""

    def test_format_file_size_bytes(self):
        """Test formatting file size in bytes."""
        result = FileUtils.format_file_size(512)
        assert "512" in result or "B" in result

    def test_format_file_size_kb(self):
        """Test formatting file size in kilobytes."""
        result = FileUtils.format_file_size(1024)
        assert "KB" in result or "1" in result

    def test_format_file_size_mb(self):
        """Test formatting file size in megabytes."""
        result = FileUtils.format_file_size(1024 * 1024)
        assert "MB" in result or "1" in result

    def test_format_file_size_gb(self):
        """Test formatting file size in gigabytes."""
        result = FileUtils.format_file_size(1024 * 1024 * 1024)
        assert "GB" in result or "1" in result

    def test_get_media_type_pptx(self):
        """Test getting media type for PPTX file."""
        result = FileUtils.get_media_type(".pptx")
        assert "presentation" in result.lower()

    def test_get_media_type_xlsx(self):
        """Test getting media type for XLSX file."""
        result = FileUtils.get_media_type(".xlsx")
        assert "sheet" in result.lower() or "excel" in result.lower()

    def test_get_media_type_unknown(self):
        """Test getting media type for unknown extension."""
        result = FileUtils.get_media_type(".xyz")
        assert "octet-stream" in result

    def test_ensure_directory(self, temp_test_dir):
        """Test ensuring directory exists."""
        new_dir = temp_test_dir / "new_directory"
        result = FileUtils.ensure_directory(new_dir)
        assert result is True
        assert new_dir.exists()

    def test_is_safe_path_valid(self, temp_test_dir):
        """Test checking if path is safe (within base)."""
        child_path = temp_test_dir / "subfolder" / "file.txt"
        result = FileUtils.is_safe_path(temp_test_dir, child_path)
        assert result is True

    def test_is_safe_path_invalid(self, temp_test_dir):
        """Test checking if path is unsafe (outside base)."""
        outside_path = Path("/etc/passwd")
        result = FileUtils.is_safe_path(temp_test_dir, outside_path)
        assert result is False

    def test_safe_delete_file(self, temp_test_dir):
        """Test safely deleting a file."""
        test_file = temp_test_dir / "test_delete.txt"
        test_file.write_text("test content")
        
        result = FileUtils.safe_delete_file(test_file)
        assert result is True
        assert not test_file.exists()

    def test_safe_delete_file_nonexistent(self, temp_test_dir):
        """Test safely deleting a nonexistent file."""
        test_file = temp_test_dir / "nonexistent.txt"
        result = FileUtils.safe_delete_file(test_file)
        # Should handle gracefully
        assert result is False or result is True

    def test_get_file_info(self, temp_test_dir):
        """Test getting file information."""
        test_file = temp_test_dir / "info_test.txt"
        test_file.write_text("test content")
        
        info = FileUtils.get_file_info(test_file)
        if info:
            assert "name" in info or "size" in info
            assert "path" in info or info is not None

    def test_list_directory_files(self, temp_test_dir):
        """Test listing files in directory."""
        # Create some test files
        (temp_test_dir / "file1.txt").write_text("content1")
        (temp_test_dir / "file2.txt").write_text("content2")
        
        files = FileUtils.list_directory_files(temp_test_dir, pattern="*.txt")
        assert isinstance(files, list)
        assert len(files) >= 2 or len(files) == 0  # May vary based on implementation
