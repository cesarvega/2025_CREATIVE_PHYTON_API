"""Tests for download utilities."""

from pathlib import Path

import pytest
from fastapi import HTTPException

from app.utils.download_utils import (
    build_api_download_url,
    create_download_token,
    decode_download_token,
    guess_media_type,
    is_subpath,
    validate_download_path,
)


def test_is_subpath():
    """Test subpath checking."""
    parent = Path("/home/user/projects")
    child = Path("/home/user/projects/myapp")
    non_child = Path("/home/other/path")
    
    assert is_subpath(child, parent) is True
    assert is_subpath(non_child, parent) is False


def test_guess_media_type():
    """Test media type guessing."""
    assert guess_media_type(Path("test.pptx")) == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    assert guess_media_type(Path("test.xlsx")) == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert guess_media_type(Path("test.pdf")) == "application/pdf"
    assert guess_media_type(Path("test.jpg")) == "image/jpeg"
    assert guess_media_type(Path("test.unknown")) == "application/octet-stream"


def test_create_and_decode_token(tmp_path):
    """Test token creation and decoding."""
    # Create a temporary file
    test_file = tmp_path / "test.txt"
    test_file.write_text("test content")
    
    # Create token
    token = create_download_token(test_file)
    assert isinstance(token, str)
    assert len(token) > 0
    
    # Token should be URL-safe (no padding =)
    assert "=" not in token


def test_decode_invalid_token():
    """Test decoding invalid tokens."""
    with pytest.raises(HTTPException) as exc_info:
        decode_download_token("invalid_token_!!!!")
    assert exc_info.value.status_code == 400


def test_build_api_download_url(tmp_path):
    """Test building download URLs."""
    test_file = tmp_path / "test.pptx"
    test_file.write_text("test")
    
    url = build_api_download_url(test_file)
    assert url.startswith("/api/presentations/files/")
    assert len(url) > len("/api/presentations/files/")


def test_build_api_download_url_custom_prefix(tmp_path):
    """Test building download URLs with custom prefix."""
    test_file = tmp_path / "test.pptx"
    test_file.write_text("test")
    
    url = build_api_download_url(test_file, endpoint_prefix="/custom/download")
    assert url.startswith("/custom/download/")


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
