"""Unit tests for email service."""

import pytest

from app.constants import PresentationType, ProjectType
from app.services.email_service import (
    _get_email_mappings,
    _resolve_recipient,
    send_presentation_emails,
)


@pytest.mark.unit
class TestEmailService:
    """Test suite for email notification service."""

    def test_get_email_mappings(self):
        """Test that email mappings are correctly configured."""
        mappings = _get_email_mappings()
        
        assert PresentationType.NORMAL in mappings
        assert PresentationType.NONPROPRIETARY in mappings
        assert PresentationType.PHONETICS in mappings
        assert PresentationType.KATAKANA in mappings

    def test_resolve_recipient_normal(self):
        """Test resolving recipient for Normal presentation type."""
        recipient = _resolve_recipient(PresentationType.NORMAL, ProjectType.NW)
        assert recipient is not None
        assert "@brandinstitute.com" in recipient

    def test_resolve_recipient_nonproprietary(self):
        """Test resolving recipient for Nonproprietary presentation type."""
        recipient = _resolve_recipient(PresentationType.NONPROPRIETARY, ProjectType.NW)
        assert recipient is not None
        assert "nonproprietary" in recipient.lower() or "chicago" in recipient.lower()

    def test_resolve_recipient_nsr_project(self):
        """Test that NSR project type maps to nonproprietary email."""
        recipient = _resolve_recipient(PresentationType.NORMAL, ProjectType.NSR)
        assert recipient is not None
        assert "nonproprietary" in recipient.lower() or "chicago" in recipient.lower()

    def test_send_presentation_emails_success(self):
        """Test sending presentation emails."""
        result = send_presentation_emails(
            presentation_id=123,
            presentation_type=PresentationType.NORMAL,
            project_type=ProjectType.NW,
            display_name="Test Project",
            user_name="test_user",
            participant_vote=False,
            is_wide_ppt=False,
        )
        
        # Email service currently just logs, should return True
        assert result is True

    def test_send_presentation_emails_nonproprietary(self):
        """Test sending emails for nonproprietary presentation."""
        result = send_presentation_emails(
            presentation_id=124,
            presentation_type=PresentationType.NONPROPRIETARY,
            project_type=ProjectType.NW,
            display_name="Test Nonproprietary",
            user_name="test_user",
        )
        
        assert result is True
