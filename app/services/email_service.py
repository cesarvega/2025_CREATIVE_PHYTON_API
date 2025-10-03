"""Email notification service placeholder."""

from __future__ import annotations

from typing import Optional

from app.config.settings import settings
from app.constants import PresentationType, ProjectType
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


# Email mappings by presentation type (use settings for actual addresses)
def _get_email_mappings() -> dict[str, str]:
    """Get email mappings from settings."""
    return {
        PresentationType.NORMAL: settings.email_creative,
        PresentationType.NORMAL_NO_NEUTRAL: settings.email_creative,
        PresentationType.PHONETICS: settings.email_creative,
        PresentationType.KATAKANA: settings.email_creative,
        PresentationType.KATAKANA_BIG_JAP: settings.email_creative,
        PresentationType.NONPROPRIETARY: settings.email_nonproprietary,
        PresentationType.TAGLINE: settings.email_creative,
        PresentationType.DESIGN: settings.email_creative,
    }


def send_presentation_emails(
    presentation_id: int,
    presentation_type: str,
    project_type: str,
    display_name: str,
    user_name: str,
    participant_vote: bool = False,
    is_wide_ppt: bool = False,
) -> bool:
    """Send notification emails for a presentation.
    
    Args:
        presentation_id: The ID of the presentation.
        presentation_type: Type of presentation.
        project_type: Type of project.
        display_name: Display name of the presentation.
        user_name: Name of the user who created it.
        participant_vote: Whether participant voting is enabled.
        is_wide_ppt: Whether it's a widescreen presentation.
        
    Returns:
        True if email sent successfully, False otherwise.
    """
    try:
        recipient = _resolve_recipient(presentation_type, project_type)
        if not recipient:
            logger.info("No recipient configured for presentation type %s", presentation_type)
            return True

        subject = f"[{project_type.upper()}] {display_name} ready"
        body = _build_body(
            presentation_id=presentation_id,
            display_name=display_name,
            user_name=user_name,
            presentation_type=presentation_type,
            project_type=project_type,
            participant_vote=participant_vote,
            is_wide_ppt=is_wide_ppt,
        )

        _send_email(recipient, subject, body)
        return True
    except Exception as exc:  # pylint: disable=broad-except
        logger.error("Error sending notification email: %s", exc)
        return False


def _resolve_recipient(presentation_type: str, project_type: str) -> Optional[str]:
    """Resolve the email recipient based on presentation and project types."""
    if presentation_type == PresentationType.NONPROPRIETARY or project_type.lower() == ProjectType.NSR:
        return settings.email_nonproprietary
    return _get_email_mappings().get(presentation_type, settings.email_creative)


def _build_body(
    *,
    presentation_id: int,
    display_name: str,
    user_name: str,
    presentation_type: str,
    project_type: str,
    participant_vote: bool,
    is_wide_ppt: bool,
) -> str:
    """Build the email body."""
    body = f"""
Hello team,

The presentation "{display_name}" has been generated.

ID: {presentation_id}
Project: {project_type.upper()}
Type: {presentation_type}
Created by: {user_name}
Wide Screen PPT: {'Yes' if is_wide_ppt else 'No'}
Participant Vote: {'Enabled' if participant_vote else 'Disabled'}

Please review it in BI Guidelines.

Regards,
Creative Macros System
"""
    return body.strip()


def _send_email(recipient: str, subject: str, body: str) -> None:
    """Send an email (currently just logs it)."""
    logger.info("Email would be sent to %s with subject '%s'", recipient, subject)
    logger.debug("Email body:\n%s", body)


# Backward compatibility: keep the class but make it a thin wrapper
class EmailService:
    """Deprecated: Use send_presentation_emails() directly."""

    @property
    def DEFAULT_CREATIVE(self) -> str:
        """Deprecated: Use settings.email_creative directly."""
        return settings.email_creative

    @property
    def DEFAULT_NONPROPRIETARY(self) -> str:
        """Deprecated: Use settings.email_nonproprietary directly."""
        return settings.email_nonproprietary

    @property
    def email_mappings(self) -> dict[str, str]:
        """Deprecated: Use _get_email_mappings() directly."""
        return _get_email_mappings()

    def send_presentation_emails(self, *args, **kwargs) -> bool:
        """Deprecated: Use send_presentation_emails() directly."""
        return send_presentation_emails(*args, **kwargs)


# Singleton instance for backward compatibility
email_service = EmailService()
