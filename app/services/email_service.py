"""Email notification service placeholder."""

from __future__ import annotations

from typing import Optional

from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class EmailService:
    """Service for sending notification emails."""

    DEFAULT_CREATIVE = "creative@brandinstitute.com"
    DEFAULT_NONPROPRIETARY = "Chicago-Nonproprietary@brandinstitute.com"

    def __init__(self) -> None:
        self.email_mappings = {
            "Normal": self.DEFAULT_CREATIVE,
            "Normal-NoNeutral": self.DEFAULT_CREATIVE,
            "Phonetics": self.DEFAULT_CREATIVE,
            "Katakana": self.DEFAULT_CREATIVE,
            "Katakana_BigJap": self.DEFAULT_CREATIVE,
            "Nonproprietary": self.DEFAULT_NONPROPRIETARY,
            "Tagline": self.DEFAULT_CREATIVE,
            "Design": self.DEFAULT_CREATIVE,
        }

    def send_presentation_emails(
        self,
        presentation_id: int,
        presentation_type: str,
        project_type: str,
        display_name: str,
        user_name: str,
        participant_vote: bool = False,
        is_wide_ppt: bool = False,
    ) -> bool:
        try:
            recipient = self._resolve_recipient(presentation_type, project_type)
            if not recipient:
                logger.info("No recipient configured for presentation type %s", presentation_type)
                return True

            subject = f"[{project_type.upper()}] {display_name} ready"
            body = self._build_body(
                presentation_id=presentation_id,
                display_name=display_name,
                user_name=user_name,
                presentation_type=presentation_type,
                project_type=project_type,
                participant_vote=participant_vote,
                is_wide_ppt=is_wide_ppt,
            )

            self._send_email(recipient, subject, body)
            return True
        except Exception as exc:  # pylint: disable=broad-except
            logger.error("Error sending notification email: %s", exc)
            return False

    def _resolve_recipient(self, presentation_type: str, project_type: str) -> Optional[str]:
        if presentation_type == "Nonproprietary" or project_type.lower() in {"nsr"}:
            return self.DEFAULT_NONPROPRIETARY
        return self.email_mappings.get(presentation_type, self.DEFAULT_CREATIVE)

    @staticmethod
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

    def _send_email(self, recipient: str, subject: str, body: str) -> None:
        logger.info("Email would be sent to %s with subject '%s'", recipient, subject)
        logger.debug("Email body:\n%s", body)


email_service = EmailService()
