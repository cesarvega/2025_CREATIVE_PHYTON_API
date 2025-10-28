"""Orchestrator service for BSR report generation."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, List

from app.services.bsr_reports_service import bsr_reports_service
from app.services.bsr_excel_report_generator import bsr_excel_report_generator
from app.services.bsr_word_report_generator import bsr_word_report_generator
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


@dataclass
class BSRReportResult:
    excel_path: Path
    word_path: Path
    warnings: List[str] = field(default_factory=list)


class BSRReportOrchestratorService:
    """Orchestrates BSR Excel and Word report generation."""

    def generate_bsr_reports(self, presentation_id: int, display_name: Optional[str] = None) -> BSRReportResult:
        # 1. Get presentation info from bsr_Master
        project, db_display = bsr_reports_service.get_presentation_info(presentation_id)
        project_name = project or display_name or db_display or f"BSR_{presentation_id}"
        warnings: List[str] = []

        logger.info("Generating BSR reports for presentation_id=%d, project_name=%s", presentation_id, project_name)

        # 2. Generate Excel report
        excel_path = bsr_excel_report_generator.generate_excel_report(presentation_id, project_name)

        # 3. Generate Word report
        try:
            word_path = bsr_word_report_generator.generate_word_report(presentation_id, project_name)
        except Exception as e:
            logger.error("Word report generation failed: %s", str(e), exc_info=True)
            warnings.append(f"Word report generation failed: {str(e)}")
            # Still return with excel_path; set word_path to excel's folder with placeholder
            word_path = excel_path.with_suffix(".doc")

        # 4. Return paths to both files
        return BSRReportResult(excel_path=excel_path, word_path=word_path, warnings=warnings)


# Global instance
bsr_report_orchestrator_service = BSRReportOrchestratorService()
