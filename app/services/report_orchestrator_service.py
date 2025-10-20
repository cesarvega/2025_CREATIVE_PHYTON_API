"""Orchestrator service for NW report generation."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from app.models.nw_reports_models import (
    DownloadResultsRequest,
    DownloadResultsResponse,
    ReportType,
)
from app.services.analytics_report_generator import analytics_report_generator
from app.services.bi_guidelines_service import bi_guidelines_service
from app.services.excel_report_generator import excel_report_generator
from app.services.word_report_generator import word_report_generator
from app.utils.download_utils import build_api_download_url
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class ReportOrchestratorService:
    """Orchestrates the generation of different types of NW reports."""

    def generate_report(
        self, request: DownloadResultsRequest
    ) -> DownloadResultsResponse:
        """Generate a report based on the request type.

        Args:
            request: Download results request with report type and parameters

        Returns:
            DownloadResultsResponse with file path and download token

        Raises:
            ValueError: If report type is not supported or presentation not found
            FileNotFoundError: If required templates are missing
        """
        logger.info(
            "Starting report generation: type=%s, presentation_id=%d",
            request.report_type,
            request.presentation_id,
        )

        # Get presentation info
        presentation_info = bi_guidelines_service.get_project_info(
            request.presentation_id
        )

        if not presentation_info:
            raise ValueError(
                f"Presentation with ID {request.presentation_id} not found"
            )

        project_name = presentation_info.get("Project", "Unknown")
        display_name = presentation_info.get("DisplayName", "Unknown")

        logger.info(
            "Generating report for project='%s', display_name='%s'",
            project_name,
            display_name,
        )

        warnings = []
        file_paths = []
        primary_file_path: Optional[Path] = None

        try:
            # For NW download results, always generate both Excel and Word reports
            if request.report_type == ReportType.EXCEL or request.report_type == ReportType.WORD:
                # Generate Excel report
                excel_path = self._generate_excel_report(
                    request, project_name, display_name
                )
                file_paths.append(excel_path)
                logger.info("Excel report generated: %s", excel_path)

                # Generate Word report
                try:
                    word_path = self._generate_word_report(
                        request, display_name
                    )
                    file_paths.append(word_path)
                    logger.info("Word report generated: %s", word_path)
                except Exception as e:
                    logger.error("Error generating Word report: %s", str(e), exc_info=True)
                    warnings.append(f"Word report generation failed: {str(e)}")

                # Primary file is based on what was requested
                if request.report_type == ReportType.EXCEL:
                    primary_file_path = excel_path
                else:
                    primary_file_path = file_paths[-1] if file_paths else excel_path

            elif request.report_type == ReportType.ANALYTICS:
                primary_file_path = self._generate_analytics_report(
                    request, project_name, display_name
                )
                file_paths.append(primary_file_path)

            elif request.report_type == ReportType.BSR:
                # BSR report generation
                raise NotImplementedError(
                    "BSR report generation is not yet implemented."
                )

            else:
                raise ValueError(f"Unsupported report type: {request.report_type}")

        except FileNotFoundError as e:
            logger.error("Template file not found: %s", str(e))
            warnings.append(f"Template file missing: {str(e)}")
            raise

        except NotImplementedError as e:
            logger.warning("Report type not implemented: %s", str(e))
            raise

        except Exception as e:
            logger.error("Error generating report: %s", str(e), exc_info=True)
            raise

        # Build download token for primary file
        download_token = None
        if primary_file_path and primary_file_path.exists():
            download_token = build_api_download_url(primary_file_path)

        # Build message
        if len(file_paths) > 1:
            message = f"Generated {len(file_paths)} reports: Excel and Word"
        else:
            message = f"{request.report_type.value.upper()} report generated successfully"

        return DownloadResultsResponse(
            success=True,
            message=message,
            file_path=str(primary_file_path) if primary_file_path else None,
            file_name=primary_file_path.name if primary_file_path else None,
            download_token=download_token,
            report_type=request.report_type,
            presentation_id=request.presentation_id,
            warnings=warnings,
        )

    def _generate_excel_report(
        self,
        request: DownloadResultsRequest,
        project_name: str,
        display_name: str,
    ) -> Path:
        """Generate Excel report.

        Args:
            request: Download results request
            project_name: Project name
            display_name: Display name

        Returns:
            Path to generated Excel file
        """
        logger.info("Generating Excel report")

        file_path = excel_report_generator.generate_excel_report(
            presentation_id=request.presentation_id,
            project_name=project_name,
            display_name=display_name,
            include_votes=request.include_votes,
            include_participants=request.include_participants,
        )

        return file_path

    def _generate_word_report(
        self,
        request: DownloadResultsRequest,
        display_name: str,
    ) -> Path:
        """Generate Word report.

        Args:
            request: Download results request
            display_name: Display name

        Returns:
            Path to generated Word file
        """
        logger.info("Generating Word report")

        # Default to phonetics mode (True)
        # This matches the C# implementation where isPhonetics is typically True
        file_path = word_report_generator.generate_word_report(
            presentation_id=request.presentation_id,
            display_name=display_name,
            is_phonetics=True,
        )

        return file_path

    def _generate_analytics_report(
        self,
        request: DownloadResultsRequest,
        project_name: str,
        display_name: str,
    ) -> Path:
        """Generate Analytics Excel report.

        Args:
            request: Download results request
            project_name: Project name
            display_name: Display name

        Returns:
            Path to generated Analytics Excel file
        """
        logger.info("Generating Analytics report")

        file_path = analytics_report_generator.generate_analytics_report(
            presentation_id=request.presentation_id,
            project_name=project_name,
            display_name=display_name,
        )

        return file_path


# Global service instance
report_orchestrator_service = ReportOrchestratorService()
