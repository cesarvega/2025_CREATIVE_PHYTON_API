"""Orchestrator service for NW report generation."""

from __future__ import annotations

import zipfile
from datetime import datetime
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
from app.utils.download_utils import create_download_token
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class ReportOrchestratorService:
    """Orchestrates the generation of different types of NW reports."""

    def generate_report(
        self, request: DownloadResultsRequest
    ) -> DownloadResultsResponse:
        """Generate NW results (Excel and Word) for a presentation.

        Always generates both Excel and Word reports. The response's primary
        download token points to the Excel file for backward compatibility.

        Args:
            request: Download results request with presentation ID.

        Returns:
            DownloadResultsResponse with primary file path and download token.

        Raises:
            ValueError: If presentation not found
            FileNotFoundError: If required templates are missing
        """
        logger.info(
            "Starting report generation for presentation_id=%d",
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
            # Always generate both Excel and Word reports
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
                word_path = None

            # Primary file is Excel for backward compatibility
            primary_file_path = excel_path

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

        # Create ZIP file with both reports
        zip_file_path = None
        download_token = None

        if len(file_paths) > 1:
            # Multiple files: create ZIP
            try:
                zip_file_path = self._create_zip_archive(
                    file_paths, display_name
                )
                logger.info("Created ZIP archive: %s", zip_file_path)

                # Token points to ZIP file
                download_token = create_download_token(zip_file_path)
                primary_file_path = zip_file_path
                message = f"Generated {len(file_paths)} reports (Excel and Word) in ZIP archive"

            except Exception as e:
                logger.error("Error creating ZIP: %s", str(e), exc_info=True)
                warnings.append(f"Failed to create ZIP archive: {str(e)}")
                # Fallback to Excel only
                if excel_path and excel_path.exists():
                    download_token = create_download_token(excel_path)
                    primary_file_path = excel_path
                message = "Report generated (ZIP creation failed)"
        else:
            # Single file: return directly
            if primary_file_path and primary_file_path.exists():
                download_token = create_download_token(primary_file_path)
            message = "Report generated successfully"

        # Build individual tokens for backward compatibility
        excel_token = None
        word_token = None
        if excel_path and excel_path.exists():
            excel_token = create_download_token(excel_path)
        if 'word_path' in locals() and word_path and word_path.exists():
            word_token = create_download_token(word_path)

        return DownloadResultsResponse(
            success=True,
            message=message,
            file_path=str(primary_file_path) if primary_file_path else None,
            file_name=primary_file_path.name if primary_file_path else None,
            download_token=download_token,
            excel_download_token=excel_token,
            word_download_token=word_token,
            excel_file_name=excel_path.name if excel_path else None,
            word_file_name=word_path.name if 'word_path' in locals() and word_path else None,
            report_type=ReportType.EXCEL if not zip_file_path else ReportType.EXCEL,
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

    def _create_zip_archive(
        self, file_paths: list[Path], display_name: str
    ) -> Path:
        """Create a ZIP archive containing multiple report files.

        Args:
            file_paths: List of file paths to include in the ZIP
            display_name: Display name for naming the ZIP file

        Returns:
            Path to the created ZIP file

        Raises:
            Exception: If ZIP creation fails
        """
        # Create ZIP filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"{display_name}_Reports_{timestamp}.zip"

        # Use the same directory as the first file
        if file_paths:
            zip_dir = file_paths[0].parent
        else:
            # Fallback to downloads directory
            from app.config.settings import settings
            zip_dir = settings.nw_files_dir / "downloads"
            zip_dir.mkdir(parents=True, exist_ok=True)

        zip_path = zip_dir / zip_filename

        logger.info("Creating ZIP archive: %s", zip_path)

        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path in file_paths:
                    if file_path and file_path.exists():
                        # Add file to ZIP with just its filename (no path)
                        zipf.write(file_path, arcname=file_path.name)
                        logger.debug("Added to ZIP: %s", file_path.name)
                    else:
                        logger.warning("File not found, skipping: %s", file_path)

            logger.info("ZIP archive created successfully: %s (size: %d bytes)",
                       zip_path, zip_path.stat().st_size)

            return zip_path

        except Exception as e:
            logger.error("Failed to create ZIP archive: %s", str(e), exc_info=True)
            # Clean up partial ZIP if it exists
            if zip_path.exists():
                try:
                    zip_path.unlink()
                except:
                    pass
            raise

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
