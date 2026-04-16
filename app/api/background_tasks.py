"""Background task handlers for long-running presentation operations."""

import time
from typing import Optional
import pythoncom
import asyncio

from app.models.presentation_models import CreatePresentationMetadata, BSRCreatePresentationMetadata
from app.services.presentation_service import presentation_service
from app.services.report_orchestrator_service import report_orchestrator_service
from app.services.feedback_template_generator import feedback_template_generator
from app.services.feedback_template_generator_openxml import feedback_template_generator_openxml
from app.services.bsr_report_orchestrator_service import bsr_report_orchestrator_service as bsr_report_orchestrator
from app.config.settings import settings
from app.utils.task_manager import task_manager, TaskStatus
from app.utils.logging_utils import get_logger
from app.utils.com_manager import com_manager

logger = get_logger(__name__)


def _create_presentation_background(
    task_id: str,
    metadata: CreatePresentationMetadata,
    excel_content: bytes,
    excel_filename: str,
    pptx_content: bytes,
    pptx_filename: str,
):
    """Background task for creating presentations."""
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    try:
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=10)

        # Create request object
        request = metadata.to_service_request(
            excel_content=excel_content,
            excel_filename=excel_filename,
            pptx_content=pptx_content,
            pptx_filename=pptx_filename,
        )

        logger.info(
            "[Task %s] Starting presentation creation for project: %s, display_name: %s",
            task_id,
            metadata.project,
            metadata.display_name,
        )

        # Create progress callback for service to use
        def progress_callback(progress: int):
            """Callback for service to report progress."""
            task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=progress)

        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=20)
        start_time = time.time()

        # Execute presentation creation with COM concurrency control
        with com_manager.acquire(f"Create Presentation - {metadata.project}/{metadata.display_name}"):
            result = presentation_service.create_presentation(request, progress_callback=progress_callback)

        processing_time = time.time() - start_time

        logger.info(
            "[Task %s] Presentation created successfully. ID: %s, Slides: %d, Time: %.2fs",
            task_id,
            result.presentation_id,
            result.total_slides,
            processing_time,
        )

        # Mark as completed with result
        task_manager.update_status(
            task_id,
            TaskStatus.COMPLETED,
            progress=100,
            result=result.model_dump()
        )

    except asyncio.CancelledError:
        # Gracefully handle cancellation during server reload/shutdown
        logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error="Task cancelled during server shutdown"
        )
        return
    except Exception as e:
        logger.error("[Task %s] Error creating presentation: %s", task_id, str(e), exc_info=True)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error=f"Failed to create presentation: {str(e)}"
        )
    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()


def _generate_backup_background(
    task_id: str,
    presentation_id: int,
):
    """Background task for generating backup presentations."""
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    try:
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=10)

        logger.info("[Task %s] Starting backup generation for presentation_id: %d", task_id, presentation_id)

        # Create progress callback
        def progress_callback(progress: int):
            task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=progress)

        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=20)

        # Generate backup with COM concurrency control
        with com_manager.acquire(f"Generate Backup - Presentation {presentation_id}"):
            result = presentation_service.generate_backup_presentation(presentation_id, progress_callback=progress_callback)

        logger.info("[Task %s] Backup generated successfully for presentation %d", task_id, presentation_id)

        # Mark as completed
        task_manager.update_status(
            task_id,
            TaskStatus.COMPLETED,
            progress=100,
            result=result
        )

    except asyncio.CancelledError:
        logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error="Task cancelled during server shutdown"
        )
        return
    except Exception as e:
        logger.error("[Task %s] Error generating backup: %s", task_id, str(e), exc_info=True)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error=f"Failed to generate backup: {str(e)}"
        )
    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()


def _create_bsr_presentation_background(
    task_id: str,
    metadata: BSRCreatePresentationMetadata,
    pptx_content: bytes,
    pptx_filename: str,
):
    """Background task for creating BSR presentations."""
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    try:
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=10)

        logger.info(
            "[Task %s] Starting BSR presentation creation: project=%s, display=%s",
            task_id,
            metadata.project_name,
            metadata.display_name,
        )

        # Create progress callback
        def progress_callback(progress: int):
            task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=progress)

        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=20)
        start_time = time.time()

        # Create BSR presentation with COM concurrency control
        with com_manager.acquire(f"Create BSR - {metadata.project_name}/{metadata.display_name}"):
            result = presentation_service.create_bsr_presentation(
                project_name=metadata.project_name,
                display_name=metadata.display_name,
                slide_number=metadata.slide_number,
                presentation_type=metadata.presentation_type,
                user_name=metadata.user_name,
                is_wide_ppt=metadata.is_wide_ppt,
                pptx_content=pptx_content,
                pptx_filename=pptx_filename,
                categories=metadata.categories,
                overwrite_existing=metadata.overwrite_existing,
                progress_callback=progress_callback,
            )

        processing_time = time.time() - start_time

        logger.info(
            "[Task %s] BSR presentation created. ID: %s, Slides: %d, Time: %.2fs",
            task_id,
            result.get("presentation_id"),
            result.get("total_slides", 0),
            processing_time,
        )

        # Mark as completed
        task_manager.update_status(
            task_id,
            TaskStatus.COMPLETED,
            progress=100,
            result={
                "presentation_id": result.get("presentation_id"),
                "total_slides": result.get("total_slides", 0),
                "categories_added": result.get("categories_added", 0),
                "overwritten": result.get("overwritten", False),
                "processing_time_seconds": processing_time,
            }
        )

    except asyncio.CancelledError:
        logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error="Task cancelled during server shutdown"
        )
        return
    except Exception as e:
        logger.error("[Task %s] Error creating BSR presentation: %s", task_id, str(e), exc_info=True)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error=f"Failed to create BSR presentation: {str(e)}"
        )
    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()


def _download_results_background(
    task_id: str,
    presentation_id: int,
):
    """Background task for generating NW reports (Excel + Word)."""
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    try:
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=10)

        logger.info("[Task %s] Starting report generation for presentation_id: %d", task_id, presentation_id)

        from app.models.nw_reports_models import DownloadResultsRequest

        request = DownloadResultsRequest(presentation_id=presentation_id)

        # Create progress callback
        def progress_callback(progress: int):
            task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=progress)

        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=20)

        # Generate reports with COM concurrency control
        with com_manager.acquire(f"Download Results - Presentation {presentation_id}"):
            result = report_orchestrator_service.generate_report(request, progress_callback=progress_callback)

        logger.info("[Task %s] Reports generated for presentation %d", task_id, presentation_id)

        # Create download tokens and ZIP
        from pathlib import Path as _Path
        from app.utils.download_utils import create_download_token
        import zipfile
        from datetime import datetime

        excel_token = None
        word_token = None
        zip_path = None
        zip_token = None

        try:
            # Get file paths from result
            excel_path = _Path(result.excel_file) if hasattr(result, 'excel_file') and result.excel_file else None
            word_path = _Path(result.word_file) if hasattr(result, 'word_file') and result.word_file else None

            # Create download tokens
            if excel_path and excel_path.exists():
                excel_token = create_download_token(excel_path)
            if word_path and word_path.exists():
                word_token = create_download_token(word_path)

            # Create ZIP if both files exist
            if excel_path and excel_path.exists() and word_path and word_path.exists():
                downloads_dir = excel_path.parent
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                zip_filename = f"NW_Reports_{presentation_id}_{timestamp}.zip"
                zip_path = downloads_dir / zip_filename

                with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    zf.write(excel_path, arcname=excel_path.name)
                    zf.write(word_path, arcname=word_path.name)

                zip_token = create_download_token(zip_path)
                logger.info("[Task %s] Created ZIP file: %s", task_id, zip_path)

        except Exception as zip_exc:
            logger.warning("[Task %s] Could not create tokens/ZIP: %s", task_id, str(zip_exc))

        # Mark as completed
        result_dict = result.model_dump() if hasattr(result, 'model_dump') else {}
        result_dict.update({
            "excel_download_token": excel_token,
            "word_download_token": word_token,
            "zip_file": str(zip_path) if zip_path else None,
            "zip_download_token": zip_token,
        })

        task_manager.update_status(
            task_id,
            TaskStatus.COMPLETED,
            progress=100,
            result=result_dict
        )

    except asyncio.CancelledError:
        logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error="Task cancelled during server shutdown"
        )
        return
    except Exception as e:
        logger.error("[Task %s] Error generating reports: %s", task_id, str(e), exc_info=True)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error=f"Failed to generate reports: {str(e)}"
        )
    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()


def _create_feedback_template_background(
    task_id: str,
    presentation_id: int,
    display_name: str,
):
    """Background task for generating feedback templates."""
    # Determine which generator to use based on feature flag
    use_openxml = settings.use_openxml_feedback

    # Only initialize COM if using legacy COM-based generator
    if not use_openxml:
        pythoncom.CoInitialize()

    try:
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=10)

        logger.info(
            "[Task %s] Starting feedback template generation (mode: %s) for presentation_id: %d",
            task_id,
            "OpenXML" if use_openxml else "COM",
            presentation_id
        )

        # Create progress callback
        def progress_callback(progress: int):
            task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=progress)

        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=20)

        # Generate feedback template
        from pathlib import Path
        from app.utils.download_utils import create_download_token

        if use_openxml:
            # OpenXML mode: No COM concurrency control needed (faster, no COM overhead)
            logger.info("[Task %s] Using OpenXML generator (5-10x faster)", task_id)
            file_path = feedback_template_generator_openxml.generate_feedback_template(
                presentation_id=presentation_id,
                display_name=display_name,
                progress_callback=progress_callback,
            )
        else:
            # COM mode: Use COM concurrency control (legacy)
            logger.info("[Task %s] Using COM generator (legacy mode)", task_id)
            with com_manager.acquire(f"Feedback Template - Presentation {presentation_id}"):
                file_path = feedback_template_generator.generate_feedback_template(
                    presentation_id=presentation_id,
                    display_name=display_name,
                    progress_callback=progress_callback,
                )

        # Create download token
        download_token = None
        if file_path and file_path.exists():
            download_token = create_download_token(file_path)

        logger.info("[Task %s] Feedback template generated: %s", task_id, file_path)

        # Mark as completed
        task_manager.update_status(
            task_id,
            TaskStatus.COMPLETED,
            progress=100,
            result={
                "file_path": str(file_path),
                "file_name": file_path.name,
                "download_token": download_token,
                "generation_mode": "OpenXML" if use_openxml else "COM",
            }
        )

    except asyncio.CancelledError:
        logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error="Task cancelled during server shutdown"
        )
        return
    except Exception as e:
        logger.error("[Task %s] Error generating feedback template: %s", task_id, str(e), exc_info=True)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error=f"Failed to generate feedback template: {str(e)}"
        )
    finally:
        # Only uninitialize COM if we initialized it
        if not use_openxml:
            pythoncom.CoUninitialize()


def _create_feedback_template_with_backup_background(
    task_id: str,
    presentation_id: int,
    display_name: str,
    with_sound_files: bool = False,
):
    """Background task for generating feedback template and zipping with backup if available."""
    # Determine which generator to use based on feature flag
    use_openxml = settings.use_openxml_feedback

    # Only initialize COM if using legacy COM-based generator
    if not use_openxml:
        pythoncom.CoInitialize()

    try:
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=10)

        logger.info(
            "[Task %s] Starting feedback+backup bundle (mode: %s) for presentation_id: %d",
            task_id,
            "OpenXML" if use_openxml else "COM",
            presentation_id,
        )

        # Create progress callback
        def progress_callback(progress: int):
            # Reserve last 10% for backup lookup/zip
            capped = min(int(progress), 90)
            task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=capped)

        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=20)

        from pathlib import Path
        from datetime import datetime
        import zipfile

        from app.config.db import get_connection_scope
        from app.utils.download_utils import create_download_token
        from app.utils.nw_data_utils import sanitize_filename
        from app.utils.path_utils import get_nw_downloads_dir, resolve_project_output

        # Generate feedback template
        if use_openxml:
            # OpenXML mode: No COM concurrency control needed (faster)
            logger.info("[Task %s] Using OpenXML generator for feedback template", task_id)
            feedback_path = feedback_template_generator_openxml.generate_feedback_template(
                presentation_id=presentation_id,
                display_name=display_name,
                progress_callback=progress_callback,
                with_sound_files=with_sound_files,
            )
        else:
            # COM mode: Use COM concurrency control (legacy)
            logger.info("[Task %s] Using COM generator for feedback template (legacy)", task_id)
            with com_manager.acquire(f"Feedback Template Bundle - Presentation {presentation_id}"):
                feedback_path = feedback_template_generator.generate_feedback_template(
                    presentation_id=presentation_id,
                    display_name=display_name,
                    progress_callback=progress_callback,
                )

        feedback_token = None
        if feedback_path and feedback_path.exists():
            feedback_token = create_download_token(feedback_path)

        # Locate backup PPTX if it exists
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=92)
        backup_path: Path | None = None
        main_ppt_filename = None

        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    """
                    SELECT DisplayName, MainPptFileName
                    FROM [BI_GUIDELINES].[dbo].[nw_Master]
                    WHERE PresentationId = ?
                    """,
                    (presentation_id,),
                )
                row = cursor.fetchone()
                if row:
                    display_name_db = row[0] or display_name
                    main_ppt_filename = row[1] or None
                else:
                    display_name_db = display_name

            output_base, _ = resolve_project_output(
                display_name_db,
                "NW",
                fallback_subdir="generated_presentations",
            )

            if main_ppt_filename:
                candidate = output_base / Path(main_ppt_filename).name
                if candidate.exists():
                    backup_path = candidate

            # Fallback: pick the most recent PPTX in the folder
            if backup_path is None and output_base.exists():
                candidates = sorted(
                    output_base.glob("*.pptx"),
                    key=lambda p: p.stat().st_mtime,
                    reverse=True,
                )
                if candidates:
                    backup_path = candidates[0]
        except Exception as e:
            logger.warning("[Task %s] Could not locate backup PPTX: %s", task_id, str(e))
            backup_path = None

        # If backup exists, zip both artifacts; else return feedback only
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=96)
        output_path: Path = feedback_path
        download_token = feedback_token
        bundle_created = False
        bundle_warning = None

        if backup_path and backup_path.exists():
            try:
                downloads_dir = get_nw_downloads_dir()
                downloads_dir.mkdir(parents=True, exist_ok=True)

                safe_display_name = sanitize_filename(display_name)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                zip_filename = f"Feedback_Backup_{safe_display_name}_{timestamp}.zip"
                zip_path = downloads_dir / zip_filename

                with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    zf.write(feedback_path, arcname=feedback_path.name)
                    zf.write(backup_path, arcname=backup_path.name)

                output_path = zip_path
                download_token = create_download_token(zip_path)
                bundle_created = True
                logger.info("[Task %s] Created feedback+backup ZIP: %s", task_id, zip_path)
            except Exception as zip_exc:
                bundle_warning = f"Failed to create ZIP with backup: {str(zip_exc)}"
                logger.warning("[Task %s] %s", task_id, bundle_warning)

        # Mark as completed
        task_manager.update_status(
            task_id,
            TaskStatus.COMPLETED,
            progress=100,
            result={
                "file_path": str(output_path),
                "file_name": output_path.name,
                "download_token": download_token,
                "feedback_file_path": str(feedback_path) if feedback_path else None,
                "feedback_file_name": feedback_path.name if feedback_path else None,
                "feedback_download_token": feedback_token,
                "backup_file_path": str(backup_path) if backup_path and backup_path.exists() else None,
                "backup_file_name": backup_path.name if backup_path and backup_path.exists() else None,
                "backup_included": bundle_created,
                "with_sound_files": with_sound_files,
                "warning": bundle_warning,
                "generation_mode": "OpenXML" if use_openxml else "COM",
            },
        )

    except asyncio.CancelledError:
        logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error="Task cancelled during server shutdown",
        )
        return
    except Exception as e:
        logger.error(
            "[Task %s] Error generating feedback+backup bundle: %s",
            task_id,
            str(e),
            exc_info=True,
        )
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error=f"Failed to generate feedback+backup bundle: {str(e)}",
        )
    finally:
        # Only uninitialize COM if we initialized it
        if not use_openxml:
            pythoncom.CoUninitialize()


def _generate_bsr_report_background(
    task_id: str,
    presentation_id: int,
    display_name: Optional[str],
):
    """Background task for generating BSR reports."""
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    try:
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=10)

        logger.info("[Task %s] Starting BSR report generation for presentation_id: %d", task_id, presentation_id)

        # Create progress callback
        def progress_callback(progress: int):
            task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=progress)

        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=20)
        start_time = time.time()

        # Generate BSR reports with COM concurrency control
        with com_manager.acquire(f"BSR Reports - Presentation {presentation_id}"):
            result = bsr_report_orchestrator.generate_bsr_reports(
                presentation_id=presentation_id,
                display_name=display_name,
                progress_callback=progress_callback,
            )

        processing_time = time.time() - start_time

        # Build download tokens
        from pathlib import Path as _Path
        from app.utils.download_utils import create_download_token
        import zipfile
        from datetime import datetime

        excel_token = None
        word_token = None
        zip_path = None
        zip_token = None

        try:
            excel_path = _Path(result.excel_path)
            word_path = _Path(result.word_path) if result.word_path else None

            if excel_path.exists():
                excel_token = create_download_token(excel_path)
            if word_path and word_path.exists():
                word_token = create_download_token(word_path)

            # Create ZIP
            if excel_path.exists() and word_path and word_path.exists():
                downloads_dir = excel_path.parent
                base_name = display_name or excel_path.stem or f"BSR_{presentation_id}"
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                zip_path = downloads_dir / f"{base_name}_BSR_Reports_{timestamp}.zip"

                with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    zf.write(excel_path, arcname=excel_path.name)
                    zf.write(word_path, arcname=word_path.name)

                zip_token = create_download_token(zip_path)

        except Exception as zip_exc:
            logger.warning("[Task %s] Could not create tokens/ZIP: %s", task_id, str(zip_exc))

        logger.info("[Task %s] BSR reports generated in %.2fs", task_id, processing_time)

        # Mark as completed
        task_manager.update_status(
            task_id,
            TaskStatus.COMPLETED,
            progress=100,
            result={
                "excel_file": str(result.excel_path),
                "word_file": str(result.word_path),
                "excel_download_token": excel_token,
                "word_download_token": word_token,
                "zip_file": str(zip_path) if zip_path else None,
                "zip_download_token": zip_token,
                "processing_time_seconds": processing_time,
                "warnings": result.warnings,
            }
        )

    except asyncio.CancelledError:
        logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error="Task cancelled during server shutdown"
        )
        return
    except Exception as e:
        logger.error("[Task %s] Error generating BSR reports: %s", task_id, str(e), exc_info=True)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error=f"Failed to generate BSR reports: {str(e)}"
        )
    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()


def _build_presentation_files_background(
    task_id: str,
    metadata_dict: dict,
    excel_content: bytes,
    excel_filename: str,
):
    """Background task for building presentation files from template."""
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    try:
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=10)

        from app.models.presentation_models import PresentationBuildMetadata

        # Reconstruct metadata object
        metadata = PresentationBuildMetadata(**metadata_dict)

        logger.info(
            "[Task %s] Starting presentation build for: %s/%s",
            task_id,
            metadata.project,
            metadata.display_name,
        )

        # Create progress callback
        def progress_callback(progress: int):
            task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=progress)

        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=20)
        start_time = time.time()

        # Build presentation with COM concurrency control
        build_request = metadata.to_creation_request(
            excel_content=excel_content,
            excel_filename=excel_filename,
            pptx_filename=metadata.base_template,
        )
        build_options = metadata.to_build_options()

        with com_manager.acquire(f"Build Presentation - {metadata.project}/{metadata.display_name}"):
            artifacts = presentation_service.build_presentation_files(
                build_request=build_request,
                options=build_options,
                progress_callback=progress_callback,
            )

        processing_time = time.time() - start_time

        from app.utils.download_utils import build_api_download_url

        logger.info(
            "[Task %s] Presentation built successfully in %.2fs, slides: %d",
            task_id,
            processing_time,
            artifacts.total_slides,
        )

        # Mark as completed
        task_manager.update_status(
            task_id,
            TaskStatus.COMPLETED,
            progress=100,
            result={
                "message": f"Presentation assembled in {processing_time:.2f}s",
                "total_slides": artifacts.total_slides,
                "slide_start": artifacts.slide_start,
                "slide_end": artifacts.slide_end,
                "printable_pptx": str(artifacts.printable_path) if artifacts.printable_path else None,
                "macro_pptx": str(artifacts.macro_path) if artifacts.macro_path else None,
                "download_urls": artifacts.download_urls,
                "api_downloads": {
                    key: build_api_download_url(path)
                    for key, path in (
                        ("printable", artifacts.printable_path),
                        ("macro", artifacts.macro_path),
                    )
                    if path is not None
                },
                "summary": artifacts.summary,
                "warnings": artifacts.warnings,
            }
        )

    except asyncio.CancelledError:
        logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error="Task cancelled during server shutdown"
        )
        return
    except Exception as e:
        logger.error("[Task %s] Error building presentation: %s", task_id, str(e), exc_info=True)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error=f"Failed to build presentation: {str(e)}"
        )
    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()
