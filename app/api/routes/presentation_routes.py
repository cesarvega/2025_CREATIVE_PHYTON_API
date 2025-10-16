"""Routes for presentation creation and assembly orchestration."""

import base64
import time
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.api.dependencies import (
    parse_build_metadata,
    parse_presentation_metadata,
)
from app.config.settings import settings
from app.models.presentation_models import (
    CreatePresentationMetadata,
    CreatePresentationRequest,
    CreatePresentationResponse,
    PresentationBuildMetadata,
    PresentationBuildResponse,
)
from app.models.nw_reports_models import (
    DownloadResultsRequest,
    DownloadResultsResponse,
)
from app.services.presentation_service import presentation_service
from app.services.report_orchestrator_service import report_orchestrator_service
from app.utils.download_utils import (
    build_api_download_url,
    decode_download_token,
    guess_media_type,
)
from app.utils.logging_utils import get_logger

router = APIRouter(prefix="/presentations", tags=["Presentation Creation"])
logger = get_logger(__name__)


@router.post(
    "/create",
    response_model=CreatePresentationResponse,
    summary="Create a complete presentation from Excel + PPTX inputs with physical file generation",
    description=(
        "Upload the Excel candidate workbook, a base PPTX template, and a metadata JSON payload "
        "to generate a fully populated presentation. This endpoint:\n\n"
        "1. **Processes Excel data** - Extracts names, categories, groups, and rationales\n"
        "2. **Converts PPTX to images** - Generates JPG images from each slide of the original PPTX\n"
        "3. **Generates slide metadata** - Creates detailed slide information for database storage\n"
        "4. **Applies template rotation** (optional) - Rotates through specified templates for variety\n"
        "5. **Generates physical PowerPoint file** - Combines original PPTX slides with template-generated slides:\n"
        "   - Original PPTX slides (before `page_number`)\n"
        "   - Template-generated slides from Excel data (groups, individuals, multi-name slides)\n"
        "   - Original PPTX slides (after generated slides)\n"
        "6. **Persists to database** - Saves presentation metadata and slide details to nw_Master and nw_Details\n"
        "7. **Returns complete response** - Includes presentation ID, generated file paths, and processing stats\n\n"
        "The generated PowerPoint file is a complete, ready-to-present deck that seamlessly integrates "
        "your original slides with dynamically generated content from the Excel data."
    ),
    response_description=(
        "Creation status with presentation ID, total slide count, processing time, "
        "and paths to generated PowerPoint files (.pptx and optional .pptm)."
    ),
    responses={
        400: {
            "description": "Invalid or missing files provided in the multipart request.",
            "content": {
                "application/json": {
                    "example": {"detail": "Excel file must be .xlsx or .xls"}
                }
            },
        },
        422: {
            "description": "Metadata payload failed validation.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": [
                            {
                                "loc": ["metadata", "project"],
                                "msg": "Field required",
                                "type": "missing",
                            }
                        ]
                    }
                }
            },
        },
        500: {
            "description": "Unexpected error during orchestration.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to create presentation: PowerPoint automation unavailable"
                    }
                }
            },
        },
    },
)
async def create_presentation(
    metadata: CreatePresentationMetadata = Depends(parse_presentation_metadata),
    excel_file: UploadFile = File(..., description="Excel file (.xlsx or .xls)"),
    pptx_file: UploadFile = File(..., description="PowerPoint file (.pptx)"),
) -> CreatePresentationResponse:
    """Create a presentation by combining JSON metadata with uploaded template files.

    The client must submit a `metadata` field containing a JSON object with all
    presentation parameters (project name, display name, background settings, flags,
    etc.) together with two file uploads: the Excel candidate sheet (`excel_file`) and
    the PowerPoint template (`pptx_file`).

    **Optional Physical PowerPoint Generation:**
    If `generate_physical_pptx=true` in metadata, the service will also generate a
    complete physical PowerPoint file combining:
    - Original slides from the uploaded PPTX (before generated content)
    - Dynamically generated slides from Excel data using templates
    - Original slides from the uploaded PPTX (after generated content)
    
    This requires additional template configuration fields:
    - `template_pack`: Template directory name (e.g., "BackgroundTemplates")
    - `base_template`, `multi_template`, `group_template`, `separator_template`, `summary_template`
    - `include_macro_version`: Generate .pptm file (default: false)

    The service then executes the full pipeline:
    Excel processing → PPTX conversion → slide generation → template application →
    database persistence → [optional] physical PowerPoint generation.
    """
    try:
        # Validate Excel file
        if not excel_file.filename or not excel_file.filename.lower().endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="Excel file must be .xlsx or .xls")
        
        # Validate PPTX file
        if not pptx_file.filename or not pptx_file.filename.lower().endswith(".pptx"):
            raise HTTPException(status_code=400, detail="PPTX file must be .pptx")
        
        # Read file contents with size validation
        excel_content = await excel_file.read()
        pptx_content = await pptx_file.read()
        
        # Validate sizes
        max_excel_size = 10 * 1024 * 1024  # 10MB
        max_pptx_size = 50 * 1024 * 1024   # 50MB
        
        if len(excel_content) == 0:
            raise HTTPException(status_code=400, detail="Empty Excel file provided")
        if len(excel_content) > max_excel_size:
            raise HTTPException(status_code=413, detail="Excel file too large. Maximum size is 10MB")
            
        if len(pptx_content) == 0:
            raise HTTPException(status_code=400, detail="Empty PPTX file provided")
        if len(pptx_content) > max_pptx_size:
            raise HTTPException(status_code=413, detail="PPTX file too large. Maximum size is 50MB")

        # Create request object from metadata + files
        request: CreatePresentationRequest = metadata.to_service_request(
            excel_content=excel_content,
            excel_filename=excel_file.filename,
            pptx_content=pptx_content,
            pptx_filename=pptx_file.filename,
        )

        logger.info(
            "Starting presentation creation for project: %s, display_name: %s",
            metadata.project,
            metadata.display_name,
        )

        start_time = time.time()

        # Orchestrate complete presentation creation
        result = presentation_service.create_presentation(request)

        processing_time = time.time() - start_time

        logger.info(
            "Presentation created successfully. ID: %s, Slides: %d, Time: %.2fs",
            result.presentation_id,
            result.total_slides,
            processing_time,
        )

        return result

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error creating presentation: %s", str(e))
        raise HTTPException(
            status_code=500, detail=f"Failed to create presentation: {str(e)}"
        ) from e


@router.post(
    "/createTemplate",
    response_model=PresentationBuildResponse,
    summary="Assemble presentation files directly from an Excel candidate sheet",
    description=(
        "Generate PPTX artifacts using metadata and an Excel workbook without uploading a template file. "
        "The metadata defines template packs, slide ranges, and output behaviour. Optionally returns base64 "
        "content or download URLs for the generated files."
    ),
    response_description="Generated artifact metadata including slide range and file locations.",
    responses={
        400: {
            "description": "Invalid Excel file or missing templates on disk.",
            "content": {
                "application/json": {
                    "example": {"detail": "Excel file must be .xlsx or .xls"}
                }
            },
        },
        422: {
            "description": "Metadata payload failed validation (slide ranges, template pack, etc.).",
            "content": {
                "application/json": {
                    "example": {
                        "detail": [
                            {
                                "loc": ["metadata", "slide_end"],
                                "msg": "slide_end must be greater than or equal to slide_start",
                                "type": "value_error",
                            }
                        ]
                    }
                }
            },
        },
        500: {
            "description": "Unexpected error while composing the presentation files.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to assemble presentation from provided inputs."
                    }
                }
            },
        },
    },
)
async def build_presentation_files(
    metadata: PresentationBuildMetadata = Depends(parse_build_metadata),
    excel_file: UploadFile = File(..., description="Excel file (.xlsx or .xls)"),
) -> PresentationBuildResponse:
    try:
        # Validate Excel file
        if not excel_file.filename or not excel_file.filename.lower().endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="Excel file must be .xlsx or .xls")
        
        # Read file content with size validation
        excel_content = await excel_file.read()
        
        # Validate size
        max_excel_size = 10 * 1024 * 1024  # 10MB
        if len(excel_content) == 0:
            raise HTTPException(status_code=400, detail="Empty Excel file provided")
        if len(excel_content) > max_excel_size:
            raise HTTPException(status_code=413, detail="Excel file too large. Maximum size is 10MB")

        build_request = metadata.to_creation_request(
            excel_content=excel_content,
            excel_filename=excel_file.filename,
            pptx_filename=metadata.base_template,
        )
        build_options = metadata.to_build_options()

        start_time = time.time()
        artifacts = presentation_service.build_presentation_files(
            build_request=build_request,
            options=build_options,
        )
        processing_time = time.time() - start_time

        printable_path = (
            str(artifacts.printable_path)
            if artifacts.printable_path is not None
            else None
        )
        macro_path = (
            str(artifacts.macro_path) if artifacts.macro_path is not None else None
        )

        printable_base64 = None
        macro_base64 = None

        if metadata.return_bytes:
            if artifacts.printable_path and artifacts.printable_path.exists():
                printable_base64 = base64.b64encode(
                    artifacts.printable_path.read_bytes()
                ).decode("utf-8")
            if artifacts.macro_path and artifacts.macro_path.exists():
                macro_base64 = base64.b64encode(
                    artifacts.macro_path.read_bytes()
                ).decode("utf-8")

        logger.info(
            "Presentation assembly finished for %s/%s in %.2fs (slides=%d)",
            metadata.project,
            metadata.display_name,
            processing_time,
            artifacts.total_slides,
        )

        return PresentationBuildResponse(
            message=f"Presentation assembled in {processing_time:.2f}s",
            total_slides=artifacts.total_slides,
            slide_start=artifacts.slide_start,
            slide_end=artifacts.slide_end,
            printable_pptx=printable_path,
            macro_pptx=macro_path,
            printable_base64=printable_base64,
            macro_base64=macro_base64,
            download_urls=artifacts.download_urls,
            api_downloads={
                key: build_api_download_url(path)
                for key, path in (
                    ("printable", artifacts.printable_path),
                    ("macro", artifacts.macro_path),
                )
                if path is not None
            },
            summary=artifacts.summary,
            warnings=artifacts.warnings,
        )

    except HTTPException:
        raise
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:  # pylint: disable=broad-except
        logger.error("Error assembling presentation: %s", exc)
        raise HTTPException(
            status_code=500,
            detail="Failed to assemble presentation from provided inputs.",
        ) from exc


@router.get(
    "/exists",
    summary="Check if a presentation exists",
    response_description="Returns a boolean indicating if the presentation exists.",
    responses={
        200: {
            "description": "Check completed successfully.",
            "content": {
                "application/json": {
                    "example": {"exists": True}
                }
            },
        },
        500: {
            "description": "Database error during check.",
            "content": {
                "application/json": {
                    "example": {"detail": "Database error while checking for presentation."}
                }
            },
        },
    },
)
async def check_presentation_exists(
    project_name: str,
    display_name: str,
    exclude_id: Optional[int] = None,
) -> dict[str, bool]:
    """
    Check if a presentation with the given project name and display name already exists.
    An optional `exclude_id` can be provided to exclude a specific presentation ID
    from the check, which is useful for update operations.
    """
    try:
        exists = presentation_service.presentation_exists(
            project_name=project_name,
            display_name=display_name,
            exclude_id=exclude_id,
        )
        return {"exists": exists}
    except HTTPException:
        raise

@router.get(
    "/files/{token}",
    summary="Download a generated presentation artifact",
    responses={
        200: {
            "description": "Binary contents of the requested presentation artifact.",
            "content": {"application/octet-stream": {"schema": {"type": "string", "format": "binary"}}},
        },
        400: {"description": "Invalid download token."},
        403: {"description": "Token does not reference an authorized file."},
        404: {"description": "Requested file not found."},
    },
)
async def download_generated_file(token: str) -> FileResponse:
    """Download a generated presentation file using a secure token.
    
    Args:
        token: The secure download token encoding the file path.
        
    Returns:
        FileResponse with the requested file.
    """
    file_path = decode_download_token(token)
    return FileResponse(
        path=file_path,
        filename=file_path.name,
        media_type=guess_media_type(file_path),
    )


@router.post(
    "/test-slide-generation",
    summary="Test slide generation with multi-candidate layout",
    description=(
        "Test endpoint to verify slide generation logic in isolation. "
        "Allows testing the multi-candidate name layout without going through the full pipeline."
    ),
)
async def test_slide_generation(
    template_name: str = "template_default_withgroups2019.pptx",
    names: str = "John Doe##Jane Smith##Bob Johnson##Alice Williams##Charlie Brown##David Miller##Emma Davis##Frank Wilson##Grace Martinez##Henry Anderson##Ivy Thomas##Jack Taylor##Kelly Moore##Liam Jackson##Mia White##Noah Harris##Olivia Martin##Peter Thompson##Quinn Garcia##Rachel Robinson##Steve Clark##Tina Rodriguez##Uma Lewis##Victor Lee##Wendy Walker##Xavier Hall##Yara Allen##Zack Young",
) -> dict:
    """Test slide generation with specified template and names.

    Args:
        template_name: Name of the template file to use (must exist in templates directory)
        names: Delimited string of names to layout (use ## as delimiter)

    Returns:
        Result of the slide generation test including success status and file path
    """
    try:
        from app.services.pptx_builder_service import PPTXBuilderService, tokenize_delimited_block
        from app.utils.path_utils import resolve_project_output
        import pythoncom

        # Initialize COM for this thread
        pythoncom.CoInitialize()

        try:
            # Parse names
            names_list = tokenize_delimited_block(names)
            if not names_list:
                raise HTTPException(status_code=400, detail="No names provided")

            logger.info("Testing slide generation with %d names", len(names_list))

            # Initialize builder service
            templates_dir = Path(settings.templates_dir) if hasattr(settings, 'templates_dir') else Path("templates")
            subdir = "BackgroundDefaultTemplate"
            templates_dir = templates_dir / subdir
            template_path = templates_dir / template_name
            
            print(template_path)

            if not template_path.exists():
                raise HTTPException(
                    status_code=404,
                    detail=f"Template not found: {template_name}. Available templates are in {templates_dir}"
                )

            builder = PPTXBuilderService(
                templates_dir=templates_dir,
                base_template_path=template_path,
            )

            # Create output directory
            output_dir, _ = resolve_project_output("test_slide_generation", None)
            output_dir.mkdir(parents=True, exist_ok=True)

            # Generate a test presentation with just one slide
            from pptx import Presentation
            import win32com.client

            # Load template
            prs = Presentation(str(template_path))

            # Use first slide or add one
            if prs.slides:
                slide = prs.slides[0]
            else:
                slide = prs.slides.add_slide(prs.slide_layouts[5])

            # Save the presentation first
            test_output_path = output_dir / f"test_slide_{int(time.time())}.pptx"
            prs.save(str(test_output_path))

            # Now open with COM automation to test the layout logic
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1

            presentation = powerpoint.Presentations.Open(str(test_output_path.absolute()))
            com_slide = presentation.Slides(1)

            # Collect placeholder shapes (use Name Candidate for template_default_withgroups2019.pptx)
            placeholder_shapes = builder._collect_placeholder_shapes(com_slide, "Name Candidate")

            # Debug: List all text shapes in the slide
            logger.info("=== DIAGNOSTIC: Analyzing slide shapes ===")
            all_text_shapes = list(builder._iter_text_shapes(com_slide))
            logger.info("Total text shapes found: %d", len(all_text_shapes))

            for idx, shape in enumerate(all_text_shapes[:20]):  # Limit to first 20
                try:
                    text = shape.TextFrame.TextRange.Text
                    logger.info("  Shape %d: '%s'", idx + 1, text[:50] if text else "(empty)")
                except Exception as e:
                    logger.info("  Shape %d: (no text - %s)", idx + 1, str(e))

            if not placeholder_shapes:
                logger.warning("No 'Name Candidate' placeholder shapes found in template")
                warnings_list = ["No 'Name Candidate' placeholder shapes found in template. Check logs for all text shapes found."]
            else:
                logger.info("Found %d 'Name Candidate' placeholder shapes", len(placeholder_shapes))
                warnings_list = []

                # Log the placeholder shapes found
                for idx, pshape in enumerate(placeholder_shapes):
                    try:
                        ptext = pshape.TextFrame.TextRange.Text
                        logger.info("  Placeholder %d: '%s'", idx + 1, ptext[:30])
                    except:
                        pass

            # Test the simplified placeholder assignment
            builder._assign_names_to_placeholders(
                placeholder_shapes=placeholder_shapes or [],
                names=names_list,
                context="test slide",
            )
            success = True

            # Save and close properly
            try:
                presentation.Save()
            except Exception as e:
                logger.warning("Could not save presentation: %s", e)

            try:
                presentation.Close()
            except Exception as e:
                logger.warning("Could not close presentation: %s", e)

            try:
                powerpoint.Quit()
            except Exception as e:
                logger.warning("Could not quit PowerPoint: %s", e)

            logger.info("Test slide generation completed: success=%s", success)

            return {
                "success": success,
                "names_count": len(names_list),
                "names": names_list[:10],  # First 10 names for brevity
                "placeholder_count": len(placeholder_shapes) if placeholder_shapes else 0,
                "output_path": str(test_output_path),
                "warnings": warnings_list,
                "message": "Slide generation test completed. Check the output file to verify the layout."
            }

        finally:
            # Uninitialize COM
            pythoncom.CoUninitialize()

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error in test slide generation: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to test slide generation: {str(e)}"
        ) from e


@router.post(
    "/test-table-layout",
    summary="Test table layout with simple list of names",
    description="Simple endpoint to test dynamic table layout with just a list of names",
)
async def test_table_layout(names: list[str]) -> dict:
    """Test table layout with a simple list of names.

    Args:
        names: List of candidate names

    Returns:
        Result of the table generation including success status and file path
    """
    try:
        from app.services.pptx_builder_service import PPTXBuilderService
        from app.utils.path_utils import resolve_project_output
        import pythoncom

        # Initialize COM for this thread
        pythoncom.CoInitialize()

        try:
            if not names:
                raise HTTPException(status_code=400, detail="No names provided")

            logger.info("Testing table layout with %d names", len(names))

            # Initialize builder service
            templates_dir = Path(settings.templates_dir) if hasattr(settings, 'templates_dir') else Path("templates")
            subdir = "BackgroundDefaultTemplate"
            templates_dir = templates_dir / subdir
            template_name = "template_default_withgroups2019.pptx"
            template_path = templates_dir / template_name

            if not template_path.exists():
                raise HTTPException(
                    status_code=404,
                    detail=f"Template not found: {template_name}"
                )

            builder = PPTXBuilderService(
                templates_dir=templates_dir,
                base_template_path=template_path,
            )

            # Create output directory
            output_dir, _ = resolve_project_output("test_table_layout", None)
            output_dir.mkdir(parents=True, exist_ok=True)

            # Generate test presentation
            from pptx import Presentation
            import win32com.client

            # Load template
            prs = Presentation(str(template_path))

            # Use first slide or add one
            if prs.slides:
                slide = prs.slides[0]
            else:
                slide = prs.slides.add_slide(prs.slide_layouts[5])

            # Save the presentation first
            test_output_path = output_dir / f"table_test_{int(time.time())}.pptx"
            prs.save(str(test_output_path))

            # Open with COM automation
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1

            presentation = powerpoint.Presentations.Open(str(test_output_path.absolute()))
            com_slide = presentation.Slides(1)

            # Collect placeholder shapes
            placeholder_shapes = builder._collect_placeholder_shapes(com_slide, "Name Candidate")

            logger.info("Found %d placeholder shapes", len(placeholder_shapes))

            # Test the simplified placeholder assignment
            cleaned_names = [n for n in names if n]
            builder._assign_names_to_placeholders(
                placeholder_shapes=placeholder_shapes or [],
                names=cleaned_names,
                context="test table",
            )
            success = True

            # Save and close properly
            try:
                presentation.Save()
            except Exception as e:
                logger.warning("Could not save: %s", e)

            try:
                presentation.Close()
            except Exception as e:
                logger.warning("Could not close: %s", e)

            try:
                powerpoint.Quit()
            except Exception as e:
                logger.warning("Could not quit: %s", e)

            logger.info("Table layout test completed: success=%s", success)

            return {
                "success": success,
                "names_count": len(names),
                "output_path": str(test_output_path),
                "message": "Table layout test completed. Check the output file."
            }

        finally:
            pythoncom.CoUninitialize()

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error in test table layout: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to test table layout: {str(e)}"
        ) from e


@router.post(
    "/download-results",
    response_model=DownloadResultsResponse,
    summary="Download NW presentation results in various formats",
    description=(
        "Generate and download presentation results in different formats:\n\n"
        "- **Excel**: Complete NW Results workbook with retained names, newly created names, "
        "roots/concepts to explore/avoid, notes, and optional votes and participants sheets\n"
        "- **Word**: NW Report document with formatted tables and charts (requires template)\n"
        "- **Analytics**: NW Analytics workbook with project-specific and region-specific metrics\n"
        "- **BSR**: BSR download with specialized Excel and Word templates\n\n"
        "The endpoint:\n"
        "1. Retrieves data from stored procedures based on presentation ID\n"
        "2. Processes and transforms data (vote conversion, name grouping, Unicode handling)\n"
        "3. Generates the requested report type\n"
        "4. Returns a download token for secure file access\n\n"
        "**Data Processing:**\n"
        "- Grouped names (delimited by ## or $$) are split into separate rows in Excel\n"
        "- Numeric votes (-1, 0, 1) are converted to text (Negative, Neutral, Positive)\n"
        "- Unicode characters are normalized for compatibility\n\n"
        "**Template Requirements:**\n"
        "- Analytics reports require: C:/Templates/CreativeMacros/NW_Analytics/NWAnalytics.xlsx\n"
        "- Word reports require: Nomenclature Workshop Report_Template_new_*.doc\n"
        "- BSR reports require: BSR_EXCEL_TEMPLATE.xls and BSR_WORD_TEMPLATE.docx"
    ),
    response_description=(
        "Report generation status with file path, filename, and secure download token. "
        "Use the download token with the /presentations/files/{token} endpoint to retrieve the file."
    ),
    responses={
        200: {
            "description": "Report generated successfully.",
            "content": {
                "application/json": {
                    "example": {
                        "success": True,
                        "message": "EXCEL report generated successfully",
                        "file_path": "C:/output/TestProject/NW_Results_TestPresentation_20250116_143022.xlsx",
                        "file_name": "NW_Results_TestPresentation_20250116_143022.xlsx",
                        "download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                        "report_type": "excel",
                        "presentation_id": 12345,
                        "generated_at": "2025-01-16T14:30:22.123456",
                        "warnings": []
                    }
                }
            },
        },
        400: {
            "description": "Invalid request parameters.",
            "content": {
                "application/json": {
                    "example": {"detail": "Presentation with ID 12345 not found"}
                }
            },
        },
        404: {
            "description": "Required template file not found.",
            "content": {
                "application/json": {
                    "example": {"detail": "Analytics template not found at: C:/Templates/CreativeMacros/NW_Analytics/NWAnalytics.xlsx"}
                }
            },
        },
        500: {
            "description": "Unexpected error during report generation.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Failed to generate report: Database connection error"
                    }
                }
            },
        },
        501: {
            "description": "Report type not yet implemented.",
            "content": {
                "application/json": {
                    "example": {
                        "detail": "Word report generation is not yet implemented. This requires Word COM automation or python-docx integration."
                    }
                }
            },
        },
    },
)
async def download_results(request: DownloadResultsRequest) -> DownloadResultsResponse:
    """Generate and download NW presentation results.

    This endpoint orchestrates the generation of various NW report types:

    **Excel Report (report_type: "excel"):**
    - Generates a complete workbook with multiple sheets
    - Includes retained names, newly created names, roots/concepts, and notes
    - Optionally includes voting data and participant information
    - Data is retrieved from stored procedures:
      - nw_dlRetainedNames_withRecraft
      - nw_CombineNewNames
      - nw_dlRootsOrConceptsToExplore
      - nw_dlRootsOrConceptsToAvoid
      - nw_dlOpenNotes
      - nw_Votesbygroups (if include_votes=true)
      - nw_VotedParticipants (if include_participants=true)

    **Analytics Report (report_type: "analytics"):**
    - Generates from NWAnalytics.xlsx template
    - Fills Project-Specific and Region-Specific sheets
    - Data from stored procedures:
      - NW_ProjectAnalytics
      - NW_RegionSpecificAnalytics

    **Word Report (report_type: "word"):** [Not Yet Implemented]
    - Would generate formatted Word document from template
    - Would use stored procedures:
      - nw_wdValuesToReplace (for placeholders)
      - nw_wdGetResults or nw_wdGetResults_Phonetics
      - Would include pie charts and formatted tables

    **BSR Report (report_type: "bsr"):** [Not Yet Implemented]
    - Would generate BSR-specific Excel and Word documents
    - Would use stored procedures:
      - bsr_GetExcelReport
      - Plus direct queries to bsr_ProjectConcepts and BSR_PageComments tables

    Args:
        request: Download results request containing:
            - presentation_id: The presentation to generate report for
            - report_type: Type of report (excel, word, analytics, bsr)
            - summary_type: Optional, for Word phonetics reports
            - mobile_link: Optional, for BSR downloads
            - include_votes: Whether to include votes sheet
            - include_participants: Whether to include participants sheet

    Returns:
        DownloadResultsResponse with:
            - success: Boolean indicating if generation succeeded
            - message: Descriptive message
            - file_path: Full path to generated file
            - file_name: Name of generated file
            - download_token: Secure token for downloading the file
            - report_type: Type of report generated
            - presentation_id: ID of presentation
            - generated_at: Timestamp of generation
            - warnings: List of any warnings during generation

    Raises:
        HTTPException 400: If presentation ID is invalid or not found
        HTTPException 404: If required template files are missing
        HTTPException 500: If report generation fails
        HTTPException 501: If report type is not yet implemented
    """
    try:
        logger.info(
            "Received download results request: presentation_id=%d, report_type=%s",
            request.presentation_id,
            request.report_type,
        )

        # Generate report using orchestrator service
        response = report_orchestrator_service.generate_report(request)

        logger.info(
            "Report generated successfully: %s for presentation %d",
            request.report_type,
            request.presentation_id,
        )

        return response

    except ValueError as e:
        logger.error("Invalid request: %s", str(e))
        raise HTTPException(status_code=400, detail=str(e)) from e

    except FileNotFoundError as e:
        logger.error("Template file not found: %s", str(e))
        raise HTTPException(status_code=404, detail=str(e)) from e

    except NotImplementedError as e:
        logger.error("Feature not implemented: %s", str(e))
        raise HTTPException(status_code=501, detail=str(e)) from e

    except HTTPException:
        raise

    except Exception as e:
        logger.error("Error generating report: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to generate report: {str(e)}"
        ) from e
