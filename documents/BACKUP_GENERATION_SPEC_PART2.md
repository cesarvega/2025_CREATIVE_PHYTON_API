# DOCUMENTACIÓN COMPLETA - SISTEMA DE GENERACIÓN DE BACKUP (PARTE 2)
## Continuación: Niveles 4 y 5

**Nota:** Este archivo es la continuación de `BACKUP_GENERATION_COMPLETE_SPECIFICATION.md`

---

# NIVEL 4: IMPLEMENTACIÓN TÉCNICA

## 4.1 Código Completo del Endpoint

### 4.1.1 Endpoint: `POST /api/presentations/generate-backup`

**Archivo:** `app/api/routes/presentation_routes.py:2292`

```python
from fastapi import APIRouter, BackgroundTasks, HTTPException
from app.models.presentation_models import GenerateBackupRequest
from app.models.response_models import TaskCreatedResponse
from app.utils.task_manager import task_manager, TaskStatus
from app.utils.concurrency_manager_v2 import concurrency_manager
from app.api.background_tasks import _generate_backup_background
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)
router = APIRouter()


def check_concurrency_limit():
    """Verifica si se ha alcanzado el límite de concurrencia."""
    stats = concurrency_manager.get_stats()
    if stats['queue_size'] >= concurrency_manager.queue_size:
        raise HTTPException(
            status_code=503,
            detail=f"Server is at capacity. Queue size: {stats['queue_size']}/{concurrency_manager.queue_size}"
        )


@router.post(
    "/generate-backup",
    response_model=TaskCreatedResponse,
    status_code=202,
    summary="Generate backup presentation from saved files",
    description=(
        "Generates a PowerPoint backup from previously saved Excel and PPTX files.\n\n"
        "This endpoint enables deferred backup generation. Users who didn't create "
        "a backup during initial presentation creation can generate it later.\n\n"
        "**Process:**\n"
        "1. Retrieves presentation metadata from database\n"
        "2. Loads saved Excel and PPTX files from disk\n"
        "3. Regenerates the complete PowerPoint presentation\n"
        "4. Saves as backup_YYYYMMDD_HHMMSS.pptx\n\n"
        "**Returns:** Task information for async processing\n"
        "**Query status:** GET /api/presentations/tasks/{task_id}"
    ),
    responses={
        202: {
            "description": "Backup generation task created successfully",
            "model": TaskCreatedResponse
        },
        404: {"description": "Presentation not found"},
        503: {"description": "Server at capacity"},
        500: {"description": "Server error"}
    }
)
async def generate_backup(
    background_tasks: BackgroundTasks,
    request: GenerateBackupRequest
) -> TaskCreatedResponse:
    """Generate a backup PowerPoint presentation from saved files."""
    try:
        # 1. Check concurrency limit before accepting the task
        check_concurrency_limit()

        # 2. Create background task
        task_id = task_manager.create_task(
            task_type="generate_backup",
            description=f"Generating backup for presentation_id: {request.presentation_id}",
            metadata={"presentation_id": request.presentation_id}
        )

        logger.info(
            "Created background task %s for backup generation: presentation_id=%d",
            task_id,
            request.presentation_id,
        )

        # 3. Submit task to concurrency manager
        submission_result = await concurrency_manager.submit_task(
            task_id=task_id,
            task_type="generate_backup",
            func=_generate_backup_background,
            args=(task_id, request.presentation_id),
            priority=TaskPriority.NORMAL,
        )

        # 4. Return task information with queue position
        task_info = task_manager.get_task(task_id)

        return TaskCreatedResponse(
            task_id=task_id,
            status=submission_result["status"],
            message=f"Backup generation task {'started' if submission_result.get('will_start_immediately') else 'queued'} for presentation {request.presentation_id}",
            task_type="generate_backup",
            created_at=task_info["created_at"],
            status_url=f"/api/presentations/tasks/{task_id}",
            position=submission_result.get("position"),
            estimated_wait_seconds=submission_result.get("estimated_wait_seconds")
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Error queueing backup generation: %s", str(e), exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Failed to queue backup generation: {str(e)}"
        ) from e
```

### 4.1.2 Endpoint: `GET /api/presentations/tasks/{task_id}`

```python
@router.get(
    "/tasks/{task_id}",
    summary="Get task status",
    description="Retrieve status and result of a background task"
)
async def get_task_status(task_id: str):
    """Get status of a background task."""
    task_info = task_manager.get_task(task_id)

    if not task_info:
        raise HTTPException(
            status_code=404,
            detail=f"Task {task_id} not found"
        )

    return task_info
```

---

## 4.2 Código Completo del Background Worker

### 4.2.1 Background Task Function

**Archivo:** `app/api/background_tasks.py:100`

```python
import asyncio
import pythoncom
from app.utils.logging_utils import get_logger
from app.utils.task_manager import task_manager, TaskStatus
from app.utils.com_manager import com_manager
from app.services.presentation_service import presentation_service

logger = get_logger(__name__)


def _generate_backup_background(
    task_id: str,
    presentation_id: int,
):
    """Background task for generating backup presentations.

    Args:
        task_id: Unique task identifier
        presentation_id: ID of the presentation to backup

    Notes:
        - Initializes COM for thread
        - Acquires COM lock to limit concurrency
        - Updates task progress throughout
        - Handles cancellation and errors
        - Cleans up COM on completion
    """
    # Initialize COM for this thread
    pythoncom.CoInitialize()

    try:
        # Update initial status
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=10)

        logger.info(
            "[Task %s] Starting backup generation for presentation_id: %d",
            task_id,
            presentation_id
        )

        # Create progress callback
        def progress_callback(progress: int):
            """Callback to update task progress."""
            task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=progress)

        # Update to 20% before COM acquisition
        task_manager.update_status(task_id, TaskStatus.PROCESSING, progress=20)

        # Generate backup with COM concurrency control
        # This blocks if 2 COM operations are already running
        with com_manager.acquire(f"Generate Backup - Presentation {presentation_id}"):
            result = presentation_service.generate_backup_presentation(
                presentation_id,
                progress_callback=progress_callback
            )

        logger.info(
            "[Task %s] Backup generated successfully for presentation %d",
            task_id,
            presentation_id
        )

        # Mark as completed with result
        task_manager.update_status(
            task_id,
            TaskStatus.COMPLETED,
            progress=100,
            result=result
        )

    except asyncio.CancelledError:
        # Handle shutdown gracefully
        logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error="Task cancelled during server shutdown"
        )
        return

    except Exception as e:
        # Handle all other errors
        logger.error(
            "[Task %s] Error generating backup: %s",
            task_id,
            str(e),
            exc_info=True
        )
        task_manager.update_status(
            task_id,
            TaskStatus.FAILED,
            error=f"Failed to generate backup: {str(e)}"
        )

    finally:
        # ALWAYS uninitialize COM for this thread
        pythoncom.CoUninitialize()
```

---

## 4.3 Código Completo del Servicio Principal

### 4.3.1 Service: `generate_backup_presentation()`

**Archivo:** `app/services/presentation_service.py:750`

```python
from pathlib import Path
from fastapi import HTTPException
from app.config.db import create_connection
from app.utils.logging_utils import get_logger
from app.utils.path_utils import resolve_project_output

logger = get_logger(__name__)


class PresentationService:
    """Service for orchestrating presentation operations."""

    def generate_backup_presentation(
        self,
        presentation_id: int,
        progress_callback=None
    ) -> Dict[str, Any]:
        """
        Generate a backup PowerPoint presentation from saved Excel and PPTX files.

        This method retrieves a previously created presentation from the database,
        reads the saved original Excel and PowerPoint files, and regenerates the
        complete PowerPoint backup file.

        Args:
            presentation_id: The ID of the presentation to generate backup for
            progress_callback: Optional callback function to report progress (int 0-100)

        Returns:
            Dict containing:
                - presentation_id: ID of the presentation
                - printable_path: Path to generated .pptx file
                - macro_path: Path to generated .pptm file (if applicable)
                - total_slides: Total number of slides in the presentation
                - warnings: List of any warnings during generation

        Raises:
            HTTPException: If presentation not found, files missing, or generation fails
        """
        try:
            logger.info("Generating backup for presentation ID: %d", presentation_id)

            # ================================================================
            # STEP 1: Retrieve presentation information from database (30-40%)
            # ================================================================
            if progress_callback:
                progress_callback(30)

            with create_connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(
                        """
                        SELECT
                            Project,
                            DisplayName,
                            MainPptFileName,
                            NameCandidateFileName,
                            NameCandidateBGType,
                            NameCandidateBGName,
                            NameCandidateStartingSlide,
                            PresentationType,
                            UploadedBy,
                            BSRDisplayName,
                            isParticipantsVote,
                            isWideScreenPPT,
                            isAWSLinkReq
                        FROM [BI_GUIDELINES].[dbo].[nw_Master]
                        WHERE PresentationId = ?
                        """,
                        (presentation_id,)
                    )

                    row = cursor.fetchone()

                    if not row:
                        raise HTTPException(
                            status_code=404,
                            detail=f"Presentation with ID {presentation_id} not found"
                        )

                    # Extract presentation metadata
                    project = row[0]
                    display_name = row[1]
                    main_ppt_filename = row[2]
                    excel_filename = row[3]
                    background_type = row[4] or "Default"
                    background_name = row[5] or "Default"
                    page_number = row[6] or 1
                    presentation_type = row[7] or "Normal"
                    user_name = row[8] or ""
                    mobile_link_bsr = row[9] or ""
                    participant_vote = row[10] or 0
                    is_wide_ppt = row[11] or 0
                    is_aws_email = row[12] or 0

                    logger.info(
                        "Retrieved presentation info: project=%s, display=%s, type=%s, excel_file=%s, ppt_file=%s",
                        project,
                        display_name,
                        presentation_type,
                        excel_filename,
                        main_ppt_filename
                    )

            # ================================================================
            # STEP 2: Resolve file paths
            # ================================================================
            # Extract just the folder name if display_name contains path separators
            # This prevents path duplication issues when the display_name was stored with a path
            from pathlib import Path as PathLib
            clean_display_name = PathLib(display_name).name if "/" in display_name or "\\" in display_name else display_name

            output_base, _ = resolve_project_output(
                clean_display_name,
                "NW",  # Default to NW project type
                fallback_subdir="generated_presentations",
            )

            # ================================================================
            # STEP 2a: Load saved Excel file
            # ================================================================
            # Use the filename from database, or fallback to original_data.xlsx for backwards compatibility
            excel_filename_to_use = excel_filename if excel_filename else "original_data.xlsx"
            excel_path = output_base / excel_filename_to_use

            logger.info(
                "Looking for Excel file: excel_filename_from_db=%s, path=%s",
                excel_filename,
                excel_path
            )

            if not excel_path.exists():
                # Try fallback for old presentations
                excel_path_fallback = output_base / "original_data.xlsx"
                if excel_path_fallback.exists():
                    excel_path = excel_path_fallback
                    logger.info("Using fallback Excel filename: original_data.xlsx")
                else:
                    # List all Excel files in the directory to help diagnose the issue
                    excel_files = list(output_base.glob("*.xlsx"))
                    logger.error(
                        "Excel file not found. Expected: %s, Available files: %s",
                        excel_path,
                        [f.name for f in excel_files]
                    )
                    raise HTTPException(
                        status_code=404,
                        detail=f"Original Excel file not found: {excel_path}. Available files: {[f.name for f in excel_files]}"
                    )

            # ================================================================
            # STEP 2b: Load saved PowerPoint template file
            # ================================================================
            pptx_path = output_base / main_ppt_filename if main_ppt_filename else None
            if not pptx_path or not pptx_path.exists():
                # Try alternate naming
                pptx_files = list(output_base.glob("*.pptx"))
                if pptx_files:
                    pptx_path = pptx_files[0]
                else:
                    raise HTTPException(
                        status_code=404,
                        detail=f"Original PowerPoint file not found in: {output_base}"
                    )

            logger.info("Found Excel file: %s", excel_path)
            logger.info("Found PPTX file: %s", pptx_path)

            # ================================================================
            # STEP 3: Read file contents (40-50%)
            # ================================================================
            if progress_callback:
                progress_callback(40)

            with open(excel_path, "rb") as f:
                excel_content = f.read()

            with open(pptx_path, "rb") as f:
                pptx_content = f.read()

            # ================================================================
            # STEP 4: Create request object from database metadata (50-55%)
            # ================================================================
            if progress_callback:
                progress_callback(50)

            request = CreatePresentationRequest(
                project=project,
                display_name=display_name,
                background_type=background_type,
                background_name=background_name,
                page_number=page_number,
                presentation_type=presentation_type,
                user_name=user_name,
                mobile_link_bsr=mobile_link_bsr,
                participant_vote=participant_vote,
                is_wide_ppt=is_wide_ppt,
                is_aws_email=is_aws_email,
                excel_file=excel_content,
                excel_filename=excel_path.name,
                pptx_file=pptx_content,
                pptx_filename=pptx_path.name,
                create_backup=1,  # Always generate backup for this endpoint
                has_groups=True,  # Default assumption
                test_name_order="Default",  # Use Default for backup generation
                project_type="NW",  # Default to NW
            )

            # ================================================================
            # STEP 5: Process Excel file (55-65%)
            # ================================================================
            if progress_callback:
                progress_callback(55)

            logger.info("Processing Excel file for backup generation")
            excel_data = self._process_excel_file(request)

            # ================================================================
            # STEP 6: Load existing PPTX image data (65-70%)
            # ================================================================
            if progress_callback:
                progress_callback(65)

            # DO NOT reconvert - that would delete the entire folder including the Excel file
            logger.info("Loading existing PPTX images for backup generation")
            pptx_data = self._load_existing_pptx_images(output_base, pptx_path.name)

            # ================================================================
            # STEP 7: Generate slides data (70-75%)
            # ================================================================
            if progress_callback:
                progress_callback(70)

            logger.info("Generating slides data for backup")
            slides_data = self._generate_slides_from_excel(
                excel_data=excel_data,
                pptx_data=pptx_data,
                request=request,
            )

            # ================================================================
            # STEP 8: Generate physical PowerPoint backup (75-90%)
            # ================================================================
            if progress_callback:
                progress_callback(75)

            logger.info("Generating physical PowerPoint backup")
            ppt_files = self._generate_physical_powerpoint(
                slides_data=slides_data,
                excel_data=excel_data,
                request=request,
                pptx_data=pptx_data,
            )

            logger.info(
                "Backup generation completed: %s (total slides: %s)",
                ppt_files.get("printable_path", "N/A"),
                ppt_files.get("total_slides", "0"),
            )

            # ================================================================
            # STEP 9: Update the database with the generated PowerPoint filename (90-95%)
            # ================================================================
            if progress_callback:
                progress_callback(90)

            printable_path = ppt_files.get("printable_path", "")
            if printable_path:
                from pathlib import Path as PathLib
                generated_filename = PathLib(printable_path).name

                logger.info(
                    "Updating database with generated PowerPoint filename: %s",
                    generated_filename
                )

                with create_connection() as conn:
                    with conn.cursor() as cursor:
                        cursor.execute(
                            """
                            UPDATE [BI_GUIDELINES].[dbo].[nw_Master]
                            SET MainPptFileName = ?
                            WHERE PresentationId = ?
                            """,
                            (generated_filename, presentation_id)
                        )
                        conn.commit()
                        logger.info("Database updated successfully with new PowerPoint filename")

            # ================================================================
            # STEP 10: Return result
            # ================================================================
            return {
                "presentation_id": presentation_id,
                "printable_path": ppt_files.get("printable_path", ""),
                "macro_path": ppt_files.get("macro_path", ""),
                "total_slides": ppt_files.get("total_slides", "0"),
                "warnings": ppt_files.get("warnings", "None"),
            }

        except HTTPException:
            raise
        except Exception as e:
            logger.error(
                "Error generating backup for presentation %d: %s",
                presentation_id,
                str(e),
                exc_info=True
            )
            raise HTTPException(
                status_code=500,
                detail=f"Failed to generate backup: {str(e)}"
            ) from e
```

---

## 4.4 Configuración y Settings

### 4.4.1 Settings File

**Archivo:** `app/config/settings.py`

```python
from pydantic_settings import BaseSettings
from typing import Optional


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    # Database Configuration
    sql_connection_string: str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=your-server;DATABASE=BI_GUIDELINES;"
        "UID=your-user;PWD=your-password"
    )
    daymaster_connection_string: str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=your-server;DATABASE=DayMaster;"
        "UID=your-user;PWD=your-password"
    )

    # Concurrency Configuration
    concurrency_max_workers: int = 5
    concurrency_queue_size: int = 100
    task_timeout_seconds: int = 1800  # 30 minutes

    # COM Concurrency
    com_max_concurrent: int = 2

    # File Paths
    nw_slides_root: str = "C:/inetpub/wwwroot/nw_slides"
    bsr_slides_root: str = "C:/inetpub/wwwroot/bsr_slides"
    templates_root: str = "C:/inetpub/wwwroot/templates"

    # API Configuration
    api_title: str = "CreativePython API"
    api_version: str = "1.0.0"

    class Config:
        env_file = ".env"
        case_sensitive = False


settings = Settings()
```

### 4.4.2 Environment Variables (.env)

```bash
# Database
SQL_CONNECTION_STRING="DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=BI_GUIDELINES;UID=sa;PWD=YourPassword"
DAYMASTER_CONNECTION_STRING="DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=DayMaster;UID=sa;PWD=YourPassword"

# Concurrency
CONCURRENCY_MAX_WORKERS=5
CONCURRENCY_QUEUE_SIZE=100
TASK_TIMEOUT_SECONDS=1800
COM_MAX_CONCURRENT=2

# File Paths
NW_SLIDES_ROOT="C:/inetpub/wwwroot/nw_slides"
BSR_SLIDES_ROOT="C:/inetpub/wwwroot/bsr_slides"
TEMPLATES_ROOT="C:/inetpub/wwwroot/templates"
```

---

## 4.5 Dependencias y Requirements

### 4.5.1 requirements.txt

```txt
# Web Framework
fastapi==0.104.1
uvicorn[standard]==0.24.0
python-multipart==0.0.6

# Database
pyodbc==5.0.1

# Excel Processing
openpyxl==3.1.2

# PowerPoint Processing (Python)
python-pptx==0.6.23

# PowerPoint Automation (Windows COM)
pywin32==306

# Image Processing
Pillow==10.1.0

# Validation
pydantic==2.5.0
pydantic-settings==2.1.0

# Utilities
python-dotenv==1.0.0

# Async
asyncio==3.4.3
```

### 4.5.2 Instalación de Dependencias

```bash
# Crear entorno virtual
python -m venv venv

# Activar entorno virtual (Windows)
venv\Scripts\activate

# Instalar dependencias
pip install -r requirements.txt

# IMPORTANTE: Instalar pywin32 post-install scripts
python venv/Scripts/pywin32_postinstall.py -install
```

---

## 4.6 Inicialización de la Aplicación

### 4.6.1 main.py

```python
import asyncio
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.config.settings import settings
from app.api.routes import presentation_routes
from app.utils.concurrency_manager_v2 import concurrency_manager
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan manager - startup and shutdown."""

    # STARTUP
    logger.info("Starting application...")

    # Start concurrency manager
    concurrency_manager.start()
    logger.info("Concurrency manager started")

    yield

    # SHUTDOWN
    logger.info("Shutting down application...")

    # Stop concurrency manager (wait for tasks to finish)
    concurrency_manager.stop(wait=True, timeout=10.0)
    logger.info("Concurrency manager stopped")


# Create FastAPI application
app = FastAPI(
    title=settings.api_title,
    version=settings.api_version,
    lifespan=lifespan
)

# CORS Configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, specify exact origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routers
app.include_router(
    presentation_routes.router,
    prefix="/api/presentations",
    tags=["presentations"]
)


@app.get("/")
async def root():
    """Health check endpoint."""
    return {
        "status": "healthy",
        "service": settings.api_title,
        "version": settings.api_version
    }


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,  # Development only
        log_level="info"
    )
```

---

# NIVEL 5: CASOS EDGE Y TESTING

## 5.1 Casos Edge y Manejo de Errores

### 5.1.1 Presentación No Existe

**Escenario:** Usuario solicita backup de presentation_id que no existe en BD

**Manejo:**
```python
# En generate_backup_presentation()
if not row:
    raise HTTPException(
        status_code=404,
        detail=f"Presentation with ID {presentation_id} not found"
    )
```

**Response:**
```json
{
    "detail": "Presentation with ID 99999 not found"
}
```

### 5.1.2 Archivo Excel No Encontrado

**Escenario:** Excel original fue borrado del file system

**Manejo:**
```python
if not excel_path.exists():
    # Try fallback
    excel_path_fallback = output_base / "original_data.xlsx"
    if excel_path_fallback.exists():
        excel_path = excel_path_fallback
    else:
        # List available files for debugging
        excel_files = list(output_base.glob("*.xlsx"))
        raise HTTPException(
            status_code=404,
            detail=f"Original Excel file not found: {excel_path}. Available files: {[f.name for f in excel_files]}"
        )
```

**Response:**
```json
{
    "detail": "Original Excel file not found: C:/path/candidates_data.xlsx. Available files: ['other_file.xlsx']"
}
```

### 5.1.3 Archivo PPTX No Encontrado

**Escenario:** PPTX original fue borrado

**Manejo:**
```python
if not pptx_path or not pptx_path.exists():
    # Try to find ANY pptx file
    pptx_files = list(output_base.glob("*.pptx"))
    if pptx_files:
        pptx_path = pptx_files[0]  # Use first one found
    else:
        raise HTTPException(
            status_code=404,
            detail=f"Original PowerPoint file not found in: {output_base}"
        )
```

### 5.1.4 Límite de Concurrencia Excedido

**Escenario:** Demasiadas tareas en cola

**Manejo:**
```python
def check_concurrency_limit():
    stats = concurrency_manager.get_stats()
    if stats['queue_size'] >= concurrency_manager.queue_size:
        raise HTTPException(
            status_code=503,
            detail=f"Server is at capacity. Queue size: {stats['queue_size']}/{concurrency_manager.queue_size}"
        )
```

**Response:**
```json
{
    "detail": "Server is at capacity. Queue size: 100/100"
}
```

**Cliente debe:**
- Esperar y reintentar (exponential backoff)
- Consultar `/api/presentations/stats` para ver capacidad

### 5.1.5 PowerPoint COM Falla

**Escenario:** PowerPoint no está instalado o COM falla

**Manejo:**
```python
try:
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
except Exception as e:
    raise Exception(f"Failed to initialize PowerPoint COM: {e}")
```

**Warnings en Response:**
```json
{
    "warnings": "Error during composition: Failed to initialize PowerPoint COM: ..."
}
```

### 5.1.6 Excel Corrupto

**Escenario:** Excel no puede ser leído por openpyxl

**Manejo:**
```python
try:
    wb = load_workbook(BytesIO(file_content), data_only=True)
except Exception as e:
    raise HTTPException(
        status_code=400,
        detail=f"Failed to read Excel file: {e}"
    )
```

### 5.1.7 Timeout de Tarea

**Escenario:** Tarea excede 30 minutos (task_timeout_seconds)

**Manejo:**
- WorkerPool automáticamente termina tarea
- Task status → FAILED
- Error: "Task timeout exceeded"

### 5.1.8 Path Traversal Attack

**Escenario:** Malicious display_name: `../../etc/passwd`

**Prevención:**
```python
def sanitize_folder_name(name: str) -> str:
    # Remove path traversal
    name = name.replace('..', '')

    # Remove invalid characters
    invalid_chars = r'[<>:"|?*\x00-\x1f]'
    sanitized = re.sub(invalid_chars, '_', name)

    return sanitized
```

### 5.1.9 Shutdown Durante Procesamiento

**Escenario:** Server shutdown mientras tarea se ejecuta

**Manejo:**
```python
except asyncio.CancelledError:
    logger.warning("[Task %s] Background task cancelled during shutdown", task_id)
    task_manager.update_status(
        task_id,
        TaskStatus.FAILED,
        error="Task cancelled during server shutdown"
    )
    return  # Exit gracefully
```

---

## 5.2 Tests Unitarios

### 5.2.1 Test: Procesamiento de Excel

```python
import pytest
from io import BytesIO
from openpyxl import Workbook
from app.services.excel_service import process_excel_file


def test_process_excel_basic():
    """Test basic Excel processing."""
    # Create test workbook
    wb = Workbook()
    ws = wb.active

    # Headers
    ws['A1'] = 'Type'
    ws['B1'] = 'Category'
    ws['C1'] = 'Name'
    ws['D1'] = 'Rationale'

    # Data
    ws['A2'] = ''
    ws['B2'] = 'Technology'
    ws['C2'] = 'TechNova'
    ws['D2'] = 'Modern tech brand'

    ws['A3'] = 'A'
    ws['B3'] = ''
    ws['C3'] = 'Innovation Names'
    ws['D3'] = 'Group A Header'

    # Save to bytes
    excel_bytes = BytesIO()
    wb.save(excel_bytes)
    excel_bytes.seek(0)

    # Process
    result = process_excel_file(
        file_content=excel_bytes.getvalue(),
        is_phonetics=False,
        has_groups=True,
        test_name_order="Default"
    )

    # Assertions
    assert result.total_rows_processed == 2
    assert result.candidate_count == 1
    assert result.group_count == 1
    assert result.has_groups == True
    assert result.lst_names[0] == 'TechNova'
    assert result.lst_names[1] == 'Innovation Names'
    assert result.lst_types[1] == 'A'


def test_process_excel_notation_extraction():
    """Test name/notation extraction from parentheses."""
    wb = Workbook()
    ws = wb.active

    ws['A1'] = 'Type'
    ws['C1'] = 'Name'

    ws['C2'] = 'TechNova (tek-NOH-vuh)'

    excel_bytes = BytesIO()
    wb.save(excel_bytes)
    excel_bytes.seek(0)

    result = process_excel_file(
        file_content=excel_bytes.getvalue()
    )

    assert result.lst_names[0] == 'TechNova'
    assert result.lst_notations[0] == 'tek-NOH-vuh'


def test_randomize_order():
    """Test randomization of test names."""
    wb = Workbook()
    ws = wb.active

    ws['A1'] = 'Type'
    ws['C1'] = 'Name'

    # Add 10 rows
    for i in range(2, 12):
        ws[f'C{i}'] = f'Name{i-1}'

    excel_bytes = BytesIO()
    wb.save(excel_bytes)
    excel_bytes.seek(0)

    result = process_excel_file(
        file_content=excel_bytes.getvalue(),
        test_name_order="Randomize"
    )

    # Order should be different
    original_order = [f'Name{i}' for i in range(1, 11)]
    assert result.lst_names != original_order  # Very likely
    assert len(result.lst_names) == 10
```

### 5.2.2 Test: Generación de Slides

```python
import pytest
from app.services.presentation_service import PresentationService
from app.models.excel_models import ProcessedExcelData
from app.models.response_models import PPTXConversionResponse
from app.models.presentation_models import CreatePresentationRequest


def test_generate_slides_from_excel():
    """Test slide generation logic."""
    service = PresentationService()

    # Create test data
    excel_data = ProcessedExcelData(
        lst_types=['', 'A', ''],
        lst_categories=['Tech', '', 'Health'],
        lst_names=['TechNova', 'Group A', 'HealthPlus'],
        lst_rationales=['Modern', 'Header', 'Leader'],
        lst_notations=['', '', ''],
        lst_kana=['', '', ''],
        lst_logos=['', '', ''],
        lst_name_sub_groups=['', '', ''],
        lst_group_letters=['', 'A', 'A']
    )
    excel_data.__post_init__()

    pptx_data = PPTXConversionResponse(
        message="Test",
        conversion_id="test_123",
        project_type="NW",
        images=["/static/slide1.jpg", "/static/slide2.jpg"],
        thumbnails=[],
        titles=["Slide 1", "Slide 2"],
        total_images=2,
        pptx_file="test.pptx"
    )

    request = CreatePresentationRequest(
        excel_file=b"",
        pptx_file=b"",
        excel_filename="test.xlsx",
        pptx_filename="test.pptx",
        project="TEST_PROJECT",
        display_name="Test",
        user_name="tester",
        page_number=2,  # Insert after first image
        project_type="NW"
    )

    # Generate slides
    result = service._generate_slides_from_excel(
        excel_data=excel_data,
        pptx_data=pptx_data,
        request=request
    )

    # Assertions
    details = result['details']

    # Should have:
    # 1. First image (prefix)
    # 2. TechNova (Name slide)
    # 3. Group A (Group slide)
    # 4. HealthPlus (Name slide)
    # 5. Summary slide (NW project)
    # 6. Second image (tail)

    assert len(details) == 6

    assert details[0].slide_type == "Image"
    assert details[0].slide_number == 1

    assert details[1].slide_type == "Name"
    assert details[1].name == "TechNova"
    assert details[1].slide_number == 2

    assert details[2].slide_type == "Group"
    assert details[2].group_name == "Group A"

    assert details[3].slide_type == "Name"
    assert details[3].name == "HealthPlus"

    assert details[4].slide_type == "Summary"

    assert details[5].slide_type == "Image"
```

### 5.2.3 Test: Path Sanitization

```python
import pytest
from app.utils.path_utils import sanitize_folder_name


def test_sanitize_folder_name_basic():
    """Test basic sanitization."""
    assert sanitize_folder_name("Normal Name") == "Normal_Name"
    assert sanitize_folder_name("Name-With-Dashes") == "Name-With-Dashes"


def test_sanitize_folder_name_path_traversal():
    """Test path traversal prevention."""
    assert ".." not in sanitize_folder_name("../../etc/passwd")
    assert sanitize_folder_name("../../etc/passwd") == "etc_passwd"


def test_sanitize_folder_name_invalid_chars():
    """Test invalid character replacement."""
    assert sanitize_folder_name("Name<>:") == "Name___"
    assert sanitize_folder_name("Name|*?") == "Name___"


def test_sanitize_folder_name_length_limit():
    """Test length limiting."""
    long_name = "A" * 300
    result = sanitize_folder_name(long_name)
    assert len(result) <= 200
```

---

## 5.3 Tests de Integración

### 5.3.1 Test: E2E Backup Generation

```python
import pytest
from fastapi.testclient import TestClient
from main import app
from app.config.db import create_connection

client = TestClient(app)


@pytest.fixture
def test_presentation_id():
    """Create test presentation in database."""
    with create_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO nw_Master
                (Project, DisplayName, PresentationType, UploadedBy)
                VALUES (?, ?, ?, ?)
                """,
                ("TEST_PROJECT", "Test Presentation", "Normal", "tester")
            )
            conn.commit()

            cursor.execute("SELECT @@IDENTITY")
            presentation_id = cursor.fetchone()[0]

    yield presentation_id

    # Cleanup
    with create_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(
                "DELETE FROM nw_Master WHERE PresentationId = ?",
                (presentation_id,)
            )
            conn.commit()


def test_generate_backup_endpoint(test_presentation_id):
    """Test complete backup generation flow."""
    # 1. Submit backup request
    response = client.post(
        "/api/presentations/generate-backup",
        json={"presentation_id": test_presentation_id}
    )

    assert response.status_code == 202
    data = response.json()
    assert "task_id" in data
    assert data["task_type"] == "generate_backup"

    task_id = data["task_id"]

    # 2. Poll for completion
    import time
    max_wait = 60  # seconds
    start = time.time()

    while time.time() - start < max_wait:
        status_response = client.get(f"/api/presentations/tasks/{task_id}")
        status_data = status_response.json()

        if status_data["status"] == "completed":
            # Success!
            assert "result" in status_data
            result = status_data["result"]
            assert result["presentation_id"] == test_presentation_id
            assert "printable_path" in result
            break

        elif status_data["status"] == "failed":
            pytest.fail(f"Task failed: {status_data.get('error')}")

        time.sleep(1)
    else:
        pytest.fail("Task did not complete within timeout")
```

---

## 5.4 Mejores Prácticas de Producción

### 5.4.1 Logging

**Configuración:**
```python
import logging
import sys

def setup_logging():
    """Configure application logging."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('logs/app.log'),
            logging.StreamHandler(sys.stdout)
        ]
    )

    # Reduce noise from libraries
    logging.getLogger("uvicorn.access").setLevel(logging.WARNING)
    logging.getLogger("fastapi").setLevel(logging.WARNING)
```

### 5.4.2 Monitoring

**Health Check:**
```python
@app.get("/health")
async def health_check():
    """Detailed health check."""
    try:
        # Test database connection
        with create_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute("SELECT 1")

        db_healthy = True
    except Exception as e:
        db_healthy = False

    stats = concurrency_manager.get_stats()

    return {
        "status": "healthy" if db_healthy else "degraded",
        "database": "connected" if db_healthy else "disconnected",
        "concurrency": {
            "active_workers": stats["active_workers"],
            "queue_size": stats["queue_size"],
            "capacity": f"{stats['queue_size']}/{concurrency_manager.queue_size}"
        },
        "com": {
            "active": com_manager.get_active_count(),
            "max": com_manager.get_max_concurrent()
        }
    }
```

### 5.4.3 Scaling Considerations

**Para producción con múltiples servidores:**

1. **Usar Redis para Task Manager:**
```python
import redis
from typing import Dict, Any

class RedisTaskManager:
    def __init__(self):
        self.redis = redis.Redis(
            host='localhost',
            port=6379,
            decode_responses=True
        )

    def create_task(self, task_type: str, description: str, metadata: Dict) -> str:
        task_id = str(uuid.uuid4())
        self.redis.hset(f"task:{task_id}", mapping={
            "task_id": task_id,
            "task_type": task_type,
            "status": "pending",
            "description": description,
            "metadata": json.dumps(metadata),
            "created_at": datetime.now().isoformat()
        })
        return task_id
```

2. **Usar Message Queue (RabbitMQ/Celery):**
```python
from celery import Celery

celery_app = Celery(
    'tasks',
    broker='redis://localhost:6379/0',
    backend='redis://localhost:6379/1'
)

@celery_app.task
def generate_backup_task(presentation_id: int):
    """Celery task for backup generation."""
    return presentation_service.generate_backup_presentation(presentation_id)
```

3. **Load Balancer con Sticky Sessions:**
- Para mantener consistencia de task tracking
- O usar shared state (Redis)

---

## 5.5 Troubleshooting Guide

### Problem: "PowerPoint COM not responding"

**Síntomas:**
- Tareas quedan stuck en processing
- Timeout errors

**Solución:**
```bash
# Kill hanging PowerPoint processes
taskkill /F /IM POWERPNT.EXE

# Check COM manager capacity
GET /api/presentations/health
```

### Problem: "Queue is full"

**Síntomas:**
- 503 errors
- "Server at capacity"

**Solución:**
```python
# Increase queue size in settings
CONCURRENCY_QUEUE_SIZE=200

# Or increase workers
CONCURRENCY_MAX_WORKERS=10
```

### Problem: "Database connection timeout"

**Síntomas:**
- 500 errors
- "Failed to connect to database"

**Solución:**
```python
# Increase connection pool
# In db.py:
_MAX_POOL_SIZE = 20

# Or add retry logic:
import tenacity

@tenacity.retry(
    stop=tenacity.stop_after_attempt(3),
    wait=tenacity.wait_exponential(multiplier=1, min=1, max=10)
)
def create_connection():
    # ...
```

---

## 5.6 Checklist de Implementación

### Para AI Agent que Implementa Esto:

**✅ Base de Datos:**
- [ ] Crear tablas `nw_Master`, `nw_Details`, `nw_Templates`
- [ ] Configurar connection string
- [ ] Testear conexión con pyodbc

**✅ File System:**
- [ ] Crear estructura de carpetas en C:/inetpub/wwwroot/
- [ ] Configurar permisos de escritura
- [ ] Testear `resolve_project_output()`

**✅ Dependencias:**
- [ ] Instalar Python 3.9+
- [ ] Instalar todas las dependencias de requirements.txt
- [ ] Instalar PowerPoint en Windows
- [ ] Ejecutar `pywin32_postinstall.py`

**✅ Código Core:**
- [ ] Implementar `excel_service.py`
- [ ] Implementar `presentation_service.py`
- [ ] Implementar `pptx_builder_service.py` (COM logic)
- [ ] Implementar endpoints

**✅ Concurrency:**
- [ ] Implementar `task_manager.py`
- [ ] Implementar `concurrency_manager_v2.py`
- [ ] Implementar `com_manager.py`
- [ ] Implementar `background_tasks.py`

**✅ Testing:**
- [ ] Tests unitarios de Excel processing
- [ ] Tests unitarios de slide generation
- [ ] Tests de integración E2E
- [ ] Load testing con múltiples requests concurrentes

**✅ Deployment:**
- [ ] Configurar logging
- [ ] Configurar health checks
- [ ] Configurar monitoring
- [ ] Documentar API (OpenAPI/Swagger)

---

## 5.7 Documentación de API (OpenAPI)

Al ejecutar la aplicación, FastAPI genera automáticamente documentación interactiva en:

- **Swagger UI:** http://localhost:8000/docs
- **ReDoc:** http://localhost:8000/redoc

Esta documentación incluye:
- Todos los endpoints
- Request/Response schemas
- Ejemplos de uso
- Try-it-out functionality

---

# CONCLUSIÓN

Esta especificación completa proporciona:

1. ✅ **Arquitectura completa** del sistema
2. ✅ **Todos los modelos de datos** (DB + Python)
3. ✅ **Lógica de negocio detallada** con algoritmos
4. ✅ **Código fuente completo** de todos los componentes
5. ✅ **Manejo de casos edge** y errores
6. ✅ **Tests** unitarios e integración
7. ✅ **Guía de deployment** y troubleshooting

**Un agente AI puede recrear este sistema completo usando esta documentación.**

---

**Archivo generado:** 2025-12-12
**Total de páginas estimadas:** ~150-200 páginas en formato impreso
**Líneas de código documentadas:** ~3,000+ líneas
