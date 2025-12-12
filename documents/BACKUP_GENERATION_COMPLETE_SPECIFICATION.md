# DOCUMENTACIÓN COMPLETA - SISTEMA DE GENERACIÓN DE BACKUP
## Especificación Técnica para Recreación en Otra Aplicación

**Versión:** 1.0
**Fecha:** 2025-12-12
**Sistema Origen:** CreativePythonAPI - Presentation Backup Generation
**Propósito:** Documentación completa para que un agente AI pueda recrear esta funcionalidad

---

# TABLA DE CONTENIDOS

1. [NIVEL 1: ARQUITECTURA DEL SISTEMA](#nivel-1-arquitectura-del-sistema)
2. [NIVEL 2: CONTRATOS DE DATOS](#nivel-2-contratos-de-datos)
3. [NIVEL 3: LÓGICA DE NEGOCIO](#nivel-3-lógica-de-negocio)
4. [NIVEL 4: IMPLEMENTACIÓN TÉCNICA](#nivel-4-implementación-técnica)
5. [NIVEL 5: CASOS EDGE Y TESTING](#nivel-5-casos-edge-y-testing)

---

# NIVEL 1: ARQUITECTURA DEL SISTEMA

## 1.1 Visión General del Sistema

El sistema de generación de backup de presentaciones es un servicio asíncrono que permite regenerar presentaciones PowerPoint completas a partir de archivos originales guardados (Excel + PPTX) y metadatos almacenados en base de datos.

### 1.1.1 Componentes Principales

```
┌─────────────────────────────────────────────────────────────┐
│                    CLIENTE (Frontend)                        │
└────────────────────┬────────────────────────────────────────┘
                     │ HTTP POST
                     │ /api/presentations/generate-backup
                     │ {presentation_id: 123}
                     ▼
┌─────────────────────────────────────────────────────────────┐
│                  ENDPOINT (FastAPI)                          │
│  - Validación de request                                     │
│  - Verificación de límites de concurrencia                   │
│  - Creación de tarea en TaskManager                          │
│  - Envío a ConcurrencyManager                                │
└────────────────────┬────────────────────────────────────────┘
                     │ Retorna inmediatamente
                     │ TaskCreatedResponse {task_id, status, position}
                     │
                     │ Ejecución en background
                     ▼
┌─────────────────────────────────────────────────────────────┐
│              BACKGROUND TASK WORKER                          │
│  - Inicializa COM (pythoncom.CoInitialize)                   │
│  - Adquiere lock COM (max 2 concurrentes)                    │
│  - Ejecuta servicio de generación                            │
│  - Actualiza progreso (10% → 100%)                           │
│  - Cleanup COM (pythoncom.CoUninitialize)                    │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│           PRESENTATION SERVICE (Core Logic)                  │
│                                                              │
│  1. Recuperar metadata de BD (30-40%)                        │
│     ├─ Query SQL Server: nw_Master                           │
│     └─ Validar existencia de presentación                    │
│                                                              │
│  2. Resolver rutas de archivos (40%)                         │
│     ├─ Limpiar display_name                                  │
│     ├─ Resolver output folder                                │
│     ├─ Localizar Excel file (con fallback)                   │
│     └─ Localizar PPTX file (con fallback)                    │
│                                                              │
│  3. Leer contenido de archivos (40-50%)                      │
│     ├─ Leer Excel bytes                                      │
│     └─ Leer PPTX bytes                                       │
│                                                              │
│  4. Crear request object (50-55%)                            │
│     └─ CreatePresentationRequest con create_backup=1         │
│                                                              │
│  5. Procesar Excel (55-65%)                                  │
│     ├─ Derivar is_phonetics de presentation_type             │
│     ├─ process_excel_file()                                  │
│     └─ Retornar ProcessedExcelData                           │
│                                                              │
│  6. Cargar imágenes PPTX existentes (65-70%)                 │
│     ├─ _load_existing_pptx_images()                          │
│     ├─ Glob *.jpg en output folder                           │
│     └─ Construir URLs de imágenes                            │
│                                                              │
│  7. Generar slides data (70-75%)                             │
│     ├─ _generate_slides_from_excel()                         │
│     ├─ Procesar cada fila del Excel                          │
│     ├─ Identificar tipo de slide (grupo/categoría/individual)│
│     ├─ Crear DetailItem para cada slide                      │
│     └─ Retornar PresentationData                             │
│                                                              │
│  8. Generar PowerPoint físico (75-90%)                       │
│     ├─ _generate_physical_powerpoint()                       │
│     ├─ Guardar PPTX original en disco                        │
│     ├─ compose_presentation_with_original_slides()           │
│     │   └─ Usa COM para combinar slides                      │
│     └─ Retornar paths de archivos generados                  │
│                                                              │
│  9. Actualizar BD con filename (90-95%)                      │
│     ├─ Extraer nombre de archivo generado                    │
│     ├─ UPDATE nw_Master SET MainPptFileName                  │
│     └─ Commit transacción                                    │
│                                                              │
│  10. Retornar resultado (100%)                               │
│      └─ {presentation_id, printable_path, total_slides}      │
└─────────────────────────────────────────────────────────────┘
```

### 1.1.2 Flujo de Datos

```
[Excel File (disk)]  ──┐
                       │
[PPTX File (disk)]  ───┼──► [process_excel_file] ──► ProcessedExcelData
                       │                                      │
[nw_Master (DB)]  ─────┘                                      │
                                                              │
                       ┌──────────────────────────────────────┘
                       │
                       ▼
          [_generate_slides_from_excel]
                       │
                       ├──► DetailItem[] (slide metadata)
                       │
                       ▼
         [_generate_physical_powerpoint]
                       │
                       ├──► Uses COM (win32com.client)
                       ├──► Opens original PPTX
                       ├──► Generates slides from templates
                       ├──► Combines at page_number position
                       │
                       ▼
              [backup_YYYYMMDD_HHMMSS.pptx]
                       │
                       ▼
              [UPDATE nw_Master in DB]
```

### 1.1.3 Stack Tecnológico

**Backend Framework:**
- FastAPI (async web framework)
- Python 3.9+

**Base de Datos:**
- SQL Server (via pyodbc)
- Database: BI_GUIDELINES
- Tablas principales: nw_Master, nw_Details

**Procesamiento de Office:**
- openpyxl (lectura de Excel)
- python-pptx (manipulación de PPTX - limitado)
- win32com.client / pythoncom (COM automation para PowerPoint - Windows only)

**Concurrencia:**
- asyncio (async/await)
- threading (para workers y locks)
- Custom TaskQueue + WorkerPool

**Almacenamiento:**
- File system (archivos PPTX, Excel, imágenes JPG)
- Estructura: `C:/inetpub/wwwroot/nw_slides/{project}/{display_name}/`

---

## 1.2 Arquitectura de Concurrencia

### 1.2.1 Componentes de Concurrencia

```
┌──────────────────────────────────────────────────────────────┐
│                    CONCURRENCY MANAGER                        │
│                                                              │
│  ┌────────────────┐         ┌──────────────────┐            │
│  │  TaskQueue     │────────►│  WorkerPool      │            │
│  │  (priority)    │         │  (N workers)     │            │
│  │  max: 100      │         │  max: 5          │            │
│  └────────────────┘         └──────────────────┘            │
│         │                            │                       │
│         │                            │                       │
│         │                            ▼                       │
│         │                   ┌─────────────────┐              │
│         │                   │  Worker Thread  │              │
│         │                   │  - Picks task   │              │
│         │                   │  - Executes     │              │
│         └──────────────────►│  - Updates      │              │
│                             └─────────────────┘              │
└──────────────────────────────────────────────────────────────┘
                                      │
                                      ▼
                         ┌────────────────────────┐
                         │   COM MANAGER          │
                         │   Semaphore(max=2)     │
                         │   - Limits concurrent  │
                         │     COM operations     │
                         └────────────────────────┘
                                      │
                                      ▼
                         ┌────────────────────────┐
                         │  PowerPoint COM        │
                         │  win32com.client       │
                         │  - Thread-specific     │
                         │  - CoInitialize()      │
                         │  - CoUninitialize()    │
                         └────────────────────────┘
```

### 1.2.2 Límites de Concurrencia

**Configuración por defecto (settings):**
```python
concurrency_max_workers = 5        # Máximo de workers simultáneos
concurrency_queue_size = 100       # Máximo de tareas en cola
task_timeout_seconds = 1800        # Timeout por tarea (30 min)
com_max_concurrent = 2             # Máximo de operaciones COM simultáneas
```

**Razón para COM limit = 2:**
- PowerPoint COM automation es pesado en recursos
- Múltiples instancias simultáneas causan inestabilidad
- 2 es el número seguro para la mayoría de sistemas Windows

### 1.2.3 Task Manager (In-Memory)

**Almacenamiento:**
- Diccionario en memoria: `Dict[task_id, task_info]`
- Thread-safe usando `threading.Lock()`
- **NOTA:** En producción con múltiples workers/servidores, usar Redis

**Estructura de Task:**
```python
{
    "task_id": "uuid-string",
    "task_type": "generate_backup",
    "status": "pending|processing|completed|failed",
    "description": "Generating backup for presentation_id: 123",
    "progress": 0-100,
    "result": None | Dict,
    "error": None | str,
    "metadata": {"presentation_id": 123},
    "created_at": "2025-12-12T10:30:00",
    "started_at": None | "2025-12-12T10:30:05",
    "completed_at": None | "2025-12-12T10:32:00"
}
```

---

## 1.3 Diagrama de Secuencia Completo

```
Cliente          Endpoint         TaskMgr    ConcurrMgr    BgWorker    Service         DB          FileSystem      COM
  │                │                 │            │            │           │            │              │             │
  │ POST /backup   │                 │            │            │           │            │              │             │
  ├───────────────►│                 │            │            │           │            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │ create_task()   │            │            │           │            │              │             │
  │                ├────────────────►│            │            │           │            │              │             │
  │                │◄────────────────┤            │            │           │            │              │             │
  │                │   task_id       │            │            │           │            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │ submit_task()   │            │            │           │            │              │             │
  │                ├────────────────────────────►│            │           │            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │     enqueue task        │           │            │              │             │
  │                │                 │            ├───────────►│           │            │              │             │
  │                │                 │            │            │           │            │              │             │
  │◄───────────────┤                 │            │            │           │            │              │             │
  │  202 Accepted  │                 │            │            │           │            │              │             │
  │  {task_id,     │                 │            │            │           │            │              │             │
  │   position: 2} │                 │            │            │           │            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │     worker picks task  │            │              │             │
  │                │                 │            │            ├──────────►│            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │  CoInitialize()        │              │             │
  │                │                 │            │            │           ├───────────────────────────────────────►│
  │                │                 │            │            │           │            │              │             │
  │                │                 │      update(10%)        │           │            │              │             │
  │                │                 │◄───────────────────────────────────┤            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │ generate_backup()      │              │             │
  │                │                 │            │            │           ├───────────►│              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │           │  SELECT    │              │             │
  │                │                 │            │            │           ├───────────►│              │             │
  │                │                 │            │            │           │◄───────────┤              │             │
  │                │                 │            │            │           │  metadata  │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │      update(40%)        │           │            │              │             │
  │                │                 │◄───────────────────────────────────┤            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │           │  read Excel│              │             │
  │                │                 │            │            │           ├────────────────────────►│              │
  │                │                 │            │            │           │◄──────────────────────────┤             │
  │                │                 │            │            │           │  bytes     │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │           │  read PPTX │              │             │
  │                │                 │            │            │           ├────────────────────────►│              │
  │                │                 │            │            │           │◄──────────────────────────┤             │
  │                │                 │            │            │           │  bytes     │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │      update(65%)        │           │            │              │             │
  │                │                 │◄───────────────────────────────────┤            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │  acquire_COM_lock()    │              │             │
  │                │                 │            │            │           ├───────────────────────────────────────►│
  │                │                 │            │            │           │◄───────────────────────────────────────┤
  │                │                 │            │            │           │  (blocking if 2 already active)        │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │           │  compose_presentation()   │             │
  │                │                 │            │            │           ├───────────────────────────────────────►│
  │                │                 │            │            │           │            │              │  PowerPoint.│
  │                │                 │            │            │           │            │              │  Application│
  │                │                 │            │            │           │◄───────────────────────────────────────┤
  │                │                 │            │            │           │  backup.pptx created      │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │  release_COM_lock()    │              │             │
  │                │                 │            │            │           ├───────────────────────────────────────►│
  │                │                 │            │            │           │            │              │             │
  │                │                 │      update(90%)        │           │            │              │             │
  │                │                 │◄───────────────────────────────────┤            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │           │  UPDATE    │              │             │
  │                │                 │            │            │           ├───────────►│              │             │
  │                │                 │            │            │           │◄───────────┤              │             │
  │                │                 │            │            │           │  COMMIT    │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │      update(100%)       │           │            │              │             │
  │                │                 │      result={...}       │           │            │              │             │
  │                │                 │◄───────────────────────────────────┤            │              │             │
  │                │                 │            │            │           │            │              │             │
  │                │                 │            │            │  CoUninitialize()      │              │             │
  │                │                 │            │            │           ├───────────────────────────────────────►│
  │                │                 │            │            │           │            │              │             │
  │  GET /tasks/   │                 │            │            │           │            │              │             │
  │  {task_id}     │                 │            │            │           │            │              │             │
  ├───────────────►│                 │            │            │           │            │              │             │
  │                │  get_task()     │            │            │           │            │              │             │
  │                ├────────────────►│            │            │           │            │              │             │
  │                │◄────────────────┤            │            │           │            │              │             │
  │◄───────────────┤                 │            │            │           │            │              │             │
  │  200 OK        │                 │            │            │           │            │              │             │
  │  {status:      │                 │            │            │           │            │              │             │
  │   completed,   │                 │            │            │           │            │              │             │
  │   result:{...}}│                 │            │            │           │            │              │             │
```

---

# NIVEL 2: CONTRATOS DE DATOS

## 2.1 Esquemas de Base de Datos

### 2.1.1 Tabla: nw_Master

**Propósito:** Almacena metadatos de presentaciones maestras

```sql
CREATE TABLE [BI_GUIDELINES].[dbo].[nw_Master] (
    -- Primary Key
    PresentationId INT IDENTITY(1,1) PRIMARY KEY,

    -- Project Information
    Project NVARCHAR(255) NOT NULL,                    -- Identificador del proyecto
    DisplayName NVARCHAR(255) NOT NULL,                -- Nombre de la presentación

    -- File References
    MainPptFileName NVARCHAR(255) NULL,                -- Archivo PPTX generado (ej: backup_20250112_143022.pptx)
    NameCandidateFileName NVARCHAR(255) NULL,          -- Archivo Excel original

    -- Background Configuration
    NameCandidateBGType NVARCHAR(50) NULL,             -- Tipo de fondo: 'Default', 'Rotate', 'Image'
    NameCandidateBGName NVARCHAR(255) NULL,            -- Nombre del fondo (puede ser pipe-separated: 'BMW_1|BrandDNA')

    -- Slide Configuration
    NameCandidateStartingSlide INT NULL DEFAULT 1,     -- Página donde insertar slides generados

    -- Presentation Metadata
    PresentationType NVARCHAR(50) NULL,                -- Tipo: 'Normal', 'Phonetics', 'BSR', etc.
    UploadedBy NVARCHAR(100) NULL,                     -- Usuario que creó
    BSRDisplayName NVARCHAR(255) NULL,                 -- Link móvil BSR

    -- Flags
    isParticipantsVote BIT NULL DEFAULT 0,             -- Habilitar voto de participantes
    isWideScreenPPT BIT NULL DEFAULT 0,                -- 0=4:3, 1=16:9
    isAWSLinkReq BIT NULL DEFAULT 0,                   -- Requerir links AWS

    -- Timestamps
    CreatedDate DATETIME NULL DEFAULT GETDATE(),
    ModifiedDate DATETIME NULL,

    -- Indexes
    INDEX IX_Project_DisplayName (Project, DisplayName)
);
```

### 2.1.2 Tabla: nw_Details

**Propósito:** Almacena detalles de cada slide individual

```sql
CREATE TABLE [BI_GUIDELINES].[dbo].[nw_Details] (
    -- Primary Key
    DetailId INT IDENTITY(1,1) PRIMARY KEY,

    -- Foreign Key
    PresentationId INT NOT NULL,
    FOREIGN KEY (PresentationId) REFERENCES nw_Master(PresentationId) ON DELETE CASCADE,

    -- Slide Information
    SlideNumber INT NOT NULL,                          -- Número de slide (1-based)
    SlideType NVARCHAR(50) NULL,                       -- 'Group', 'Name', 'Image', 'Summary'
    SlideBGFileName NVARCHAR(500) NULL,                -- Path a imagen de fondo o slide
    SlideDescription NVARCHAR(500) NULL,               -- Descripción del slide

    -- Group Information
    GroupName NVARCHAR(255) NULL,                      -- Nombre del grupo actual
    GroupLetter NVARCHAR(10) NULL,                     -- Letra del grupo: 'A', 'B', 'C'

    -- Candidate Information
    Category NVARCHAR(255) NULL,                       -- Categoría del candidato
    Name NVARCHAR(500) NULL,                           -- Nombre del candidato
    Rationale NVARCHAR(MAX) NULL,                      -- Razón/descripción
    Notation NVARCHAR(500) NULL,                       -- Notación fonética
    Kana NVARCHAR(500) NULL,                           -- Representación en kana (japonés)
    LogoFileName NVARCHAR(255) NULL,                   -- Logo asociado

    -- Template Information
    TemplateId INT NULL,                               -- ID de template usado (0 = imagen)
    NameSubGroup NVARCHAR(255) NULL,                   -- Subgrupo del nombre

    -- Indexes
    INDEX IX_PresentationId_SlideNumber (PresentationId, SlideNumber)
);
```

**Ejemplo de datos:**

```sql
-- Registro en nw_Master
INSERT INTO nw_Master (
    Project, DisplayName, MainPptFileName, NameCandidateFileName,
    NameCandidateBGType, NameCandidateBGName, NameCandidateStartingSlide,
    PresentationType, UploadedBy, isParticipantsVote, isWideScreenPPT
) VALUES (
    'NW_PROJECT_2025',
    'Name Evaluation Q1',
    'backup_20250112_143022.pptx',
    'candidates_data.xlsx',
    'Default',
    'Default',
    1,
    'Normal',
    'analyst',
    1,
    0
);

-- Registros en nw_Details (para PresentationId = 12345)
-- Slide 1: Imagen de PPTX original
INSERT INTO nw_Details VALUES (12345, 1, 'Image', '/static/nw/NW_PROJECT_2025/slide_1.jpg', 'Title Slide', '', '', '', '', '', '', '', '', 0, '', '');

-- Slide 2: Header de grupo
INSERT INTO nw_Details VALUES (12345, 2, 'Group', '', 'Group A Header', 'Innovation Names', 'A', '', '', '', '', '', '', 1, '', 'A');

-- Slide 3: Nombre individual
INSERT INTO nw_Details VALUES (12345, 3, 'Name', '', 'Candidate 1', '', 'A', 'Technology', 'TechNova', 'Modern tech-focused brand', 'tek-NOH-vuh', '', 'techlogo.png', 2, '', 'A');
```

---

## 2.2 Modelos de Datos (Python/Pydantic)

### 2.2.1 Request Models

#### GenerateBackupRequest

```python
from pydantic import BaseModel, Field

class GenerateBackupRequest(BaseModel):
    """Request para generar backup de una presentación existente."""

    presentation_id: int = Field(
        ...,
        description="ID de la presentación en nw_Master",
        gt=0,
        example=12345
    )
```

#### CreatePresentationRequest (usado internamente)

```python
from typing import Optional
from pydantic import BaseModel, Field

class CreatePresentationRequest(BaseModel):
    """Request interno para el servicio de presentación."""

    # Archivos binarios
    excel_file: bytes = Field(..., description="Contenido del Excel")
    pptx_file: bytes = Field(..., description="Contenido del PPTX")
    excel_filename: str = Field("data.xlsx", description="Nombre del archivo Excel")
    pptx_filename: str = Field("presentation.pptx", description="Nombre del archivo PPTX")

    # Metadata del proyecto
    project: str = Field(..., description="Identificador del proyecto")
    display_name: str = Field(..., description="Nombre de la presentación")
    presentation_type: str = Field("Normal", description="Tipo de presentación")
    user_name: str = Field(..., description="Usuario creador")
    mobile_link_bsr: Optional[str] = Field(default=None)

    # Configuración de slides
    page_number: int = Field(1, ge=1, description="Página de inserción")
    background_type: str = Field("Default", description="Tipo de fondo")
    background_name: str = Field("", description="Nombre del fondo")

    # Flags
    participant_vote: int = Field(1, ge=0, le=1)
    is_wide_ppt: int = Field(0, ge=0, le=1, description="0=4:3, 1=16:9")
    is_aws_email: int = Field(0, ge=0, le=1)
    amazon_link_required: int = Field(0, ge=0, le=1)

    # Procesamiento Excel
    has_groups: bool = Field(True, description="El Excel contiene grupos")
    test_name_order: str = Field("Default", description="'Default', 'Randomize', 'Randomize_top_5'")

    # Control
    project_type: str = Field("NW", description="'NW', 'BSR', 'NSR', 'DW'")
    create_backup: int = Field(0, ge=0, le=1, description="Crear archivo físico")
    overwrite_existing: bool = Field(False)
```

### 2.2.2 Response Models

#### GenerateBackupResponse

```python
from typing import Optional, List, Dict
from pydantic import BaseModel, Field, ConfigDict

class GenerateBackupResponse(BaseModel):
    """Response del endpoint de generación de backup."""

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "success": True,
                "message": "Backup presentation generated successfully",
                "presentation_id": 12345,
                "printable_path": "C:/inetpub/wwwroot/nw2/nw_slides/TestProject/Presentations/backup_20250123_143022.pptx",
                "file_name": "backup_20250123_143022.pptx",
                "download_token": "eyJ0eXAiOiJKV1QiLCJhbGc...",
                "total_slides": 24,
                "warnings": "None",
                "missing_files": []
            }
        }
    )

    success: bool = Field(..., description="Indica si la generación fue exitosa")
    message: str = Field(..., description="Mensaje descriptivo")
    presentation_id: int = Field(..., description="ID de la presentación")

    printable_path: Optional[str] = Field(
        default=None,
        description="Path completo al archivo PPTX generado (null si falló)"
    )
    file_name: Optional[str] = Field(
        default=None,
        description="Nombre del archivo generado (null si falló)"
    )
    download_token: Optional[str] = Field(
        default=None,
        description="Token JWT para descargar el archivo"
    )
    total_slides: Optional[str] = Field(
        default=None,
        description="Total de slides generados (null si falló)"
    )
    warnings: str = Field(
        default="None",
        description="Warnings encontrados durante generación"
    )
    missing_files: List[Dict[str, str]] = Field(
        default_factory=list,
        description="Lista de archivos faltantes. Cada item: {file_type, expected_path, instructions}"
    )
```

#### TaskCreatedResponse

```python
from datetime import datetime
from pydantic import BaseModel, Field

class TaskCreatedResponse(BaseModel):
    """Response cuando se crea una tarea asíncrona."""

    task_id: str = Field(..., description="UUID de la tarea")
    status: str = Field(..., description="'pending', 'queued', 'processing'")
    message: str = Field(..., description="Mensaje descriptivo")
    task_type: str = Field(..., description="Tipo de tarea: 'generate_backup'")

    created_at: str = Field(..., description="ISO timestamp de creación")
    status_url: str = Field(..., description="URL para consultar estado: /api/presentations/tasks/{task_id}")

    position: Optional[int] = Field(
        default=None,
        description="Posición en la cola (null si ya se está ejecutando)"
    )
    estimated_wait_seconds: Optional[int] = Field(
        default=None,
        description="Tiempo estimado de espera en segundos"
    )
```

### 2.2.3 Data Models

#### ProcessedExcelData

```python
from dataclasses import dataclass, field
from typing import List

@dataclass
class ProcessedExcelData:
    """Datos procesados del archivo Excel."""

    # Listas paralelas (mismo índice = misma fila)
    lst_types: List[str] = field(default_factory=list)           # Marcador de grupo: 'A', 'B', 'C', o ''
    lst_categories: List[str] = field(default_factory=list)      # Categoría del candidato
    lst_names: List[str] = field(default_factory=list)           # Nombre del candidato
    lst_rationales: List[str] = field(default_factory=list)      # Razón/descripción
    lst_notations: List[str] = field(default_factory=list)       # Notación fonética
    lst_kana: List[str] = field(default_factory=list)            # Kana (japonés)
    lst_logos: List[str] = field(default_factory=list)           # Nombre de archivo de logo
    lst_name_sub_groups: List[str] = field(default_factory=list) # Subgrupo (Group1 o Group2)
    lst_group_letters: List[str] = field(default_factory=list)   # Letra de grupo actual

    # Metadata calculada
    total_rows_processed: int = 0
    candidate_count: int = 0          # Filas que no son marcadores de grupo
    group_count: int = 0              # Filas que son marcadores de grupo (A, B, C)
    has_groups: bool = False
    is_phonetics: bool = False

    # Índices útiles
    group_row_indexes: List[int] = field(default_factory=list)      # [2, 5, 8] si hay grupos en esas filas
    candidate_row_indexes: List[int] = field(default_factory=list)  # [0, 1, 3, 4, 6, 7]
    rotation_reset_indexes: List[int] = field(default_factory=list) # Puntos donde resetear rotación de backgrounds
```

**Ejemplo de datos:**

```python
ProcessedExcelData(
    lst_types=['', '', 'A', '', '', 'B', '', ''],
    lst_categories=['Tech', 'Tech', '', 'Health', 'Health', '', 'Finance', 'Finance'],
    lst_names=['TechNova', 'InnovateTech', 'Innovation', 'HealthPlus', 'VitaCare', 'Wellness', 'MoneyWise', 'FinSecure'],
    lst_rationales=['Modern tech brand', 'Tech innovation', 'Group A header', 'Healthcare leader', 'Vital care', 'Group B header', 'Financial wisdom', 'Secure finances'],
    lst_notations=['tek-NOH-vuh', 'in-NOH-vayt-tek', '', 'helth-plus', 'VEE-tuh-care', '', 'MUN-ee-wize', 'fin-see-KYOOR'],
    lst_kana=['', '', '', '', '', '', '', ''],
    lst_logos=['tech1.png', 'tech2.png', '', 'health1.png', 'health2.png', '', 'fin1.png', 'fin2.png'],
    lst_name_sub_groups=['', '', '', '', '', '', '', ''],
    lst_group_letters=['', '', 'A', 'A', 'A', 'B', 'B', 'B'],
    total_rows_processed=8,
    candidate_count=6,
    group_count=2,
    has_groups=True,
    is_phonetics=False,
    group_row_indexes=[2, 5],
    candidate_row_indexes=[0, 1, 3, 4, 6, 7],
    rotation_reset_indexes=[0, 2, 5]
)
```

#### DetailItem

```python
from typing import Optional
from pydantic import BaseModel, Field

class DetailItem(BaseModel):
    """Representa un slide individual en nw_Details."""

    slide_number: int = Field(..., ge=1, description="Número de slide (1-based)")
    slide_type: str = Field(..., description="'Group', 'Name', 'Image', 'Summary'")

    slide_bg_file_name: Optional[str] = Field(
        default=None,
        description="Path a imagen de fondo o URL del slide"
    )
    slide_description: Optional[str] = Field(
        default=None,
        description="Descripción del slide"
    )

    group_name: Optional[str] = Field(
        default=None,
        description="Nombre del grupo actual"
    )
    category: Optional[str] = Field(
        default=None,
        description="Categoría del candidato"
    )
    name: Optional[str] = Field(
        default=None,
        description="Nombre del candidato"
    )
    rationale: Optional[str] = Field(
        default=None,
        description="Razón/descripción"
    )
    notation: Optional[str] = Field(
        default=None,
        description="Notación fonética"
    )
    kana: Optional[str] = Field(
        default=None,
        description="Representación en kana"
    )
    logo_filename: Optional[str] = Field(
        default=None,
        description="Archivo de logo"
    )
    template_id: int = Field(..., description="ID de template usado (0=imagen)")
    name_sub_group: Optional[str] = Field(
        default=None,
        description="Subgrupo del nombre"
    )
    group_letter: Optional[str] = Field(
        default="",
        description="Letra del grupo: 'A', 'B', 'C'"
    )
```

#### PPTXConversionResponse

```python
from typing import List
from pydantic import BaseModel, Field

class PPTXConversionResponse(BaseModel):
    """Resultado de convertir PPTX a imágenes."""

    message: str = Field(..., description="Mensaje de estado")
    conversion_id: str = Field(..., description="ID único de la conversión (folder name)")
    project_type: str = Field(..., description="'NW', 'BSR', etc.")

    images: List[str] = Field(
        default_factory=list,
        description="Lista de URLs/paths a imágenes JPG de cada slide"
    )
    thumbnails: List[str] = Field(
        default_factory=list,
        description="Lista de URLs/paths a thumbnails"
    )
    titles: List[str] = Field(
        default_factory=list,
        description="Lista de títulos extraídos de cada slide"
    )

    total_images: int = Field(..., description="Total de imágenes generadas")
    pptx_file: str = Field(..., description="Nombre del archivo PPTX original")
```

---

## 2.3 Estructura de Archivos Excel

### 2.3.1 Formato del Excel

**Archivo:** `candidates_data.xlsx`
**Hoja:** Primera hoja del workbook
**Encabezados:** Fila 1 (HEADER_ROW_INDEX = 1)
**Datos:** Fila 2 en adelante

**Columnas:**

| Col | Letra | Nombre          | Tipo   | Descripción                                      | Ejemplo              |
|-----|-------|-----------------|--------|--------------------------------------------------|----------------------|
| 1   | A     | Type/Marker     | String | Marcador de grupo: 'A', 'B', 'C', o vacío       | 'A', '', 'B'         |
| 2   | B     | Category        | String | Categoría del candidato                          | 'Technology'         |
| 3   | C     | Name            | String | Nombre del candidato (puede incluir notación)   | 'TechNova (tek-NOH)' |
| 4   | D     | Rationale       | String | Descripción/razón del candidato                  | 'Modern tech brand'  |
| 5   | E     | Kana            | String | Pronunciación en kana (para japonés)             | 'テクノバ'            |
| 6   | F     | Logo            | String | Nombre de archivo de logo                        | 'techlogo.png'       |
| 7   | G     | Group1/SubGroup | String | Subgrupo primario                                | 'Premium'            |
| 8   | H     | Group2          | String | Subgrupo secundario (fallback si G está vacío)  | 'Standard'           |

**Ejemplo de Excel:**

```
Row | A   | B          | C                    | D                      | E    | F           | G       | H
----|-----|------------|----------------------|------------------------|------|-------------|---------|--------
1   | Type| Category   | Name                 | Rationale              | Kana | Logo        | Group1  | Group2
2   |     | Technology | TechNova (tek-NOH)   | Modern tech brand      |      | tech1.png   |         |
3   |     | Technology | InnovateTech         | Innovation focused     |      | tech2.png   |         |
4   | A   |            | Innovation Names     | Group A Header         |      |             |         |
5   |     | Healthcare | HealthPlus           | Healthcare leader      |      | health1.png |         |
6   |     | Healthcare | VitaCare             | Vital care provider    |      | health2.png |         |
7   | B   |            | Wellness Names       | Group B Header         |      |             |         |
8   |     | Finance    | MoneyWise            | Financial wisdom       |      | fin1.png    |         |
9   |     | Finance    | FinSecure            | Secure finances        |      | fin2.png    |         |
```

### 2.3.2 Reglas de Procesamiento

**1. Detección de Marcadores de Grupo:**
- Si columna A contiene solo una letra (A-Z), es un marcador de grupo
- Regex: `^\s*([A-Za-z])\s*(?:[:\)\.\-–—]|\Z)`
- Ejemplos válidos: "A", "A:", "A)", "A -", "B."
- Ejemplos inválidos: "Group A", "A1", "AA"

**2. Extracción de Name y Notation:**
- Si Name contiene paréntesis: `TechNova (tek-NOH-vuh)`
- Extraer name: `TechNova`
- Extraer notation: `tek-NOH-vuh`
- Regex: `^(?P<name>[^\(]+?)(?:\((?P<notation>.+)\))?$`

**3. Detección de Slides de Categoría:**
- Si Category tiene valor Y todos los demás campos están vacíos (Name, Rationale, Notation, Kana, Logo)
- Se genera un slide de imagen con la categoría

**4. Detección de Grouped Slides:**
- Si Name contiene `##` o `$$`: delimitador de múltiples nombres en un slide
- Ejemplo: `Name1##Name2##Name3`
- Si Group1 (columna G) tiene valor: es un grouped slide

**5. Test Name Order (Randomización):**
- **Default:** Mantener orden original
- **Randomize:** Shuffle todas las filas
- **Randomize_top_5:** Shuffle solo las primeras 5 filas, resto en orden

---

## 2.4 Estructura del File System

### 2.4.1 Organización de Carpetas

```
C:/inetpub/wwwroot/
└── nw_slides/                          # Root para proyectos NW
    └── {project}/                      # Nombre del proyecto (ej: NW_PROJECT_2025)
        └── {display_name}/             # Nombre de la presentación (sanitized)
            ├── original_data.xlsx      # Excel original guardado
            ├── presentation.pptx       # PPTX original guardado
            ├── slide_1.jpg             # Slide 1 convertido a imagen
            ├── slide_2.jpg             # Slide 2 convertido a imagen
            ├── ...
            ├── Presentations/          # Subcarpeta para backups generados
            │   ├── backup_20250112_143022.pptx
            │   ├── backup_20250115_091544.pptx
            │   └── ...
            └── Categories/             # Subcarpeta para imágenes de categorías (opcional)
                ├── category_Technology.jpg
                └── category_Healthcare.jpg
```

**Ejemplo real:**
```
C:/inetpub/wwwroot/nw_slides/NW_PROJECT_2025/Name_Evaluation_Q1/
├── candidates_data.xlsx                           # 45 KB
├── template_presentation.pptx                     # 2.3 MB
├── slide_1.jpg                                    # 156 KB
├── slide_2.jpg                                    # 142 KB
├── slide_3.jpg                                    # 138 KB
├── Presentations/
│   └── backup_20250112_143022.pptx               # 8.7 MB (archivo final combinado)
└── Categories/
    └── category_5_Technology.jpg                  # 89 KB
```

### 2.4.2 Resolución de Paths

**Función: `resolve_project_output()`**

```python
def resolve_project_output(
    display_name: str,
    project_type: str,
    fallback_subdir: str = "generated_presentations"
) -> Tuple[Path, str]:
    """
    Resuelve el directorio de salida para un proyecto.

    Args:
        display_name: Nombre de la presentación (puede contener path)
        project_type: 'NW', 'BSR', 'NSR', 'DW'
        fallback_subdir: Subcarpeta si no hay project en display_name

    Returns:
        (output_base, url_root):
        - output_base: Path absoluto: C:/inetpub/wwwroot/nw_slides/Project/Display
        - url_root: URL relativo: /static/nw/Project/Display
    """
    # Limpiar display_name
    clean_name = sanitize_folder_name(display_name)

    # Determinar root según project_type
    if project_type.upper() == 'NW':
        base_root = Path("C:/inetpub/wwwroot/nw_slides")
        url_prefix = "/static/nw"
    elif project_type.upper() == 'BSR':
        base_root = Path("C:/inetpub/wwwroot/bsr_slides")
        url_prefix = "/static/bsr"
    else:
        base_root = Path("C:/inetpub/wwwroot/nw_slides")
        url_prefix = "/static/nw"

    # Si display_name contiene '/', asumir que ya incluye project
    if '/' in display_name or '\\' in display_name:
        # display_name = "NW_PROJECT_2025/Name_Evaluation_Q1"
        output_base = base_root / clean_name
    else:
        # display_name = "Name_Evaluation_Q1" → usar fallback
        output_base = base_root / fallback_subdir / clean_name

    # Crear directorio si no existe
    output_base.mkdir(parents=True, exist_ok=True)

    # Construir URL relativo
    relative_path = output_base.relative_to(base_root)
    url_root = f"{url_prefix}/{relative_path.as_posix()}"

    return output_base, url_root
```

**Función: `sanitize_folder_name()`**

```python
import re

def sanitize_folder_name(name: str) -> str:
    """
    Sanitiza un nombre para uso como carpeta.

    - Reemplaza caracteres inválidos con '_'
    - Limita longitud a 200 caracteres
    - Previene path traversal

    Args:
        name: Nombre original

    Returns:
        Nombre sanitizado
    """
    # Remover path traversal
    name = name.replace('..', '')

    # Reemplazar caracteres inválidos en Windows
    invalid_chars = r'[<>:"|?*\x00-\x1f]'
    sanitized = re.sub(invalid_chars, '_', name)

    # Reemplazar múltiples espacios/guiones con uno solo
    sanitized = re.sub(r'[\s_-]+', '_', sanitized)

    # Remover guiones bajos al inicio/final
    sanitized = sanitized.strip('_')

    # Limitar longitud
    if len(sanitized) > 200:
        sanitized = sanitized[:200]

    return sanitized
```

---

# NIVEL 3: LÓGICA DE NEGOCIO

## 3.1 Procesamiento de Excel

### 3.1.1 Función: `process_excel_file()`

**Archivo:** `app/services/excel_service.py`

```python
from openpyxl import load_workbook
from io import BytesIO
import random
import re
from typing import Tuple

# Constantes
HEADER_ROW_INDEX = 1
TYPE_COLUMN = 1        # A
CATEGORY_COLUMN = 2    # B
NAME_COLUMN = 3        # C
RATIONALE_COLUMN = 4   # D
KANA_COLUMN = 5        # E
LOGO_COLUMN = 6        # F
NAME_SUBGROUP_COLUMN = 7  # G (Group1)

GROUP_MARKERS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
GROUP_DELIMITER = "##"
_group_marker_regex = re.compile(r"^\s*([A-Za-z])\s*(?:[:\)\.\-–—]|\Z)")
_name_notation_regex = re.compile(r"^(?P<name>[^\(]+?)(?:\((?P<notation>.+)\))?$")


def process_excel_file(
    file_content: bytes,
    is_phonetics: bool = False,
    has_groups: bool = True,
    test_name_order: str = "Default"
) -> ProcessedExcelData:
    """
    Procesa archivo Excel y extrae datos de candidatos.

    Args:
        file_content: Bytes del archivo Excel
        is_phonetics: Si True, procesar columna Kana
        has_groups: Si True, buscar marcadores de grupo
        test_name_order: 'Default', 'Randomize', 'Randomize_top_5'

    Returns:
        ProcessedExcelData con listas paralelas de datos
    """
    # 1. Cargar workbook
    wb = load_workbook(BytesIO(file_content), data_only=True)
    sheet = wb.active

    # 2. Leer todas las filas (skip header)
    all_rows = []
    max_row = sheet.max_row

    for row_idx in range(HEADER_ROW_INDEX + 1, max_row + 1):
        row_data = {}

        # Leer columnas
        type_cell = sheet.cell(row_idx, TYPE_COLUMN).value
        category_cell = sheet.cell(row_idx, CATEGORY_COLUMN).value
        name_cell = sheet.cell(row_idx, NAME_COLUMN).value
        rationale_cell = sheet.cell(row_idx, RATIONALE_COLUMN).value
        kana_cell = sheet.cell(row_idx, KANA_COLUMN).value
        logo_cell = sheet.cell(row_idx, LOGO_COLUMN).value
        subgroup_cell = sheet.cell(row_idx, NAME_SUBGROUP_COLUMN).value
        subgroup2_cell = sheet.cell(row_idx, NAME_SUBGROUP_COLUMN + 1).value  # Column H

        # Limpiar valores
        row_data['type'] = _clean(type_cell)
        row_data['category'] = _clean(category_cell)
        row_data['name_raw'] = _clean(name_cell)
        row_data['rationale'] = _clean(rationale_cell)
        row_data['kana'] = _normalize_kana(kana_cell, is_phonetics)
        row_data['logo'] = _clean(logo_cell)
        row_data['subgroup'] = _clean(subgroup_cell) or _clean(subgroup2_cell)  # Fallback a H si G vacío

        # Skip fila completamente vacía
        if not any([
            row_data['type'],
            row_data['category'],
            row_data['name_raw'],
            row_data['rationale']
        ]):
            continue

        all_rows.append(row_data)

    # 3. Aplicar test_name_order (randomización)
    all_rows = _apply_test_name_order(all_rows, test_name_order)

    # 4. Procesar cada fila y extraer datos
    result = ProcessedExcelData()
    current_group_letter = ""

    for row_data in all_rows:
        # Detectar marcador de grupo
        type_val = row_data['type']
        is_group = False

        if has_groups and type_val:
            match = _group_marker_regex.match(type_val)
            if match:
                # Es un marcador de grupo (A, B, C, etc.)
                current_group_letter = match.group(1).upper()
                is_group = True
                result.lst_types.append(current_group_letter)
            else:
                result.lst_types.append("")
        else:
            result.lst_types.append("")

        # Extraer name y notation
        name_raw = row_data['name_raw']
        if is_group:
            # Para grupos, el "name" es el nombre del grupo
            name = name_raw
            notation = ""
        else:
            # Para candidatos, separar name de notation
            name, notation = _split_name_and_notation(name_raw)

        # Agregar a listas paralelas
        result.lst_categories.append(row_data['category'])
        result.lst_names.append(name)
        result.lst_rationales.append(row_data['rationale'])
        result.lst_notations.append(notation)
        result.lst_kana.append(row_data['kana'])
        result.lst_logos.append(row_data['logo'])
        result.lst_name_sub_groups.append(row_data['subgroup'])
        result.lst_group_letters.append(current_group_letter)

    # 5. Establecer metadata
    result.is_phonetics = is_phonetics
    # __post_init__ calculará automáticamente:
    # - total_rows_processed
    # - candidate_count, group_count, has_groups
    # - group_row_indexes, candidate_row_indexes, rotation_reset_indexes
    result.__post_init__()

    return result


def _clean(value) -> str:
    """Limpia valor de celda."""
    if value is None or str(value).strip().lower() == "none":
        return ""
    return str(value).strip()


def _normalize_kana(value, is_phonetics: bool) -> str:
    """Normaliza valor kana."""
    value = _clean(value)
    if is_phonetics and value:
        # Escapar comillas dobles
        value = value.replace('"', '""')
    return value


def _split_name_and_notation(raw_value: str) -> Tuple[str, str]:
    """
    Divide 'Name (notation)' en dos partes.

    Ejemplos:
        'TechNova (tek-NOH-vuh)' → ('TechNova', 'tek-NOH-vuh')
        'SimpleName' → ('SimpleName', '')
    """
    if not raw_value:
        return "", ""

    raw_value = raw_value.strip()
    match = _name_notation_regex.match(raw_value)

    if not match:
        return raw_value, ""

    name = match.group("name").strip()
    notation = match.group("notation")

    return name, notation.strip() if notation else ""


def _apply_test_name_order(all_rows: list, test_name_order: str) -> list:
    """
    Aplica randomización según estrategia.

    Args:
        all_rows: Lista de filas del Excel
        test_name_order: 'Default', 'Randomize', 'Randomize_top_5'

    Returns:
        Lista reordenada
    """
    if test_name_order == "Default":
        return all_rows

    elif test_name_order == "Randomize":
        randomized = all_rows.copy()
        random.shuffle(randomized)
        return randomized

    elif test_name_order == "Randomize_top_5":
        if len(all_rows) <= 5:
            randomized = all_rows.copy()
            random.shuffle(randomized)
            return randomized
        else:
            top_5 = all_rows[:5].copy()
            rest = all_rows[5:]
            random.shuffle(top_5)
            return top_5 + rest

    return all_rows
```

---

## 3.2 Generación de Slides desde Excel

### 3.2.1 Función: `_generate_slides_from_excel()`

**Archivo:** `app/services/presentation_service.py:1268`

Esta función es el **corazón de la lógica de negocio**. Convierte los datos del Excel en una lista de `DetailItem` que representa cada slide.

```python
from pathlib import Path
from typing import List, Dict, Any

def _generate_slides_from_excel(
    self,
    *,
    excel_data: ProcessedExcelData,
    pptx_data: PPTXConversionResponse,
    request: CreatePresentationRequest,
) -> Dict[str, Any]:
    """
    Genera estructura de slides combinando PPTX original + slides del Excel.

    Orden de slides:
    1. Slides del PPTX original ANTES de page_number
    2. Slides generados del Excel (grupos, categorías, nombres)
    3. Slide de resumen (solo para NW/DW)
    4. Slides restantes del PPTX original DESPUÉS de los generados

    Args:
        excel_data: Datos procesados del Excel
        pptx_data: Imágenes del PPTX convertido
        request: Request original

    Returns:
        Dict con:
            - details: List[DetailItem]
            - total_slides: int
    """
    details: List[DetailItem] = []
    last_group_name = ""
    last_group_letter = ""

    # Obtener template ID por defecto
    default_template = self._get_template_metadata("Default")
    default_template_id = default_template["template_id"]

    # Variables de seguimiento
    ppt_images = pptx_data.images or []
    slide_number = 1

    # ========================================
    # FASE 1: Slides del PPTX ANTES de page_number
    # ========================================
    prefix_count = 0
    if request.page_number > 1 and ppt_images:
        prefix_count = min(request.page_number - 1, len(ppt_images))

        for idx in range(prefix_count):
            image_path = ppt_images[idx]
            description = Path(image_path).stem.replace("_", " ") if image_path else f"Slide {idx+1}"

            details.append(DetailItem(
                slide_number=slide_number,
                slide_type="Image",
                slide_bg_file_name=Path(image_path).as_posix() if image_path else "",
                slide_description=description,
                group_name="",
                category="",
                name="",
                rationale="",
                notation="",
                kana="",
                logo_filename="",
                template_id=0,  # 0 indica que es una imagen, no un template
                name_sub_group="",
                group_letter=""
            ))
            slide_number += 1

    # ========================================
    # FASE 2: Slides generados del Excel
    # ========================================
    slide_number = request.page_number  # Resetear a page_number
    total_rows = excel_data.total_rows_processed

    for index in range(total_rows):
        # Extraer datos de esta fila
        marker = (excel_data.lst_types[index] or "").strip().upper()
        category = (excel_data.lst_categories[index] or "").strip()
        name = (excel_data.lst_names[index] or "").strip()
        rationale = (excel_data.lst_rationales[index] or "").strip()
        notation = (excel_data.lst_notations[index] or "").strip()
        kana = (excel_data.lst_kana[index] or "").strip()
        logo = (excel_data.lst_logos[index] or "").strip()
        name_sub_group = (excel_data.lst_name_sub_groups[index] or "").strip()

        # Detectar tipo de slide
        is_group_marker = marker in GROUP_MARKERS
        has_delimiter = "##" in name or "$$" in name
        is_grouped_slide = bool(name_sub_group)

        # -------------------------------------------------
        # CASO 1: Slide de Categoría (solo imagen)
        # -------------------------------------------------
        # Condición: category tiene valor Y todos los demás están vacíos
        if category and (not name and not rationale and not notation and not kana and not logo):
            logger.info(
                "🖼️ Adding category header image slide: Index=%s, Category='%s'",
                index, category
            )

            # Generar imagen de categoría
            image_path = pptx_builder_service.generate_category_slide_image(
                category=category,
                display_name=request.display_name,
                slide_number=slide_number,
                project_type=request.project_type,
            )

            group_letter = excel_data.lst_group_letters[index] if index < len(excel_data.lst_group_letters) else ""

            details.append(DetailItem(
                slide_number=slide_number,
                slide_type="Image",
                slide_bg_file_name=image_path or "",
                slide_description=category,
                group_name=last_group_name,
                category=category,
                name=name,
                rationale=rationale,
                notation=notation,
                kana=kana,
                logo_filename=logo,
                template_id=0,  # Imagen generada
                name_sub_group=name_sub_group,
                group_letter=group_letter,
            ))
            slide_number += 1

        # -------------------------------------------------
        # CASO 2: Slide de Grupo (header: A, B, C)
        # -------------------------------------------------
        elif is_group_marker:
            logger.info(f"🟢 Adding group header slide: Index={index}, Marker='{marker}'")

            detail = self._create_group_slide(
                excel_data=excel_data,
                index=index,
                slide_number=slide_number,
                request=request,
                default_template_id=default_template_id,
            )

            last_group_name = detail.group_name or last_group_name
            last_group_letter = detail.group_letter or last_group_letter

            details.append(detail)
            slide_number += 1

        # -------------------------------------------------
        # CASO 3: Slide con Múltiples Nombres (grouped)
        # -------------------------------------------------
        elif is_grouped_slide or has_delimiter:
            logger.info(f"🟡 Adding grouped names slide: Index={index}, SubGroup='{name_sub_group}'")

            detail = self._create_individual_slide(
                excel_data=excel_data,
                index=index,
                slide_number=slide_number,
                request=request,
                current_group=last_group_name,
                default_template_id=default_template_id,
            )

            details.append(detail)
            slide_number += 1

        # -------------------------------------------------
        # CASO 4: Slide Individual Normal
        # -------------------------------------------------
        else:
            logger.info(f"🔵 Adding individual name slide: Index={index}, Name='{name}'")

            detail = self._create_individual_slide(
                excel_data=excel_data,
                index=index,
                slide_number=slide_number,
                request=request,
                current_group=last_group_name,
                default_template_id=default_template_id,
            )

            details.append(detail)
            slide_number += 1

    # ========================================
    # FASE 3: Slide de Resumen (solo NW/DW)
    # ========================================
    if request.project_type.lower() in {"nw", "dw"}:
        summary_slide = self._create_summary_slide(
            slide_number=slide_number,
            last_group=last_group_name,
            request=request,
            default_template_id=default_template_id,
            last_group_letter=last_group_letter,
        )
        details.append(summary_slide)
        slide_number += 1

    # ========================================
    # FASE 4: Slides restantes del PPTX
    # ========================================
    if ppt_images and prefix_count < len(ppt_images):
        for tail_idx, image_path in enumerate(ppt_images[prefix_count:], start=1):
            description = Path(image_path).stem.replace("_", " ") if image_path else f"Slide {prefix_count + tail_idx}"

            details.append(DetailItem(
                slide_number=slide_number,
                slide_type="Image",
                slide_bg_file_name=Path(image_path).as_posix() if image_path else "",
                slide_description=description,
                group_name=last_group_name,
                category="",
                name="",
                rationale="",
                notation="",
                kana="",
                logo_filename="",
                template_id=0,
                name_sub_group="",
                group_letter=last_group_letter
            ))
            slide_number += 1

    # ========================================
    # Retornar resultado
    # ========================================
    return {
        "details": details,
        "total_slides": len(details)
    }
```

### 3.2.2 Funciones Auxiliares

#### `_create_group_slide()`

```python
def _create_group_slide(
    self,
    *,
    excel_data: ProcessedExcelData,
    index: int,
    slide_number: int,
    request: CreatePresentationRequest,
    default_template_id: int,
) -> DetailItem:
    """
    Crea DetailItem para slide de grupo (A, B, C, etc.).

    Características:
    - slide_type = "Group"
    - group_name = nombre del grupo (lst_names[index])
    - group_letter = marcador (lst_types[index])
    - template_id = ID del template de grupo
    """
    marker = excel_data.lst_types[index].strip().upper()
    group_name = excel_data.lst_names[index].strip()

    # Obtener template específico para grupos
    group_template = self._get_template_metadata("Group")
    template_id = group_template["template_id"]

    return DetailItem(
        slide_number=slide_number,
        slide_type="Group",
        slide_bg_file_name="",  # Se aplica en el builder
        slide_description=f"Group {marker} Header",
        group_name=group_name,
        category=excel_data.lst_categories[index],
        name=group_name,
        rationale=excel_data.lst_rationales[index],
        notation=excel_data.lst_notations[index],
        kana=excel_data.lst_kana[index],
        logo_filename=excel_data.lst_logos[index],
        template_id=template_id,
        name_sub_group=excel_data.lst_name_sub_groups[index],
        group_letter=marker,
    )
```

#### `_create_individual_slide()`

```python
def _create_individual_slide(
    self,
    *,
    excel_data: ProcessedExcelData,
    index: int,
    slide_number: int,
    request: CreatePresentationRequest,
    current_group: str,
    default_template_id: int,
) -> DetailItem:
    """
    Crea DetailItem para slide individual o grouped.

    Características:
    - slide_type = "Name"
    - Usa template específico para nombres
    - Hereda group info del grupo actual
    """
    name = excel_data.lst_names[index].strip()

    # Determinar template
    # Si tiene subgrupo o delimitador, usar template "Multi"
    has_delimiter = "##" in name or "$$" in name
    has_subgroup = bool(excel_data.lst_name_sub_groups[index])

    if has_subgroup or has_delimiter:
        template = self._get_template_metadata("Multi")
    else:
        template = self._get_template_metadata("Default")

    template_id = template["template_id"]
    group_letter = excel_data.lst_group_letters[index]

    return DetailItem(
        slide_number=slide_number,
        slide_type="Name",
        slide_bg_file_name="",  # Se aplica en el builder
        slide_description=name,
        group_name=current_group,
        category=excel_data.lst_categories[index],
        name=name,
        rationale=excel_data.lst_rationales[index],
        notation=excel_data.lst_notations[index],
        kana=excel_data.lst_kana[index],
        logo_filename=excel_data.lst_logos[index],
        template_id=template_id,
        name_sub_group=excel_data.lst_name_sub_groups[index],
        group_letter=group_letter,
    )
```

#### `_create_summary_slide()`

```python
def _create_summary_slide(
    self,
    *,
    slide_number: int,
    last_group: str,
    request: CreatePresentationRequest,
    default_template_id: int,
    last_group_letter: str,
) -> DetailItem:
    """
    Crea slide de resumen (solo para NW/DW).

    Características:
    - slide_type = "Summary"
    - Se inserta al final de los slides generados
    """
    summary_template = self._get_template_metadata("Summary")
    template_id = summary_template["template_id"]

    return DetailItem(
        slide_number=slide_number,
        slide_type="Summary",
        slide_bg_file_name="",
        slide_description="Summary Slide",
        group_name=last_group,
        category="",
        name="",
        rationale="",
        notation="",
        kana="",
        logo_filename="",
        template_id=template_id,
        name_sub_group="",
        group_letter=last_group_letter,
    )
```

#### `_get_template_metadata()`

```python
def _get_template_metadata(self, template_name: str) -> Dict[str, Any]:
    """
    Obtiene metadata de un template desde caché o BD.

    Args:
        template_name: 'Default', 'Group', 'Multi', 'Summary', 'Separator'

    Returns:
        Dict con:
            - template_id: int
            - template_name: str
            - file_path: str
    """
    global _template_metadata_cache, _cache_timestamp

    # Check cache freshness (TTL: 5 minutos)
    if _cache_timestamp and (time.time() - _cache_timestamp) > 300:
        _template_metadata_cache.clear()
        _cache_timestamp = None

    # Return from cache if exists
    if template_name in _template_metadata_cache:
        return _template_metadata_cache[template_name]

    # Query database
    with create_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT TemplateId, TemplateName, TemplateFilePath
                FROM [BI_GUIDELINES].[dbo].[nw_Templates]
                WHERE TemplateName = ?
                """,
                (template_name,)
            )
            row = cursor.fetchone()

            if row:
                metadata = {
                    "template_id": row[0],
                    "template_name": row[1],
                    "file_path": row[2],
                }
            else:
                # Fallback a template por defecto
                metadata = {
                    "template_id": 1,
                    "template_name": "Default",
                    "file_path": "template_default_2019.pptx",
                }

    # Cache result
    _template_metadata_cache[template_name] = metadata
    if not _cache_timestamp:
        _cache_timestamp = time.time()

    return metadata
```

---

## 3.3 Generación de PowerPoint Físico

### 3.3.1 Función: `_generate_physical_powerpoint()`

**Archivo:** `app/services/presentation_service.py:1174`

```python
def _generate_physical_powerpoint(
    self,
    *,
    slides_data: Dict[str, Any],
    excel_data: ProcessedExcelData,
    request: CreatePresentationRequest,
    pptx_data: PPTXConversionResponse,
) -> Dict[str, str]:
    """
    Genera archivo PowerPoint físico combinando:
    1. Slides del PPTX original (como imágenes)
    2. Slides generados desde templates (Name, Group, Summary)

    Este método usa COM (win32com) para manipular PowerPoint.

    Args:
        slides_data: Dict con 'details' (List[DetailItem])
        excel_data: Datos del Excel procesado
        request: Request original
        pptx_data: Datos del PPTX convertido

    Returns:
        Dict con:
            - printable_path: str (path al .pptx)
            - macro_path: str (path al .pptm, si aplica)
            - total_slides: str (número total)
            - warnings: str (warnings concatenados)
    """
    # 1. Construir opciones de build
    build_options = PresentationBuildOptions(
        template_pack="BackgroundDefaultTemplate",
        base_template="template_default_2019.pptx",
        multi_template="template_default_withgroups2019.pptx",
        group_template="template_default_withgroup_2019.pptx",
        separator_template="template_default_seperator_2019.pptx",
        summary_template="template_default_summary2019.pptx",
        slide_start=1,
        slide_end=None,
        include_print_ready_version=True,
        include_macro_version=False,
        return_urls=False,
    )

    logger.info(
        "Generating physical PowerPoint: %d original images, %d template slides, insert at position %d",
        len(pptx_data.images) if pptx_data.images else 0,
        len(slides_data.get("details", [])),
        request.page_number,
    )

    # 2. Guardar PPTX original en disco (necesario para COM)
    output_base, _ = resolve_project_output(
        request.display_name,
        request.project_type,
        fallback_subdir="generated_presentations",
    )
    project_folder = output_base
    pptx_original_path = project_folder / request.pptx_filename

    if not pptx_original_path.exists():
        with open(pptx_original_path, "wb") as f:
            f.write(request.pptx_file)

    # 3. Llamar al builder service (usa COM)
    artifacts = pptx_builder_service.compose_presentation_with_original_slides(
        request=request,
        options=build_options,
        details=slides_data["details"],
        excel_data=excel_data,
        original_pptx_path=str(pptx_original_path.resolve()),
        page_number_insert=request.page_number,
    )

    # 4. Convertir resultado
    result = {
        "printable_path": str(artifacts.printable_path) if artifacts.printable_path else "",
        "macro_path": str(artifacts.macro_path) if artifacts.macro_path else "",
        "total_slides": str(artifacts.total_slides),
        "warnings": ", ".join(artifacts.warnings) if artifacts.warnings else "None",
    }

    logger.info(
        "Physical PowerPoint generated successfully: %d total slides, %d warnings",
        artifacts.total_slides,
        len(artifacts.warnings),
    )

    if artifacts.warnings:
        for warning in artifacts.warnings:
            logger.warning("PowerPoint generation warning: %s", warning)

    return result
```

### 3.3.2 Función: `compose_presentation_with_original_slides()`

**Archivo:** `app/services/pptx_builder_service.py`

**NOTA:** Esta función es **muy compleja** y usa extensivamente COM automation. Aquí está la lógica simplificada:

```python
import win32com.client
from pathlib import Path
from datetime import datetime
from typing import List

class PresentationBuildArtifacts:
    """Resultado de compose_presentation_with_original_slides."""
    def __init__(self):
        self.printable_path: Optional[Path] = None
        self.macro_path: Optional[Path] = None
        self.total_slides: int = 0
        self.warnings: List[str] = []


def compose_presentation_with_original_slides(
    *,
    request: CreatePresentationRequest,
    options: PresentationBuildOptions,
    details: List[DetailItem],
    excel_data: ProcessedExcelData,
    original_pptx_path: str,
    page_number_insert: int,
) -> PresentationBuildArtifacts:
    """
    Compone presentación combinando PPTX original + slides generados.

    ALGORITMO:
    1. Abrir PPTX original con PowerPoint COM
    2. Iterar sobre details:
        a. Si slide_type == "Image" → slide ya está en PPTX original
        b. Si slide_type != "Image" → generar desde template y agregar
    3. Insertar slides generados en posición page_number_insert
    4. Guardar como backup_YYYYMMDD_HHMMSS.pptx

    Args:
        request: Request original
        options: Opciones de build
        details: Lista de DetailItem (cada uno representa un slide)
        excel_data: Datos del Excel
        original_pptx_path: Path absoluto al PPTX original
        page_number_insert: Posición donde insertar slides generados

    Returns:
        PresentationBuildArtifacts con paths generados
    """
    artifacts = PresentationBuildArtifacts()

    # 1. Inicializar PowerPoint Application via COM
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = False  # No mostrar UI
    except Exception as e:
        raise Exception(f"Failed to initialize PowerPoint COM: {e}")

    try:
        # 2. Abrir presentación original
        presentation = ppt_app.Presentations.Open(
            FileName=original_pptx_path,
            ReadOnly=False,
            Untitled=True,
            WithWindow=False
        )

        # 3. Crear nueva presentación para slides generados
        generated_pres = ppt_app.Presentations.Add(WithWindow=False)

        # 4. Iterar sobre details y construir slides generados
        for detail in details:
            if detail.slide_type == "Image":
                # Skip: este slide ya está en el PPTX original
                continue

            # Generar slide desde template
            slide = _generate_slide_from_template(
                generated_pres,
                detail,
                excel_data,
                options,
                request
            )

        # 5. Combinar: insertar slides generados en posición page_number_insert
        #    Estrategia:
        #    - Copiar slides 1 a (page_number_insert - 1) del original
        #    - Insertar TODOS los slides generados
        #    - Copiar slides restantes del original

        final_pres = ppt_app.Presentations.Add(WithWindow=False)
        slide_index = 1

        # Copiar prefix slides del original
        for i in range(1, min(page_number_insert, presentation.Slides.Count + 1)):
            slide = presentation.Slides(i)
            slide.Copy()
            final_pres.Slides.Paste(Index=slide_index)
            slide_index += 1

        # Copiar slides generados
        for i in range(1, generated_pres.Slides.Count + 1):
            slide = generated_pres.Slides(i)
            slide.Copy()
            final_pres.Slides.Paste(Index=slide_index)
            slide_index += 1

        # Copiar slides restantes del original
        if page_number_insert <= presentation.Slides.Count:
            for i in range(page_number_insert, presentation.Slides.Count + 1):
                slide = presentation.Slides(i)
                slide.Copy()
                final_pres.Slides.Paste(Index=slide_index)
                slide_index += 1

        # 6. Guardar archivo final
        output_folder = Path(original_pptx_path).parent / "Presentations"
        output_folder.mkdir(exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"backup_{timestamp}.pptx"
        output_path = output_folder / output_filename

        final_pres.SaveAs(
            FileName=str(output_path.resolve()),
            FileFormat=1  # ppSaveAsDefault (PPTX)
        )

        artifacts.printable_path = output_path
        artifacts.total_slides = final_pres.Slides.Count

        # 7. Cleanup
        final_pres.Close()
        generated_pres.Close()
        presentation.Close()

    except Exception as e:
        artifacts.warnings.append(f"Error during composition: {str(e)}")
        raise

    finally:
        # 8. Quit PowerPoint
        ppt_app.Quit()

    return artifacts


def _generate_slide_from_template(
    presentation,
    detail: DetailItem,
    excel_data: ProcessedExcelData,
    options: PresentationBuildOptions,
    request: CreatePresentationRequest
):
    """
    Genera un slide individual desde template.

    LÓGICA:
    1. Cargar template PPTX según detail.slide_type
    2. Copiar primer slide del template
    3. Reemplazar placeholders con valores de detail
    4. Aplicar background según configuración
    5. Agregar a presentación

    Placeholders comunes:
    - {{NAME}}: detail.name
    - {{CATEGORY}}: detail.category
    - {{RATIONALE}}: detail.rationale
    - {{NOTATION}}: detail.notation
    - {{GROUP_NAME}}: detail.group_name
    - {{LOGO}}: detail.logo_filename (insertar imagen)
    """
    # Esta función es extremadamente compleja y específica
    # Involucra manipulación de shapes, text frames, imágenes, etc. via COM
    # Ver implementación completa en pptx_builder_service.py
    pass
```

---

Continúo en el siguiente mensaje debido al límite de tokens...
