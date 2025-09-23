# Creative Python API

A FastAPI-based API for professional PowerPoint presentation conversion to images and presentation data management with SQL Server database integration.

## 🚀 Key Features

### 📊 PowerPoint Conversion
- **Complete conversion**: Transforms PPTX files to high-quality PNG images
- **Title extraction**: Automatically extracts titles from each slide
- **Thumbnail generation**: Creates optimized thumbnails for preview
- **Multi-project support**: Separate management for "bipresents" and "nw" projects
- **Public URLs**: Generates direct links for web access

### 🗄️ Database Integration
- **SQL Server connection**: Complete integration with BI_GUIDELINES database
- **Presentation management**: Full CRUD for master and detail records
- **Stored procedures**: Integration with existing stored procedures
- **Transactions**: Transactional operations for data integrity

### 📁 File Management
- **Individual downloads**: Access to specific files by project
- **Bulk downloads**: ZIP file generation for complete downloads
- **Path management**: Intelligent path system by project type
- **Security validation**: Prevention of unauthorized path access

### 🌐 Complete RESTful API
- **Automatic documentation**: Multiple documentation interfaces available
  - **Swagger UI** at `/docs` (traditional OpenAPI interface)
  - **Scalar UI** at `/scalar` (modern, enhanced documentation)
- **CORS configured**: Support for front-end web applications
- **Advanced logging**: Colored logs system with configurable levels
- **Health checks**: Monitoring and status endpoints

## 📦 Installation

### 1. Clone the repository
```bash
git clone http://harley:8080/tfs/DefaultCollection/_git/2025-CREATIVE-PYTHON-API
cd 2025-CREATIVE-PYTHON-API
```

### 2. Create virtual environment
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Linux/Mac
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure database
Conection string to access SQL Server configured in `app/config/settings.py`:
- **Server**: 192.168.0.85
- **Database**: BI_GUIDELINES

## 🚀 Running the Application

### 🛠️ Environment Configuration

The application supports two environments with different directory configurations:

#### Development Environment (Default)
- Uses local `NW_Files/` directory structure
- Safe for local development and testing
- Configured via `.env` file

#### Production Environment at Matrix server
- Uses `C:/inetpub/wwwroot/` directory structure
- Requires appropriate system permissions
- Configured via `.env.production` or environment variables

### 📁 Directory Structure by Environment

**Development: at Local Environment**
```
NW_Files/
├── bipresents/bsr_slides/    # Bipresents projects
├── nw2/nw_slides/           # NW projects
├── PowerPoint_files/        # Original PPTX files
└── output_images/           # Generated images
```

**Production: at Matrix server**
```
C:/inetpub/wwwroot/
├── bipresents/bsr_slides/   # Bipresents projects
├── nw2/nw_slides/          # NW projects
└── CreativePythonAPI/NW_Files/
    ├── PowerPoint_files/    # Original PPTX files
    └── output_images/       # Generated images
```

### 🚀 Starting the Application

#### Development Mode
```bash
# Method 1: Using batch script
run_dev.bat

# Method 2: Manual
set ENVIRONMENT=development
python run.py

# Method 3: Using .env file (default)
python run.py
```

#### Production Mode
```bash
# Method 1: Using batch script
run_prod.bat

# Method 2: Manual
set ENVIRONMENT=production
python run.py

# Method 3: Using uvicorn directly
uvicorn app.main:app --host 0.0.0.0 --port 50100
```

## 📋 API Endpoints

### 🎯 PowerPoint Conversion (`/api/pptx/`)

#### POST `/api/pptx/convert/`
Converts a PPTX file to images and extracts titles.

**Parameters:**
- `pptx_file`: PPTX file (multipart/form-data)
- `displayName`: Project/folder name
- `projectType`: Project type (`"bipresents"` or `"nw"`)

**Response:**
```json
{
  "message": "Conversion completed successfully",
  "conversion_id": "PROJECT_EXAMPLE",
  "project_type": "nw",
  "total_images": 15,
  "images": ["http://bipresents.com/nw2/nw_slides/PROJECT_EXAMPLE/slide_1.png"],
  "thumbnails": ["http://bipresents.com/nw2/nw_slides/PROJECT_EXAMPLE/Thumbnails/slide_1.png"],
  "titles": ["Slide Title 1", "Slide Title 2"],
  "pptx_file": "presentation.pptx"
}
```

#### GET `/api/pptx/status/{project_type}/{conversion_id}`
Checks the status of a specific conversion.

#### GET `/api/pptx/list/`
Lists all available conversions.
- **Query param**: `project_type` (optional)

### 📁 Files Management (`/api/files/`)

#### GET `/api/files/download/{project_type}/{conversion_id}/{filename:path}`
Downloads a specific file from the project.

#### GET `/api/files/download-all/{project_type}/{conversion_id}`
Downloads all files from a project as ZIP.

#### DELETE `/api/files/delete/{project_type}/{conversion_id}`
Deletes all files from a conversion.

### 🗃️ BI Guidelines (`/api/bi_guidelines/`)

#### GET `/api/bi_guidelines/nw-master`
Gets the first 1000 records from the `nw_Master` table.

#### POST `/api/bi_guidelines/presentations/`
Creates a new master presentation record with details.

**Request body:**
```json
{
  "project": "PROJECT_EXAMPLE",
  "display_name": "Demo Presentation",
  "powerpoint_file": "demo.pptx",
  "excel_file": "data.xlsx",
  "background_type": "standard",
  "background_name": "theme1",
  "page_number": 1,
  "presentation_type": "research",
  "user_name": "username",
  "bsr_display_name": "Demo BSR",
  "participant_vote": 1,
  "is_wide_ppt": 1,
  "is_aws_email": 0,
  "details": [
    {
      "slide_number": 1,
      "slide_type": "title",
      "slide_description": "Main title",
      "template_id": 1
    }
  ]
}
```

### 🔍 System Endpoints

#### GET `/`
Basic API information.

#### GET `/health`
Service health check.

#### GET `/docs`
Interactive Swagger UI documentation (traditional OpenAPI interface).

#### GET `/scalar`
Modern Scalar documentation interface with enhanced UI and better user experience.

## ⚙️ Configuration

### 🌍 Environment-Based Configuration

The application uses environment variables for configuration. You can set these via:

1. **`.env` file** (for development)
2. **`.env.production` file** (for production)
3. **System environment variables**
4. **Batch scripts** (`run_dev.bat` / `run_prod.bat`)

### 📝 Available Environment Variables

```bash
# Environment Mode
ENVIRONMENT=development          # "development" or "production"

# Server Configuration
HOST=0.0.0.0
PORT=50100

# Application Configuration
APP_NAME=Report Generator API
APP_VERSION=2.0.0
DEBUG=false

# Database Configuration
SQL_CONNECTION_STRING=DRIVER={SQL Server};SERVER=192.168.0.85;...

# Logging
LOG_LEVEL=INFO
```

### 🔧 Environment-Specific Paths

The application automatically configures paths based on the `ENVIRONMENT` variable:

**Development Mode** (`ENVIRONMENT=development`):
```python
base_dir_bipresents = "NW_Files/bipresents/bsr_slides"
base_dir_nw = "NW_Files/nw2/nw_slides"
nw_files_dir = "NW_Files"
```

**Production Mode** (`ENVIRONMENT=production`):
```python
base_dir_bipresents = "C:/inetpub/wwwroot/bipresents/bsr_slides"
base_dir_nw = "C:/inetpub/wwwroot/nw2/nw_slides"
nw_files_dir = "C:/inetpub/wwwroot/CreativePythonAPI/NW_Files"
```

### Directory Structure
```
C:/inetpub/wwwroot/
├── bipresents/bsr_slides/          # Bipresents projects
├── nw2/nw_slides/                  # NW projects
└── CreativePythonAPI/
    └── NW_Files/
        ├── PowerPoint_files/       # Original PPTX files
        └── output_images/          # Generated images
```

## 🏗️ Project Architecture

```
app/
├── main.py                 # Main FastAPI application
├── config/
│   └── settings.py        # Centralized configuration
├── api/routes/
│   ├── pptx_conversion.py # PPTX conversion endpoints
│   ├── files.py           # File management
│   └── bi_guidelines.py   # Database integration
├── models/
│   ├── response_models.py # Response models
│   └── nw_master_request.py # Database models
├── services/
│   └── pptx_service.py    # PPTX conversion logic
└── utils/
    ├── logging_utils.py   # Logging utilities
    └── files_utils.py     # File utilities
```

### Local Development
```bash
python run.py
```
- **Server**: http://localhost:50100
- **Swagger UI Documentation**: http://localhost:50100/docs
- **Scalar Documentation**: http://localhost:50100/scalar (recommended)

### IIS Production
1. Configure IIS with FastCGI
2. Install Python and dependencies
3. Configure virtual directory `/CreativePythonAPI`
4. Set permissions for working folders

