# Creative Python API

A comprehensive FastAPI-based API for professional PowerPoint presentation creation, conversion, and management with SQL Server database integration.

## 🚀 Key Features

### 📊 Complete Presentation Orchestration
- **Unified API endpoint**: Single endpoint for complete presentation creation workflow
- **Excel processing**: Advanced Excel file processing with phonetic and grouping support
- **PowerPoint conversion**: Transforms PPTX files to high-quality PNG images
- **Slide generation**: Automatic slide creation from Excel data with template rotation
- **Database integration**: Direct SQL Server integration with stored procedures
- **Multi-project support**: Separate management for "bipresents" and "nw" projects

### 📈 Excel Processing & Analysis
- **Advanced slide generation**: Processes Excel files to create presentation slides
- **Multiple slide types**: Supports group slides (A-Z) and individual name evaluation slides
- **Phonetic support**: Japanese kana and notation processing
- **Grouping logic**: Intelligent grouping of data for optimal presentation structure
- **Validation**: Comprehensive input validation and error handling

### 🎨 PowerPoint Processing
- **Complete conversion**: Transforms PPTX files to high-quality PNG images
- **Title extraction**: Automatically extracts titles from each slide
- **Thumbnail generation**: Creates optimized thumbnails for preview
- **Template rotation**: Dynamic template assignment for slides
- **Background management**: Configurable background types and names

### 🗄️ Database Integration
- **SQL Server connection**: Secure integration with BI_GUIDELINES database via environment variables
- **Presentation management**: Full CRUD for master and detail records
- **Stored procedures**: Direct integration with `nw_InsertPresentationMaster_sep2025` and `nw_InsertPresentationDetail_copy`
- **Transactions**: Transactional operations for data integrity
- **Direct database calls**: Optimized performance without HTTP overhead

### 🌐 Modern RESTful API
- **Automatic documentation**: Multiple documentation interfaces available
  - **Swagger UI** at `/docs` (traditional OpenAPI interface)
  - **Scalar UI** at `/scalar` (modern, enhanced documentation)
- **CORS configured**: Environment-based CORS configuration for front-end applications
- **Advanced logging**: Colored logs system with configurable levels
- **Health checks**: Monitoring and status endpoints
- **Environment-based configuration**: Flexible configuration via `.env` files

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

### 3. Configure environment variables
```bash
# Copy the example environment file
cp .env.example .env

# Edit .env with your actual configuration values
# IMPORTANT: Update database credentials and other sensitive settings
```

### 4. Install dependencies
```bash
pip install -r requirements.txt
```

### 5. Run the application
```bash
# Development mode
python run.py

# Or use the batch files
run_dev.bat    # Development
run_prod.bat   # Production
```

## ⚙️ Configuration

### 🌍 Environment-Based Configuration

The application uses environment variables for secure and flexible configuration. The configuration is loaded from `.env` files automatically.

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

# CORS Configuration
CORS_ORIGINS=*                   # Comma-separated origins or "*" for all
CORS_METHODS=*                   # Comma-separated methods or "*" for all
CORS_HEADERS=*                   # Comma-separated headers or "*" for all

# Database Configuration (Security: Never commit real credentials!)
SQL_CONNECTION_STRING=DRIVER={SQL Server};SERVER=your_server;DATABASE=your_database;UID=your_username;PWD=your_password;TrustServerCertificate=yes;Connection Timeout=30;Encrypt=no;

# Logging Configuration
LOG_LEVEL=INFO
```

> **🔒 Security Note**: Database credentials and sensitive configuration are loaded from the `.env` file, which is ignored by Git. Copy `.env.example` to `.env` and update with your actual credentials. Never commit the `.env` file to version control!

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
└── output_images/           # Generated presentation images
```

**Production: at Matrix server**
```
C:/inetpub/wwwroot/
├── bipresents/bsr_slides/    # Bipresents projects
├── nw2/nw_slides/           # NW projects
└── CreativePythonAPI/       # Application files
```

## 📡 API Endpoints

### 🎯 Presentation Creation (Main Feature)
- `POST /api/presentation/create` - **Complete presentation orchestration**
  - Processes Excel file, converts PPTX, generates slides, and saves to database
  - Unified endpoint for the entire presentation creation workflow

### 📊 Excel Processing
- `POST /api/excel-processing/process-excel` - Process Excel files for slide generation
  - Supports phonetic processing and grouping logic
  - Returns structured data for presentation slides

### 🎨 PowerPoint Processing
- `POST /api/pptx-conversion/convert-pptx` - Convert PPTX files to images
  - Extracts slides as high-quality PNG images
  - Supports multiple projects and background configurations

### 🗄️ Database Operations
- `POST /api/bi_guidelines/presentations/` - Create presentation records
  - Direct database integration with stored procedures
  - Transactional operations for data integrity
- `GET /api/bi_guidelines/presentations/{id}` - Retrieve presentation details

### 🔍 System Endpoints
- `GET /` - API root with basic information
- `GET /health` - Health check endpoint
- `GET /docs` - Swagger UI documentation
- `GET /scalar` - Modern Scalar documentation

## 🚀 Usage Examples

### 🎯 Complete Presentation Creation
Use the main endpoint to create a complete presentation in one API call:

```bash
POST /api/presentation/create
Content-Type: multipart/form-data

# Form data:
- excel_file: [Excel file with candidate data]
- pptx_file: [PowerPoint template file]
- project: "NW_PROJECT_2025"
- display_name: "Name Evaluation Presentation"
- background_type: "standard"
- background_name: "blue_gradient"
- page_number: 1
- presentation_type: "NameEvaluation"
- user_name: "analyst"
- bsr_display_name: "Brand Study Results"
- participant_vote: 1
- is_wide_ppt: 1
- is_aws_email: 0
- project_type: "nw"
- has_groups: true
- is_phonetics: true
```

**Response:**
```json
{
  "message": "Presentation created successfully",
  "presentation_id": 12345,
  "total_slides": 25,
  "excel_data": {...},
  "pptx_data": {...},
  "processing_time_seconds": 45.67
}
```

### 📊 Individual Component Usage
You can also use individual endpoints for specific operations:

- **Excel Processing Only**: `POST /api/excel-processing/process-excel`
- **PPTX Conversion Only**: `POST /api/pptx-conversion/convert-pptx`
- **Database Operations**: `POST /api/bi_guidelines/presentations/`

## � Security & Performance

### 🛡️ Security Features
- **Environment-based configuration**: Sensitive data loaded from `.env` files
- **Database credentials**: Never hardcoded in source code
- **CORS configuration**: Environment-controlled cross-origin settings
- **Git security**: `.env` files automatically ignored by Git

### ⚡ Performance Optimizations
- **Direct database integration**: No HTTP overhead for database operations
- **Transactional operations**: Atomic database operations with rollback support
- **Efficient file processing**: Optimized Excel and PowerPoint processing
- **Connection pooling**: Database connection management for high performance

### 📊 Monitoring & Logging
- **Colored logging**: Enhanced console output with color coding
- **Configurable log levels**: Environment-based logging configuration
- **Health checks**: System monitoring endpoints
- **Performance metrics**: Processing time tracking for operations

## 🛠️ Development

### 🧪 Testing the API
1. Start the server: `python run.py`
2. Open documentation: `http://localhost:50100/docs`
3. Use the interactive Swagger UI to test endpoints
4. Check logs in the console for detailed operation information

### 🔧 Code Quality
- **Pylint integration**: Code quality checking with 9.8/10 rating
- **Type hints**: Full type annotation support
- **Pydantic models**: Data validation and serialization
- **Modular architecture**: Clean separation of concerns

# Method 2: Manual
set ENVIRONMENT=production
python run.py

# Method 3: Using uvicorn directly
uvicorn app.main:app --host 0.0.0.0 --port 50100
```

## 📋 API Endpoints

### 🎯 Unified Presentation Creation (`/api/presentation/`)

#### POST `/api/presentation/create/`
Creates a complete presentation from Excel and PPTX files with database integration.

**Parameters:**
- `excel_file`: Excel file with presentation data (multipart/form-data)
- `pptx_file`: PowerPoint template file (multipart/form-data)
- `displayName`: Project/folder name
- `projectType`: Project type (`"bipresents"` or `"nw"`)

**Response:**
```json
{
  "message": "Presentation created successfully",
  "presentation_id": "12345",
  "conversion_id": "PROJECT_EXAMPLE",
  "project_type": "nw",
  "total_slides": 15,
  "slides_created": 15,
  "database_status": "inserted",
  "processing_time": "45.67 seconds"
}
```

**Features:**
- **Complete orchestration**: Handles file upload, processing, slide generation, and database insertion
- **Transactional safety**: All operations are atomic with automatic rollback on failure
- **Progress tracking**: Real-time processing status and timing information
- **Error handling**: Comprehensive error reporting with specific failure reasons

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

### 📊 Excel Processing (`/api/excel/`)

#### POST `/api/excel/process-excel/`
Processes Excel files and generates presentation slides.

**Parameters:**
- `excel_file`: Excel file (multipart/form-data)
- `project_type`: Project type (`"bipresents"` or `"nw"`)
- `display_name`: Display name for the project

**Response:**
```json
{
  "message": "Excel processed successfully",
  "slides": [
    {
      "slide_number": 1,
      "slide_type": "title",
      "title": "Main Title",
      "content": "..."
    }
  ],
  "total_slides": 10
}
```

### ️ BI Guidelines (`/api/bi_guidelines/`)

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

# CORS Configuration
CORS_ORIGINS=*
CORS_METHODS=*
CORS_HEADERS=*

# Database Configuration (Security: Never commit real credentials!)
SQL_CONNECTION_STRING=DRIVER={SQL Server};SERVER=your_server;DATABASE=your_database;UID=your_username;PWD=your_password;TrustServerCertificate=yes;Connection Timeout=30;Encrypt=no;

# Logging
LOG_LEVEL=INFO
```

> **🔒 Security Note**: Database credentials and sensitive configuration are loaded from the `.env` file, which is ignored by Git. Copy `.env.example` to `.env` and update with your actual credentials. Never commit the `.env` file to version control!

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
│   ├── excel_processing.py # Excel processing endpoints
│   └── bi_guidelines.py   # Database integration
├── models/
│   ├── response_models.py # Response models
│   ├── nw_master_request.py # Database models
│   └── report_models.py   # Excel processing models
├── services/
│   ├── pptx_service.py    # PPTX conversion logic
│   ├── excel_service.py   # Excel processing logic
│   └── chart_service.py   # Chart generation utilities
├── test/                   # Test files directory
│   ├── test.py            # General test script
│   ├── test_excel_direct.py # Direct Excel processing tests
│   ├── test_excel_utils.py # Excel utilities tests
│   └── test_excel_venv.bat # Test environment setup
└── utils/
    ├── logging_utils.py   # Logging utilities
    └── files_utils.py     # File utilities
```

## 🚀 Recent Improvements & Features

### ✨ New Features (v2.0.0)
- **🎯 Unified Presentation API**: Single endpoint (`/api/presentation/create`) that orchestrates the complete presentation creation workflow
- **🔄 Direct Database Integration**: Replaced HTTP calls with direct stored procedure execution for improved reliability and performance
- **⚡ Transactional Operations**: Atomic database operations with automatic rollback on failure
- **📊 Enhanced Monitoring**: Colored logging, processing time tracking, and detailed operation metrics
- **🛡️ Security Hardening**: Environment-based configuration with sensitive data moved to `.env` files

### 🔒 Security Enhancements
- **Environment Variables**: All sensitive configuration (database credentials, API keys) moved to `.env` files
- **Git Security**: `.env` files automatically excluded from version control
- **CORS Configuration**: Environment-controlled cross-origin resource sharing settings
- **Input Validation**: Enhanced Pydantic models for comprehensive data validation

### ⚡ Performance Optimizations
- **Direct DB Calls**: Eliminated HTTP overhead by calling stored procedures directly
- **Connection Pooling**: Efficient database connection management
- **Optimized Processing**: Streamlined file processing and slide generation workflows
- **Memory Management**: Improved resource utilization for large file processing

### 🏗️ Architecture Improvements
- **Modular Services**: Clean separation between presentation orchestration, file processing, and database operations
- **Error Handling**: Comprehensive error reporting with specific failure reasons and recovery suggestions
- **Code Quality**: Pylint 9.8/10 rating with full type hints and modern Python practices
- **Documentation**: Updated API documentation with interactive examples and comprehensive endpoint descriptions

### 🔄 Migration Notes
- **API Changes**: The new unified `/api/presentation/create` endpoint replaces separate file processing and database insertion calls
- **Configuration**: All hardcoded values moved to environment variables - update your `.env` files accordingly
- **Dependencies**: Updated to latest stable versions of FastAPI, Pydantic, and database drivers

## 🧪 Testing

### Running Tests
```bash
# Run general tests
python -m app.test.test

# Run Excel processing tests
python -m app.test.test_excel_direct

# Run Excel utilities tests
python -m app.test.test_excel_utils

# Setup test environment (Windows)
app\test\test_excel_venv.bat
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

## 🐛 Troubleshooting

### Common Issues

#### Database Connection Errors
```
Error: Unable to connect to database
```
**Solution:**
- Verify `SQL_CONNECTION_STRING` in `.env` file
- Check database server availability
- Ensure SQL Server drivers are installed
- Confirm firewall settings allow database connections

#### File Upload Errors
```
Error: File processing failed
```
**Solution:**
- Check file permissions on `NW_Files/` directory
- Verify supported file formats (Excel: .xlsx, .xls; PowerPoint: .pptx)
- Ensure sufficient disk space for file processing
- Check antivirus software isn't blocking file operations

#### CORS Issues
```
Error: CORS policy blocked
```
**Solution:**
- Update `CORS_ORIGINS` in `.env` file to include your frontend domain
- Restart the application after configuration changes
- Check browser console for specific CORS error details

#### Performance Issues
```
Error: Processing timeout
```
**Solution:**
- Increase `Connection Timeout` in database connection string
- Check server resources (CPU, memory) during processing
- Optimize file sizes before upload
- Monitor logs for bottleneck identification

### Debug Mode
Enable debug logging by setting:
```bash
LOG_LEVEL=DEBUG
DEBUG=true
```

### Health Checks
- **GET `/health`**: Basic service availability check
- **GET `/`**: API information and version details
- **Logs**: Check console output for detailed error information

## 📞 Support

### Getting Help
1. **Check Logs**: Enable DEBUG logging and review console output
2. **Health Endpoints**: Use `/health` to verify service status
3. **Documentation**: Refer to `/scalar` for interactive API documentation
4. **Configuration**: Verify all environment variables are properly set

### Version Information
- **Current Version**: 2.0.0
- **Python Version**: 3.13+
- **FastAPI Version**: 0.115.12
- **Database**: SQL Server with pyodbc

---

**Last Updated**: December 2024
**Maintained by**: Development Team

