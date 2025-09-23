# Report Generator API

A FastAPI-based service for generating professional Word reports with dynamic content including text, tables, charts, and images.

## Features

- Generate Word documents with structured content
- Support for various components:
  - Text (titles, headings, paragraphs)
  - Tables with advanced formatting
  - Charts (pie, bar, etc.)
  - Images (local or generated)
  - Table of Contents
  - Page breaks
- PPTX to image conversion with title extraction
- Customizable headers and footers
- Consistent styling and formatting

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/report-generator-api.git
   cd report-generator-api
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. To run the service on the CLI
   ```bash
    Activate Environment: venv\Scripts\activate
    Run the app: python run.py
    Run the app in the BG pythonw run.py (NOT WORKING PROPERLY YET)
   ```
## Requirements

- Python 3.7+
- Microsoft PowerPoint (for PPTX conversion)
- Required Python packages:
  - fastapi
  - python-docx
  - matplotlib
  - uvicorn
  - python-pptx
  - pywin32 (for Windows COM integration)

## API Endpoints

### Report Generation

- `POST /generate-report/` - Generate a Word report
- `GET /download-report/{report_id}` - Download generated report
- `DELETE /delete-report/{report_id}` - Delete generated report

### PPTX Conversion

- `POST /convert-pptx/` - Convert PPTX to images and extract titles
- `GET /download/{conversion_id}/{image_name}}` - Download specific image
- `GET /download-all/{conversion_id}` - Download all images as ZIP
- `DELETE /delete/{conversion_id}` - Delete conversion files

## Usage

### Generating Reports

Send a POST request to `/generate-report/` with a JSON body following the `ReportContent` schema. Example:

```json
{
  "title": "Sample Report",
  "company": "ACME Corp",
  "components": [
    {
      "type": "title",
      "content": "Report Title",
      "align": "center"
    },
    {
      "type": "paragraph",
      "content": "This is a sample paragraph."
    },
    {
      "type": "chart",
      "chart_type": "pie",
      "data": {
        "labels": ["A", "B", "C"],
        "values": [30, 40, 30]
      },
      "title": "Sample Chart"
    }
  ]
}
```

### PPTX Conversion

Upload a PPTX file to `/convert-pptx/` to convert it to images. The API will return a list of image URLs and slide titles.

## Configuration

The API creates the following directories automatically:

- `NW_Files` in Documents folder (for output)
  - `PowerPoint_files` (for PPTX processing)
  - `output_images` (for generated content)

## Running the Server

```bash
uvicorn main:app --host 0.0.0.0 --port 8000
```

For production, consider using:
- Gunicorn with Uvicorn workers
- Proper HTTPS configuration
- Authentication

## Development Notes

- The API uses Windows COM for PPTX conversion (requires Microsoft PowerPoint)
- For non-Windows environments, PPTX conversion will not work
- Generated files are stored temporarily and should be cleaned up after download