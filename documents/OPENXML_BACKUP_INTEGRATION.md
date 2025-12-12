# OpenXML Backup Generation Integration

## Overview

A new OpenXML-based backup generation implementation has been integrated into the CreativePythonAPI as an alternative to the existing COM-based PowerPoint generation. This provides a faster, more versatile, and more reliable way to generate backup presentations.

## Key Features

### Advantages of OpenXML Implementation

1. **Faster Performance**: Direct XML manipulation is significantly faster than COM automation
2. **No PowerPoint Dependency**: Works without PowerPoint installed on the server
3. **Better Concurrency**: No COM threading limitations (max 2 concurrent operations)
4. **Cross-Platform**: Can run on Linux/Mac (in theory, though current system is Windows-based)
5. **More Reliable**: No COM initialization issues or zombie processes

### Current Implementation Status

- ✅ Feature flag added to settings
- ✅ OpenXML service module created ([openxml_pptx_service.py](../app/services/openxml_pptx_service.py))
- ✅ Integrated into presentation_service
- ✅ Dynamic slide insertion position (from database `NameCandidateStartingSlide`)
- ✅ Big Japanese mode support (based on `presentation_type == "Phonetics"`)
- ✅ Excel data conversion from existing `ProcessedExcelData` format
- ✅ PowerPoint template compatibility maintained

## How to Enable

### Using Environment Variables

Add to your `.env` file:

```env
USE_OPENXML_BACKUP=true
```

### Using Code/Settings

Modify [app/config/settings.py](../app/config/settings.py):

```python
# Backup generation settings
use_openxml_backup: bool = True  # Feature flag: True = OpenXML (faster), False = COM (legacy)
```

## Architecture

### File Structure

```
app/
├── services/
│   ├── openxml_pptx_service.py        # NEW: OpenXML implementation
│   ├── presentation_service.py         # MODIFIED: Feature flag integration
│   └── pptx_builder_service.py        # EXISTING: Legacy COM implementation
├── config/
│   └── settings.py                     # MODIFIED: Feature flag added
└── models/
    └── presentation_models.py          # EXISTING: Shared data models
```

### Data Flow

```
┌─────────────────────────────────────────────────────────────┐
│ POST /api/presentations/generate-backup                    │
│ Input: {presentation_id: int}                               │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│ presentation_service.generate_backup_presentation()         │
│ - Query database for metadata                               │
│ - Load Excel and PPTX files                                 │
│ - Process Excel data → ProcessedExcelData                   │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│ _generate_physical_powerpoint()                             │
│ Check: settings.use_openxml_backup?                         │
└────────────┬──────────────────────┬─────────────────────────┘
             │                      │
   ┌─────────┴──────────┐  ┌────────┴──────────────┐
   │ if False (Legacy)  │  │ if True (New)         │
   ▼                    │  ▼                       │
┌──────────────────┐    │  ┌──────────────────────┐│
│ COM-based        │    │  │ OpenXML-based        ││
│ pptx_builder_    │    │  │ openxml_pptx_        ││
│ service          │    │  │ service              ││
│                  │    │  │                      ││
│ - Uses win32com │    │  │ - Uses python-pptx   ││
│ - Templates      │    │  │ - Direct XML         ││
│ - Slow           │    │  │ - Fast               ││
└──────────────────┘    │  └──────────────────────┘│
                        │                          │
                        └──────────┬───────────────┘
                                   │
                                   ▼
                    ┌────────────────────────────┐
                    │ Save to Presentations/     │
                    │ backup_YYYYMMDD_HHMMSS.pptx│
                    └────────────────────────────┘
```

## Key Mappings

### 1. Slide Insertion Position

**Your Code (Hardcoded)**:
```python
TARGET_POSITION = 5  # zero-based index for the 6th slide
```

**Integrated Code (Dynamic)**:
```python
# From database column: NameCandidateStartingSlide (1-based)
insert_position = max(0, request.page_number - 1)  # Convert to 0-based
```

The slide insertion position now comes from the database, not hardcoded to 5.

### 2. Big Japanese Mode

**Your Code**:
```python
is_big_japanese: bool = True  # Parameter to /process endpoint
```

**Integrated Code**:
```python
# Derived from presentation_type in database
is_big_japanese = request.presentation_type.lower() == "phonetics"
```

Presentation types in database:
- `"Phonetics"` → `is_big_japanese = True` (emphasize katakana)
- `"Normal"` → `is_big_japanese = False` (standard mode)
- `"BSR-Japan"`, `"NSR-Japan"` → Also False (different Japanese handling)

### 3. Excel Data Format

**Your Code Expected**:
```python
List[List[Any]] with headers:
["SEQ", "Category", "Name", "Rationale", "Pronunciation", "Katakana", "Group1", "Group2"]
```

**System Provides**:
```python
ProcessedExcelData {
    lst_types: List[str]           # SEQ column
    lst_categories: List[str]      # Category
    lst_names: List[str]           # Name
    lst_rationales: List[str]      # Rationale
    lst_notations: List[str]       # Pronunciation
    lst_kana: List[str]            # Katakana
    lst_name_sub_groups: List[str] # Group1
}
```

**Conversion** happens in `openxml_pptx_service._convert_processed_excel_to_rows()`.

### 4. PowerPoint Template

**Your Code**:
```python
pptx_file: UploadFile  # User uploads PPTX
```

**Integrated Code**:
```python
request.pptx_file: bytes  # Original PPTX template from database storage
# Stored at: C:\inetpub\wwwroot\nw_slides\{display_name}\template.pptx
```

The original uploaded PPTX template is used as the base for background copying.

## Implementation Details

### OpenXML Service Methods

#### Main Entry Point
```python
openxml_pptx_service.generate_backup_presentation(
    pptx_bytes: bytes,           # Original PowerPoint template
    excel_data: ProcessedExcelData,  # Processed Excel data
    insert_position: int = 5,    # Zero-based insert position
    is_big_japanese: bool = False,  # Phonetics mode flag
) -> bytes  # Returns generated PPTX as bytes
```

#### Slide Rendering Methods

1. **Category Slides**: `_render_category_slide()`
   - Triggered when SEQ column contains letters (A, B, C, etc.)
   - Centers category text
   - Supports `$` separator for two-line layout

2. **Single Name Slides**: `_render_name_rationale_slide()`
   - One candidate per slide
   - Shows: Category, Name (or Katakana if big_japanese), Pronunciation, Rationale
   - Includes placeholders for New Names, Comments, and Sentiment

3. **Group Slides**: `_render_group_slide()`
   - Multiple candidates in checkbox grid layout
   - Automatically scales: 3, 4, or 5 columns based on count
   - Groups identified by Group1/Group2 columns

### Background Copying

The OpenXML implementation preserves slide backgrounds from the original template:

```python
def _copy_background(source_slide, target_slide):
    """Copy background XML from source to target slide."""
    # Extracts p:bg XML element and deep copies to new slide
```

This ensures generated slides match the visual style of the original template.

## Testing

### Manual Testing Steps

1. **Enable the feature flag**:
   ```bash
   # In .env file
   USE_OPENXML_BACKUP=true
   ```

2. **Restart the API server**:
   ```bash
   # Stop current server
   # Start server
   uvicorn app.main:app --reload --host 0.0.0.0 --port 50100
   ```

3. **Trigger backup generation**:
   ```bash
   POST http://localhost:50100/api/presentations/generate-backup
   {
     "presentation_id": 123
   }
   ```

4. **Check logs for confirmation**:
   ```
   INFO - Using OpenXML backup generation (feature flag enabled)
   INFO - Presentation type: Phonetics, is_big_japanese: True
   INFO - Converting page_number 6 to zero-based insert position 5
   INFO - Inserted 25 slides using OpenXML
   INFO - OpenXML backup saved to: C:\inetpub\wwwroot\nw_slides\...\Presentations\backup_20251212_143522.pptx
   ```

### Comparison Testing

To compare COM vs OpenXML implementations:

1. Generate backup with `use_openxml_backup = False` (COM)
2. Generate backup with `use_openxml_backup = True` (OpenXML)
3. Compare:
   - File size
   - Number of slides
   - Visual appearance
   - Generation time
   - Any errors/warnings

## Limitations & Differences

### OpenXML Implementation

**Supported**:
- ✅ Category slides
- ✅ Single name/rationale slides
- ✅ Group slides (multi-candidate checkbox grids)
- ✅ Background preservation
- ✅ Phonetics/Big Japanese mode
- ✅ Dynamic slide insertion position
- ✅ Text wrapping and font scaling

**Not Supported** (compared to COM version):
- ❌ Template placeholders (e.g., "Name Candidate 1", "Name Candidate 2")
- ❌ Macro-enabled presentations (.pptm files)
- ❌ Logo image insertion from database
- ❌ Summary slides
- ❌ Advanced template types (BSR, NSR specific layouts)

### When to Use Each Implementation

**Use OpenXML** (`use_openxml_backup = True`) when:
- Performance is critical
- You have high concurrency requirements
- You want simpler, cleaner slide layouts
- Templates are not heavily customized

**Use COM** (`use_openxml_backup = False`) when:
- You need template placeholder replacement
- You require macro-enabled presentations
- You need logo insertion from database
- You need BSR/NSR specific layouts
- You need summary slides
- Maximum compatibility with existing workflows is required

## Future Enhancements

### Potential Improvements

1. **Template Placeholder Support**
   - Parse placeholder shapes from template PPTX
   - Replace text in named placeholders (like COM version)
   - Support "Name Candidate 1-18" dynamic placeholders

2. **Logo Insertion**
   - Query logo filenames from Excel data
   - Insert logo images at specific positions
   - Support logo resizing/positioning

3. **Summary Slides**
   - Generate final summary slide
   - List all selected candidates
   - Include voting statistics

4. **Advanced Template Types**
   - BSR-specific layouts
   - NSR-specific layouts
   - Custom template packs

5. **Performance Optimization**
   - Parallel slide generation
   - Cached template parsing
   - Async file I/O

## Troubleshooting

### Common Issues

**Issue**: "Missing expected column: seq"
- **Cause**: Excel data doesn't have required headers
- **Fix**: Ensure Excel has headers: SEQ, Category, Name, Rationale

**Issue**: "PowerPoint has no slides to copy a background from"
- **Cause**: Template PPTX file is empty or corrupted
- **Fix**: Check that template.pptx exists and has at least 1 slide

**Issue**: Feature flag not working
- **Cause**: `.env` file not loaded or server not restarted
- **Fix**: Restart server, verify `.env` is in correct location

**Issue**: Slides inserted at wrong position
- **Cause**: `NameCandidateStartingSlide` in database is incorrect
- **Fix**: Update database value (1-based index, e.g., 6 for 6th slide)

## Configuration Reference

### Settings.py

```python
# Backup generation settings
use_openxml_backup: bool = False  # Feature flag
```

**Environment Variable**: `USE_OPENXML_BACKUP`
- Type: Boolean
- Default: `False` (uses COM)
- Values: `true`, `false`, `1`, `0`

### Database Fields Used

| Table | Column | Purpose |
|-------|--------|---------|
| nw_Master | PresentationId | Primary key for presentation |
| nw_Master | DisplayName | Project folder name |
| nw_Master | PresentationType | "Phonetics" → is_big_japanese |
| nw_Master | NameCandidateStartingSlide | Insert position (1-based) |
| nw_Master | NameCandidateFileName | Excel filename to load |
| nw_Master | MainPptFileName | Template PPTX filename |

## Code References

### Key Files

- [app/services/openxml_pptx_service.py](../app/services/openxml_pptx_service.py) - OpenXML implementation
- [app/services/presentation_service.py](../app/services/presentation_service.py#L1287-L1384) - Integration point
- [app/config/settings.py](../app/config/settings.py#L190-L191) - Feature flag

### Key Methods

- `openxml_pptx_service.generate_backup_presentation()` - Main entry point
- `presentation_service._generate_physical_powerpoint_openxml()` - Integration wrapper
- `openxml_pptx_service._build_slides_from_excel()` - Slide generation logic

## Support

For issues or questions:
1. Check logs for detailed error messages
2. Verify database has correct metadata
3. Ensure Excel and PPTX files exist in expected locations
4. Test with feature flag disabled to compare behavior

## Changelog

### 2025-12-12 - Initial Integration

- Added `use_openxml_backup` feature flag to settings
- Created `openxml_pptx_service.py` with your OpenXML code
- Integrated into `presentation_service._generate_physical_powerpoint()`
- Made slide insertion position dynamic (from database)
- Mapped `is_big_japanese` to `presentation_type == "Phonetics"`
- Added data conversion from `ProcessedExcelData` to row format
- Preserved background copying functionality
- Maintained compatibility with existing COM-based flow
