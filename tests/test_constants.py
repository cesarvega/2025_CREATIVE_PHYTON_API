"""Unit tests for application constants."""

import pytest

from app.constants import (
    DbLimits,
    Defaults,
    FileExtension,
    HttpStatus,
    PresentationType,
    ProjectType,
    SlideType,
    TemplateFilename,
    WordTemplateType,
)


@pytest.mark.unit
class TestPresentationType:
    """Test suite for presentation type constants."""

    def test_presentation_types_defined(self):
        """Test that all presentation types are defined."""
        assert PresentationType.NORMAL == "Normal"
        assert PresentationType.NONPROPRIETARY == "Nonproprietary"
        assert PresentationType.PHONETICS == "Phonetics"
        assert PresentationType.KATAKANA == "Katakana"
        assert PresentationType.TAGLINE == "Tagline"

    def test_presentation_types_all(self):
        """Test that all() returns all presentation types."""
        all_types = PresentationType.all()
        assert isinstance(all_types, list)
        assert len(all_types) > 0
        assert PresentationType.NORMAL in all_types
        assert PresentationType.NONPROPRIETARY in all_types


@pytest.mark.unit
class TestProjectType:
    """Test suite for project type constants."""

    def test_project_types_defined(self):
        """Test that all project types are defined."""
        assert ProjectType.BIPRESENTS == "bipresents"
        assert ProjectType.BSR == "bsr"
        assert ProjectType.NSR == "nsr"
        assert ProjectType.NW == "nw"

    def test_project_types_all(self):
        """Test that all() returns all project types."""
        all_types = ProjectType.all()
        assert isinstance(all_types, list)
        assert len(all_types) == 4
        assert ProjectType.NW in all_types


@pytest.mark.unit
class TestSlideType:
    """Test suite for slide type constants."""

    def test_slide_types_defined(self):
        """Test that all slide types are defined."""
        assert SlideType.NAME_EVALUATION == "NameEvaluation"
        assert SlideType.BSR_GROUP == "BSR_Group"
        assert SlideType.NSR_GROUP == "NSR_Group"

    def test_slide_types_all(self):
        """Test that all() returns all slide types."""
        all_types = SlideType.all()
        assert isinstance(all_types, list)
        assert SlideType.NAME_EVALUATION in all_types


@pytest.mark.unit
class TestTemplateFilename:
    """Test suite for template filename constants."""

    def test_word_templates_defined(self):
        """Test that Word template filenames are defined."""
        assert "Rationales" in TemplateFilename.WORD_RATIONALES
        assert "KataKana" in TemplateFilename.WORD_KATAKANA
        assert "Phonetics" in TemplateFilename.WORD_PHONETICS

    def test_pptx_templates_defined(self):
        """Test that PPTX template filenames are defined."""
        assert "2019" in TemplateFilename.PPTX_DEFAULT_2019
        assert "groups" in TemplateFilename.PPTX_WITH_GROUPS_2019.lower()
        assert "summary" in TemplateFilename.PPTX_SUMMARY_2019.lower()


@pytest.mark.unit
class TestFileExtension:
    """Test suite for file extension constants."""

    def test_excel_extensions(self):
        """Test Excel file extensions."""
        assert FileExtension.XLSX == ".xlsx"
        assert FileExtension.XLS == ".xls"

    def test_powerpoint_extensions(self):
        """Test PowerPoint file extensions."""
        assert FileExtension.PPTX == ".pptx"
        assert FileExtension.PPTM == ".pptm"

    def test_word_extensions(self):
        """Test Word document file extensions."""
        assert FileExtension.DOC == ".doc"
        assert FileExtension.DOCX == ".docx"

    def test_image_extensions(self):
        """Test image file extensions."""
        assert FileExtension.JPG == ".jpg"
        assert FileExtension.PNG == ".png"


@pytest.mark.unit
class TestHttpStatus:
    """Test suite for HTTP status code constants."""

    def test_success_codes(self):
        """Test success HTTP status codes."""
        assert HttpStatus.OK == 200
        assert HttpStatus.CREATED == 201

    def test_client_error_codes(self):
        """Test client error HTTP status codes."""
        assert HttpStatus.BAD_REQUEST == 400
        assert HttpStatus.NOT_FOUND == 404
        assert HttpStatus.PAYLOAD_TOO_LARGE == 413

    def test_server_error_codes(self):
        """Test server error HTTP status codes."""
        assert HttpStatus.INTERNAL_SERVER_ERROR == 500


@pytest.mark.unit
class TestDbLimits:
    """Test suite for database limit constants."""

    def test_string_length_limits(self):
        """Test database string field length limits."""
        assert DbLimits.MAX_PROJECT_NAME_LENGTH > 0
        assert DbLimits.MAX_DISPLAY_NAME_LENGTH > 0
        assert DbLimits.MAX_USER_NAME_LENGTH > 0
        assert DbLimits.MAX_FILE_PATH_LENGTH > 0


@pytest.mark.unit
class TestDefaults:
    """Test suite for default value constants."""

    def test_chart_defaults(self):
        """Test chart default values."""
        assert Defaults.CHART_WIDTH > 0
        assert Defaults.CHART_HEIGHT > 0
        assert Defaults.CHART_DPI > 0

    def test_font_defaults(self):
        """Test font default values."""
        assert Defaults.FONT_NAME == "Calibri"
        assert Defaults.FONT_SIZE > 0

    def test_excel_defaults(self):
        """Test Excel default values."""
        assert Defaults.MAX_EXCEL_ROWS > 0


@pytest.mark.unit
class TestWordTemplateType:
    """Test suite for Word template type constants."""

    def test_word_template_types(self):
        """Test Word template type values."""
        assert WordTemplateType.NORMAL == "normal"
        assert WordTemplateType.PHONETICS == "phonetics"
        assert WordTemplateType.KATAKANA == "katakana"
        assert WordTemplateType.NONPROPRIETARY == "nonproprietary"
