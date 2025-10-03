"""Unit tests for Word document generation service."""

from pathlib import Path
from unittest.mock import MagicMock, Mock, patch

import pytest

from app.constants import WordTemplateType
from app.models.excel_models import ProcessedExcelData
from app.services.word_service import _resolve_template_path, generate_feedback_document


@pytest.mark.unit
class TestWordService:
    """Test suite for Word document generation."""

    def test_resolve_template_path_normal(self):
        """Test resolving template path for normal type."""
        with patch('app.services.word_service.TEMPLATES_ROOT') as mock_root:
            mock_template = Mock(spec=Path)
            mock_template.exists.return_value = True
            mock_root.__truediv__.return_value = mock_template
            
            path = _resolve_template_path(
                presentation_type=WordTemplateType.NORMAL,
                project_type="nw"
            )
            
            assert path is not None

    def test_resolve_template_path_phonetics(self):
        """Test resolving template path for phonetics type."""
        with patch('app.services.word_service.TEMPLATES_ROOT') as mock_root:
            mock_template = Mock(spec=Path)
            mock_template.exists.return_value = True
            mock_root.__truediv__.return_value = mock_template
            
            path = _resolve_template_path(
                presentation_type=WordTemplateType.PHONETICS,
                project_type="nw"
            )
            
            assert path is not None

    def test_resolve_template_path_katakana(self):
        """Test resolving template path for katakana type."""
        with patch('app.services.word_service.TEMPLATES_ROOT') as mock_root:
            mock_template = Mock(spec=Path)
            mock_template.exists.return_value = True
            mock_root.__truediv__.return_value = mock_template
            
            path = _resolve_template_path(
                presentation_type=WordTemplateType.KATAKANA,
                project_type="nw"
            )
            
            assert path is not None

    def test_resolve_template_path_not_found(self):
        """Test that missing template raises error."""
        with patch('app.services.word_service.TEMPLATES_ROOT') as mock_root:
            mock_template = Mock(spec=Path)
            mock_template.exists.return_value = False
            mock_root.__truediv__.return_value = mock_template
            
            with pytest.raises(FileNotFoundError):
                _resolve_template_path(
                    presentation_type=WordTemplateType.NORMAL,
                    project_type="nw"
                )

    @patch('app.services.word_service.Document', None)
    @patch('app.services.word_service._copy_template_to_output')
    @patch('app.services.word_service._resolve_template_path')
    def test_generate_feedback_document_no_docx(
        self, mock_resolve, mock_copy, temp_test_dir
    ):
        """Test that generating document without python-docx still works."""
        excel_data = ProcessedExcelData(
            lst_categories=["Cat1"],
            lst_names=["Name1"],
            lst_rationales=["Reason1"],
            lst_types=[""],
            lst_notations=[""],
            lst_kana=[""],
            lst_logos=[""],
            lst_name_sub_groups=[""],
        )
        
        mock_template = temp_test_dir / "template.doc"
        mock_output = temp_test_dir / "output.doc"
        mock_resolve.return_value = mock_template
        mock_copy.return_value = mock_output
        
        result = generate_feedback_document(
            presentation_id=1,
            project_type="nw",
            presentation_type=WordTemplateType.NORMAL,
            excel_data=excel_data,
            display_name="Test",
            user_name="test_user"
        )
        
        # Should return path even when Document is not available
        assert result == str(mock_output)

    @patch('app.services.word_service._populate_docx')
    @patch('app.services.word_service._copy_template_to_output')
    @patch('app.services.word_service._resolve_template_path')
    @patch('app.services.word_service.Document')
    def test_generate_feedback_document_basic(
        self, mock_doc, mock_resolve, mock_copy, mock_populate, temp_test_dir
    ):
        """Test basic document generation."""
        excel_data = ProcessedExcelData(
            lst_categories=["Category1"],
            lst_names=["TestName"],
            lst_rationales=["Test Rationale"],
            lst_types=[""],
            lst_notations=[""],
            lst_kana=[""],
            lst_logos=[""],
            lst_name_sub_groups=[""],
        )
        
        mock_template = temp_test_dir / "template.docx"
        mock_output = temp_test_dir / "output.docx"
        mock_resolve.return_value = mock_template
        mock_copy.return_value = mock_output
        
        result = generate_feedback_document(
            presentation_id=1,
            project_type="nw",
            presentation_type=WordTemplateType.NORMAL,
            excel_data=excel_data,
            display_name="TestProject",
            user_name="test_user"
        )
        
        # Should return path to generated document
        assert result == str(mock_output)
        mock_populate.assert_called_once()

    def test_generate_feedback_document_design_mode(self):
        """Test that design mode skips Word generation."""
        excel_data = ProcessedExcelData(
            lst_categories=[],
            lst_names=[],
            lst_rationales=[],
            lst_types=[],
            lst_notations=[],
            lst_kana=[],
            lst_logos=[],
            lst_name_sub_groups=[],
        )
        
        result = generate_feedback_document(
            presentation_id=1,
            project_type="design",
            presentation_type=WordTemplateType.NORMAL,
            excel_data=excel_data,
            display_name="TestProject",
            user_name="test_user"
        )
        
        # Design mode should return empty string
        assert result == ""
