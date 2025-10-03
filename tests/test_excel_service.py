"""Unit tests for the Excel processing service."""

import io

import pytest
from openpyxl import Workbook

from app.services.excel_service import excel_processing_service


@pytest.mark.unit
class TestExcelProcessingService:
    """Test suite for Excel processing functionality."""

    def test_process_excel_without_groups(self, sample_excel_bytes):
        """Test processing Excel file without groups."""
        result = excel_processing_service.process_excel_file(
            file_content=sample_excel_bytes,
            is_phonetics=False,
            has_groups=False,
        )

        assert result is not None
        assert result.total_rows_processed > 0
        assert len(result.lst_categories) > 0
        assert len(result.lst_names) > 0
        assert len(result.lst_rationales) > 0

    def test_process_excel_with_groups(self, sample_excel_with_groups):
        """Test processing Excel file with groups."""
        result = excel_processing_service.process_excel_file(
            file_content=sample_excel_with_groups,
            is_phonetics=False,
            has_groups=True,
        )

        assert result is not None
        assert result.total_rows_processed > 0
        assert len(result.lst_name_sub_groups) > 0

    def test_process_excel_with_phonetics(self, sample_excel_bytes):
        """Test processing Excel with phonetics enabled."""
        result = excel_processing_service.process_excel_file(
            file_content=sample_excel_bytes,
            is_phonetics=True,
            has_groups=False,
        )

        assert result is not None
        assert result.total_rows_processed > 0

    def test_process_empty_excel(self):
        """Test that processing empty Excel raises proper error."""
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["Type", "Category", "Name", "Rationale"])
        
        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)
        
        # Empty Excel should raise ValueError
        with pytest.raises(ValueError, match="does not contain valid candidates"):
            excel_processing_service.process_excel_file(
                file_content=buffer.getvalue(),
                is_phonetics=False,
                has_groups=False,
            )

    def test_processed_data_structure(self, sample_excel_bytes):
        """Test that processed data has correct structure."""
        result = excel_processing_service.process_excel_file(
            file_content=sample_excel_bytes,
            is_phonetics=False,
            has_groups=False,
        )

        assert hasattr(result, 'lst_categories')
        assert hasattr(result, 'lst_names')
        assert hasattr(result, 'lst_rationales')
        assert hasattr(result, 'lst_types')
        assert hasattr(result, 'lst_kana')
        assert hasattr(result, 'lst_logos')
        assert hasattr(result, 'lst_name_sub_groups')
        assert hasattr(result, 'total_rows_processed')
        
        assert len(result.lst_categories) == len(result.lst_names)
        assert len(result.lst_names) == len(result.lst_rationales)

    def test_max_item_number(self, sample_excel_bytes):
        """Test that lst_max_item_number is correctly calculated."""
        result = excel_processing_service.process_excel_file(
            file_content=sample_excel_bytes,
            is_phonetics=False,
            has_groups=False,
        )
        
        # lst_max_item_number is a property that returns max(len - 1, 0)
        expected = max(len(result.lst_types) - 1, 0)
        assert result.lst_max_item_number == expected
