"""
Excel processing service for transforming Excel data into processed arrays.
"""

# pylint: disable=broad-exception-caught

import re
from io import BytesIO

import openpyxl
from fastapi import HTTPException

from app.models.excel_models import ProcessedExcelData
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class ExcelProcessingService:
    """Service for processing Excel files and returning processed data arrays."""

    def __init__(self):
        self.max_rows = 1000

        # Arrays for Excel data processing
        self.lst_categories = []
        self.lst_names = []
        self.lst_rationales = []
        self.lst_notations = []
        self.lst_types = []
        self.lst_kana = []
        self.lst_logos = []
        self.lst_name_sub_groups = []
        self.lst_max_item_number = 0

    def process_excel_file(
        self,
        file_content: bytes,
        is_phonetics: bool = False,
        has_groups: bool = False,
    ) -> ProcessedExcelData:
        """
        Process Excel file and return processed data arrays.

        Args:
            file_content: Raw bytes of the Excel file
            is_phonetics: Whether to use phonetics processing
            has_groups: Whether to process groups

        Returns:
            ProcessedExcelData: Processed arrays for Excel data
        """
        try:
            # Reset arrays
            self._reset_arrays()

            # Load Excel data into arrays
            if not is_phonetics and not has_groups:
                self._load_excel_basic(file_content)
            elif is_phonetics and not has_groups:
                self._load_excel_phonetics(file_content)
            elif not is_phonetics and has_groups:
                self._load_excel_with_groups(file_content)
            else:  # is_phonetics and has_groups
                self._load_excel_phonetics_with_groups(file_content)

            # Create response with processed data
            processed_data = ProcessedExcelData(
                lst_categories=self.lst_categories[: self.lst_max_item_number + 1],
                lst_names=self.lst_names[: self.lst_max_item_number + 1],
                lst_rationales=self.lst_rationales[: self.lst_max_item_number + 1],
                lst_notations=self.lst_notations[: self.lst_max_item_number + 1],
                lst_types=self.lst_types[: self.lst_max_item_number + 1],
                lst_kana=self.lst_kana[: self.lst_max_item_number + 1],
                lst_logos=self.lst_logos[: self.lst_max_item_number + 1],
                lst_name_sub_groups=self.lst_name_sub_groups[
                    : self.lst_max_item_number + 1
                ],
                lst_max_item_number=self.lst_max_item_number,
                total_rows_processed=self.lst_max_item_number + 1,
            )

            logger.info(
                "Successfully processed Excel file. Processed %d rows.",
                processed_data.total_rows_processed,
            )
            return processed_data

        except Exception as e:
            logger.error("Error processing Excel file: %s", str(e))
            raise HTTPException(
                status_code=400, detail=f"Failed to process Excel file: {str(e)}"
            ) from e

    def _reset_arrays(self):
        """Reset all data arrays to initial state."""
        self.lst_max_item_number = 0
        self.lst_categories = [""] * self.max_rows
        self.lst_names = [""] * self.max_rows
        self.lst_rationales = [""] * self.max_rows
        self.lst_notations = [""] * self.max_rows
        self.lst_types = [""] * self.max_rows
        self.lst_kana = [""] * self.max_rows
        self.lst_logos = [""] * self.max_rows
        self.lst_name_sub_groups = [""] * self.max_rows

    def _load_excel_basic(self, file_content: bytes):
        """Load Excel file using basic processing."""
        try:
            workbook = openpyxl.load_workbook(BytesIO(file_content))
            worksheet = workbook.active

            # Find last row with data
            numrows = 0
            keep_checking = True

            while keep_checking and numrows < 10000:  # Prevent infinite loop
                cell_a = worksheet[f"A{numrows + 1}"].value
                cell_b = worksheet[f"B{numrows + 1}"].value

                if (cell_a is None or str(cell_a).strip() == "") and (
                    cell_b is None or str(cell_b).strip() == ""
                ):
                    keep_checking = False
                else:
                    numrows += 1

            # Set max item number
            if numrows > 0:
                self.lst_max_item_number = numrows - 2  # Skip header rows

            # Process rows starting from row 2 (skip headers)
            for i in range(2, numrows + 1):
                try:
                    # Column A: Type
                    self.lst_types[i - 2] = str(worksheet[f"A{i}"].value or "")

                    # Column B: Category
                    self.lst_categories[i - 2] = self._clean_str(
                        str(worksheet[f"B{i}"].value or "")
                    )

                    # Column C: Name + Notation (needs processing)
                    name_with_notation = str(worksheet[f"C{i}"].value or "")
                    name, notation = self._get_test_name(name_with_notation)
                    self.lst_names[i - 2] = name
                    self.lst_notations[i - 2] = notation

                    # Column D: Rationale
                    self.lst_rationales[i - 2] = self._clean_str(
                        str(worksheet[f"D{i}"].value or "")
                    )

                    # Column E: Kana
                    self.lst_kana[i - 2] = str(worksheet[f"E{i}"].value or "")

                    # Column F: Logo
                    self.lst_logos[i - 2] = str(worksheet[f"F{i}"].value or "")

                except Exception as e:
                    # Justified: Excel processing can raise various exceptions (IO, parsing, etc.)
                    logger.warning("Error processing row %d: %s", i, str(e))
                    continue

            workbook.close()

        except Exception as e:
            logger.error("Error in _load_excel_basic: %s", str(e))
            raise

    def _load_excel_phonetics(self, file_content: bytes):
        """Load Excel file with phonetics processing."""
        try:
            workbook = openpyxl.load_workbook(BytesIO(file_content))
            worksheet = workbook.active

            # Find last row with data (same logic as basic)
            numrows = 0
            keep_checking = True

            while keep_checking and numrows < 10000:
                cell_a = worksheet[f"A{numrows + 1}"].value
                cell_b = worksheet[f"B{numrows + 1}"].value

                if (cell_a is None or str(cell_a).strip() == "") and (
                    cell_b is None or str(cell_b).strip() == ""
                ):
                    keep_checking = False
                else:
                    numrows += 1

            if numrows > 0:
                self.lst_max_item_number = numrows - 2

            # Process rows with phonetics-specific handling
            for i in range(2, numrows + 1):
                try:
                    # Column A: Type
                    self.lst_types[i - 2] = str(worksheet[f"A{i}"].value or "")

                    # Column B: Category
                    self.lst_categories[i - 2] = self._clean_str(
                        str(worksheet[f"B{i}"].value or "")
                    )

                    # Column C: Name + Notation with phonetics processing
                    name_with_notation = str(worksheet[f"C{i}"].value or "")
                    name, notation = self._get_test_name_phonetics(name_with_notation)
                    self.lst_names[i - 2] = name
                    self.lst_notations[i - 2] = notation

                    # Column D: Rationale
                    self.lst_rationales[i - 2] = self._clean_str(
                        str(worksheet[f"D{i}"].value or "")
                    )

                    # Column E: Kana (phonetics specific processing)
                    kana_value = str(worksheet[f"E{i}"].value or "")
                    self.lst_kana[i - 2] = self._process_phonetics_kana(kana_value)

                    # Column F: Logo
                    self.lst_logos[i - 2] = str(worksheet[f"F{i}"].value or "")

                except Exception as e:
                    # Justified: Excel processing can raise various exceptions (IO, parsing, etc.)
                    logger.warning("Error processing phonetics row %d: %s", i, str(e))
                    continue

            workbook.close()

        except Exception as e:
            logger.error("Error in _load_excel_phonetics: %s", str(e))
            raise

    def _load_excel_with_groups(self, file_content: bytes):
        """Load Excel file with groups processing (isPhonetics, hasGroups)."""
        try:
            workbook = openpyxl.load_workbook(BytesIO(file_content))
            worksheet = workbook.active

            # Find last row with data
            numrows = 0
            keep_checking = True

            while keep_checking and numrows < 10000:
                cell_a = worksheet[f"A{numrows + 1}"].value
                cell_b = worksheet[f"B{numrows + 1}"].value

                if (cell_a is None or str(cell_a).strip() == "") and (
                    cell_b is None or str(cell_b).strip() == ""
                ):
                    keep_checking = False
                else:
                    numrows += 1

            if numrows > 0:
                self.lst_max_item_number = numrows - 2

            # Process rows with group handling
            for i in range(2, numrows + 1):
                try:
                    self.lst_types[i - 2] = str(worksheet[f"A{i}"].value or "")
                    self.lst_categories[i - 2] = self._clean_str(
                        str(worksheet[f"B{i}"].value or "")
                    )

                    name_with_notation = str(worksheet[f"C{i}"].value or "")
                    name, notation = self._get_test_name(name_with_notation)
                    self.lst_names[i - 2] = name
                    self.lst_notations[i - 2] = notation

                    self.lst_rationales[i - 2] = self._clean_str(
                        str(worksheet[f"D{i}"].value or "")
                    )
                    self.lst_kana[i - 2] = str(worksheet[f"E{i}"].value or "")
                    self.lst_logos[i - 2] = str(worksheet[f"F{i}"].value or "")

                    # Column G: NameSubGroup (for groups)
                    self.lst_name_sub_groups[i - 2] = str(
                        worksheet[f"G{i}"].value or ""
                    )

                except Exception as e:
                    # Justified: Excel processing can raise various exceptions (IO, parsing, etc.)
                    logger.warning("Error processing row %d: %s", i, str(e))
                    continue

            workbook.close()

        except Exception as e:
            logger.error("Error in _load_excel_with_groups: %s", str(e))
            raise

    def _load_excel_phonetics_with_groups(self, file_content: bytes):
        """Load Excel file with both phonetics and groups processing - combined logic."""
        try:
            workbook = openpyxl.load_workbook(BytesIO(file_content))
            worksheet = workbook.active

            # Find last row with data
            numrows = 0
            keep_checking = True

            while keep_checking and numrows < 10000:
                cell_a = worksheet[f"A{numrows + 1}"].value
                cell_b = worksheet[f"B{numrows + 1}"].value

                if (cell_a is None or str(cell_a).strip() == "") and (
                    cell_b is None or str(cell_b).strip() == ""
                ):
                    keep_checking = False
                else:
                    numrows += 1

            if numrows > 0:
                self.lst_max_item_number = numrows - 2

            # Process rows with combined phonetics and groups handling
            for i in range(2, numrows + 1):
                try:
                    self.lst_types[i - 2] = str(worksheet[f"A{i}"].value or "")
                    self.lst_categories[i - 2] = self._clean_str(
                        str(worksheet[f"B{i}"].value or "")
                    )

                    # Combined phonetics + groups processing for name
                    name_with_notation = str(worksheet[f"C{i}"].value or "")
                    name, notation = self._get_test_name_phonetics(name_with_notation)
                    self.lst_names[i - 2] = name
                    self.lst_notations[i - 2] = notation

                    self.lst_rationales[i - 2] = self._clean_str(
                        str(worksheet[f"D{i}"].value or "")
                    )

                    # Kana with phonetics processing
                    kana_value = str(worksheet[f"E{i}"].value or "")
                    self.lst_kana[i - 2] = self._process_phonetics_kana(kana_value)

                    self.lst_logos[i - 2] = str(worksheet[f"F{i}"].value or "")

                    # Column G: NameSubGroup (for groups)
                    self.lst_name_sub_groups[i - 2] = str(
                        worksheet[f"G{i}"].value or ""
                    )

                except Exception as e:
                    # Justified: Excel processing can raise various exceptions (IO, parsing, etc.)
                    logger.warning("Error processing combined row %d: %s", i, str(e))
                    continue

            workbook.close()

        except Exception as e:
            logger.error("Error in _load_excel_phonetics_with_groups: %s", str(e))
            raise

    def _get_test_name(self, name_str: str) -> tuple[str, str]:
        """
        Extract name and notation from combined string.

        Returns:
            tuple: (clean_name, notation)
        """
        if not name_str or name_str.strip() == "":
            return "", ""

        name_str = name_str.strip()

        # Common notation patterns
        notation_patterns = [
            r"\(c\)",
            r"®",
            r"™",
            r"©",
            r"℗",
            r"℠",  # Trademark/copyright symbols
            r"\bTM\b",
            r"\bR\b",
            r"\bC\b",  # TM, R, C notations
            r"\d+",  # Numbers that might be notations
        ]

        name = name_str
        notation = ""

        # Check for notation patterns at the end of the string
        for pattern in notation_patterns:
            match = re.search(f"{pattern}\\s*$", name_str, re.IGNORECASE)
            if match:
                notation = match.group().strip()
                name = re.sub(
                    f"{pattern}\\s*$", "", name_str, flags=re.IGNORECASE
                ).strip()
                break

        # If no pattern found, check if the string ends with parentheses
        if not notation and name_str.endswith(")"):
            paren_start = name_str.rfind("(")
            if paren_start > 0:
                potential_notation = name_str[paren_start:]
                # Check if parentheses contain notation-like content
                if any(
                    char in potential_notation.lower()
                    for char in ["c", "r", "tm", "®", "™", "©"]
                ):
                    notation = potential_notation
                    name = name_str[:paren_start].strip()

        # Clean up the name
        name = self._clean_str(name)

        return name, notation

    def _get_test_name_phonetics(self, name_str: str) -> tuple[str, str]:
        """
        Extract name and notation with phonetics-specific processing.
        """
        # Start with basic processing
        name, notation = self._get_test_name(name_str)

        # Additional phonetics processing
        if name:
            # Handle special phonetics characters or formatting
            # This might include specific Japanese character handling
            name = self._process_phonetics_name(name)

        return name, notation

    def _process_phonetics_kana(self, kana_str: str) -> str:
        """
        Process Kana column with phonetics-specific logic.
        """
        if not kana_str or kana_str.strip() == "":
            return ""

        kana_str = kana_str.strip()

        # Basic validation for Japanese characters
        # You might want to add more sophisticated Japanese text validation here
        if self._is_japanese_text(kana_str):
            return kana_str

        # If not Japanese, return as-is but cleaned
        return self._clean_str(kana_str)

    def _process_phonetics_name(self, name_str: str) -> str:
        """
        Process name with phonetics-specific logic.
        """
        if not name_str:
            return ""

        # Handle any phonetics-specific name processing
        # This might include character normalization, special handling for Japanese names, etc.

        # For now, just return cleaned string
        return self._clean_str(name_str)

    def _is_japanese_text(self, text: str) -> bool:
        """
        Check if text contains Japanese characters (Hiragana, Katakana, Kanji).
        """
        if not text:
            return False

        # Japanese character ranges
        japanese_ranges = [
            (0x3040, 0x309F),  # Hiragana
            (0x30A0, 0x30FF),  # Katakana
            (0x4E00, 0x9FFF),  # Kanji
        ]

        for char in text:
            code = ord(char)
            if any(start <= code <= end for start, end in japanese_ranges):
                return True

        return False

    def _clean_str(self, s: str) -> str:
        """
        Clean string by replacing single quotes with backticks.
        """
        if s and "'" in s:
            return s.replace("'", "`")
        return s


# Create service instance
excel_processing_service = ExcelProcessingService()
