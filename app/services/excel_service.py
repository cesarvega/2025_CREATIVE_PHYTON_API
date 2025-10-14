"""Excel processing service that mirrors the original VB.NET clsExcel behavior."""

from __future__ import annotations

import re
from io import BytesIO
from typing import Tuple

from openpyxl import load_workbook

from app.models.excel_models import ProcessedExcelData
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


HEADER_ROW_INDEX = 1
TYPE_COLUMN = 1                   # Column A: Group marker (A, B, C, etc.)
CATEGORY_COLUMN = 2               # Column B: Category
NAME_COLUMN = 3                   # Column C: Name
RATIONALE_COLUMN = 4              # Column D: Rationale
KANA_COLUMN = 5                   # Column E: Kana/Pronunciation
LOGO_COLUMN = 6                   # Column F: Logo
NAME_SUBGROUP_COLUMN = 7          # Column G: Group1 (primary subgroup column)
                                  # Column H: Group2 (fallback if G is empty, index 7)

GROUP_MARKERS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
GROUP_DELIMITER = "##"
# Accept a leading single letter as a valid group marker, optionally followed by
# punctuation or whitespace. Examples matched: "A", "A:", "A)", "A -", "A."
# Examples NOT matched: "Group A", "A1", "AA".
_group_marker_regex = re.compile(r"^\s*([A-Za-z])\s*(?:[:\)\.\-–—]|\Z)")

_name_notation_regex = re.compile(r"^(?P<name>[^\(]+?)(?:\((?P<notation>.+)\))?$")


def _clean(value: str) -> str:
    if value is None or str(value).strip().lower() == "none":
        return ""
    return str(value).strip()


def _normalize_kana(value: str, is_phonetics: bool) -> str:
    value = _clean(value)
    if is_phonetics and value:
        value = value.replace('"', '""')
    return value


def _split_name_and_notation(raw_value: str) -> Tuple[str, str]:
    if not raw_value:
        return "", ""

    raw_value = raw_value.strip()
    match = _name_notation_regex.match(raw_value)
    if not match:
        return raw_value, ""

    name = match.group("name").strip()
    notation = match.group("notation")
    return name, notation.strip() if notation else ""


def process_excel_file(
    file_content: bytes,
    is_phonetics: bool = False,
    has_groups: bool = False,
) -> ProcessedExcelData:
    """
    Parse the Excel workbook and return a structure equivalent to clsExcel.LoadExcelFile.

    Optimized with efficient row filtering and minimal memory allocations.
    """
    workbook = load_workbook(filename=BytesIO(file_content), data_only=True)
    sheet = workbook.active

    lst_types = []
    lst_categories = []
    lst_names = []
    lst_rationales = []
    lst_notations = []
    lst_kana = []
    lst_logos = []
    lst_name_sub_groups = []

    # Collect all rows first - optimized with generator and early filtering
    all_rows = []
    for row in sheet.iter_rows(min_row=HEADER_ROW_INDEX + 1, max_col=NAME_SUBGROUP_COLUMN + 1):
        # Quick check: skip empty rows early
        if all(cell.value is None or str(cell.value).strip() == "" for cell in row[:LOGO_COLUMN]):
            continue
        all_rows.append(row)

    if not has_groups:
        for row in all_rows:
            raw_type = _clean(str(row[TYPE_COLUMN - 1].value) if len(row) >= TYPE_COLUMN else "")
            raw_category = _clean(str(row[CATEGORY_COLUMN - 1].value) if len(row) >= CATEGORY_COLUMN else "")
            raw_name = _clean(str(row[NAME_COLUMN - 1].value) if len(row) >= NAME_COLUMN else "")
            raw_rationale = _clean(str(row[RATIONALE_COLUMN - 1].value) if len(row) >= RATIONALE_COLUMN else "")
            raw_kana = _normalize_kana(str(row[KANA_COLUMN - 1].value) if len(row) >= KANA_COLUMN else "", is_phonetics)
            raw_logo = _clean(str(row[LOGO_COLUMN - 1].value) if len(row) >= LOGO_COLUMN else "")
            
            # Extract subgroup: Check column G (Group1) first
            raw_name_sub_group = _clean(
                str(row[NAME_SUBGROUP_COLUMN - 1].value) if len(row) >= NAME_SUBGROUP_COLUMN else ""
            )
            # If column G is empty, fallback to column H (Group2)
            if not raw_name_sub_group and len(row) >= NAME_SUBGROUP_COLUMN + 1:
                raw_name_sub_group = _clean(str(row[NAME_SUBGROUP_COLUMN].value))

            name, notation = _split_name_and_notation(raw_name)

            if has_groups:
                m = _group_marker_regex.match(raw_type)
                marker = m.group(1).upper() if m else ""
            else:
                marker = ""

            lst_types.append(marker)
            lst_categories.append(raw_category)
            lst_names.append(name)
            lst_rationales.append(raw_rationale)
            lst_notations.append(notation)
            lst_kana.append(raw_kana)
            lst_logos.append(raw_logo)
            lst_name_sub_groups.append(raw_name_sub_group)
    
    else:
        # NEW: Group processing logic
        processed_indices = set()
        
        for i, row in enumerate(all_rows):
            if i in processed_indices:
                continue
            
            # Extract subgroup: Check column G (Group1) first
            raw_name_sub_group = _clean(
                str(row[NAME_SUBGROUP_COLUMN - 1].value) if len(row) >= NAME_SUBGROUP_COLUMN else ""
            )
            # If column G is empty, fallback to column H (Group2)
            if not raw_name_sub_group and len(row) >= NAME_SUBGROUP_COLUMN + 1:
                raw_name_sub_group = _clean(str(row[NAME_SUBGROUP_COLUMN].value))
            
            delimiter = GROUP_DELIMITER
            
            # Process row data
            raw_type = _clean(str(row[TYPE_COLUMN - 1].value) if len(row) >= TYPE_COLUMN else "")
            raw_category = _clean(str(row[CATEGORY_COLUMN - 1].value) if len(row) >= CATEGORY_COLUMN else "")
            raw_name = _clean(str(row[NAME_COLUMN - 1].value) if len(row) >= NAME_COLUMN else "")
            raw_rationale = _clean(str(row[RATIONALE_COLUMN - 1].value) if len(row) >= RATIONALE_COLUMN else "")
            raw_kana = _normalize_kana(str(row[KANA_COLUMN - 1].value) if len(row) >= KANA_COLUMN else "", is_phonetics)
            raw_logo = _clean(str(row[LOGO_COLUMN - 1].value) if len(row) >= LOGO_COLUMN else "")
            
            name, notation = _split_name_and_notation(raw_name)
            
            if raw_type:
                m = _group_marker_regex.match(raw_type)
                marker = m.group(1).upper() if m else ""
            else:
                marker = ""
            
            if not raw_name_sub_group:
                # Single row, no grouping
                lst_types.append(marker)
                lst_categories.append(raw_category)
                lst_names.append(name)
                lst_rationales.append(raw_rationale)
                lst_notations.append(notation)
                lst_kana.append(raw_kana)
                lst_logos.append(raw_logo)
                lst_name_sub_groups.append("")
                processed_indices.add(i)
            else:
                # Group multiple rows with same sub_group
                grouped_names = [name]
                grouped_rationales = [raw_rationale] if raw_rationale else []
                processed_indices.add(i)
                
                # Look for other rows with same sub_group
                for j in range(i + 1, len(all_rows)):
                    if j in processed_indices:
                        continue
                    
                    other_row = all_rows[j]
                    # Extract subgroup from other row: Check column G (Group1) first
                    other_sub_group = _clean(
                        str(other_row[NAME_SUBGROUP_COLUMN - 1].value) if len(other_row) >= NAME_SUBGROUP_COLUMN else ""
                    )
                    # If column G is empty, fallback to column H (Group2)
                    if not other_sub_group and len(other_row) >= NAME_SUBGROUP_COLUMN + 1:
                        other_sub_group = _clean(str(other_row[NAME_SUBGROUP_COLUMN].value))
                    
                    if other_sub_group == raw_name_sub_group:
                        other_name = _clean(str(other_row[NAME_COLUMN - 1].value) if len(other_row) >= NAME_COLUMN else "")
                        other_rationale = _clean(str(other_row[RATIONALE_COLUMN - 1].value) if len(other_row) >= RATIONALE_COLUMN else "")
                        other_name_clean, _ = _split_name_and_notation(other_name)
                        
                        grouped_names.append(other_name_clean)
                        grouped_rationales.append(other_rationale)
                        processed_indices.add(j)
                
                # Combine with delimiter
                combined_name = delimiter.join(grouped_names) + delimiter
                combined_rationale = delimiter.join(grouped_rationales) + delimiter
                
                lst_types.append(marker)
                lst_categories.append(raw_category)
                lst_names.append(combined_name)
                lst_rationales.append(combined_rationale)
                lst_notations.append(notation)
                lst_kana.append(raw_kana)
                lst_logos.append(raw_logo)
                lst_name_sub_groups.append(raw_name_sub_group)

    if not lst_names:
        raise ValueError("The Excel file does not contain valid candidates")

    excel_data = ProcessedExcelData(
        lst_types=lst_types,
        lst_categories=lst_categories,
        lst_names=lst_names,
        lst_rationales=lst_rationales,
        lst_notations=lst_notations,
        lst_kana=lst_kana,
        lst_logos=lst_logos,
        lst_name_sub_groups=lst_name_sub_groups,
        is_phonetics=is_phonetics,
    )

    logger.info(
        "Excel processed: %s rows, %s candidates, %s groups",
        excel_data.total_rows_processed,
        excel_data.candidate_count,
        excel_data.group_count,
    )
    return excel_data

# Singleton service
class ExcelProcessingService:
    def process_excel_file(
        self, file_content: bytes, is_phonetics: bool = False, has_groups: bool = False
    ) -> ProcessedExcelData:
        return process_excel_file(file_content, is_phonetics=is_phonetics, has_groups=has_groups)


excel_processing_service = ExcelProcessingService()
