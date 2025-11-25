"""Models representing Excel parsing results for the FastAPI parity implementation."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List

from pydantic import BaseModel, Field


@dataclass
class ProcessedExcelData:
    """Container that mirrors the arrays produced by the original VB.NET clsExcel class."""

    lst_types: List[str] = field(default_factory=list)
    lst_categories: List[str] = field(default_factory=list)
    lst_names: List[str] = field(default_factory=list)
    lst_rationales: List[str] = field(default_factory=list)
    lst_notations: List[str] = field(default_factory=list)
    lst_kana: List[str] = field(default_factory=list)
    lst_logos: List[str] = field(default_factory=list)
    lst_name_sub_groups: List[str] = field(default_factory=list)
    lst_group_letters: List[str] = field(default_factory=list)  # Group letter (A, B, C) for each row

    # Derived metadata
    total_rows_processed: int = 0
    candidate_count: int = 0
    group_count: int = 0
    has_groups: bool = False
    is_phonetics: bool = False

    # Optional helpers used during PPT/Word generation
    group_row_indexes: List[int] = field(default_factory=list)
    candidate_row_indexes: List[int] = field(default_factory=list)
    rotation_reset_indexes: List[int] = field(default_factory=list)

    def __post_init__(self) -> None:
        self.total_rows_processed = len(self.lst_types)
        self.candidate_row_indexes = [
            idx
            for idx, marker in enumerate(self.lst_types)
            if marker is None
            or marker == ""
            or not marker.strip()
            or marker.strip().upper() not in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        ]
        self.group_row_indexes = [
            idx
            for idx, marker in enumerate(self.lst_types)
            if marker and marker.strip().upper() in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        ]
        self.candidate_count = len(self.candidate_row_indexes)
        self.group_count = len(self.group_row_indexes)
        self.has_groups = self.group_count > 0

        # Rotation should reset after each group marker and at the beginning.
        rotation_points: List[int] = [idx for idx in self.group_row_indexes]
        if 0 not in rotation_points:
            rotation_points.insert(0, 0)
        self.rotation_reset_indexes = sorted(set(rotation_points))

    @property
    def lst_max_item_number(self) -> int:
        """Match VB convention where the arrays are 0-based and expose the max index."""
        return max(len(self.lst_types) - 1, 0)

    def row_dict(self, index: int) -> Dict[str, Any]:
        """Return the row content as a dictionary for convenience."""
        return {
            "type": self._safe_get(self.lst_types, index),
            "category": self._safe_get(self.lst_categories, index),
            "name": self._safe_get(self.lst_names, index),
            "rationale": self._safe_get(self.lst_rationales, index),
            "notation": self._safe_get(self.lst_notations, index),
            "kana": self._safe_get(self.lst_kana, index),
            "logo": self._safe_get(self.lst_logos, index),
            "name_sub_group": self._safe_get(self.lst_name_sub_groups, index),
            "group_letter": self._safe_get(self.lst_group_letters, index),
        }

    def as_dict(self) -> Dict[str, Any]:
        """Serialize to a JSON-friendly dict for API responses."""
        return {
            "lst_types": self.lst_types,
            "lst_categories": self.lst_categories,
            "lst_names": self.lst_names,
            "lst_rationales": self.lst_rationales,
            "lst_notations": self.lst_notations,
            "lst_kana": self.lst_kana,
            "lst_logos": self.lst_logos,
            "lst_name_sub_groups": self.lst_name_sub_groups,
            "lst_group_letters": self.lst_group_letters,
            "total_rows_processed": self.total_rows_processed,
            "candidate_count": self.candidate_count,
            "group_count": self.group_count,
            "has_groups": self.has_groups,
            "is_phonetics": self.is_phonetics,
            "group_row_indexes": self.group_row_indexes,
            "candidate_row_indexes": self.candidate_row_indexes,
            "rotation_reset_indexes": self.rotation_reset_indexes,
        }

    @staticmethod
    def _safe_get(values: List[str], index: int) -> str:
        try:
            return values[index]
        except IndexError:
            return ""




class ProcessedExcelDataModel(BaseModel):
    """Pydantic representation of processed Excel data for API responses."""

    lst_types: List[str] = Field(default_factory=list)
    lst_categories: List[str] = Field(default_factory=list)
    lst_names: List[str] = Field(default_factory=list)
    lst_rationales: List[str] = Field(default_factory=list)
    lst_notations: List[str] = Field(default_factory=list)
    lst_kana: List[str] = Field(default_factory=list)
    lst_logos: List[str] = Field(default_factory=list)
    lst_name_sub_groups: List[str] = Field(default_factory=list)
    lst_group_letters: List[str] = Field(default_factory=list)
    total_rows_processed: int = 0
    candidate_count: int = 0
    group_count: int = 0
    has_groups: bool = False
    is_phonetics: bool = False
    group_row_indexes: List[int] = Field(default_factory=list)
    candidate_row_indexes: List[int] = Field(default_factory=list)
    rotation_reset_indexes: List[int] = Field(default_factory=list)

    @classmethod
    def from_processed(cls, data: "ProcessedExcelData") -> "ProcessedExcelDataModel":
        return cls(**data.as_dict())


class ExcelProcessingResponse(BaseModel):
    """Response schema for Excel processing endpoint."""

    message: str
    data: ProcessedExcelDataModel
    processing_id: str

    @classmethod
    def from_processed(
        cls, *, processed: "ProcessedExcelData", message: str, processing_id: str
    ) -> "ExcelProcessingResponse":
        return cls(
            message=message,
            data=ProcessedExcelDataModel.from_processed(processed),
            processing_id=processing_id,
        )


__all__ = [
    "ProcessedExcelData",
    "ProcessedExcelDataModel",
    "ExcelProcessingResponse",
]
