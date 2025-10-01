"""Word document generation service."""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Dict, Optional

try:  # pragma: no cover - optional dependency
    from docx import Document  # type: ignore
except ImportError:  # pragma: no cover
    Document = None  # type: ignore

from app.config import TEMPLATES_ROOT, get_project_root
from app.models.excel_models import ProcessedExcelData
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


_TEMPLATE_MAP: Dict[str, str] = {
    "normal": "InputDocumentRationales.doc",
    "normal-noneutral": "InputDocumentRationales.doc",
    "katakana": "InputDocumentRationales_KataKana.doc",
    "katakana_bigjap": "InputDocumentRationales_KataKana.doc",
    "phonetics": "InputDocumentRationales_Phonetics.doc",
    "nonproprietary": "InputDocumentRationales.doc",
    "tagline": "InputDocumentRationales.doc",
    "design": "",
}

_BSR_TEMPLATE = Path("BSR_TEMPLATE") / "BSR_WORD_TEMPLATE.docx"


class WordService:
    def generate_feedback_template(
        self,
        presentation_id: int,
        project_type: str,
        presentation_type: str,
        excel_data: ProcessedExcelData,
        display_name: str,
        user_name: str,
    ) -> str:
        presentation_type_key = presentation_type.lower()
        if project_type.lower() in {"design"} or presentation_type_key == "design":
            logger.info("Design mode detected; skipping Word template generation")
            return ""

        template_path = self._resolve_template_path(project_type, presentation_type)
        if not template_path:
            logger.info(
                "No Word template found for project type '%s' and presentation type '%s'",
                project_type,
                presentation_type,
            )
            return ""

        output_path = self._copy_template_to_output(project_type, display_name, template_path)

        if output_path.suffix.lower() == ".docx" and Document is not None:
            self._populate_docx(output_path, excel_data)
        else:
            logger.info(
                "Template %s copied without modifications (format %s)",
                template_path,
                output_path.suffix,
            )

        return str(output_path)

    def _resolve_template_path(self, project_type: str, presentation_type: str) -> Optional[Path]:
        project_type_lower = project_type.lower()
        if project_type_lower in {"bsr", "nsr"}:
            template = TEMPLATES_ROOT / _BSR_TEMPLATE
        else:
            template_name = _TEMPLATE_MAP.get(presentation_type.lower())
            if not template_name:
                return None
            template = TEMPLATES_ROOT / template_name

        if not template.exists():
            raise FileNotFoundError(f"Template not found: {template}")
        return template

    def _copy_template_to_output(self, project_type: str, display_name: str, template_path: Path) -> Path:
        project_root = get_project_root(project_type)
        downloads_dir = project_root / display_name / "Downloads"
        downloads_dir.mkdir(parents=True, exist_ok=True)

        output_extension = template_path.suffix
        output_path = downloads_dir / f"{display_name}{output_extension}"
        shutil.copy2(template_path, output_path)
        logger.info("Template copied to %s", output_path)
        return output_path

    def _populate_docx(self, output_path: Path, excel_data: ProcessedExcelData) -> None:
        if Document is None:
            return
        logger.info("Populating DOCX document with Excel data")
        document = Document(str(output_path))

    # Clear any previous table when using controlled placeholders (optional)
        document.add_heading("Candidate Results", level=1)
        table = document.add_table(rows=1, cols=4)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Name"
        hdr_cells[1].text = "Category"
        hdr_cells[2].text = "Rationale"
        hdr_cells[3].text = "Subgroup"

        for idx in excel_data.candidate_row_indexes:
            row_info = excel_data.row_dict(idx)
            row_cells = table.add_row().cells
            row_cells[0].text = row_info["name"]
            row_cells[1].text = row_info["category"]
            row_cells[2].text = row_info["rationale"]
            row_cells[3].text = row_info["name_sub_group"]

        document.save(str(output_path))


word_service = WordService()
