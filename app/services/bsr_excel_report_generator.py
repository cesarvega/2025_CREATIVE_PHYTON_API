"""Service for generating BSR Excel reports using COM automation."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import List, Dict, Set

import win32com.client  # type: ignore

from app.services.bsr_reports_service import bsr_reports_service
from app.utils.logging_utils import get_logger
from app.utils.nw_data_utils import sanitize_filename
from app.utils.path_utils import get_nw_downloads_dir
from app.config.settings import settings

logger = get_logger(__name__)


class BSRExcelReportGenerator:
    """Generate BSR Excel reports from templates."""

    def __init__(self) -> None:
        self.excel = None

    def _get_template_path(self) -> Path:
        return settings.app_dir / "templates" / "BSR_TEMPLATE" / "bsr_excel_template_new.xls"

    def _derive_source(self, row: Dict) -> str:
        # Logic: Mobile/PC/Moderator
        src = (row.get("SOURCE") or row.get("source") or "").strip()
        is_mobile = str(row.get("ISMOBILE") or row.get("is_mobile") or row.get("IsMobile") or "0").strip()
        if src.lower() == "moderator":
            return "Moderator"
        if is_mobile in ("1", "true", "True", 1, True):
            return "Mobile"
        return "PC"

    def _collect_categories(self, data: List[Dict]) -> List[str]:
        cats: Set[str] = set()
        for r in data:
            raw = r.get("Categories") or r.get("categories") or ""
            if raw:
                for c in str(raw).split(","):
                    cname = c.strip()
                    if cname:
                        cats.add(cname)
        return sorted(cats)

    def generate_excel_report(self, presentation_id: int, project_name: str) -> Path:
        """Generate BSR Excel report and save as .xls in downloads folder."""
        template = self._get_template_path()
        if not template.exists():
            raise FileNotFoundError(f"Excel template not found: {template}")

        data = bsr_reports_service.get_excel_report_data(presentation_id)
        logger.info("Excel generator received %d rows of data", len(data))
        if data:
            logger.info("First row keys: %s", list(data[0].keys()))
            logger.info("First row values: %s", data[0])
        categories = self._collect_categories(data)
        logger.info("Categories detected: %s", ", ".join(categories) if categories else "<none>")

        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False

        workbook = None
        try:
            workbook = self.excel.Workbooks.Open(str(template))
            # Find sheet "NAME CANDIDATES"
            sheet = None
            for ws in workbook.Worksheets:
                if str(ws.Name).strip().lower() == "name candidates":
                    sheet = ws
                    break
            if sheet is None:
                sheet = workbook.Worksheets(1)

            # Header row: set category headers from col 6 onward
            start_col = 6
            col = start_col
            for cname in categories:
                sheet.Cells(1, col).Value = cname
                col += 1

            # Fill rows starting at row 2
            row_idx = 2
            for r in data:
                # Column A (1): NAME CANDIDATE
                name = r.get("NAME") or r.get("name") or ""
                if name:
                    name = str(name).strip()
                sheet.Cells(row_idx, 1).Value = name
                # Column B (2): RATIONAL
                sheet.Cells(row_idx, 2).Value = r.get("Rationale") or ""
                # Column C (3): DETAILS - shows PC/Mobile/Moderator
                sheet.Cells(row_idx, 3).Value = self._derive_source(r)
                # Column D (4): SOURCE - original source (user name)
                sheet.Cells(row_idx, 4).Value = r.get("SOURCE") or r.get("source") or ""
                # Column E (5): DATE/TIME CREATED
                created = r.get("createddate") or r.get("created_date") or r.get("CreatedDate")
                dt_val = None
                if created:
                    try:
                        if isinstance(created, str):
                            # Try common formats
                            for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
                                try:
                                    dt_val = datetime.strptime(created.split(".")[0], fmt)
                                    break
                                except Exception:
                                    pass
                        else:
                            dt_val = created
                    except Exception:
                        dt_val = None
                if dt_val:
                    sheet.Cells(row_idx, 5).Value = dt_val.strftime("%m/%d/%Y %H:%M:%S")
                else:
                    sheet.Cells(row_idx, 5).Value = ""

                # Categories marks (starting at column F = 6)
                raw_cat = r.get("Categories") or r.get("categories") or ""
                row_cats = [c.strip() for c in str(raw_cat).split(",") if c.strip()]
                if categories:
                    for i, cname in enumerate(categories):
                        val = "X" if cname in row_cats else ""
                        sheet.Cells(row_idx, start_col + i).Value = val

                row_idx += 1

            # Apply borders to used range
            try:
                used = sheet.UsedRange
                used.Borders.LineStyle = 1  # xlContinuous
            except Exception:
                pass

            # Save
            downloads = get_nw_downloads_dir()
            downloads.mkdir(parents=True, exist_ok=True)
            safe = sanitize_filename(project_name or f"BSR_{presentation_id}")
            out_path = downloads / f"{safe}.xls"
            workbook.SaveAs(str(out_path))
            workbook.Close(SaveChanges=False)
            logger.info("Excel report saved to: %s", out_path)
            return out_path
        finally:
            try:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if self.excel is not None:
                    self.excel.Quit()
            except Exception:
                pass


# Global instance
bsr_excel_report_generator = BSRExcelReportGenerator()
