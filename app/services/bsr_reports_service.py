"""Service for BSR reports data retrieval."""

from __future__ import annotations

from typing import List, Optional, Tuple

from app.config.db import get_connection_scope
from app.models.bsr_reports_models import (
    BSRExcelReportRow,
    BSRProjectClient,
    BSRProjectConcept,
    BSRSlideNote,
)
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class BSRReportsService:
    """Service for retrieving BSR report data."""

    def _fetch_all(self, cursor) -> Tuple[List[str], List[tuple]]:
        if cursor.description is None:
            return [], []
        columns = [c[0] for c in cursor.description]
        rows = cursor.fetchall()
        return columns, rows

    def get_presentation_info(self, presentation_id: int) -> Tuple[Optional[str], Optional[str]]:
        """Get (project_name, display_name) from bsr_Master."""
        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    "SELECT TOP 1 project, displayname FROM [BI_GUIDELINES].[dbo].[bsr_Master] WHERE PresentationId = ?",
                    (presentation_id,),
                )
                row = cursor.fetchone()
                if row:
                    try:
                        cols = [c[0] for c in cursor.description]
                        data = dict(zip(cols, row))
                        return data.get("project"), data.get("displayname")
                    except Exception:
                        return row[0] if len(row) > 0 else None, row[1] if len(row) > 1 else None
                return None, None
        except Exception as e:
            logger.error("Error fetching presentation info for %d: %s", presentation_id, str(e), exc_info=True)
            return None, None

    def get_excel_report_data(self, presentation_id: int) -> List[dict]:
        """Execute bsr_GetExcelReport SP and return list of dict rows."""
        try:
            with get_connection_scope(timeout=60) as cursor:
                # Support either CALL or EXEC depending on server config
                try:
                    cursor.execute("{CALL bsr_GetExcelReport(?)}", (presentation_id,))
                except Exception:
                    cursor.execute("EXEC bsr_GetExcelReport @ProjectID = ?", (presentation_id,))

                columns, rows = self._fetch_all(cursor)
                results: List[dict] = [dict(zip(columns, r)) for r in rows]

                logger.info("bsr_GetExcelReport returned %d rows for %d", len(results), presentation_id)
                if results:
                    logger.info("Columns returned: %s", columns)
                    logger.info("First row sample: %s", results[0])
                return results
        except Exception as e:
            logger.error("Error executing bsr_GetExcelReport for %d: %s", presentation_id, str(e), exc_info=True)
            return []

    def get_slide_notes(self, presentation_id: int) -> List[BSRSlideNote]:
        """Get slide notes and comments."""
        try:
            with get_connection_scope(timeout=60) as cursor:
                cursor.execute(
                    """
                    SELECT D.SlideDescription, PC.Comments
                    FROM BSR_PageComments PC
                    INNER JOIN BSR_Details D
                        ON PC.PresentationId = D.PresentationId
                        AND PC.SlideNumber = D.SlideNumber
                    WHERE PC.PresentationId = ?
                    ORDER BY PC.SlideNumber
                    """,
                    (presentation_id,),
                )
                cols, rows = self._fetch_all(cursor)
                out: List[BSRSlideNote] = []
                for r in rows:
                    rd = dict(zip(cols, r))
                    out.append(
                        BSRSlideNote(
                            slide_description=str(rd.get("SlideDescription", "") or ""),
                            comments=(rd.get("Comments") or None),
                        )
                    )
                return out
        except Exception as e:
            logger.error("Error fetching slide notes for %d: %s", presentation_id, str(e), exc_info=True)
            return []

    def get_project_attributes(self, presentation_id: int) -> List[str]:
        """Get project attributes HTML (cleaned)."""
        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    """
                    SELECT REPLACE(REPLACE(REPLACE(HTML,'\n\n',''),'\n\t',''),'\n','') AS HTML
                    FROM bsr_ProjectConcepts WHERE projectid = ?
                    """,
                    (presentation_id,),
                )
                cols, rows = self._fetch_all(cursor)
                idx = cols.index("HTML") if "HTML" in cols else 0
                return [str(r[idx]) for r in rows if r and r[idx] is not None]
        except Exception as e:
            logger.error("Error fetching project attributes for %d: %s", presentation_id, str(e), exc_info=True)
            return []

    def get_key_concepts(self, presentation_id: int) -> List[BSRProjectConcept]:
        """Get key concepts with HTML."""
        try:
            with get_connection_scope(timeout=30) as cursor:
                cursor.execute(
                    """
                    SELECT concept,
                           REPLACE(REPLACE(REPLACE(HTML,'\n\n',''),'\n\t',''),'\n','') AS HTML
                    FROM bsr_ProjectConcepts WHERE projectid = ?
                    ORDER BY concept
                    """,
                    (presentation_id,),
                )
                cols, rows = self._fetch_all(cursor)
                out: List[BSRProjectConcept] = []
                for r in rows:
                    d = dict(zip(cols, r))
                    out.append(
                        BSRProjectConcept(
                            concept=d.get("concept"),
                            html=str(d.get("HTML", "") or ""),
                        )
                    )
                return out
        except Exception as e:
            logger.error("Error fetching key concepts for %d: %s", presentation_id, str(e), exc_info=True)
            return []

    def get_client_name(self, project_name: str) -> str:
        """Get client name from DayMaster.Projects by project_name."""
        try:
            if not project_name:
                return ""
            with get_connection_scope(timeout=30, use_daymaster=True) as cursor:
                # Query DayMaster Projects table
                try:
                    cursor.execute(
                        "SELECT TOP 1 client FROM dbo.Projects WHERE ProjectName = ?",
                        (project_name,),
                    )
                except Exception:
                    cursor.execute(
                        "SELECT TOP 1 client FROM Projects WHERE ProjectName = ?",
                        (project_name,),
                    )
                row = cursor.fetchone()
                if row:
                    try:
                        cols = [c[0] for c in cursor.description]
                        d = dict(zip(cols, row))
                        return str(d.get("client", "") or "")
                    except Exception:
                        return str(row[0]) if row and row[0] is not None else ""
                return ""
        except Exception as e:
            logger.error("Error fetching client name for '%s': %s", project_name, str(e), exc_info=True)
            return ""


# Global instance
bsr_reports_service = BSRReportsService()
