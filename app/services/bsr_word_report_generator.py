"""Service for generating BSR Word reports using COM automation."""

from __future__ import annotations

import time
from datetime import datetime
from pathlib import Path
from typing import List

import win32com.client  # type: ignore
import win32clipboard  # type: ignore

from app.services.bsr_reports_service import bsr_reports_service
from app.utils.logging_utils import get_logger
from app.utils.nw_data_utils import sanitize_filename
from app.utils.path_utils import get_nw_downloads_dir
from app.config.settings import settings

logger = get_logger(__name__)

WD_FORMAT_DOCUMENT = 0  # .doc


class BSRWordReportGenerator:
    """Generate BSR Word reports from templates."""

    def __init__(self) -> None:
        self.word = None

    def _get_template_path(self) -> Path:
        return settings.app_dir / "templates" / "BSR_TEMPLATE" / "bsr_word_template.docx"

    def generate_word_report(self, presentation_id: int, project_name: str) -> Path:
        """Generate BSR Word report."""
        template = self._get_template_path()
        if not template.exists():
            raise FileNotFoundError(f"Word template not found: {template}")

        # Data
        notes = bsr_reports_service.get_slide_notes(presentation_id)
        attrs = bsr_reports_service.get_project_attributes(presentation_id)
        concepts = bsr_reports_service.get_key_concepts(presentation_id)
        client = bsr_reports_service.get_client_name(project_name)

        header_text = f"{client} {project_name} {datetime.now().strftime('%m/%d/%Y')}".strip()

        # Build HTML blocks
        notes_html = self._build_notes_html(notes)
        attrs_html = "<br/>".join(attrs) if attrs else ""
        keycon_html = self._build_keycon_html(concepts)

        # Word automation
        doc = None
        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = 0

            doc = self.word.Documents.Open(str(template))

            # Replace bookmarks
            self._replace_bookmark_with_text(doc, "HEADER", header_text)
            if notes_html:
                self._replace_bookmark_with_html(doc, "NOTES", notes_html)
            if attrs_html:
                self._replace_bookmark_with_html(doc, "ATTR", attrs_html)
            if keycon_html:
                self._replace_bookmark_with_html(doc, "KEYCON", keycon_html)

            # Delete remaining bookmarks
            self._cleanup_bookmarks(doc)

            # Save
            downloads = get_nw_downloads_dir()
            downloads.mkdir(parents=True, exist_ok=True)
            safe = sanitize_filename(project_name or f"BSR_{presentation_id}")
            out_path = downloads / f"{safe}.doc"
            doc.SaveAs2(str(out_path), FileFormat=WD_FORMAT_DOCUMENT)
            doc.Close()
            logger.info("Word report saved to: %s", out_path)
            return out_path
        finally:
            try:
                if doc is not None:
                    doc.Close()
            except Exception:
                pass
            try:
                if self.word is not None:
                    self.word.Quit()
            except Exception:
                pass

    def _build_notes_html(self, notes: List) -> str:
        parts: List[str] = []
        for n in notes:
            sd = (n.slide_description or "").strip()
            cm = (n.comments or "").strip()
            if sd:
                parts.append(f"<h3>{sd}</h3>")
            if cm:
                parts.append(f"<p>{cm}</p>")
        return "".join(parts)

    def _build_keycon_html(self, concepts: List) -> str:
        parts: List[str] = []
        for c in concepts:
            title = (c.concept or "").strip()
            html = (c.html or "").strip()
            if title:
                parts.append(f"<h3>{title}</h3>")
            if html:
                parts.append(html)
        return "".join(parts)

    def _replace_bookmark_with_text(self, doc, bookmark_name: str, text: str) -> None:
        try:
            if doc.Bookmarks.Exists(bookmark_name):
                bm = doc.Bookmarks(bookmark_name)
                bm.Range.Text = text
        except Exception as e:
            logger.warning("Failed replacing text bookmark %s: %s", bookmark_name, str(e))

    def _replace_bookmark_with_html(self, doc, bookmark_name: str, html_content: str):
        """Replace bookmark with HTML content using clipboard."""
        try:
            if not doc.Bookmarks.Exists(bookmark_name):
                return
            rng = doc.Bookmarks(bookmark_name).Range
            self._copy_html_to_clipboard(html_content)
            time.sleep(1)
            rng.Paste()
        except Exception as e:
            logger.warning("Paste HTML failed for bookmark %s: %s; falling back to text", bookmark_name, str(e))
            try:
                rng = doc.Bookmarks(bookmark_name).Range
                rng.Text = html_content
            except Exception:
                pass

    def _build_cf_html(self, html: str) -> bytes:
        # Build CF_HTML per spec
        html_bytes = html.encode("utf-8")
        prefix = (
            "Version:0.9\r\n" \
            "StartHTML:00000097\r\n" \
            "EndHTML:{end_html:08d}\r\n" \
            "StartFragment:00000131\r\n" \
            "EndFragment:{end_frag:08d}\r\n"
        )
        wrapper_start = b"<html><body><!--StartFragment-->"
        wrapper_end = b"<!--EndFragment--></body></html>"
        full_html = wrapper_start + html_bytes + wrapper_end
        end_html = 97 + len("EndHTML:00000000\r\nStartFragment:00000131\r\nEndFragment:00000000\r\n") + len(full_html)
        end_frag = 131 + len(html_bytes)
        header = prefix.format(end_html=end_html, end_frag=end_frag).encode("ascii")
        return header + full_html

    def _copy_html_to_clipboard(self, html: str) -> None:
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                fmt = win32clipboard.RegisterClipboardFormat("HTML Format")
                cf_html = self._build_cf_html(html)
                win32clipboard.SetClipboardData(fmt, cf_html)
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            logger.warning("Error setting clipboard HTML: %s", str(e))

    def _cleanup_bookmarks(self, doc) -> None:
        try:
            count = doc.Bookmarks.Count
            for i in range(count, 0, -1):
                try:
                    doc.Bookmarks(i).Delete()
                except Exception:
                    pass
        except Exception as e:
            logger.warning("Error cleaning bookmarks: %s", str(e))


# Global instance
bsr_word_report_generator = BSRWordReportGenerator()
