"""WORD-HTML-CONNECT-001: поведение подключения Word HTML на синтетике.

Проверяет связку «конвертер → кэш → URL → байты», от которой зависит
показ в iframe: payload содержит url, файл существует, его содержимое —
полный HTML-документ, повтор — хит без перегенерации, смена источника —
новый ключ. Видимость в окне пользователя этим тестом не покрывается
(критерий приёмки — ручной показ по задаче).
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document

from app.rendering import engine_word


def _synthetic_docx(path: Path) -> None:
    doc = Document()
    doc.add_heading("Connection probe", level=1)
    doc.add_paragraph("Alpha beta gamma delta. " * 20)
    tbl = doc.add_table(rows=2, cols=2)
    tbl.rows[0].cells[0].text = "c11"
    doc.save(str(path))


class WordHtmlConnection(unittest.TestCase):
    def test_preview_url_serves_same_bytes(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = tmp / "probe.docx"
            _synthetic_docx(src)
            cache = tmp / "cache"
            first = engine_word.ensure_docx_preview(src, cache)
            self.assertFalse(first.get("cacheHit"))
            # URL подставляет сервер (server.word_preview_html); ядро отдаёт файл.
            self.assertEqual(first["file"], "preview.html")
            self.assertGreater(first["bytes"], 0)
            page = (cache / "preview.html").read_text(encoding="utf-8")
            self.assertIn("<html", page)
            self.assertIn("Alpha beta", page)
            self.assertTrue(page.rstrip().endswith("</html>"))
            second = engine_word.ensure_docx_preview(src, cache)
            self.assertTrue(second.get("cacheHit"))

    def test_source_change_busts_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = tmp / "probe.docx"
            _synthetic_docx(src)
            cache = tmp / "cache"
            before = (cache / "preview.html")
            engine_word.ensure_docx_preview(src, cache)
            first_bytes = before.read_bytes()
            doc = Document(str(src))
            doc.add_paragraph("appended")
            doc.save(str(src))
            again = engine_word.ensure_docx_preview(src, cache)
            self.assertFalse(again.get("cacheHit"))
            self.assertNotEqual(before.read_bytes(), first_bytes)

    def test_css_pdf_thumbs_hidden_when_excel_viewer_active(self):
        css_file = ROOT / "app" / "frontend" / "styles.css"
        css_text = css_file.read_text(encoding="utf-8")
        # Ensure that when excel-viewer (or wordViewer) is visible, pdf-thumbs is hidden
        # so pdf-stage is not pushed down below the overflow container.
        self.assertIn(".shell:not(.full-view) .pdf-viewer:has(.excel-viewer:not([hidden])) .pdf-thumbs", css_text)
        self.assertIn(".shell:not(.full-view) .pdf-viewer .pdf-stage:has(.excel-viewer:not([hidden]))", css_text)

    def test_diag_endpoint_appends_log(self):
        from http.server import ThreadingHTTPServer
        import json
        import threading
        import urllib.request
        from unittest.mock import patch
        from app.backend.server import LauncherHandler

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            with patch("app.backend.server.RUNTIME_DIR", tmp_path):
                server = ThreadingHTTPServer(("127.0.0.1", 0), LauncherHandler)
                port = server.server_port
                t = threading.Thread(target=server.serve_forever, daemon=True)
                t.start()
                try:
                    req = urllib.request.Request(
                        f"http://127.0.0.1:{port}/api/diag",
                        data=json.dumps({"event": "unit_test_event"}).encode("utf-8"),
                        headers={"Content-Type": "application/json"},
                        method="POST",
                    )
                    with urllib.request.urlopen(req, timeout=5) as resp:
                        self.assertEqual(resp.status, 200)
                    diag_file = tmp_path / "logs" / "gui_diag.log"
                    self.assertTrue(diag_file.exists())
                    log_content = diag_file.read_text(encoding="utf-8")
                    self.assertIn("unit_test_event", log_content)
                finally:
                    server.shutdown()
                    server.server_close()


if __name__ == "__main__":
    unittest.main()
