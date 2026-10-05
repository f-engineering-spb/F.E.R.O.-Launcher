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


if __name__ == "__main__":
    unittest.main()
