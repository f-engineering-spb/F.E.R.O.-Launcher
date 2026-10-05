"""Интеграция Word → PDF → PNG: маршрут, повторное использование, блокировки.

Только искусственные файлы и подмена внешнего Word-вызова.
Реальный договор проверяется отдельно локальным тестом (не здесь).
Запуск: C:\\Python314\\python.exe -m unittest tests.test_word_png_integration -v
(из корня проекта, PYTHONPATH=корень проекта)
"""
from __future__ import annotations

import concurrent.futures
import json
import re
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document

from app.backend import server
from app.rendering import engine_word


def _make_docx(path: Path, paras: int = 3) -> Path:
    doc = Document()
    for i in range(paras):
        doc.add_paragraph(f"SYN-PARA-{i}")
    doc.save(str(path))
    return path


def _fake_com_builder(calls: list, delay: float = 0.0, fail=None):
    """Подмена word_to_pdf_via_com: строит настоящий 3-страничный PDF через fitz."""
    import fitz

    def fake(src, dst_pdf, convert_script, timeout=120):
        calls.append(str(src))
        if fail == "error":
            raise RuntimeError("Word не смог конвертировать документ в PDF")
        if fail == "timeout":
            raise subprocess.TimeoutExpired("powershell", timeout)
        if delay:
            time.sleep(delay)
        out = fitz.open()
        for i in range(3):
            page = out.new_page()
            page.insert_text((72, 120), f"SYN-PAGE-{i + 1}")
        out.save(str(dst_pdf))
        out.close()
        return {"pdf": str(dst_pdf), "bytes": dst_pdf.stat().st_size, "convertMs": 1.0}

    return fake


class WordPngIntegration(unittest.TestCase):
    def _patch_dirs(self, tmp: Path):
        word_dir = tmp / "word"
        pdf_dir = tmp / "pdf"
        word_dir.mkdir(parents=True, exist_ok=True)
        pdf_dir.mkdir(parents=True, exist_ok=True)
        script = tmp / "convert_word_to_pdf.ps1"
        script.write_text("# dummy", encoding="utf-8")
        return word_dir, pdf_dir, script

    def test_01_paginated_route_single_pdf_150(self):
        """Маршрут Word: один COM-экспорт → постраничные PNG 150 dpi из того же PDF."""
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = _make_docx(tmp / "route.docx")
            word_dir, pdf_dir, script = self._patch_dirs(tmp)
            calls = []
            with mock.patch.object(server, "WORD_CACHE_DIR", word_dir), \
                 mock.patch.object(server, "PDF_CACHE_DIR", pdf_dir), \
                 mock.patch.object(server, "WORD_CONVERT_SCRIPT", script), \
                 mock.patch.object(engine_word, "word_to_pdf_via_com",
                                   side_effect=_fake_com_builder(calls)):
                payload = server.render_word(src, dpi=150)
            self.assertEqual(payload["pages"], 3)
            self.assertEqual(len(payload["items"]), 3)
            self.assertEqual(len(calls), 1)
            self.assertEqual(payload["dpi"], 150)
            self.assertEqual([it["page"] for it in payload["items"]], [1, 2, 3])
            for item in payload["items"]:
                self.assertTrue(item["url"])

    def test_02_reuse_no_reexport_no_reraster(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = _make_docx(tmp / "reuse.docx")
            word_dir, pdf_dir, script = self._patch_dirs(tmp)
            calls = []
            with mock.patch.object(server, "WORD_CACHE_DIR", word_dir), \
                 mock.patch.object(server, "PDF_CACHE_DIR", pdf_dir), \
                 mock.patch.object(server, "WORD_CONVERT_SCRIPT", script), \
                 mock.patch.object(engine_word, "word_to_pdf_via_com",
                                   side_effect=_fake_com_builder(calls)):
                first = server.render_word(src, dpi=150)
                second = server.render_word(src, dpi=150)
            self.assertFalse(first["convertCacheHit"])
            self.assertTrue(second["convertCacheHit"])
            self.assertEqual(len(calls), 1)
            self.assertTrue(second["cacheHit"])
            self.assertEqual(second["newRenderedPages"], 0)

    def test_03_on_demand_300_only_selected_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = _make_docx(tmp / "sel300.docx")
            word_dir, pdf_dir, script = self._patch_dirs(tmp)
            calls = []
            with mock.patch.object(server, "WORD_CACHE_DIR", word_dir), \
                 mock.patch.object(server, "PDF_CACHE_DIR", pdf_dir), \
                 mock.patch.object(server, "WORD_CONVERT_SCRIPT", script), \
                 mock.patch.object(engine_word, "word_to_pdf_via_com",
                                   side_effect=_fake_com_builder(calls)):
                payload = server.render_word(src, dpi=150)
                page2 = server.render_pdf_page(Path(payload["convertedPdfPath"]),
                                               page=2, dpi=300)
            self.assertEqual(page2["dpi"], 300)
            self.assertEqual(page2["page"], 2)
            only300 = sorted(p.name for p in (pdf_dir / page2["cacheKey"]).glob("page-*.png"))
            self.assertEqual(only300, ["page-2.png"])

    def test_04_source_change_invalidates(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = _make_docx(tmp / "inv.docx")
            word_dir, pdf_dir, script = self._patch_dirs(tmp)
            calls = []
            with mock.patch.object(server, "WORD_CACHE_DIR", word_dir), \
                 mock.patch.object(server, "PDF_CACHE_DIR", pdf_dir), \
                 mock.patch.object(server, "WORD_CONVERT_SCRIPT", script), \
                 mock.patch.object(engine_word, "word_to_pdf_via_com",
                                   side_effect=_fake_com_builder(calls)):
                server.render_word(src, dpi=150)
                later = src.stat().st_mtime_ns + 2_000_000_000
                import os
                os.utime(src, ns=(later, later))
                again = server.render_word(src, dpi=150)
            self.assertEqual(len(calls), 2)
            self.assertFalse(again["convertCacheHit"])

    def test_05_parallel_same_doc_single_export(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = _make_docx(tmp / "par.docx")
            word_dir, pdf_dir, script = self._patch_dirs(tmp)
            calls = []
            with mock.patch.object(server, "WORD_CACHE_DIR", word_dir), \
                 mock.patch.object(server, "PDF_CACHE_DIR", pdf_dir), \
                 mock.patch.object(server, "WORD_CONVERT_SCRIPT", script), \
                 mock.patch.object(engine_word, "word_to_pdf_via_com",
                                   side_effect=_fake_com_builder(calls, delay=0.8)):
                with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
                    results = list(pool.map(lambda _: server.render_word(src, dpi=150),
                                            range(4)))
            self.assertEqual(len(calls), 1)
            self.assertTrue(all(r["pages"] == 3 for r in results))

    def test_06_export_error_surfaced_others_unaffected(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            bad = _make_docx(tmp / "bad.docx")
            good = _make_docx(tmp / "good.docx")
            word_dir, pdf_dir, script = self._patch_dirs(tmp)
            calls = []

            def selective(src, dst_pdf, convert_script, timeout=120):
                calls.append(str(src))
                if Path(src).name == "bad.docx":
                    raise RuntimeError("Word не смог конвертировать документ в PDF")
                return _fake_com_builder([])(src, dst_pdf, convert_script, timeout)

            with mock.patch.object(server, "WORD_CACHE_DIR", word_dir), \
                 mock.patch.object(server, "PDF_CACHE_DIR", pdf_dir), \
                 mock.patch.object(server, "WORD_CONVERT_SCRIPT", script), \
                 mock.patch.object(engine_word, "word_to_pdf_via_com", side_effect=selective):
                with self.assertRaises(RuntimeError):
                    server.render_word(bad, dpi=150)
                with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
                    fut_ok = pool.submit(server.render_word, good, 150)
                    try:
                        pool.submit(server.render_word, bad, 150).result(timeout=60)
                        bad_raised = False
                    except RuntimeError:
                        bad_raised = True
                    good_res = fut_ok.result(timeout=60)
            self.assertTrue(bad_raised)
            self.assertEqual(good_res["pages"], 3)

    def test_07_export_timeout_propagates(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = _make_docx(tmp / "tmo.docx")
            word_dir, pdf_dir, script = self._patch_dirs(tmp)
            with mock.patch.object(server, "WORD_CACHE_DIR", word_dir), \
                 mock.patch.object(server, "PDF_CACHE_DIR", pdf_dir), \
                 mock.patch.object(server, "WORD_CONVERT_SCRIPT", script), \
                 mock.patch.object(engine_word, "word_to_pdf_via_com",
                                   side_effect=_fake_com_builder([], fail="timeout")):
                with self.assertRaises(subprocess.TimeoutExpired):
                    server.render_word(src, dpi=150)

    def test_08_frontend_docx_paginated_route(self):
        """DOCX ведёт на постраничный WORD-путь; быстрый HTML остаётся фолбэком."""
        app_js = (ROOT / "app" / "frontend" / "app.js").read_text(encoding="utf-8")
        self.assertRegex(app_js, r'if\s*\(\s*ext\s*===\s*"DOCX"\s*\)\s*\{\s*await\s+showWordPreviewPaginated\s*\(')
        self.assertIn("async function showWordPreviewPaginated", app_js)
        self.assertIn('previewType: "WORD"', app_js)
        self.assertIn("async function showWordPreviewFast", app_js)
        self.assertIn("Подготовка предпросмотра Word…", app_js)
        self.assertIn("const myEpoch = state.renderEpoch + 1;", app_js)


if __name__ == "__main__":
    unittest.main()
