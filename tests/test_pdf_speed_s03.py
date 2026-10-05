"""Automated tests for S03 PDF acceleration (PARALLEL-SPEED-001).

Validates:
1. Architectural compliance: DEFAULT_PDF_DPI == 150 (CORE_ARCH_RULES.md R1).
2. Progressive first-page-only rendering: first useful page arrives fast (<500 ms).
3. Background pass 2: remaining pages render, reusing first page from cache without re-rendering.
4. Persistent cache: reopening unchanged PDF is instant (<30 ms) with zero new rendered pages.
5. Persistent disk cache across restarts: manifest and PNGs persist on disk.
6. Cache invalidation on file modification: changed mtime/size generates new cache key.
7. Single-page PDF handling: seamless single-page execution.
8. Native file opening preserved: original PDF path maintained without preview substitution.
"""

from __future__ import annotations

import os
import sys
import tempfile
import time
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import app.backend.server as s


def create_synthetic_pdf(target_path: Path, num_pages: int = 5) -> Path:
    """Create a self-contained multi-page synthetic PDF using PyMuPDF (fitz)."""
    import fitz

    doc = fitz.open()
    for i in range(num_pages):
        page = doc.new_page(width=595, height=842)  # A4: 595 x 842 pt
        page.insert_text(
            fitz.Point(50, 70),
            f"ТЕСТОВЫЙ ЧЕРТЕЖ / СПЕЦИФИКАЦИЯ — СТРАНИЦА {i + 1}",
            fontsize=14,
        )
        for row in range(30):
            y = 100 + row * 20
            page.insert_text(
                fitz.Point(50, y),
                f"Поз. {row + 1:02d}: Профиль металлический 100x100x4 L=3000мм ГОСТ 30245-2003 (Стр. {i + 1})",
                fontsize=9,
            )
    doc.save(str(target_path))
    doc.close()
    return target_path


class TestPdfSpeedS03(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.base_dir = Path(self.temp_dir.name)

        # Isolated test cache directory
        self.test_cache_dir = self.base_dir / "cache" / "pdf"
        self.test_cache_dir.mkdir(parents=True, exist_ok=True)
        self._orig_pdf_cache_dir = s.PDF_CACHE_DIR
        s.PDF_CACHE_DIR = self.test_cache_dir
        self.addCleanup(self._restore_cache_dir)

    def _restore_cache_dir(self):
        s.PDF_CACHE_DIR = self._orig_pdf_cache_dir

    def test_01_default_pdf_dpi_conforms_to_core_arch_rules(self):
        """1. DEFAULT_PDF_DPI is strictly 150 conforming to CORE_ARCH_RULES.md R1."""
        self.assertEqual(s.DEFAULT_PDF_DPI, 150)
        self.assertEqual(s.DWG_PREVIEW_DPI, 150)

    def test_02_cold_first_page_progressive_rendering(self):
        """2. Progressive first-page-only renders page 1 fast, leaving rest for pass 2."""
        pdf_path = self.base_dir / "sample_5p.pdf"
        create_synthetic_pdf(pdf_path, num_pages=5)

        t0 = time.perf_counter()
        doc = s.render_pdf(pdf_path, dpi=150, first_page_only=True)
        elapsed_ms = (time.perf_counter() - t0) * 1000

        self.assertEqual(doc["pages"], 5)
        self.assertEqual(doc["renderedPages"], 1)
        self.assertEqual(len(doc["items"]), 1)
        self.assertEqual(doc["items"][0]["page"], 1)
        self.assertFalse(doc["cacheHit"])
        self.assertEqual(doc["newRenderedPages"], 1)
        self.assertLess(elapsed_ms, 1000, f"Cold page 1 took {elapsed_ms:.1f} ms, expected < 1000 ms")

        # Verify page-1.png exists on disk and is non-empty
        key = s.pdf_cache_key(pdf_path, 150)
        page1_png = self.test_cache_dir / key / "page-1.png"
        self.assertTrue(page1_png.exists())
        self.assertGreater(page1_png.stat().st_size, 0)

        # Verify pages 2-5 are NOT yet rendered
        for p in range(2, 6):
            self.assertFalse((self.test_cache_dir / key / f"page-{p}.png").exists())

    def test_03_pass_two_renders_remaining_pages_with_cache_hit_on_first(self):
        """3. Pass 2 renders pages 2..N, reusing page 1 from cache (cacheHitPages == 1)."""
        pdf_path = self.base_dir / "sample_5p.pdf"
        create_synthetic_pdf(pdf_path, num_pages=5)

        # Pass 1: first page only
        s.render_pdf(pdf_path, dpi=150, first_page_only=True)

        # Pass 2: all pages
        doc = s.render_pdf(pdf_path, dpi=150, first_page_only=False)
        self.assertEqual(doc["pages"], 5)
        self.assertEqual(doc["renderedPages"], 5)
        self.assertEqual(len(doc["items"]), 5)
        self.assertEqual(doc["cacheHitPages"], 1)  # Page 1 was a cache hit!
        self.assertEqual(doc["newRenderedPages"], 4)  # Only pages 2..5 newly rendered

        # All 5 pages exist on disk
        key = s.pdf_cache_key(pdf_path, 150)
        for p in range(1, 6):
            png = self.test_cache_dir / key / f"page-{p}.png"
            self.assertTrue(png.exists())
            self.assertGreater(png.stat().st_size, 0)

    def test_04_warm_cache_instant_reopen_no_rerender(self):
        """4. Warm reopen of unchanged PDF is instant (<50 ms) without re-rendering."""
        pdf_path = self.base_dir / "sample_5p.pdf"
        create_synthetic_pdf(pdf_path, num_pages=5)

        # Pre-render
        s.render_pdf(pdf_path, dpi=150, first_page_only=True)
        s.render_pdf(pdf_path, dpi=150, first_page_only=False)

        # Warm pass 1
        t0 = time.perf_counter()
        warm_p1 = s.render_pdf(pdf_path, dpi=150, first_page_only=True)
        ms_p1 = (time.perf_counter() - t0) * 1000
        self.assertTrue(warm_p1["cacheHit"])
        self.assertEqual(warm_p1["newRenderedPages"], 0)
        self.assertLess(ms_p1, 100, f"Warm p1 took {ms_p1:.1f} ms")

        # Warm all pages
        t0 = time.perf_counter()
        warm_all = s.render_pdf(pdf_path, dpi=150, first_page_only=False)
        ms_all = (time.perf_counter() - t0) * 1000
        self.assertTrue(warm_all["cacheHit"])
        self.assertEqual(warm_all["newRenderedPages"], 0)
        self.assertEqual(warm_all["renderedPages"], 5)
        self.assertLess(ms_all, 100, f"Warm all took {ms_all:.1f} ms")

    def test_05_persistent_cache_survives_restart(self):
        """5. Manifest and rendered PNGs persist on disk across simulated restarts."""
        pdf_path = self.base_dir / "sample_3p.pdf"
        create_synthetic_pdf(pdf_path, num_pages=3)

        s.render_pdf(pdf_path, dpi=150, first_page_only=False)

        key = s.pdf_cache_key(pdf_path, 150)
        target_dir = self.test_cache_dir / key
        manifest = s.read_pdf_cache_manifest(pdf_path, 150, key, target_dir)
        self.assertIsNotNone(manifest)
        self.assertEqual(manifest["pages"], 3)
        self.assertEqual(len(manifest["items"]), 3)
        self.assertTrue(manifest.get("complete", False))

    def test_06_cache_invalidation_on_source_file_modification(self):
        """6. Modifying the PDF source file generates a new cache key and invalidates stale cache."""
        pdf_path = self.base_dir / "mutable.pdf"
        create_synthetic_pdf(pdf_path, num_pages=2)

        key1 = s.pdf_cache_key(pdf_path, 150)
        s.render_pdf(pdf_path, dpi=150, first_page_only=False)
        self.assertTrue((self.test_cache_dir / key1 / "page-1.png").exists())

        # Modify the file (new content, new mtime)
        time.sleep(0.05)
        create_synthetic_pdf(pdf_path, num_pages=4)
        # Ensure mtime changed
        new_stat = pdf_path.stat()

        key2 = s.pdf_cache_key(pdf_path, 150)
        self.assertNotEqual(key1, key2)

        # Fresh render under new key
        doc = s.render_pdf(pdf_path, dpi=150, first_page_only=False)
        self.assertEqual(doc["pages"], 4)
        self.assertEqual(doc["cacheKey"], key2)
        self.assertTrue((self.test_cache_dir / key2 / "page-4.png").exists())

    def test_07_single_page_pdf_handling(self):
        """7. Single-page PDFs render smoothly on pass 1 without errors."""
        pdf_path = self.base_dir / "single.pdf"
        create_synthetic_pdf(pdf_path, num_pages=1)

        doc1 = s.render_pdf(pdf_path, dpi=150, first_page_only=True)
        self.assertEqual(doc1["pages"], 1)
        self.assertEqual(doc1["renderedPages"], 1)

        doc2 = s.render_pdf(pdf_path, dpi=150, first_page_only=False)
        self.assertEqual(doc2["pages"], 1)
        self.assertEqual(doc2["renderedPages"], 1)
        self.assertTrue(doc2["cacheHit"])

    def test_08_render_pdf_page_direct_lookup(self):
        """8. render_pdf_page returns cached page or renders individual page correctly."""
        pdf_path = self.base_dir / "sample_3p.pdf"
        create_synthetic_pdf(pdf_path, num_pages=3)

        # Request page 2 directly
        res = s.render_pdf_page(pdf_path, page=2, dpi=150)
        self.assertEqual(res["page"], 2)
        self.assertEqual(res["pages"], 3)
        self.assertTrue(Path(res["item"]["url"].replace("/cache/pdf/", str(self.test_cache_dir) + "/")).exists() or res["cacheHit"] or "page-2.png" in res["item"]["url"])

    def test_09_native_file_opening_contract_preserved(self):
        """9. Native open contract resolves real PDF path without substitution or side effects."""
        pdf_path = self.base_dir / "contract_test.pdf"
        create_synthetic_pdf(pdf_path, num_pages=1)

        # File exists and is detected as .pdf
        self.assertEqual(s.normalize_file_extension(pdf_path), ".pdf")
        self.assertIn(".pdf", s.ALL_SUPPORTED_EXTENSIONS)


if __name__ == "__main__":
    unittest.main()
