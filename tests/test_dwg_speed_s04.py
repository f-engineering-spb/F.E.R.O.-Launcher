"""Тесты S04 DWG: DWG → PDF → PNG (accoreconsole, кэш, дедупликация, нативное открытие)."""

import os
import shutil
import tempfile
import threading
import time
import unittest
from pathlib import Path

from app.rendering import dispatcher, engine_dwg, engine_dwg_pdf
from app.backend.server import render_dwg_model, dwg_to_model_pdf

SAMPLE_A01 = Path(r"C:\Program Files\Autodesk\AutoCAD 2024\Sample\Sheet Sets\Architectural\A-01.dwg")
SAMPLE_MODEL = Path(r"C:\Program Files\Autodesk\AutoCAD 2024\Sample\ru-RU\Dynamic Blocks\Architectural - Metric.dwg")
SAMPLE_MULTI = Path(r"C:\Program Files\Autodesk\AutoCAD 2024\Sample\Sheet Sets\Civil\Site Grading Plan.dwg")


class TestDwgEngine(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_dwg_s04_"))

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_01_find_accoreconsole(self):
        """Проверка обнаружения консоли AutoCAD accoreconsole.exe."""
        accore = engine_dwg.find_accoreconsole()
        self.assertIsNotNone(accore, "accoreconsole.exe должен быть найден на машине с установленным AutoCAD")
        self.assertTrue(accore.is_file(), f"Файл {accore} должен существовать")
        self.assertTrue(accore.name.lower().endswith("accoreconsole.exe"))

    def test_02_dwg_cache_key_deterministic(self):
        """Детерминированность ключа кэша и обновление при изменении источника."""
        if not SAMPLE_A01.is_file():
            self.skipTest("Sample A-01.dwg не найден")

        key1 = engine_dwg.dwg_file_key(SAMPLE_A01)
        key2 = engine_dwg.dwg_file_key(SAMPLE_A01)
        self.assertEqual(key1, key2, "Ключ одного и того же файла должен быть идентичен")

        # Создаем временную копию
        temp_copy = self.temp_dir / "test_copy.dwg"
        shutil.copy2(SAMPLE_A01, temp_copy)
        key_copy1 = engine_dwg.dwg_file_key(temp_copy)

        # Модифицируем mtime
        new_time = time.time() + 100
        os.utime(temp_copy, (new_time, new_time))
        key_copy2 = engine_dwg.dwg_file_key(temp_copy)
        self.assertNotEqual(key_copy1, key_copy2, "Ключ кэша обязан измениться при изменении mtime файла")

    def test_03_export_dwg_paper_layout(self):
        """Конвертация DWG с листами Paper Space (A-01.dwg) через accoreconsole."""
        if not SAMPLE_A01.is_file():
            self.skipTest("Sample A-01.dwg не найден")

        pdf_path, hit = engine_dwg.ensure_dwg_pdf(SAMPLE_A01, cache_root_dir=self.temp_dir)
        self.assertTrue(pdf_path.is_file(), "PDF должен быть создан")
        self.assertGreater(pdf_path.stat().st_size, 1024, "Размер PDF должен быть больше 1 КБ")
        self.assertFalse(hit, "Первый вызов должен быть холодным (hit=False)")

        # Проверка повторного вызова (warm cache)
        t0 = time.perf_counter()
        pdf_path_warm, hit_warm = engine_dwg.ensure_dwg_pdf(SAMPLE_A01, cache_root_dir=self.temp_dir)
        warm_ms = (time.perf_counter() - t0) * 1000.0
        self.assertTrue(hit_warm, "Второй вызов должен быть из кэша (hit=True)")
        self.assertEqual(pdf_path, pdf_path_warm)
        self.assertLess(warm_ms, 50.0, f"Время кэш-хита должно быть < 50 мс, получено: {warm_ms:.2f} мс")

    def test_04_export_dwg_model_space(self):
        """Конвертация DWG только с Model Space (Architectural - Metric.dwg)."""
        if not SAMPLE_MODEL.is_file():
            self.skipTest("Sample Architectural - Metric.dwg не найден")

        pdf_path, hit = engine_dwg.ensure_dwg_pdf(SAMPLE_MODEL, cache_root_dir=self.temp_dir)
        self.assertTrue(pdf_path.is_file(), "PDF для пространства модели должен быть создан")
        self.assertGreater(pdf_path.stat().st_size, 1024)

    def test_05_get_dwg_preview_fast_first_page(self):
        """Быстрый первый просмотр первой страницы через dispatcher и engine_dwg."""
        if not SAMPLE_A01.is_file():
            self.skipTest("Sample A-01.dwg не найден")

        prev = engine_dwg.get_dwg_preview(SAMPLE_A01, cache_root_dir=self.temp_dir, dpi=96)
        self.assertEqual(prev.get("type"), "dwg")
        self.assertIn("firstPage", prev)
        fp = prev["firstPage"]
        self.assertTrue(Path(fp["path"]).is_file(), "PNG первой страницы должен существовать")
        self.assertGreater(fp["width"], 1000, "Разрешение листа должно быть читаемым (> 1000px)")
        self.assertGreater(fp["height"], 1000)

        # Проверка повторного вызова превью (дисковый кэш PNG)
        t0 = time.perf_counter()
        prev_warm = engine_dwg.get_dwg_preview(SAMPLE_A01, cache_root_dir=self.temp_dir, dpi=96)
        warm_prev_ms = (time.perf_counter() - t0) * 1000.0
        self.assertTrue(prev_warm["firstPage"]["cacheHit"], "Повторный вызов превью должен иметь cacheHit=True")
        self.assertLess(warm_prev_ms, 100.0, f"Warm preview должен отдаваться мгновенно, получено {warm_prev_ms:.2f} мс")

    def test_06_in_flight_deduplication(self):
        """Дедупликация одновременных запросов одного и того же чертежа."""
        if not SAMPLE_A01.is_file():
            self.skipTest("Sample A-01.dwg не найден")

        temp_copy = self.temp_dir / "dedup_test.dwg"
        shutil.copy2(SAMPLE_A01, temp_copy)

        results = []
        errors = []

        def worker():
            try:
                p, h = engine_dwg.ensure_dwg_pdf(temp_copy, cache_root_dir=self.temp_dir)
                results.append((p, h))
            except Exception as e:
                errors.append(e)

        t1 = threading.Thread(target=worker)
        t2 = threading.Thread(target=worker)
        t1.start()
        t2.start()
        t1.join()
        t2.join()

        self.assertEqual(len(errors), 0, f"Ошибок быть не должно: {errors}")
        self.assertEqual(len(results), 2, "Оба потока должны вернуть результат")
        self.assertEqual(results[0][0], results[1][0], "Оба потока должны вернуть один и тот же PDF")

    def test_07_dispatcher_dwg_route(self):
        """Проверка маршрута POST /api/preview через dispatcher."""
        if not SAMPLE_A01.is_file():
            self.skipTest("Sample A-01.dwg не найден")

        payload = dispatcher.get_file_preview(SAMPLE_A01, runtime_dir=self.temp_dir)
        self.assertEqual(payload.get("type"), "dwg")
        self.assertEqual(payload.get("name"), SAMPLE_A01.name)
        self.assertIn("firstPage", payload)
        self.assertIn("pdfPath", payload)

    def test_08_server_render_dwg_model_progressive(self):
        """Проверка server.render_dwg_model с флагом first_page_only."""
        if not SAMPLE_A01.is_file():
            self.skipTest("Sample A-01.dwg не найден")

        doc_first = render_dwg_model(SAMPLE_A01, dpi=96, first_page_only=True)
        self.assertEqual(doc_first.get("sourceType"), "DWG")
        self.assertEqual(len(doc_first.get("items", [])), 1, "При first_page_only=True должен отдаваться только 1 лист")

        doc_all = render_dwg_model(SAMPLE_A01, dpi=96, first_page_only=False)
        self.assertEqual(doc_all.get("sourceType"), "DWG")
        self.assertGreaterEqual(len(doc_all.get("items", [])), 1)


if __name__ == "__main__":
    unittest.main()
