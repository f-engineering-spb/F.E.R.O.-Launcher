"""Воспроизводимый базовый сценарий Excel: синтетическая книга, без данных клиента.

Проверяет инварианты маршрута (не конкретные значения замеров):
- список листов без загрузки книги;
- первый запрошенный лист готовит всю книгу;
- повтор — чистый дисковый кэш-хит без openpyxl;
- лимит строк применяется с флагом limited в метаданных.
Запуск: C:\\Python314\\python.exe -m unittest tests.test_excel_baseline -v
"""
from __future__ import annotations

import sys
import tempfile
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl

from app.rendering import engine_excel


def _synthetic_book(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "First"
    for r in range(1, 51):
        for c in range(1, 6):
            ws.cell(row=r, column=c, value=r * 10 + c)
    ws2 = wb.create_sheet("Second")
    ws2["A1"] = 1
    ws2.merge_cells("A2:B3")
    big = wb.create_sheet("Big")
    for r in range(1, engine_excel.MAX_ROWS + 101):
        big.cell(row=r, column=1, value=r)
    wb.save(str(path))


class ExcelBaselineScenario(unittest.TestCase):
    def test_baseline_invariants(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = tmp / "syn.xlsx"
            t0 = time.perf_counter()
            _synthetic_book(src)
            build_ms = (time.perf_counter() - t0) * 1000
            cache = tmp / "cache"
            t1 = time.perf_counter()
            names = engine_excel.get_sheet_names(src)
            names_ms = (time.perf_counter() - t1) * 1000
            self.assertEqual(len(names), 3)
            t2 = time.perf_counter()
            metas = engine_excel.ensure_workbook_rendered(src, cache)
            first_ms = (time.perf_counter() - t2) * 1000
            self.assertEqual(len(metas), 3)
            self.assertEqual(len(list(cache.glob("sheet-*-*.html"))), 3)
            big = next(m for m in metas if m["rows"] > engine_excel.MAX_ROWS - 100)
            self.assertTrue(big["limited"])
            t3 = time.perf_counter()
            metas2 = engine_excel.ensure_workbook_rendered(src, cache)
            re_ms = (time.perf_counter() - t3) * 1000
            self.assertEqual(len(metas2), 3)
            print(f"\n[excel-baseline] build={build_ms:.0f}ms names={names_ms:.1f}ms "
                  f"first={first_ms:.0f}ms repeat={re_ms:.1f}ms")


if __name__ == "__main__":
    unittest.main()
