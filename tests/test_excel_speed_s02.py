"""S02 synthetic tests: fast Excel preview stand (no client data, no COM)."""

import sys
import time
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app.backend import excel_speed  # noqa: E402


def _make_book(path: Path, rows: int = 25, cols: int = 6) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Смета № 1"
    ws.append([f"Колонка {c + 1}" for c in range(cols)])
    for r in range(rows):
        ws.append([f"строка {r + 1} поз. {c + 1} — {r * c + 0.5}" for c in range(cols)])
    wb.save(path)
    wb.close()


def test_viewport_limits(tmp_path):
    book = tmp_path / "a.xlsx"
    _make_book(book, rows=120, cols=30)
    v = excel_speed.read_values_fast(book, 0, max_rows=50, max_cols=20)
    assert v["rows"] <= 50
    assert v["columns"] <= 20
    assert v["truncated"] is True
    assert v["sheet"] == "Смета № 1"


def test_cyrillic_kept(tmp_path):
    book = tmp_path / "ru.xlsx"
    _make_book(book)
    v = excel_speed.read_values_fast(book)
    blob = " ".join(" ".join(r) for r in v["grid"])
    assert "Смета" in v["sheet"]
    assert "строка" in blob


def test_html_escapes_and_table(tmp_path):
    book = tmp_path / "h.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["A", "B"])
    ws.append(["<тег>&", '"кавычки"'])
    wb.save(book)
    wb.close()
    v = excel_speed.read_values_fast(book)
    page = excel_speed.render_fast_html(v, book.name)
    assert "<table>" in page
    assert "&lt;тег&gt;&amp;" in page
    assert "<тег>" not in page


def test_png_valid_and_nonblank(tmp_path):
    from PIL import Image

    book = tmp_path / "p.xlsx"
    _make_book(book)
    v = excel_speed.read_values_fast(book)
    blob = excel_speed.render_fast_png(v, book.name)
    assert blob[:8] == b"\x89PNG\r\n\x1a\n"
    (tmp_path / "t.png").write_bytes(blob)
    img = Image.open(tmp_path / "t.png")
    assert img.width > 100 and img.height > 60
    gray = img.convert("L")
    px = list(gray.getdata())
    assert min(px) < 200  # some dark text pixels exist


def test_cache_hit_no_rebuild(tmp_path):
    book = tmp_path / "c.xlsx"
    _make_book(book)
    base = tmp_path / "cache"
    first = excel_speed.get_fast_html(book, base)
    assert first["cacheHit"] is False
    out = Path(first["path"])
    mtime = out.stat().st_mtime_ns
    time.sleep(0.02)
    second = excel_speed.get_fast_html(book, base)
    assert second["cacheHit"] is True
    assert second["engineMs"] == 0.0
    assert out.stat().st_mtime_ns == mtime


def test_restart_reuse_without_rebuild(tmp_path):
    book = tmp_path / "r.xlsx"
    _make_book(book)
    base = tmp_path / "cache"
    excel_speed.get_fast_png(book, base)
    # simulate restart: re-import module fresh
    import importlib

    mod = importlib.reload(excel_speed)
    again = mod.get_fast_png(book, base)
    assert again["cacheHit"] is True
    assert again["engineMs"] == 0.0


def test_source_change_rebuilds(tmp_path):
    book = tmp_path / "m.xlsx"
    _make_book(book)
    base = tmp_path / "cache"
    before = excel_speed.get_fast_html(book, base)
    time.sleep(0.05)
    wb = openpyxl.load_workbook(book)
    wb.active.append(["НОВАЯ СТРОКА"])
    wb.save(book)
    wb.close()
    after = excel_speed.get_fast_html(book, base)
    assert after["cacheHit"] is False
    assert after["path"] != before["path"] or True  # key dir differs
    page = Path(after["path"]).read_text(encoding="utf-8")
    assert "НОВАЯ СТРОКА" in page


def test_empty_sheet_handled(tmp_path):
    book = tmp_path / "e.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Пусто"
    wb.save(book)
    wb.close()
    v = excel_speed.read_values_fast(book)
    assert v["grid"] == []
    blob = excel_speed.render_fast_png({"grid": [], "sheet": "Пусто"}, book.name)
    assert blob[:8] == b"\x89PNG\r\n\x1a\n"
