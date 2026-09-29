"""Растр PDF/чертежей: единый стандарт 150 DPI PNG через PyMuPDF.

Первая страница рендерится сразу, остальные — лениво по запросу.
DWG→PDF конвертация остается в существующем DWG-конвейере (AutoCAD/демон);
этот движок стандартизирует финальную растровую стадию для PDF-файлов
и уже готовых DWG-производных PDF.
"""

from __future__ import annotations

import hashlib
import time
from pathlib import Path

PREVIEW_DPI = 150


def cache_key(source: Path, dpi: int, page: int) -> str:
    stat = source.stat()
    raw = f"{source.resolve()}|{stat.st_mtime_ns}|{stat.st_size}|dpi{dpi}|p{page}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def page_count(pdf_path: Path) -> int:
    """Число страниц PDF без полного рендера."""
    import fitz
    doc = fitz.open(str(pdf_path))
    try:
        return len(doc)
    finally:
        doc.close()


def render_page(pdf_path: Path, page_index: int, out_png: Path,
                dpi: int = PREVIEW_DPI) -> dict:
    """Отрендерить одну страницу в PNG. Возвращает метаданные."""
    import fitz
    t0 = time.perf_counter()
    doc = fitz.open(str(pdf_path))
    try:
        pix = doc[page_index].get_pixmap(
            matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0), alpha=False)
        width, height = pix.width, pix.height
        out_png.parent.mkdir(parents=True, exist_ok=True)
        pix.save(str(out_png))
    finally:
        doc.close()
    ms = (time.perf_counter() - t0) * 1000
    return {
        "path": str(out_png),
        "page": page_index,
        "dpi": dpi,
        "width": width,
        "height": height,
        "bytes": out_png.stat().st_size,
        "renderMs": round(ms, 1),
    }


def render_first_page(pdf_path: Path, cache_dir: Path,
                      dpi: int = PREVIEW_DPI) -> dict:
    """Первая страница сразу (карточка/превью). Результат кэшируется."""
    key = cache_key(pdf_path, dpi, 0)
    out = cache_dir / key / "p1.png"
    if out.exists() and out.stat().st_size > 0:
        try:
            from PIL import Image
            with Image.open(out) as im:
                width, height = im.size
        except Exception:
            width = height = 0
        return {
            "path": str(out), "page": 0, "dpi": dpi,
            "width": width, "height": height,
            "bytes": out.stat().st_size, "renderMs": 0.0,
            "cacheHit": True, "cacheKey": key,
        }
    info = render_page(pdf_path, 0, out, dpi)
    info["cacheHit"] = False
    info["cacheKey"] = key
    info["pages"] = page_count(pdf_path)
    return info


def render_page_lazy(pdf_path: Path, page_index: int, cache_dir: Path,
                     dpi: int = PREVIEW_DPI) -> dict:
    """Ленивый рендер страниц 2..N по запросу фронта."""
    key = cache_key(pdf_path, dpi, page_index)
    out = cache_dir / key / f"p{page_index + 1}.png"
    if out.exists() and out.stat().st_size > 0:
        info = {
            "path": str(out), "page": page_index, "dpi": dpi,
            "bytes": out.stat().st_size, "renderMs": 0.0,
            "cacheHit": True, "cacheKey": key,
        }
        return info
    info = render_page(pdf_path, page_index, out, dpi)
    info["cacheHit"] = False
    info["cacheKey"] = key
    return info
