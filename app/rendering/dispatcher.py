"""Единый шлюз превью: расширение → чистый движок или карточка нативного открытия.

Фронт дергает один маршрут POST /api/preview {path} и по полю type решает:
встроить превью (pdf/sheet/word_html/photo/text) либо показать заглушку
с кнопкой «Открыть в программе» (native_app → POST /api/open-file).

Движки намеренно не импортируют server.py (нет циклов): пути к кэшу
передаются явно, по умолчанию — runtime/cache/rendering.
"""

from __future__ import annotations

import hashlib
from pathlib import Path

from . import GOLDEN_EXTENSIONS, NATIVE_APP_EXTENSIONS
from . import engine_dwg, engine_dwg_pdf, engine_excel, engine_media, engine_word

RENDER_CACHE_NS = "rendering"


def repo_runtime_dir() -> Path:
    """Найти runtime/ от корня репозитория (зеркалит логику server.py)."""
    p = Path(__file__).resolve()
    if "_internal" in p.parts:
        idx = p.parts.index("_internal")
        return Path(*p.parts[:idx]) / "runtime"
    # app/rendering/dispatcher.py -> parents[2] == корень репозитория
    return p.parents[2] / "runtime"


def cache_root(runtime_dir: Path | None = None) -> Path:
    base = Path(runtime_dir) if runtime_dir else repo_runtime_dir()
    root = base / "cache" / RENDER_CACHE_NS
    root.mkdir(parents=True, exist_ok=True)
    return root


def file_key(path: Path, tag: str) -> str:
    stat = path.stat()
    raw = f"{path.resolve()}|{stat.st_mtime_ns}|{stat.st_size}|{tag}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def native_app_card(path: Path) -> dict:
    """Заглушка для специфических форматов: метаданные + команда открытия."""
    stat = path.stat()
    return {
        "type": "native_app",
        "name": path.name,
        "path": str(path),
        "extension": path.suffix.casefold().lstrip("."),
        "bytes": stat.st_size,
        "openAction": "/api/open-file",
        "openPayload": {"path": str(path), "action": "default"},
        "hint": "Формат открывается в нативной программе (кнопка «Открыть»)",
    }


def get_file_preview(path: str | Path, runtime_dir: Path | None = None,
                     **params) -> dict:
    """Главная точка входа. Всегда возвращает JSON-сериализуемый dict."""
    src = Path(str(path)).expanduser()
    if not src.exists():
        return {"type": "error", "error": f"Файл не найден: {src}"}
    if not src.is_file() or src.name.startswith("~$") or src.name.startswith(".~"):
        return {"type": "error", "error": f"Не превью-файл: {src.name}"}

    ext = src.suffix.casefold()
    root = cache_root(runtime_dir)

    if ext in NATIVE_APP_EXTENSIONS:
        return native_app_card(src)
    if ext not in GOLDEN_EXTENSIONS:
        return {"type": "unsupported", "name": src.name, "extension": ext,
                "fallback": native_app_card(src)}

    if ext == ".pdf":
        key = file_key(src, f"pdf-{engine_dwg_pdf.PREVIEW_DPI}")
        info = engine_dwg_pdf.render_first_page(src, root / "pdf")
        return {"type": "pdf", "name": src.name, "path": str(src),
                "pages": info.get("pages", engine_dwg_pdf.page_count(src)),
                "firstPage": info, "cacheKey": info.get("cacheKey", key)}

    if ext == ".dwg":
        dpi = int(params.get("dpi") or engine_dwg.DEFAULT_DWG_DPI)
        return engine_dwg.get_dwg_preview(src, cache_root_dir=root, dpi=dpi)

    if ext in (".xlsx", ".xlsm"):
        key = file_key(src, "excel-html")
        cache_dir = root / "excel" / "html" / key
        metas = engine_excel.ensure_workbook_rendered(src, cache_dir)
        sheets = [{"index": i, "name": m["name"], **m} for i, m in enumerate(metas)]
        return {"type": "workbook", "name": src.name, "path": str(src),
                "sheets": sheets, "cacheKey": key}

    if ext == ".xls":
        key = file_key(src, "xls-html")
        cache_dir = root / "xls" / "html" / key
        metas = engine_excel.ensure_xls_rendered(src, cache_dir)
        sheets = [{"index": i, "name": m["name"], **m} for i, m in enumerate(metas)]
        return {"type": "workbook", "name": src.name, "path": str(src),
                "engine": "xlrd", "sheets": sheets, "cacheKey": key}

    if ext == ".docx":
        key = file_key(src, "word-html")
        cache_dir = root / "word" / "html" / key
        meta = engine_word.ensure_docx_preview(src, cache_dir)
        return {"type": "word_html", "name": src.name, "path": str(src),
                "cacheKey": key, **meta}

    if ext in (".jpg", ".jpeg", ".png", ".gif"):
        key = file_key(src, "photo")
        dst = root / "photo" / key / f"preview{ext if ext != '.jpeg' else '.jpg'}"
        info = engine_media.photo_preview(src, dst)
        return {"type": "photo", "name": src.name, "path": str(src),
                "cacheKey": key, **info}

    if ext in (".txt", ".log"):
        key = file_key(src, "text")
        dst = root / "text" / key / "preview.html"
        info = engine_media.text_preview(src, dst)
        return {"type": "text", "name": src.name, "path": str(src),
                "cacheKey": key, **info}

    return {"type": "unsupported", "name": src.name, "extension": ext,
            "fallback": native_app_card(src)}
