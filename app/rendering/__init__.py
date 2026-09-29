"""Модульное ядро рендеринга F-Engineering Launcher.

Золотая 5-ка форматов (DWG, PDF, XLSX/XLS, DOCX, JPG/PNG/TXT) рендерится
внутри окна мгновенно чистыми движками без запуска тяжелых COM-процессов.
Специфические форматы (.mpp, .mov, .mp4, .zip, .rvt и др.) отдают
карточку-заглушку с открытием в нативной программе.

Архитектура:
    dispatcher.get_file_preview()  — единый шлюз (точка входа фронта)
      ├─ engine_dwg_pdf            — 150 DPI PNG через PyMuPDF
      ├─ engine_excel              — XLSX (1 разбор, все вкладки) / XLS (xlrd)
      ├─ engine_word               — DOCX→HTML (python-docx), COM только для .doc
      └─ engine_media              — фото (Pillow + safe EXIF) / TXT (автокодировка)

Все движки — чистые функции + дисковый кэш, без зависимостей от server.py.
"""

from __future__ import annotations

__version__ = "1.0.0"

GOLDEN_EXTENSIONS = frozenset({
    ".dwg", ".pdf",
    ".xlsx", ".xlsm", ".xls",
    ".docx",
    ".jpg", ".jpeg", ".png", ".gif",
    ".txt", ".log",
})

NATIVE_APP_EXTENSIONS = frozenset({
    ".mpp", ".mpt",
    ".mov", ".mp4", ".m4v", ".avi", ".mkv", ".webm", ".mpeg", ".mpg",
    ".zip", ".rar", ".7z",
    ".rvt", ".rfa", ".nwc", ".nwd", ".ifc",
    ".doc", ".rtf", ".odt",
    ".ppt", ".pptx", ".vsd", ".vsdx",
})

PREVIEW_DPI = 150


def _lazy(name):
    import importlib
    return importlib.import_module(f"{__name__}.{name}")


def __getattr__(name):
    if name == "dispatcher":
        return _lazy("dispatcher")
    if name in {"engine_dwg_pdf", "engine_excel", "engine_word", "engine_media"}:
        return _lazy(name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


__all__ = [
    "GOLDEN_EXTENSIONS",
    "NATIVE_APP_EXTENSIONS",
    "PREVIEW_DPI",
    "dispatcher",
    "engine_dwg_pdf",
    "engine_excel",
    "engine_word",
    "engine_media",
]
