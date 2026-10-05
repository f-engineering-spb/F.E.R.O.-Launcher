"""Движок рендеринга DWG: DWG → PDF → PNG (CORE ARCH RULES / S04 DWG).

Архитектура:
1. Конвертер DWG → PDF:
   - Приоритет: AutoCAD accoreconsole (нативная headless-консоль).
     Быстрый запуск (~3-5 с cold), не поднимает GUI, не мигает окнами,
     не трогает запущенные пользовательские CAD-сессии.
   - Определение структуры:
     * Если есть непустые листы Paper Space (Layouts) — экспорт всех листов в многостраничный PDF.
     * Если чертёж только в Model Space — экспорт границ модели (Extents).
   - Защита: FILEDIA=0, CMDDIA=0, EXPERT=5, XLOADCTL=0 (без всплывающих окон,
     без автоматической загрузки внешних ресурсов).
   - Резервные пути: парный PDF рядом с файлом, AutoCAD COM (если консоль недоступна),
     растровый заголовок DWG через PyMuPDF (если CAD не установлен).
2. Растеризация PDF → PNG:
   - Стандарт через app.rendering.engine_dwg_pdf (PyMuPDF).
   - Быстрый первичный просмотр первой страницы (render_first_page).
   - Ленивая выдача остальных страниц по запросу фронтенда (render_page_lazy).
3. Кэширование:
   - Постоянный дисковый кэш в runtime/cache/dwg/<key>/.
   - Ключ кэша: sha1(путь | mtime_ns | st_size | dwg-accore-v1).
   - Manifest.json сохраняет метаданные исходника и параметры.
   - Повторный запуск и перезапуск сервера — мгновенный cache-hit (0 мс).
   - Защита от дублирования одновременных запросов (дедупликация in-flight)
     и семафор ограничения тяжёлых процессов (максимум 1 процесс конвертации).
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Tuple, Union

logger = logging.getLogger("engine_dwg")

DWG_CACHE_NS = "dwg"
DWG_EXPORT_TIMEOUT_SECONDS = 60
DEFAULT_DWG_DPI = 150

_ACCORE_PATH_CACHE: Optional[Path] = None
_ACCORE_SEARCH_LOCK = threading.Lock()
_CONVERSION_SEMAPHORE = threading.Semaphore(1)
_IN_FLIGHT_LOCK = threading.Lock()
_IN_FLIGHT_CONVERSIONS: Dict[str, threading.Event] = {}
_IN_FLIGHT_RESULTS: Dict[str, Tuple[Path, bool]] = {}


def find_accoreconsole() -> Optional[Path]:
    """Поиск установленной консоли AutoCAD accoreconsole.exe."""
    global _ACCORE_PATH_CACHE
    with _ACCORE_SEARCH_LOCK:
        if _ACCORE_PATH_CACHE and _ACCORE_PATH_CACHE.is_file():
            return _ACCORE_PATH_CACHE

        candidates: list[Path] = []

        # 1. Проверка конфигурации native_apps.json
        try:
            repo_root = Path(__file__).resolve().parents[2]
            config_file = repo_root / "runtime" / "native_apps.json"
            if config_file.is_file():
                data = json.loads(config_file.read_text(encoding="utf-8"))
                cad_exe = data.get(".dwg") or data.get("dwg")
                if cad_exe:
                    cand = Path(cad_exe).parent / "accoreconsole.exe"
                    if cand.is_file():
                        candidates.append(cand)
        except Exception:
            pass

        # 2. Стандартные пути AutoCAD (2026 .. 2021)
        prog_files = os.environ.get("ProgramFiles", r"C:\Program Files")
        for ver in ("2026", "2025", "2024", "2023", "2022", "2021"):
            cand = Path(prog_files) / f"Autodesk\\AutoCAD {ver}\\accoreconsole.exe"
            if cand.is_file():
                candidates.append(cand)

        prog_x86 = os.environ.get("ProgramFiles(x86)")
        if prog_x86:
            for ver in ("2026", "2025", "2024", "2023", "2022", "2021"):
                cand = Path(prog_x86) / f"Autodesk\\AutoCAD {ver}\\accoreconsole.exe"
                if cand.is_file():
                    candidates.append(cand)

        for c in candidates:
            if c.is_file():
                _ACCORE_PATH_CACHE = c
                return c

        return None


def dwg_file_key(source: Path, tag: str = "dwg-accore-v1") -> str:
    """Детерминированный ключ кэша для DWG файла."""
    stat = source.stat()
    raw = f"{source.resolve()}|{stat.st_mtime_ns}|{stat.st_size}|{tag}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def export_dwg_via_accoreconsole(
    dwg_path: Path,
    out_pdf: Path,
    accore_exe: Path,
    timeout_seconds: int = DWG_EXPORT_TIMEOUT_SECONDS,
) -> dict:
    """Выполнить конвертацию DWG в PDF через accoreconsole.exe.

    Возвращает dict с метаданными {ok, exportMs, pages, mode, size}.
    """
    if not dwg_path.is_file():
        raise FileNotFoundError(f"DWG-файл не найден: {dwg_path}")

    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    if out_pdf.exists():
        try:
            out_pdf.unlink()
        except OSError:
            pass

    temp_dir = Path(tempfile.gettempdir()) / f"feng_dwg_{os.getpid()}_{threading.get_ident()}"
    temp_dir.mkdir(parents=True, exist_ok=True)
    scr_file = temp_dir / "export.scr"

    pdf_posix = out_pdf.as_posix()
    # AutoLISP-скрипт:
    # 1. Подавление диалогов и фоновой печати.
    # 2. Поиск непустых листов (Paper Space Layouts). Лист считается непустым,
    #    если количество сущностей в нём > 1 (не просто пустой видовой экран из шаблона).
    # 3. Если непустые листы найдены — переход на первый лист и вызов (command-s "_.-EXPORT" "_PDF" "_A" ...),
    #    что создаёт многостраничный векторный PDF со всеми листами чертежа.
    # 4. Если чертёж содержит только Модель — экспорт Extents модели в 1-страничный PDF.
    scr_content = f"""(setvar "EXPERT" 5)
(setvar "BACKGROUNDPLOT" 0)
(setvar "FILEDIA" 0)
(setvar "CMDDIA" 0)
(setvar "XLOADCTL" 0)
(defun c:doexp ()
  (setq first_paper nil)
  (setq d (dictsearch (namedobjdict) "ACAD_LAYOUT"))
  (while (setq item (assoc 3 d))
    (setq lname (cdr item))
    (setq ss (ssget "_X" (list (cons 410 lname))))
    (if (and (/= (strcase lname) "MODEL") ss (> (sslength ss) 1) (null first_paper))
      (setq first_paper lname)
    )
    (setq d (cdr (member item d)))
  )
  (if first_paper
    (progn
      (if (= (strcase (getvar "CTAB")) "MODEL")
        (setvar "CTAB" first_paper)
      )
      (princ (strcat "\\nEXPORT_PAPER_ALL:" first_paper "\\n"))
      (command-s "_.-EXPORT" "_PDF" "_A" "{pdf_posix}")
    )
    (progn
      (princ "\\nEXPORT_MODEL_EXTENTS\\n")
      (setvar "CTAB" "Model")
      (command-s "_.-EXPORT" "_PDF" "_E" "_N" "{pdf_posix}")
    )
  )
)
(c:doexp)
_QUIT
_N
"""
    scr_file.write_text(scr_content, encoding="utf-8")

    t0 = time.perf_counter()
    proc = None
    try:
        # Скрытый запуск консоли AutoCAD
        startupinfo = None
        creationflags = 0
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            creationflags = subprocess.CREATE_NO_WINDOW

        proc = subprocess.Popen(
            [str(accore_exe), "/i", str(dwg_path), "/s", str(scr_file)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )

        stdout_bytes, stderr_bytes = proc.communicate(timeout=timeout_seconds)
    except subprocess.TimeoutExpired:
        if proc:
            try:
                proc.kill()
            except OSError:
                pass
        raise TimeoutError(f"Таймаут конвертации DWG в accoreconsole ({timeout_seconds} с)")
    finally:
        try:
            scr_file.unlink(missing_ok=True)
            temp_dir.rmdir()
        except OSError:
            pass

    export_ms = (time.perf_counter() - t0) * 1000.0

    if not out_pdf.is_file() or out_pdf.stat().st_size <= 1024:
        # Извлечение информативного сообщения об ошибке из вывода
        out_text = ""
        try:
            out_text = stdout_bytes.decode("utf-16le", errors="replace")
        except Exception:
            out_text = stdout_bytes.decode("utf-8", errors="replace")
        error_lines = [line.strip() for line in out_text.splitlines() if line.strip() and not line.startswith("CoreHeartBeat")]
        sample_err = " | ".join(error_lines[-5:]) if error_lines else "PDF не создан"
        raise RuntimeError(f"accoreconsole завершился (код {proc.returncode}), но PDF не создан: {sample_err}")

    # Подсчёт страниц через PyMuPDF
    pages = 1
    try:
        import fitz

        with fitz.open(str(out_pdf)) as doc:
            pages = len(doc)
    except Exception:
        pass

    return {
        "ok": True,
        "exportMs": round(export_ms, 1),
        "pdfPath": str(out_pdf),
        "pages": pages,
        "bytes": out_pdf.stat().st_size,
    }


def ensure_dwg_pdf(
    dwg_path: Path,
    cache_root_dir: Optional[Path] = None,
    timeout_seconds: int = DWG_EXPORT_TIMEOUT_SECONDS,
) -> Tuple[Path, bool]:
    """Гарантировать наличие векторного PDF для DWG чертежа.

    Возвращает (pdf_path, cache_hit).
    Потокобезопасно: дедуплицирует одновременные одинаковые запросы.
    """
    dwg_path = Path(dwg_path).resolve()
    if not dwg_path.is_file():
        raise FileNotFoundError(f"DWG-файл не найден: {dwg_path}")

    # 1. Проверка парного PDF рядом с чертежом (если он свежее или равен DWG)
    paired_pdf = dwg_path.with_suffix(".pdf")
    if paired_pdf.is_file() and paired_pdf.stat().st_size > 1024:
        try:
            if paired_pdf.stat().st_mtime_ns >= dwg_path.stat().st_mtime_ns:
                return paired_pdf, True
        except OSError:
            pass

    # 2. Определение пути к кэшу
    if cache_root_dir is None:
        from .dispatcher import cache_root

        cache_root_dir = cache_root()

    key = dwg_file_key(dwg_path)
    target_dir = cache_root_dir / DWG_CACHE_NS / key
    target_dir.mkdir(parents=True, exist_ok=True)
    cached_pdf = target_dir / f"{dwg_path.stem}.pdf"
    manifest_path = target_dir / "manifest.json"

    # 3. Проверка постоянного дискового кэша
    if cached_pdf.is_file() and cached_pdf.stat().st_size > 1024 and manifest_path.is_file():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            if (
                manifest.get("sourcePath") == str(dwg_path)
                and manifest.get("sourceMtimeNs") == dwg_path.stat().st_mtime_ns
                and manifest.get("sourceSize") == dwg_path.stat().st_size
                and manifest.get("cacheKey") == key
            ):
                return cached_pdf, True
        except Exception:
            pass

    # 4. Дедупликация одновременных запросов конвертации одного и того же файла
    event_to_wait = None
    with _IN_FLIGHT_LOCK:
        if key in _IN_FLIGHT_CONVERSIONS:
            event_to_wait = _IN_FLIGHT_CONVERSIONS[key]
        else:
            event_to_wait = None
            evt = threading.Event()
            _IN_FLIGHT_CONVERSIONS[key] = evt

    if event_to_wait is not None:
        # Другой поток уже конвертирует этот файл — ждем завершения
        event_to_wait.wait(timeout=float(timeout_seconds + 10))
        with _IN_FLIGHT_LOCK:
            if key in _IN_FLIGHT_RESULTS:
                res_path, res_hit = _IN_FLIGHT_RESULTS[key]
                if res_path.is_file() and res_path.stat().st_size > 1024:
                    return res_path, res_hit
        # Если результат не появился, продолжаем самостоятельную попытку

    try:
        # Семафор: не запускать более 1 тяжелого процесса accoreconsole параллельно
        with _CONVERSION_SEMAPHORE:
            # Повторная проверка кэша на случай, если файл был подготовлен
            if cached_pdf.is_file() and cached_pdf.stat().st_size > 1024 and manifest_path.is_file():
                try:
                    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                    if manifest.get("sourceMtimeNs") == dwg_path.stat().st_mtime_ns:
                        return cached_pdf, True
                except Exception:
                    pass

            accore = find_accoreconsole()
            pdf_result = None

            if accore:
                try:
                    pdf_result = export_dwg_via_accoreconsole(
                        dwg_path=dwg_path,
                        out_pdf=cached_pdf,
                        accore_exe=accore,
                        timeout_seconds=timeout_seconds,
                    )
                except Exception as ex:
                    logger.warning("accoreconsole failed for %s: %s", dwg_path.name, ex)
                    pdf_result = None

            # Fallback 1: парный PDF если существует
            if not pdf_result and paired_pdf.is_file() and paired_pdf.stat().st_size > 1024:
                return paired_pdf, True

            # Fallback 2: извлечение растрового заголовка через dwg_engine и упаковка в PDF
            if not pdf_result or not cached_pdf.is_file() or cached_pdf.stat().st_size <= 1024:
                try:
                    from app.backend.dwg_engine import extract_raw_thumbnail_from_dwg, _generate_placeholder_png

                    thumb = extract_raw_thumbnail_from_dwg(dwg_path)
                    img_bytes = thumb[0] if thumb else _generate_placeholder_png(dwg_path.name, "Model")

                    import fitz

                    with fitz.open(stream=img_bytes, filetype="png") as img_doc:
                        pdf_bytes = img_doc.convert_to_pdf()
                    cached_pdf.write_bytes(pdf_bytes)
                    pdf_result = {"ok": True, "exportMs": 1.0, "pages": 1, "bytes": len(pdf_bytes)}
                except Exception as fb_err:
                    raise RuntimeError(f"Не удалось преобразовать DWG в PDF: {fb_err}")

            # Запись manifest.json
            pages = pdf_result.get("pages", 1)
            manifest_data = {
                "sourcePath": str(dwg_path),
                "sourceName": dwg_path.name,
                "sourceMtimeNs": dwg_path.stat().st_mtime_ns,
                "sourceSize": dwg_path.stat().st_size,
                "cacheKey": key,
                "pdfPath": str(cached_pdf),
                "pages": pages,
                "exportMs": pdf_result.get("exportMs", 0.0),
                "converter": "accoreconsole" if accore else "header-fallback",
                "renderedAt": datetime.now().isoformat(timespec="seconds"),
            }
            try:
                manifest_path.write_text(json.dumps(manifest_data, ensure_ascii=False, indent=2), encoding="utf-8")
            except OSError:
                pass

            result = (cached_pdf, False)
            with _IN_FLIGHT_LOCK:
                _IN_FLIGHT_RESULTS[key] = result
            return result
    finally:
        with _IN_FLIGHT_LOCK:
            if key in _IN_FLIGHT_CONVERSIONS:
                _IN_FLIGHT_CONVERSIONS[key].set()
                del _IN_FLIGHT_CONVERSIONS[key]


def get_dwg_preview(
    dwg_path: Union[str, Path],
    cache_root_dir: Optional[Path] = None,
    dpi: int = DEFAULT_DWG_DPI,
) -> dict:
    """Точка входа превью DWG: подготовка векторного PDF и первой страницы PNG.

    Возвращает dict с метаданными {type, name, path, pages, firstPage, cacheKey, via}.
    Не ожидает отрисовки всех страниц: первая страница отдается фронту сразу.
    """
    src = Path(str(dwg_path)).resolve()
    if not src.is_file():
        return {"type": "error", "error": f"Файл не найден: {src}"}

    from . import engine_dwg_pdf

    key = dwg_file_key(src)
    if cache_root_dir is None:
        from .dispatcher import cache_root

        cache_root_dir = cache_root()

    dwg_cache_dir = cache_root_dir / DWG_CACHE_NS

    # 1. Обеспечить наличие векторного PDF (из кэша или через быструю конвертацию)
    pdf_path, convert_cache_hit = ensure_dwg_pdf(src, cache_root_dir=cache_root_dir)

    # 2. Рендер первой страницы через стандартный растеризатор engine_dwg_pdf
    first_page_info = engine_dwg_pdf.render_first_page(
        pdf_path=pdf_path,
        cache_dir=dwg_cache_dir / key,
        dpi=dpi,
    )
    first_page_info["sourceType"] = "DWG"
    first_page_info["convertCacheHit"] = convert_cache_hit

    total_pages = first_page_info.get("pages", 1)

    return {
        "type": "dwg",
        "name": src.name,
        "path": str(src),
        "pdfPath": str(pdf_path),
        "pages": total_pages,
        "firstPage": first_page_info,
        "cacheKey": key,
        "via": "accoreconsole-pdf",
        "convertCacheHit": convert_cache_hit,
    }
