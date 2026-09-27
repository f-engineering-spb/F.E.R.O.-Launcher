from __future__ import annotations

import argparse
import atexit
import hashlib
import html
import io
import json
import mimetypes
import os
import re
import shlex
import shutil
import struct
import subprocess
import sys
import threading
import time
import traceback
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse
import xml.etree.ElementTree as ET
import zipfile

try:
    import openpyxl
except ImportError:  # pragma: no cover - reported through the local API
    openpyxl = None


if getattr(sys, 'frozen', False):
    REPO_ROOT = Path(sys.executable).resolve().parent
else:
    # Проверяем, не лежит ли файл внутри _internal
    p = Path(__file__).resolve()
    if "_internal" in p.parts:
        idx = p.parts.index("_internal")
        REPO_ROOT = Path(*p.parts[:idx])
    else:
        REPO_ROOT = p.parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))
_backend_dir = str(Path(__file__).resolve().parent)
if _backend_dir not in sys.path:
    sys.path.insert(0, _backend_dir)
RUNTIME_DIR = REPO_ROOT / "runtime"
MANIFESTS_DIR = RUNTIME_DIR / "manifests"
FRONTEND_DIR = REPO_ROOT / "app" / "frontend"
PDF_CACHE_DIR = RUNTIME_DIR / "cache" / "pdf"
WORD_CACHE_DIR = RUNTIME_DIR / "cache" / "word"
EXCEL_CACHE_DIR = RUNTIME_DIR / "cache" / "excel"
DWG_CACHE_DIR = RUNTIME_DIR / "cache" / "dwg"
WORD_CONVERT_SCRIPT = REPO_ROOT / "scripts" / "convert_word_to_pdf.ps1"
EXCEL_CONVERT_SCRIPT = REPO_ROOT / "scripts" / "convert_excel_to_pdf.ps1"
EXCEL_XLS_CONVERT_SCRIPT = REPO_ROOT / "scripts" / "convert_xls_to_xlsx.ps1"
DWG_RENDER_SCRIPT = REPO_ROOT / "scripts" / "render_dwg_model_space.ps1"
DWG_SMART_RENDER_SCRIPT = REPO_ROOT / "scripts" / "render_dwg_smart.ps1"
VERSION = "0.4.0-v3-pdf-render"
APP_VERSION = "3.1.0"
APP_BRANCH = "dvg-main"
SKIP_DIR_NAMES = {".git", "__pycache__", "node_modules", ".venv", "venv"}


def get_launcher_version_info() -> dict:
    """Вернуть {version, branch} из app.shared.version_info или version.json.

    Не падает при отсутствии файла: возвращает значения по умолчанию
    для ветки dvg-main.
    """
    try:
        from app.shared.version_info import get_version_payload

        payload = get_version_payload()
        if isinstance(payload, dict):
            version = str(payload.get("version") or APP_VERSION)
            branch = str(payload.get("git_branch") or payload.get("branch") or APP_BRANCH)
            return {"version": version, "branch": branch}
    except Exception:
        pass
    try:
        manifest_path = REPO_ROOT / "version.json"
        with manifest_path.open("r", encoding="utf-8") as manifest_file:
            payload = json.load(manifest_file)
        if isinstance(payload, dict):
            version = str(payload.get("version") or APP_VERSION)
            branch = str(payload.get("git_branch") or payload.get("branch") or APP_BRANCH)
            return {"version": version, "branch": branch}
    except Exception:
        pass
    return {"version": APP_VERSION, "branch": APP_BRANCH}
DEFAULT_PDF_DPI = 300
PDF_PAGE_TIMEOUT_SECONDS = 25
PDF_DOCUMENT_TIMEOUT_SECONDS = 600
WORD_CONVERT_TIMEOUT_SECONDS = 120
EXCEL_CONVERT_TIMEOUT_SECONDS = 180
DWG_RENDER_TIMEOUT_SECONDS = 600
DWG_MODEL_PAGE_TIMEOUT_SECONDS = 120
DWG_DAEMON_SCRIPT = REPO_ROOT / "scripts" / "render_dwg_daemon.ps1"
DWG_DAEMON_DIR = DWG_CACHE_DIR / "daemon"
DWG_DAEMON_START_TIMEOUT_SECONDS = 240
DWG_DAEMON_IDLE_MINUTES = 15
MAX_XLSX_ROWS = 2000
MAX_XLSX_COLS = 100
POPPLER_BIN_DIR = (
    Path.home()
    / ".cache"
    / "codex-runtimes"
    / "codex-primary-runtime"
    / "dependencies"
    / "native"
    / "poppler"
    / "Library"
    / "bin"
)
# Portable Poppler shipped inside a release archive.  Checked before the
# machine-wide cache path so a copied release works on a fresh computer.
PORTABLE_POPPLER_BIN_DIR = REPO_ROOT / "runtime" / "tools" / "poppler" / "Library" / "bin"


def object_id_for_path(path: Path) -> str:
    return hashlib.sha1(str(path.resolve()).casefold().encode("utf-8")).hexdigest()[:16]


def file_extension(path: Path) -> str:
    suffix = path.suffix.upper().lstrip(".")
    return suffix or "NO_EXT"


def build_tree(folder: Path) -> tuple[dict, dict[str, int], int, int]:
    counts: dict[str, int] = {}
    folder_count = 0
    file_count = 0
    visited_dirs: set[Path] = set()

    # Итеративный обход: стек хранит (path, parent_children_list).
    # Каждый каталог создаёт узел и кладёт своих потомков-каталогов в стек.
    root_node: dict = {"type": "folder", "name": folder.name, "path": str(folder), "children": []}
    stack: list[tuple[Path, list]] = [(folder, root_node["children"])]

    while stack:
        current, parent_children = stack.pop()
        folder_count += 1
        try:
            resolved_current = current.resolve()
            if resolved_current in visited_dirs:
                return {
                    "type": "folder",
                    "name": current.name,
                    "path": str(current),
                    "children": [],
                }
            visited_dirs.add(resolved_current)
            entries = sorted(current.iterdir(), key=lambda item: (not item.is_dir(), item.name.casefold()))
        except OSError as error:
            # Ошибка чтения каталога — добавляем узел с ошибкой
            if current != folder:
                parent_children.append({
                    "type": "folder",
                    "name": current.name,
                    "path": str(current),
                    "error": str(error),
                    "children": [],
                })
            else:
                root_node["error"] = str(error)
            continue

        # Собираем дочерние узлы для текущего каталога
        children_dirs_to_push: list[tuple[Path, list]] = []
        for entry in entries:
            if entry.name.startswith("~$") or entry.name.startswith(".~"):
                continue
            if entry.is_dir():
                if entry.name in SKIP_DIR_NAMES:
                    continue
                child_node: dict = {"type": "folder", "name": entry.name, "path": str(entry), "children": []}
                parent_children.append(child_node)
                children_dirs_to_push.append((entry, child_node["children"]))
            elif entry.is_file():
                ext = file_extension(entry)
                counts[ext] = counts.get(ext, 0) + 1
                file_count += 1
                try:
                    stat = entry.stat()
                    size = stat.st_size
                    mtime_ns = stat.st_mtime_ns
                    stat_error = None
                except OSError as stat_err:
                    size = 0
                    mtime_ns = 0
                    stat_error = str(stat_err)
                node = {
                    "type": "file",
                    "name": entry.name,
                    "path": str(entry),
                    "extension": ext,
                    "size": size,
                    "mtimeNs": mtime_ns,
                }
                if stat_error:
                    node["error"] = stat_error
                parent_children.append(node)

        # Добавляем подкаталоги в стек в обратном порядке для сохранения
        # порядка обхода (первый каталог будет обработан следующим)
        for item in reversed(children_dirs_to_push):
            stack.append(item)

    return root_node, counts, folder_count, file_count


def tree_file_signatures(tree: dict | None) -> dict[str, dict]:
    result: dict[str, dict] = {}
    if not tree:
        return result
    stack = [tree]
    while stack:
        node = stack.pop()
        if node.get("type") == "file":
            path = node.get("path") or ""
            result[path.casefold()] = {
                "path": path,
                "name": node.get("name"),
                "size": node.get("size"),
                "mtimeNs": node.get("mtimeNs"),
            }
        for child in node.get("children") or []:
            stack.append(child)
    return result


def diff_trees(old_tree: dict | None, new_tree: dict | None) -> dict:
    old = tree_file_signatures(old_tree)
    new = tree_file_signatures(new_tree)

    def signature(entry: dict) -> tuple:
        return (entry.get("size"), entry.get("mtimeNs"))

    added: list[str] = []
    changed: list[str] = []
    removed: list[str] = []
    unchanged: list[str] = []

    for key, entry in new.items():
        old_entry = old.get(key)
        if old_entry is None:
            added.append(entry["path"])
        elif signature(old_entry) != signature(entry):
            changed.append(entry["path"])
        else:
            unchanged.append(entry["path"])

    for key, entry in old.items():
        if key not in new:
            removed.append(entry["path"])

    # Эвристика переименования: если исчез ровно один файл и появился ровно
    # один новый с тем же размером, считаем это переносом/заменой имени,
    # а не «удалён + добавлен». Срабатывает только для однозначной пары,
    # чтобы не объединять случайные файлы одинакового размера.
    renamed: set[str] = set()
    if len(removed) == 1 and len(added) == 1:
        removed_path = removed[0]
        added_path = added[0]
        removed_entry = old.get(removed_path.casefold())
        added_entry = new.get(added_path.casefold())
        if (
            removed_entry is not None
            and added_entry is not None
            and removed_entry.get("size") == added_entry.get("size")
        ):
            renamed.add(added_path.casefold())
            renamed.add(removed_path.casefold())
            changed.append(added_path)
    if renamed:
        added = [path for path in added if path.casefold() not in renamed]
        removed = [path for path in removed if path.casefold() not in renamed]

    return {
        "added": sorted(added),
        "changed": sorted(changed),
        "removed": sorted(removed),
        "unchanged": sorted(unchanged),
    }


def scan_object(raw_path: str) -> dict:
    if not raw_path or not raw_path.strip():
        raise ValueError("Путь к папке объекта пустой")
    root = Path(raw_path.strip().strip('"')).expanduser()
    if not root.exists():
        raise FileNotFoundError(f"Путь не найден: {root}")
    if root.is_file():
        root = root.parent
    if not root.is_dir():
        raise NotADirectoryError(f"Это не папка: {root}")

    object_id = object_id_for_path(root)
    previous = load_manifest(object_id)
    previous_tree = (previous or {}).get("tree")

    tree, extension_counts, folder_count, file_count = build_tree(root)
    last_diff = diff_trees(previous_tree, tree)
    manifest = {
        "id": object_id,
        "name": root.name,
        "rootPath": str(root.resolve()),
        "scannedAt": datetime.now().isoformat(timespec="seconds"),
        "statistics": {
            "folders": folder_count,
            "files": file_count,
            "extensions": extension_counts,
        },
        "tree": tree,
        "lastDiff": last_diff,
    }
    MANIFESTS_DIR.mkdir(parents=True, exist_ok=True)
    (MANIFESTS_DIR / f"{manifest['id']}.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return manifest


def manifest_path(object_id: str) -> Path:
    return MANIFESTS_DIR / f"{object_id}.json"


def validate_and_filter_tree(node: dict) -> dict:
    if not isinstance(node, dict):
        return node
    # Итеративная валидация: обрабатываем все папки через стек
    stack = [node]
    while stack:
        current = stack.pop()
        if not isinstance(current, dict):
            continue
        cur_path = current.get("path") or ""
        if not os.path.exists(cur_path):
            current["is_missing"] = True
            current["error"] = "Путь недоступен на диске"
            if current.get("type") == "folder":
                current["children"] = []
            continue

        if current.get("type") == "folder":
            valid_children = []
            for child in current.get("children", []):
                if not isinstance(child, dict):
                    continue
                c_path = child.get("path") or ""
                if not os.path.exists(c_path):
                    continue
                valid_children.append(child)
                if child.get("type") == "folder":
                    stack.append(child)
            current["children"] = valid_children
    return node


def calculate_tree_statistics(tree: dict) -> dict:
    counts: dict[str, int] = {}
    folder_count = 0
    file_count = 0

    stack = [tree]
    while stack:
        n = stack.pop()
        if not isinstance(n, dict):
            continue
        if n.get("type") == "folder":
            folder_count += 1
            for child in n.get("children", []):
                stack.append(child)
        elif n.get("type") == "file":
            file_count += 1
            ext = (n.get("extension") or file_extension(Path(n.get("path", "")))).upper()
            counts[ext] = counts.get(ext, 0) + 1

    return {
        "folders": folder_count,
        "files": file_count,
        "extensions": counts,
    }


def load_manifest(object_id: str) -> dict | None:
    path = manifest_path(object_id)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def manifest_summary(manifest: dict) -> dict:
    return {
        "id": manifest.get("id"),
        "name": manifest.get("name"),
        "rootPath": manifest.get("rootPath"),
        "scannedAt": manifest.get("scannedAt"),
        "statistics": manifest.get("statistics", {}),
    }


def list_object_summaries() -> list[dict]:
    MANIFESTS_DIR.mkdir(parents=True, exist_ok=True)
    manifests: list[dict] = []
    for path in MANIFESTS_DIR.glob("*.json"):
        try:
            manifest = json.loads(path.read_text(encoding="utf-8"))
            manifests.append(manifest_summary(manifest))
        except (OSError, json.JSONDecodeError):
            continue
    return sorted(manifests, key=lambda item: str(item.get("scannedAt", "")), reverse=True)


def pdf_cache_key(path: Path, dpi: int) -> str:
    stat = path.stat()
    raw = f"{path.resolve()}|{stat.st_mtime_ns}|{stat.st_size}|{dpi}"
    return hashlib.sha1(raw.casefold().encode("utf-8")).hexdigest()[:20]


def file_cache_key(path: Path, purpose: str) -> str:
    stat = path.stat()
    raw = f"{purpose}|{path.resolve()}|{stat.st_mtime_ns}|{stat.st_size}"
    return hashlib.sha1(raw.casefold().encode("utf-8")).hexdigest()[:20]


def is_word_file(path: Path) -> bool:
    if path.name.startswith("~$") or path.name.startswith(".~"):
        return False
    return path.suffix.casefold() in {".doc", ".docx"}


def is_excel_file(path: Path) -> bool:
    if path.name.startswith("~$") or path.name.startswith(".~"):
        return False
    return path.suffix.casefold() in {".xls", ".xlsx", ".xlsm"}


def pdf_page_item(path: Path, key: str, page: int, png: Path) -> dict:
    return {
        "page": page,
        "name": f"{path.name} · стр. {page}",
        "url": f"/cache/pdf/{key}/{png.name}",
        "bytes": png.stat().st_size,
    }


def read_pdf_cache_manifest(path: Path, dpi: int, key: str, target_dir: Path) -> dict | None:
    manifest_path = target_dir / "manifest.json"
    manifest = None
    if manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            manifest = None
        if manifest and (manifest.get("path") != str(path) or manifest.get("dpi") != dpi or manifest.get("cacheKey") != key):
            manifest = None
    if not manifest:
        legacy_items = []
        for png in sorted(target_dir.glob("page-*.png")):
            try:
                page = int(png.stem.removeprefix("page-"))
            except ValueError:
                continue
            if page > 0 and png.stat().st_size > 0:
                legacy_items.append(pdf_page_item(path, key, page, png))
        if not legacy_items:
            return None
        return {
            "name": path.name,
            "path": str(path),
            "dpi": dpi,
            "pages": max(item["page"] for item in legacy_items),
            "renderedPages": len(legacy_items),
            "cacheKey": key,
            "complete": False,
            "errors": [],
            "items": legacy_items,
        }
    items = []
    for item in manifest.get("items", []):
        page = int(item.get("page") or 0)
        png = target_dir / f"page-{page}.png"
        if page > 0 and png.exists() and png.stat().st_size > 0:
            items.append(pdf_page_item(path, key, page, png))
    if not items:
        return None
    manifest["items"] = items
    return manifest


def write_pdf_cache_manifest(
    path: Path,
    dpi: int,
    key: str,
    target_dir: Path,
    page_count: int,
    items: list[dict],
    errors: list[dict],
) -> None:
    manifest = {
        "name": path.name,
        "path": str(path),
        "dpi": dpi,
        "pages": page_count,
        "renderedPages": len(items),
        "cacheKey": key,
        "complete": len(items) == page_count and not errors,
        "errors": errors,
        "items": items,
        "updatedAt": datetime.now().isoformat(timespec="seconds"),
    }
    (target_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def poppler_tool(name: str) -> str | None:
    for candidate in (PORTABLE_POPPLER_BIN_DIR, POPPLER_BIN_DIR):
        exe = candidate / f"{name}.exe"
        if exe.exists():
            return str(exe)
    found = shutil.which(name)
    if found and found.lower().endswith(".cmd"):
        return None
    return found


def hidden_process_kwargs() -> dict:
    """Флаги скрытого запуска фоновых процессов: никаких чёрных окон консоли."""
    if os.name != "nt":
        return {}
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        return {
            "startupinfo": startupinfo,
            "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
        }
    except Exception:
        return {}


def decode_folder_dialog_output(raw: bytes | None) -> str:
    """Последняя непустая строка stdout диалога выбора папки -> путь.

    Декодируем терпимо: сначала UTF-8 (скрипт выставляет
    [Console]::OutputEncoding в UTF-8 без BOM), затем OEM/ANSI консоли,
    чтобы кириллические пути не падали с UnicodeDecodeError
    (было: errors="strict" -> тихая ошибка выбора папки).
    Из строк выбираем последнюю существующую на диске: в stdout могут
    попасть предупреждения PowerShell, а путь скрипт пишет последним.
    """
    if not raw:
        return ""
    text: str | None = None
    for encoding in ("utf-8", "cp866", "cp1251"):
        try:
            text = raw.decode(encoding)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if text is None:
        text = raw.decode("utf-8", errors="replace")
    lines = [line.strip() for line in text.strip().splitlines() if line.strip()]
    for line in reversed(lines):
        try:
            if Path(line).exists():
                return line
        except (OSError, ValueError):
            continue
    return lines[-1] if lines else ""


def run_poppler(args: list[str], timeout: int) -> subprocess.CompletedProcess[str]:
    process = subprocess.Popen(
        args,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        **hidden_process_kwargs(),
    )
    try:
        stdout, stderr = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as error:
        process.kill()
        stdout, stderr = process.communicate()
        raise subprocess.TimeoutExpired(args, timeout, output=stdout, stderr=stderr) from error
    return subprocess.CompletedProcess(args=args, returncode=process.returncode, stdout=stdout, stderr=stderr)


def pdf_page_count(path: Path) -> int:
    try:
        import fitz
        doc = fitz.open(str(path))
        count = len(doc)
        doc.close()
        return count
    except Exception:
        pass

    pdfinfo = poppler_tool("pdfinfo")
    if not pdfinfo:
        raise RuntimeError("pdfinfo не найден. Нужен Poppler из runtime.")
    result = run_poppler(
        [pdfinfo, str(path)],
        timeout=60,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "pdfinfo не смог прочитать PDF")
    for line in result.stdout.splitlines():
        if line.startswith("Pages:"):
            return int(line.split(":", 1)[1].strip())
    raise RuntimeError("pdfinfo не вернул количество страниц PDF")


def word_to_pdf(path: Path) -> tuple[Path, bool]:
    if not path.exists():
        raise FileNotFoundError(f"Word-файл не найден: {path}")
    if not path.is_file() or not is_word_file(path):
        raise ValueError(f"Это не Word-файл: {path}")
    if not WORD_CONVERT_SCRIPT.exists():
        raise RuntimeError(f"Скрипт конвертации Word не найден: {WORD_CONVERT_SCRIPT}")

    key = file_cache_key(path, "word-pdf")
    target_dir = WORD_CACHE_DIR / key
    target_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = target_dir / f"{path.stem}.pdf"
    manifest_path = target_dir / "manifest.json"

    cached = False
    if pdf_path.exists() and pdf_path.stat().st_size > 0 and manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            cached = (
                manifest.get("sourcePath") == str(path)
                and manifest.get("cacheKey") == key
                and manifest.get("sourceMtimeNs") == path.stat().st_mtime_ns
                and manifest.get("sourceSize") == path.stat().st_size
            )
        except (OSError, json.JSONDecodeError):
            cached = False

    if cached:
        return pdf_path, True

    if pdf_path.exists():
        pdf_path.unlink()

    process = subprocess.run(
        [
            "powershell",
            "-STA",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(WORD_CONVERT_SCRIPT),
            "-InputPath",
            str(path),
            "-OutputPath",
            str(pdf_path),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=WORD_CONVERT_TIMEOUT_SECONDS,
        **hidden_process_kwargs(),
    )
    if process.returncode != 0:
        message = process.stderr.strip() or process.stdout.strip() or "Word не смог конвертировать документ в PDF"
        raise RuntimeError(message)
    if not pdf_path.exists() or pdf_path.stat().st_size <= 0:
        raise RuntimeError("Word не создал PDF для preview")

    manifest_path.write_text(
        json.dumps(
            {
                "sourcePath": str(path),
                "sourceName": path.name,
                "sourceMtimeNs": path.stat().st_mtime_ns,
                "sourceSize": path.stat().st_size,
                "cacheKey": key,
                "pdfPath": str(pdf_path),
                "convertedAt": datetime.now().isoformat(timespec="seconds"),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return pdf_path, False


def render_word(path: Path, dpi: int = DEFAULT_PDF_DPI, first_page_only: bool = False) -> dict:
    if not path.exists():
        return {
            "name": path.name,
            "path": str(path),
            "sourcePath": str(path),
            "sourceName": path.name,
            "sourceType": file_extension(path),
            "dpi": dpi,
            "pages": 0,
            "renderedPages": 0,
            "cacheKey": "",
            "cacheHit": False,
            "cacheHitPages": 0,
            "newRenderedPages": 0,
            "is_missing": True,
            "errors": [{"error": f"Документ не найден на диске: {path.name}", "is_missing": True}],
            "items": [],
        }
    pdf_path, convert_cache_hit = word_to_pdf(path)
    document = render_pdf(pdf_path, dpi=dpi, first_page_only=first_page_only)
    document["sourcePath"] = str(path)
    document["sourceName"] = path.name
    document["sourceType"] = file_extension(path)
    document["convertedPdfPath"] = str(pdf_path)
    document["convertCacheHit"] = convert_cache_hit
    return document


_DWG_DAEMON_LOCK = threading.Lock()
_DWG_DAEMON_PROC: subprocess.Popen | None = None


class _DwgFileError(RuntimeError):
    """Файл не конвертируется (демон отработал, CAD отказался). Повтор разово бессмысленен."""


def _dwg_daemon_files() -> tuple[Path, Path, Path]:
    """Каталог заданий демона, файл готовности, лог демона."""
    job_dir = DWG_DAEMON_DIR / "jobs"
    job_dir.mkdir(parents=True, exist_ok=True)
    return job_dir, DWG_DAEMON_DIR / "ready.json", RUNTIME_DIR / "logs" / "dwg-daemon.log"


def _dwg_daemon_alive() -> bool:
    proc = _DWG_DAEMON_PROC
    return proc is not None and proc.poll() is None


def ensure_dwg_daemon() -> bool:
    """Поднять долгоживущий CAD-конвертер (одна сессия на все файлы).

    Возвращает True, если демон готов принимать задания. При любой
    неудаче возвращает False — вызыватель откатится на разовый запуск.
    """
    global _DWG_DAEMON_PROC
    with _DWG_DAEMON_LOCK:
        if _dwg_daemon_alive():
            return True
        _DWG_DAEMON_PROC = None
        if not DWG_DAEMON_SCRIPT.exists():
            return False
        job_dir, ready_file, log_file = _dwg_daemon_files()
        for stale in (ready_file, job_dir / "daemon.fatal"):
            try:
                if stale.exists():
                    stale.unlink()
            except OSError:
                pass
        try:
            _DWG_DAEMON_PROC = subprocess.Popen(
                [
                    "powershell",
                    "-STA",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(DWG_DAEMON_SCRIPT),
                    "-JobDir",
                    str(job_dir),
                    "-ReadyFile",
                    str(ready_file),
                    "-LogFile",
                    str(log_file),
                    "-IdleMinutes",
                    str(DWG_DAEMON_IDLE_MINUTES),
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                **hidden_process_kwargs(),
            )
        except Exception as err:
            _DWG_DAEMON_PROC = None
            return False
        deadline = time.time() + DWG_DAEMON_START_TIMEOUT_SECONDS
        while time.time() < deadline:
            if _DWG_DAEMON_PROC.poll() is not None:
                _DWG_DAEMON_PROC = None
                return False
            if (job_dir / "daemon.fatal").exists():
                try:
                    _DWG_DAEMON_PROC.terminate()
                except Exception:
                    pass
                _DWG_DAEMON_PROC = None
                return False
            if ready_file.exists():
                # Проверяем, что ready от ЭТОГО запуска, а не stale от убитого демона.
                try:
                    ready_pid = int(
                        json.loads(ready_file.read_text(encoding="utf-8")).get("pid") or 0
                    )
                except Exception:
                    ready_pid = 0
                if ready_pid == _DWG_DAEMON_PROC.pid:
                    return True
            time.sleep(0.5)
        try:
            _DWG_DAEMON_PROC.kill()
        except Exception:
            pass
        _DWG_DAEMON_PROC = None
        return False


def kill_dwg_daemon(reason: str) -> None:
    """Прибить демона и только СВОЮ CAD-сессию (PID записан демоном при старте).

    Чужие запущенные AutoCAD (пользовательские) не трогаем: их PID
    старше демона и в ready.json их нет.
    """
    global _DWG_DAEMON_PROC
    with _DWG_DAEMON_LOCK:
        proc = _DWG_DAEMON_PROC
        _DWG_DAEMON_PROC = None
    acad_pid = 0
    try:
        _, ready_file, _ = _dwg_daemon_files()
        if ready_file.exists():
            info = json.loads(ready_file.read_text(encoding="utf-8"))
            acad_pid = int(info.get("acadPid") or 0)
    except Exception:
        acad_pid = 0
    if proc is not None:
        try:
            proc.terminate()
            try:
                proc.wait(timeout=10)
            except Exception:
                proc.kill()
        except Exception:
            pass
    if acad_pid > 0:
        try:
            listed = subprocess.run(
                ["tasklist", "/FI", f"PID eq {acad_pid}", "/FO", "CSV", "/NH"],
                capture_output=True,
                text=True,
                timeout=15,
                **hidden_process_kwargs(),
            )
            if "acad.exe" in listed.stdout.lower():
                subprocess.run(
                    ["taskkill", "/PID", str(acad_pid), "/F"],
                    capture_output=True,
                    timeout=15,
                    **hidden_process_kwargs(),
                )
        except Exception:
            pass


def dwg_convert_via_daemon(
    *,
    input_path: Path,
    output_path: Path,
    fallback_path: Path,
    timeout_seconds: int,
) -> subprocess.CompletedProcess:
    """Одно задание демону. Возвращает CompletedProcess как разовый запуск."""
    if not ensure_dwg_daemon():
        raise RuntimeError("dwg-daemon unavailable")
    job_dir, _, _ = _dwg_daemon_files()
    job_id = hashlib.sha1(os.urandom(16)).hexdigest()
    job_file = job_dir / f"{job_id}.job.json"
    done_file = job_dir / f"{job_id}.done.json"
    job_file.write_text(
        json.dumps(
            {
                "inputPath": str(input_path),
                "outputPath": str(output_path),
                "fallbackCachePath": str(fallback_path),
                "pythonExe": str(sys.executable),
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    deadline = time.time() + timeout_seconds
    try:
        while time.time() < deadline:
            if done_file.exists():
                payload = json.loads(done_file.read_text(encoding="utf-8"))
                if payload.get("ok"):
                    result = payload.get("result") or {}
                    stdout_line = json.dumps(
                        {
                            "ok": True,
                            "finalPath": result.get("finalPath", ""),
                            "pageCount": result.get("pageCount", 0),
                            "progId": result.get("progId", ""),
                        },
                        ensure_ascii=False,
                    )
                    return subprocess.CompletedProcess(
                        args=[], returncode=0, stdout=stdout_line + "\n", stderr=""
                    )
                raise _DwgFileError(str(payload.get("error") or "CAD daemon failed"))
            if not _dwg_daemon_alive():
                raise RuntimeError("CAD daemon died mid-job")
            time.sleep(0.25)
    finally:
        for temp_file in (job_file, done_file):
            try:
                temp_file.unlink(missing_ok=True)
            except OSError:
                pass
    kill_dwg_daemon("job-timeout")
    raise TimeoutError(f"CAD conversion timed out after {timeout_seconds}s")


def dwg_convert_oneshot(
    *,
    script_to_run: Path,
    input_path: Path,
    output_path: Path,
    fallback_path: Path,
    timeout_seconds: int,
) -> subprocess.CompletedProcess:
    """Старый разовый запуск: новый CAD на каждый файл (fallback демона)."""
    return subprocess.run(
        [
            "powershell",
            "-STA",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_to_run),
            "-InputPath",
            str(input_path),
            "-OutputPath",
            str(output_path),
            "-FallbackCachePath",
            str(fallback_path),
            "-PythonExe",
            str(sys.executable),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout_seconds,
        **hidden_process_kwargs(),
    )


def dwg_convert_process(
    *,
    path: Path,
    paired_pdf: Path,
    fallback_pdf: Path,
    script_to_run: Path,
) -> subprocess.CompletedProcess:
    """Конвертация DWG в PDF: сначала быстрый нативный accoreconsole + _.-EXPORT _PDF, при сбое — CAD-сессия."""
    # 1. ПРИОРИТЕТ: Консоль accoreconsole.exe + нативная команда _.-EXPORT _PDF.
    # Запускается за доли секунды, не поднимает GUI, не мигает окнами, не тратит память на ленту.
    try:
        proc = dwg_convert_oneshot(
            script_to_run=script_to_run,
            input_path=path,
            output_path=paired_pdf,
            fallback_path=fallback_pdf,
            timeout_seconds=DWG_RENDER_TIMEOUT_SECONDS,
        )
        if proc.returncode == 0:
            return proc
    except Exception:
        pass

    # 2. РЕЗЕРВ: CAD COM демон (если accoreconsole не найден)
    try:
        return dwg_convert_via_daemon(
            input_path=path,
            output_path=paired_pdf,
            fallback_path=fallback_pdf,
            timeout_seconds=DWG_RENDER_TIMEOUT_SECONDS,
        )
    except (TimeoutError, _DwgFileError):
        raise
    except Exception:
        return dwg_convert_oneshot(
            script_to_run=script_to_run,
            input_path=path,
            output_path=paired_pdf,
            fallback_path=fallback_pdf,
            timeout_seconds=DWG_RENDER_TIMEOUT_SECONDS,
        )


def dwg_to_model_pdf(path: Path) -> tuple[Path, bool]:
    """Экспорт чертежа DWG в многостраничный векторный PDF через AutoCAD COM.

    Сохраняет парный PDF рядом с исходным DWG-файлом (при наличии прав записи)
    либо в резервный локальный кэш, если папка защищена от записи.
    """
    if not path.exists():
        raise FileNotFoundError(f"DWG-файл не найден: {path}")
    if not path.is_file() or path.suffix.casefold() != ".dwg":
        raise ValueError(f"Это не DWG-файл: {path}")

    # 1. Проверяем парный PDF рядом с исходником DWG
    paired_pdf = path.with_suffix(".pdf")
    if paired_pdf.exists() and paired_pdf.is_file() and paired_pdf.stat().st_size > 1024:
        try:
            if paired_pdf.stat().st_mtime_ns >= path.stat().st_mtime_ns:
                return paired_pdf, True
        except OSError:
            pass

    # 2. Проверяем резервный кэш
    key = file_cache_key(path, "dwg-smart-cad-v2")
    target_dir = DWG_CACHE_DIR / key
    target_dir.mkdir(parents=True, exist_ok=True)
    fallback_pdf = target_dir / f"{path.stem}.pdf"
    manifest_path = target_dir / "manifest.json"
    if fallback_pdf.exists() and fallback_pdf.stat().st_size > 1024 and manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            if (
                manifest.get("sourcePath") == str(path)
                and manifest.get("cacheKey") == key
                and manifest.get("sourceMtimeNs") == path.stat().st_mtime_ns
                and manifest.get("sourceSize") == path.stat().st_size
            ):
                return fallback_pdf, True
        except (OSError, json.JSONDecodeError):
            pass

    script_to_run = DWG_SMART_RENDER_SCRIPT if DWG_SMART_RENDER_SCRIPT.exists() else DWG_RENDER_SCRIPT
    final_pdf = None
    try:
        process = dwg_convert_process(
            path=path,
            paired_pdf=paired_pdf,
            fallback_pdf=fallback_pdf,
            script_to_run=script_to_run,
        )
        if process.returncode != 0:
            message = process.stderr.strip() or process.stdout.strip() or "CAD-система (AutoCAD) не смогла создать PDF для чертежа"
            raise RuntimeError(message)

        if paired_pdf.exists() and paired_pdf.stat().st_size > 1024:
            final_pdf = paired_pdf
        elif fallback_pdf.exists() and fallback_pdf.stat().st_size > 1024:
            final_pdf = fallback_pdf
        else:
            for line in reversed((process.stdout or "").splitlines()):
                line = line.strip()
                if line.startswith("{") and line.endswith("}"):
                    try:
                        data = json.loads(line)
                        cand = Path(data.get("finalPath", ""))
                        if cand.exists() and cand.stat().st_size > 1024:
                            final_pdf = cand
                            break
                    except Exception:
                        pass

        if not final_pdf or not final_pdf.exists() or final_pdf.stat().st_size <= 1024:
            raise RuntimeError("CAD-система (AutoCAD) не создала PDF-файл для чертежа")
    except Exception as cad_err:
        try:
            try:
                from app.backend.dwg_engine import extract_raw_thumbnail_from_dwg, _generate_placeholder_png
            except ImportError:
                from dwg_engine import extract_raw_thumbnail_from_dwg, _generate_placeholder_png
            thumb = extract_raw_thumbnail_from_dwg(path)
            img_bytes = thumb[0] if thumb else _generate_placeholder_png(path.name, "Model")

            # Конвертируем PNG превью чертежа в PDF через fitz (PyMuPDF)
            import fitz
            img_doc = fitz.open(stream=img_bytes, filetype="png")
            pdf_bytes = img_doc.convert_to_pdf()
            img_doc.close()

            fallback_pdf.write_bytes(pdf_bytes)
            final_pdf = fallback_pdf
        except Exception as fallback_err:
            raise RuntimeError(f"Сбой рендера DWG: {cad_err} (фолбэк: {fallback_err})")

    if final_pdf == fallback_pdf:
        manifest_path.write_text(
            json.dumps(
                {
                    "sourcePath": str(path),
                    "sourceName": path.name,
                    "sourceMtimeNs": path.stat().st_mtime_ns,
                    "sourceSize": path.stat().st_size,
                    "cacheKey": key,
                    "pdfPath": str(fallback_pdf),
                    "mode": "smart-layouts-fallback-cache",
                    "renderedAt": datetime.now().isoformat(timespec="seconds"),
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
    return final_pdf, False


_DWG_BG_LOCK = threading.Lock()
_DWG_BG_ACTIVE: set[str] = set()


def render_dwg_model(path: Path, dpi: int = DEFAULT_PDF_DPI) -> dict:
    if not path.exists():
        return {
            "name": path.name,
            "path": str(path),
            "sourcePath": str(path),
            "sourceName": path.name,
            "sourceType": "DWG",
            "dpi": dpi,
            "pages": 0,
            "renderedPages": 0,
            "cacheKey": "",
            "cacheHit": False,
            "cacheHitPages": 0,
            "newRenderedPages": 0,
            "is_missing": True,
            "errors": [{"error": f"DWG-файл не найден на диске: {path.name}", "is_missing": True}],
            "items": [],
            "status": "missing",
            "convertedPdfPath": "",
            "convertCacheHit": False,
            "previewMode": "cad-smart-layouts",
        }
    if not path.is_file() or path.suffix.casefold() != ".dwg":
        raise ValueError(f"Это не DWG-файл: {path}")

    # 1. Проверяем, есть ли уже готовый векторный PDF (парный или в кэше)
    paired_pdf = path.with_suffix(".pdf")
    key = file_cache_key(path, "dwg-smart-cad-v2")
    target_dir = DWG_CACHE_DIR / key
    target_dir.mkdir(parents=True, exist_ok=True)
    fallback_pdf = target_dir / f"{path.stem}.pdf"
    manifest_path = target_dir / "manifest.json"
    has_cached_pdf = False
    if paired_pdf.exists() and paired_pdf.is_file() and paired_pdf.stat().st_size > 1024:
        try:
            if paired_pdf.stat().st_mtime_ns >= path.stat().st_mtime_ns:
                has_cached_pdf = True
        except OSError:
            pass
    if not has_cached_pdf and fallback_pdf.exists() and fallback_pdf.stat().st_size > 1024 and manifest_path.exists():
        try:
            m = json.loads(manifest_path.read_text(encoding="utf-8"))
            if (
                m.get("sourcePath") == str(path)
                and m.get("cacheKey") == key
                and m.get("sourceMtimeNs") == path.stat().st_mtime_ns
                and m.get("sourceSize") == path.stat().st_size
            ):
                has_cached_pdf = True
        except (OSError, json.JSONDecodeError):
            pass

    # Если векторного PDF нет, но в кэше лежат старые растровые картинки — очищаем их
    if not has_cached_pdf:
        for stale_png in target_dir.glob("page-*.png"):
            try:
                stale_png.unlink()
            except OSError:
                pass

    # 2. Рендерим качественную модель DWG через векторный PDF (экспорт или печать)
    pdf_path, convert_cache_hit = dwg_to_model_pdf(path)
    document = render_pdf(pdf_path, dpi=dpi, page_timeout_seconds=DWG_MODEL_PAGE_TIMEOUT_SECONDS)
    document["name"] = path.name
    document["sourcePath"] = str(path)
    document["sourceName"] = path.name
    document["sourceType"] = "DWG"
    document["convertedPdfPath"] = str(pdf_path)
    document["convertCacheHit"] = convert_cache_hit
    document["previewMode"] = "cad-smart-layouts"
    return document

def set_windows_clipboard(text: str) -> bool:
    if os.name != "nt":
        return False
    normalized_text = os.path.normpath(text)
    try:
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        user32.OpenClipboard.argtypes = [wintypes.HWND]
        user32.OpenClipboard.restype = wintypes.BOOL
        user32.EmptyClipboard.argtypes = []
        user32.EmptyClipboard.restype = wintypes.BOOL
        user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
        user32.SetClipboardData.restype = wintypes.HANDLE
        user32.CloseClipboard.argtypes = []
        user32.CloseClipboard.restype = wintypes.BOOL

        kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
        kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
        kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalLock.restype = wintypes.LPVOID
        kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalUnlock.restype = wintypes.BOOL

        GMEM_MOVEABLE = 0x0002
        CF_UNICODETEXT = 13

        opened = False
        for _ in range(10):
            if user32.OpenClipboard(None):
                opened = True
                break
            time.sleep(0.04)

        if opened:
            try:
                user32.EmptyClipboard()
                encoded = (normalized_text + "\0").encode("utf-16le")
                h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(encoded))
                if h_mem:
                    p_mem = kernel32.GlobalLock(h_mem)
                    if p_mem:
                        ctypes.memmove(p_mem, encoded, len(encoded))
                        kernel32.GlobalUnlock(h_mem)
                        if user32.SetClipboardData(CF_UNICODETEXT, h_mem):
                            return True
            finally:
                user32.CloseClipboard()
    except Exception:
        pass

    try:
        res = subprocess.run(["clip.exe"], input=normalized_text.encode("utf-16le"), check=True, timeout=3)
        return res.returncode == 0
    except Exception:
        pass

    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", "$input | Set-Clipboard"],
            input=normalized_text,
            text=True,
            timeout=3,
            check=True
        )
        return True
    except Exception:
        return False


NATIVE_APPS_CONFIG_FILE = RUNTIME_DIR / "native_apps.json"

DEFAULT_NATIVE_APPS: dict[str, Any] = {
    "useNativeApps": True,
    "defaultAction": "explorer",
    ".dwg": "",
    ".dxf": "",
    ".pdf": "",
    ".doc": "",
    ".docx": "",
    ".rtf": "",
    ".odt": "",
    ".xls": "",
    ".xlsx": "",
    ".xlsm": "",
    ".csv": "",
    ".ods": "",
    ".ppt": "",
    ".pptx": "",
    ".odp": "",
    ".jpg": "",
    ".jpeg": "",
    ".png": "",
    ".bmp": "",
    ".gif": "",
    ".webp": "",
    ".tif": "",
    ".tiff": "",
    ".ico": "",
    ".svg": "",
    ".zip": "",
    ".rar": "",
    ".7z": "",
    ".tar": "",
    ".gz": "",
    ".txt": "",
    ".log": "",
    ".ini": "",
    ".cfg": "",
    ".json": "",
    ".xml": "",
    ".yaml": "",
    ".yml": "",
    ".mp4": "",
    ".avi": "",
    ".mov": "",
    ".mkv": "",
    ".mp3": "",
    ".wav": "",
}

def load_native_apps_config() -> dict[str, Any]:
    if NATIVE_APPS_CONFIG_FILE.exists():
        try:
            with open(NATIVE_APPS_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    cfg = dict(DEFAULT_NATIVE_APPS)
                    cfg.update(data)
                    return cfg
        except Exception as err:
            pass
    return dict(DEFAULT_NATIVE_APPS)

def save_native_apps_config(data: dict[str, Any]) -> None:
    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    with open(NATIVE_APPS_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


class NativeOpenError(RuntimeError):
    """Структурированная ошибка нативного открытия файла."""

    def __init__(self, message: str, error_code: str = "native_open_failed", details: dict | None = None):
        super().__init__(message)
        self.error_code = error_code
        self.details = details or {}


def normalize_file_extension(ext_or_path: str | Path) -> str:
    """Нормализует расширение файла: всегда в нижнем регистре с ведущей точкой.

    Примеры: '.DWG' -> '.dwg', 'dwg' -> '.dwg', Path('file.DXF') -> '.dxf'.
    """
    if isinstance(ext_or_path, Path):
        ext = ext_or_path.suffix
    else:
        ext = str(ext_or_path or "").strip()
    if not ext:
        return ""
    ext = ext.strip().casefold()
    if not ext.startswith("."):
        ext = "." + ext
    return ext


SAFE_EXECUTABLE_EXTENSIONS = {".exe", ".cmd", ".bat", ".com"}
FORBIDDEN_AUTODETECT_EXES = {"aclauncher.exe", "zwlauncher.exe"}


def _clean_configured_exe_path(raw: str) -> str:
    """Очищает путь к исполняемому файлу от кавычек и параметров реестра."""
    if not raw:
        return ""
    raw = raw.strip()
    # 1. Снимаем любые внешние кавычки
    stripped = raw.strip('"').strip("'").strip()
    if stripped.lower().endswith(".exe"):
        return stripped
    # 2. Ищем путь в кавычках с .exe (например, из реестра: "C:\...\app.exe" "%1")
    m = re.search(r'"([^"]+?\.exe)"', raw, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # 3. Ищем путь с диском Windows без кавычек
    m2 = re.search(r'([a-zA-Z]:\\[^\s"]+?\.exe)', raw, re.IGNORECASE)
    if m2:
        return m2.group(1).strip()
    return stripped


def validate_configured_executable(exe_path: str) -> tuple[bool, str]:
    """Проверяет путь к исполняемому файлу из настроек пользователя.

    Возвращает (is_valid, cleaned_path).
    """
    if not exe_path or not str(exe_path).strip():
        return False, ""
    cleaned = _clean_configured_exe_path(str(exe_path).strip())
    if not cleaned:
        return False, ""
    p = Path(cleaned)
    if not p.is_file():
        return False, cleaned
    if p.suffix.casefold() not in SAFE_EXECUTABLE_EXTENSIONS:
        return False, cleaned
    return True, cleaned


def get_configured_app_for_extension(ext: str, cfg: dict | None = None) -> str:
    """Возвращает путь к настроенной программе для заданного расширения из cfg."""
    norm_ext = normalize_file_extension(ext)
    if not norm_ext:
        return ""
    if cfg is None:
        cfg = load_native_apps_config()

    # Точное нормализованное расширение (например, ".dwg")
    val = cfg.get(norm_ext)
    if val and str(val).strip():
        return str(val).strip()

    # Варианты без точки / в верхнем регистре для обратной совместимости
    no_dot = norm_ext.lstrip(".")
    for key in (no_dot, norm_ext.upper(), no_dot.upper()):
        val = cfg.get(key)
        if val and str(val).strip():
            return str(val).strip()

    # Специализированные ключи формы setting...Exe
    fallback_key = ""
    if norm_ext in {".dwg", ".dxf"}:
        fallback_key = "settingDwgExe"
    elif norm_ext in {".pdf"}:
        fallback_key = "settingPdfExe"
    elif norm_ext in {".doc", ".docx", ".rtf", ".odt"}:
        fallback_key = "settingWordExe"
    elif norm_ext in {".xls", ".xlsx", ".xlsm", ".csv", ".ods", ".xlsb"}:
        fallback_key = "settingExcelExe"
    elif norm_ext in {".ppt", ".pptx", ".odp"}:
        fallback_key = "settingPptExe"
    elif norm_ext in {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff", ".ico", ".svg"}:
        fallback_key = "settingImgExe"
    elif norm_ext in {".zip", ".rar", ".7z", ".tar", ".gz"}:
        fallback_key = "settingArchExe"
    elif norm_ext in {".txt", ".log", ".ini", ".cfg", ".json", ".xml", ".yaml", ".yml"}:
        fallback_key = "settingTxtExe"
    elif norm_ext in {".mp4", ".avi", ".mov", ".mkv", ".wmv", ".mp3", ".wav"}:
        fallback_key = "settingMediaExe"

    if fallback_key and cfg.get(fallback_key):
        val = cfg.get(fallback_key)
        if val and str(val).strip():
            return str(val).strip()

    return ""

def _force_foreground(hwnd: int, show_cmd: int = 9) -> bool:
    """Вывести окно на передний план и проверить результат.

    Прямой SetForegroundWindow фоновому процессу система режет,
    поэтому идём лесенкой: привязка потоков, затем Alt-толчок.
    Возвращает True, если окно действительно стало главным.
    """
    try:
        import ctypes

        u32 = ctypes.windll.user32
        k32 = ctypes.windll.kernel32
        try:
            u32.AllowSetForegroundWindow(0xFFFFFFFF)
        except Exception:
            pass
        try:
            u32.ShowWindow(hwnd, show_cmd)  # SW_RESTORE
        except Exception:
            pass
        try:
            if u32.SetForegroundWindow(hwnd):
                return True
        except Exception:
            pass
        # Лесенка 1: привязываемся к потоку главного окна.
        try:
            fg = u32.GetForegroundWindow()
            fg_thread = u32.GetWindowThreadProcessId(fg, None)
            my_thread = k32.GetCurrentThreadId()
            if fg_thread and u32.AttachThreadInput(my_thread, fg_thread, True):
                try:
                    u32.ShowWindow(hwnd, show_cmd)
                    if u32.SetForegroundWindow(hwnd):
                        return True
                finally:
                    try:
                        u32.AttachThreadInput(my_thread, fg_thread, False)
                    except Exception:
                        pass
        except Exception:
            pass
        # Лесенка 2: короткое нажатие Alt снимает блокировку, повторяем.
        try:
            u32.keybd_event(0x12, 0, 0, 0)
            u32.keybd_event(0x12, 0, 2, 0)
        except Exception:
            pass
        try:
            import time as _time

            _time.sleep(0.15)
            u32.ShowWindow(hwnd, show_cmd)
            if u32.SetForegroundWindow(hwnd):
                return True
            return u32.GetForegroundWindow() == hwnd
        except Exception:
            return False
    except Exception:
        return False


def _center_window_on_screen(target_hwnd) -> bool:
    """Поставить окно по центру основного экрана, не меняя размер и фокус.

    Общая helper для обоих сторожей (custom-exe и system-default).
    Развёрнутые/свёрнутые/нулевые окна пропускаем.
    """
    try:
        import ctypes

        u32 = ctypes.windll.user32

        class _Rect(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                        ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

        if u32.IsZoomed(target_hwnd) or u32.IsIconic(target_hwnd):
            return False
        rect = _Rect()
        if not u32.GetWindowRect(target_hwnd, ctypes.byref(rect)):
            return False
        width = rect.right - rect.left
        height = rect.bottom - rect.top
        if width <= 0 or height <= 0:
            return False
        screen_w = u32.GetSystemMetrics(0)
        screen_h = u32.GetSystemMetrics(1)
        if not screen_w or not screen_h:
            return False
        pos_x = max(0, (screen_w - width) // 2)
        pos_y = max(0, (screen_h - height) // 2)
        SWP_NOSIZE = 0x0001
        SWP_NOZORDER = 0x0004
        SWP_NOACTIVATE = 0x0010
        return bool(u32.SetWindowPos(target_hwnd, 0, pos_x, pos_y, 0, 0,
                                     SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE))
    except Exception:
        return False


def _bring_window_to_front(title_part: str | None, cls_name: str | None,
                           timeout: float = 120.0, match_title: str = "", show_cmd: int = 9) -> None:
    """Фоном ждём новое окно и делаем его главным. Ответ API не блокирует.

    Окно ищется среди появившихся после запуска: сначала по совпадению
    заголовка, иначе первое новое окно нужного класса. Факт проверяется,
    попытки повторяются до таймаута.
    """
    def _watch() -> None:
        try:
            import ctypes
            import time as _time

            u32 = ctypes.windll.user32
            EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

            def _snapshot() -> dict:
                found: dict = {}

                def _cb(hwnd, _):
                    try:
                        if not u32.IsWindowVisible(hwnd):
                            return True
                        length = u32.GetWindowTextLengthW(hwnd)
                        if length <= 0:
                            return True
                        buf = ctypes.create_unicode_buffer(length + 1)
                        u32.GetWindowTextW(hwnd, buf, length + 1)
                        cls = ctypes.create_unicode_buffer(256)
                        u32.GetClassNameW(hwnd, cls, 256)
                        if cls_name and cls.value != cls_name:
                            return True
                        found[int(hwnd)] = buf.value or ""
                    except Exception:
                        pass
                    return True

                u32.EnumWindows(EnumWindowsProc(_cb), 0)
                return found

            want = (match_title or "").casefold()
            before = set(_snapshot())
            deadline = _time.time() + timeout
            best = None
            last_try = 0.0
            focus_fails = 0
            while _time.time() < deadline:
                _time.sleep(0.5)
                try:
                    after = _snapshot()
                except Exception:
                    continue
                if best is not None and best not in after:
                    best = None
                fresh = {h: t for h, t in after.items() if h not in before}
                for handle, title in fresh.items():
                    if want and want in (title or "").casefold():
                        best = handle
                        break
                if best is None and fresh:
                    best = next(iter(fresh))
                if best is None:
                    continue
                now = _time.time()
                if now - last_try < 2.0:
                    continue
                last_try = now
                try:
                    _force_foreground(best, show_cmd)
                    # Позицию правим независимо от фокуса: даже неудостоенное
                    # фокуса окно останется по центру, а не в углу.
                    try:
                        _center_window_on_screen(best)
                    except Exception:
                        pass
                    _time.sleep(0.3)
                    try:
                        if u32.GetForegroundWindow() == best:
                            return
                        focus_fails += 1
                    except Exception:
                        return
                except Exception:
                    focus_fails += 1
                # Фокус не даётся подряд: дальше долбить — только мигать
                # экраном. Позиция уже выправлена выше, выходим.
                if focus_fails >= 6:
                    return
        except Exception:
            pass

    try:
        threading.Thread(target=_watch, daemon=True).start()
    except Exception:
        pass


def _bring_explorer_to_front(target_dir: str, timeout: float = 120.0) -> None:
    """Фоном доводим окно проводника на передний план, как только появится.

    Не блокирует ответ API: холодная облачная папка открывается минутами,
    а пользователь должен увидеть окно главным, а не значком в панели.
    """
    _bring_window_to_front(None, "CabinetWClass", timeout,
                           os.path.basename(os.path.normpath(target_dir)))


def bring_native_window_to_front(pid: int, exe_path: str, timeout: float = 120.0) -> None:
    """Окно нативной программы — на передний план и на весь экран.

    Окно AutoCAD дополнительно ставится по центру экрана (оно любит
    открываться в левом верхнем углу).

    Windows не даёт фоновому серверу фокус напрямую (новое окно только
    мигает в панели задач), поэтому в фоновом потоке ждём главное окно
    запущенного PID и доводим его проверенной лесенкой (_force_foreground)
    с разворотом на весь экран. Если процесс-переходник уже вышел
    (single-instance: файл подхватила старая сессия), берём главное окно
    того же exe. Ответ API не блокируем.
    """
    exe_name = os.path.basename(exe_path or "").casefold()

    def _watch() -> None:
        try:
            import ctypes
            import time as _time

            def _flog(msg: str) -> None:
                try:
                    with open(RUNTIME_DIR / "logs" / "native-front.log",
                              "a", encoding="utf-8") as log:
                        log.write(f"{_time.strftime('%H:%M:%S')} pid={pid} {msg}\n")
                except Exception:
                    pass

            u32 = ctypes.windll.user32
            k32 = ctypes.windll.kernel32
            _flog(f"watch start exe={exe_name}")
            EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
            GW_OWNER = 4

            class _Rect(ctypes.Structure):
                _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                            ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

            def _pid_of(hwnd) -> int:
                out = ctypes.c_ulong(0)
                try:
                    u32.GetWindowThreadProcessId(hwnd, ctypes.byref(out))
                except Exception:
                    pass
                return int(out.value)

            def _rect_area(hwnd) -> int:
                try:
                    rect = _Rect()
                    if u32.GetWindowRect(hwnd, ctypes.byref(rect)):
                        return max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
                except Exception:
                    pass
                return 0

            def _has_title(hwnd) -> bool:
                try:
                    return u32.GetWindowTextLengthW(hwnd) > 0
                except Exception:
                    return False

            def _center_on_screen(target_hwnd) -> bool:
                # Позицию считает общий helper; здесь только делегирование,
                # чтобы не разъезжались две копии логики.
                try:
                    return bool(_center_window_on_screen(target_hwnd))
                except Exception:
                    return False

            def _image_of(test_pid: int) -> str:
                try:
                    handle = k32.OpenProcess(0x1000, False, test_pid)
                    if not handle:
                        return ""
                    try:
                        buf = ctypes.create_unicode_buffer(260)
                        size = ctypes.c_ulong(260)
                        if k32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
                            return os.path.basename(buf.value or "").casefold()
                    finally:
                        try:
                            k32.CloseHandle(handle)
                        except Exception:
                            pass
                except Exception:
                    pass
                return ""

            def _pid_alive(test_pid: int) -> bool:
                try:
                    handle = k32.OpenProcess(0x1000, False, test_pid)
                    if not handle:
                        return False
                    try:
                        k32.CloseHandle(handle)
                    except Exception:
                        pass
                    return True
                except Exception:
                    return False

            def _main_windows(want_pid: int = 0, want_image: str = "") -> list:
                found: list = []

                def _cb(hwnd, _):
                    try:
                        if not u32.IsWindowVisible(hwnd):
                            return True
                        if u32.GetWindow(hwnd, GW_OWNER):
                            return True
                        if want_pid and _pid_of(hwnd) != want_pid:
                            return True
                        if want_image and _image_of(_pid_of(hwnd)) != want_image:
                            return True
                        found.append(hwnd)
                    except Exception:
                        pass
                    return True

                try:
                    u32.EnumWindows(EnumWindowsProc(_cb), 0)
                except Exception:
                    pass
                return found

            def _best(hwnds: list):
                titled = [h for h in hwnds if _has_title(h)]
                pool = titled or hwnds
                if not pool:
                    return None
                try:
                    return max(pool, key=_rect_area)
                except Exception:
                    return pool[0]

            hwnd = None
            deadline = _time.time() + timeout
            while _time.time() < deadline and hwnd is None:
                if not _pid_alive(pid):
                    break
                hwnd = _best(_main_windows(want_pid=pid))
                if hwnd is None:
                    _time.sleep(0.5)
            if hwnd is None and exe_name:
                for _ in range(20):
                    hwnd = _best(_main_windows(want_image=exe_name))
                    if hwnd is not None:
                        break
                    _time.sleep(0.5)
            if hwnd is None:
                _flog("no window found")
                return
            _flog(f"target hwnd={hwnd}")
            # AutoCAD любит вставать в левый верхний угол: дважды ставим
            # окно по центру (до и после доводки фокуса), пока оно грузится.
            is_cad = exe_name.startswith("acad")
            if is_cad:
                try:
                    _flog(f"centered-before={_center_on_screen(hwnd)}")
                except Exception:
                    pass
            for _ in range(3):
                try:
                    _force_foreground(hwnd, 3)
                except Exception:
                    pass
                _time.sleep(1.0)
                try:
                    fg = u32.GetForegroundWindow()
                    zoomed = bool(u32.IsZoomed(hwnd))
                    _flog(f"try fg={fg} zoomed={zoomed}")
                    if fg == hwnd:
                        break
                except Exception:
                    break
            if is_cad:
                try:
                    _flog(f"centered-after={_center_on_screen(hwnd)}")
                except Exception:
                    pass
        except Exception:
            pass

    try:
        threading.Thread(target=_watch, daemon=True).start()
    except Exception:
        pass


def open_in_explorer(path: Path) -> str:
    """Гарантированно открывает Проводник Windows с выделением файла или переходом в папку."""
    if not path.exists():
        raise FileNotFoundError(f"Файл или папка не найдены: {path}")

    resolved = os.path.normpath(str(path.resolve()))
    target_dir = resolved if path.is_dir() else os.path.dirname(resolved)

    if os.name != "nt":
        target = path if path.is_dir() else path.parent
        subprocess.Popen(["xdg-open", str(target)])
        return "xdg-open"

    try:
        import ctypes
        ctypes.windll.user32.AllowSetForegroundWindow(ctypes.c_uint32(0xFFFFFFFF))
    except Exception:
        pass

    if path.is_dir():
        try:
            # Папку открываем явным вызовом проводника отдельными
            # аргументами: склеенная строка вида 'explorer.exe "путь"'
            # через Popen без shell работает ненадёжно.
            subprocess.Popen(["explorer.exe", resolved])
            _bring_explorer_to_front(resolved)
            return "explorer-open-folder"
        except Exception:
            os.startfile(resolved)
            return "explorer-open-folder-fallback"

    # Для файла: выделяем его в Проводнике Windows.
    # Важно: "/select," и путь — ОТДЕЛЬНЫЕ аргументы. Склеенный вариант
    # вида "/select,C:\..." проводник молча игнорирует (окно либо не
    # открывается, либо открывается не та папка).
    try:
        subprocess.Popen(["explorer.exe", "/select,", resolved])
        _bring_explorer_to_front(target_dir)
        return "explorer-select"
    except Exception:
        try:
            os.startfile(target_dir)
            return "explorer-open-dir-fallback"
        except Exception:
            subprocess.Popen(["explorer.exe", target_dir])
            return "explorer-open-dir-subp"


def _shell_open(target: str, params: str = "", cwd: str = "") -> bool:
    """Универсальный вызов ShellExecuteW в нормальном видимом окне (SW_SHOWNORMAL = 1).
    Это эталонный механизм Windows Shell для открытия любых файлов и программ.
    Никогда не создаёт чёрных окон консоли и никогда не запускает окно скрытым.
    """
    if os.name != "nt":
        return False
    try:
        import ctypes
        res = ctypes.windll.shell32.ShellExecuteW(
            None,
            "open",
            str(target),
            str(params) if params else None,
            str(cwd) if cwd else None,
            1,  # SW_SHOWNORMAL = 1
        )
        # В Win32 API ShellExecute возвращает дескриптор экземпляра > 32 при успехе
        return int(res) > 32
    except Exception:
        return False


def launch_system_default(path: Path) -> str:
    """Открыть файл программой Windows по умолчанию (системная ассоциация Windows)."""
    if not path.exists():
        raise FileNotFoundError(f"Файл или папка не найдены: {path}")
    resolved = os.path.normpath(str(path.resolve()))
    if os.name != "nt":
        target = path if path.is_dir() else path.parent
        subprocess.Popen(["xdg-open", str(target)])
        return "xdg-open"

    cwd = os.path.dirname(resolved)

    # 1. ShellExecuteW("open", resolved, None, cwd, SW_SHOWNORMAL=1)
    if _shell_open(resolved, "", cwd):
        _bring_window_to_front(None, None, 120.0, path.name, 3)
        return "system-default"

    # 2. os.startfile (встроенная функция Python для Windows)
    try:
        if hasattr(os, "startfile"):
            os.startfile(resolved)
            _bring_window_to_front(None, None, 120.0, path.name, 3)
            return "system-default"
    except Exception:
        pass

    raise NativeOpenError(
        f"Не удалось открыть файл программой по умолчанию: {path.name}",
        error_code="system_default_failed",
        details={"path": resolved},
    )


def launch_custom_app(exe_path: str, file_path: Path) -> str:
    """Запустить файл в явно указанной пользователем программе (EXE).

    Запускает ТОЛЬКО указанный исполняемый файл, передавая оригинальный путь
    к файлу отдельным аргументом без объединения в командную строку.
    Никогда не подменяет на системный запуск по умолчанию и не вызывает accoreconsole.
    """
    resolved_file = os.path.normpath(str(file_path.resolve()))
    resolved_exe = os.path.normpath(str(Path(exe_path).resolve()))
    file_dir = os.path.dirname(resolved_file)

    if os.name != "nt":
        subprocess.Popen([resolved_exe, resolved_file], cwd=file_dir)
        return f"custom-app:{Path(resolved_exe).name}"

    # 1. ShellExecuteW("open", resolved_exe, f'"{resolved_file}"', file_dir, SW_SHOWNORMAL=1)
    if _shell_open(resolved_exe, f'"{resolved_file}"', file_dir):
        bring_native_window_to_front(0, resolved_exe)
        _bring_window_to_front(None, None, 120.0, file_path.name, 3)
        return f"custom-app:{Path(resolved_exe).name}"

    # 2. subprocess.Popen с DETACHED_PROCESS и CREATE_NEW_PROCESS_GROUP
    try:
        creationflags = getattr(subprocess, "DETACHED_PROCESS", 0x00000008) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
        proc = subprocess.Popen(
            [resolved_exe, resolved_file],
            cwd=file_dir,
            creationflags=creationflags,
        )
        bring_native_window_to_front(proc.pid, resolved_exe)
        _bring_window_to_front(None, None, 120.0, file_path.name, 3)
        return f"custom-app:{Path(resolved_exe).name}"
    except Exception as err:
        raise NativeOpenError(
            f"Не удалось запустить приложение {resolved_exe}: {err}",
            error_code="custom_app_launch_failed",
            details={"exe": resolved_exe, "file": resolved_file, "error": str(err)},
        )


def launch_native_file(path: Path) -> str:
    """Запустить файл в ассоциированной программе, либо открыть через системную ассоциацию Windows.

    Приоритет:
    A. Есть валидный EXE, сохранённый в Параметрах -> запускать ТОЛЬКО этот EXE отдельными аргументами.
    B. Настройка отсутствует/пуста -> системный fallback (ShellExecuteW / os.startfile).
    C. Настройка указана, но файл не существует/невалиден -> NativeOpenError("configured_app_missing") БЕЗ fallback!
    """
    if not path.exists():
        raise FileNotFoundError(f"Файл или папка не найдены: {path}")
    if path.is_dir():
        return open_in_explorer(path)

    norm_ext = normalize_file_extension(path)
    cfg = load_native_apps_config()
    use_native = bool(cfg.get("useNativeApps", True))

    configured_exe = ""
    if use_native:
        configured_exe = get_configured_app_for_extension(norm_ext, cfg)

    if configured_exe:
        is_valid, clean_exe = validate_configured_executable(configured_exe)
        if not is_valid:
            ext_label = norm_ext.upper().replace(".", "")
            raise NativeOpenError(
                f"Для файлов {ext_label} указано недоступное приложение. Проверьте Параметры → Приложения.",
                error_code="configured_app_missing",
                details={"extension": norm_ext, "configured_app": configured_exe},
            )
        return launch_custom_app(clean_exe, path)

    # Настройка отсутствует или пуста -> системный fallback
    return launch_system_default(path)

NATIVE_OPEN_LOG_LOCK = threading.Lock()
_OPEN_FILE_LOCK = threading.Lock()
_LAST_OPENED_TARGETS: dict[str, float] = {}


def append_native_open_log(event: dict) -> None:
    """Write one UTF-8 diagnostic record for a native-open attempt.

    ThreadingHTTPServer serves each request on its own thread, so concurrent
    /api/open-file calls can interleave partial JSON lines.  A lock keeps each
    record as one atomic write.
    """
    logs_dir = RUNTIME_DIR / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    line = json.dumps(event, ensure_ascii=False) + "\n"
    with NATIVE_OPEN_LOG_LOCK:
        with (logs_dir / "native-open.jsonl").open("a", encoding="utf-8") as log:
            log.write(line)


IMPORT_LOG_LOCK = threading.Lock()


def append_import_log(event: dict) -> None:
    """Записать одну UTF-8 строку диагностики загрузки объекта в файл."""
    try:
        logs_dir = RUNTIME_DIR / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        line = json.dumps(event, ensure_ascii=False) + "\n"
        with IMPORT_LOG_LOCK:
            with (logs_dir / "import.jsonl").open("a", encoding="utf-8") as log:
                log.write(line)
    except Exception:
        pass


def excel_html_cache_dir(path: Path) -> Path:
    """Return an immutable cache location for one exact workbook revision."""
    return EXCEL_CACHE_DIR / "html" / file_cache_key(path, "excel-html")


def excel_convert_xls_to_xlsx(path: Path) -> tuple[Path, bool]:
    """Convert a legacy .xls workbook to .xlsx via Excel COM, cached by source.

    openpyxl can only read the OpenXML formats, so legacy .xls books are
    converted once through the installed Excel and reused until the source
    file changes.  The converted copy is only a preview source: nothing is
    written next to the original workbook.
    """
    if path.suffix.casefold() not in {".xls"}:
        raise ValueError(f"Это не файл XLS: {path}")
    if not EXCEL_XLS_CONVERT_SCRIPT.exists():
        raise RuntimeError(f"Скрипт конвертации XLS не найден: {EXCEL_XLS_CONVERT_SCRIPT}")

    key = file_cache_key(path, "excel-xls-convert")
    target_dir = EXCEL_CACHE_DIR / "xlsx" / key
    target_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = target_dir / f"{path.stem}.xlsx"
    manifest_path = target_dir / "manifest.json"

    cached = False
    if xlsx_path.exists() and xlsx_path.stat().st_size > 0 and manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            cached = (
                manifest.get("sourcePath") == str(path)
                and manifest.get("sourceMtimeNs") == path.stat().st_mtime_ns
                and manifest.get("sourceSize") == path.stat().st_size
            )
        except (OSError, json.JSONDecodeError):
            cached = False

    if cached:
        return xlsx_path, True

    if xlsx_path.exists():
        xlsx_path.unlink()

    process = subprocess.run(
        [
            "powershell",
            "-STA",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(EXCEL_XLS_CONVERT_SCRIPT),
            "-InputPath",
            str(path),
            "-OutputPath",
            str(xlsx_path),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=EXCEL_CONVERT_TIMEOUT_SECONDS,
        **hidden_process_kwargs(),
    )
    if process.returncode != 0:
        message = process.stderr.strip() or process.stdout.strip() or "Excel не смог сконвертировать книгу XLS"
        raise RuntimeError(message)
    if not xlsx_path.exists() or xlsx_path.stat().st_size <= 0:
        raise RuntimeError("Excel не создал XLSX для preview")

    manifest_path.write_text(
        json.dumps(
            {
                "sourcePath": str(path),
                "sourceName": path.name,
                "sourceMtimeNs": path.stat().st_mtime_ns,
                "sourceSize": path.stat().st_size,
                "cacheKey": key,
                "xlsxPath": str(xlsx_path),
                "convertedAt": datetime.now().isoformat(timespec="seconds"),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return xlsx_path, False


def excel_preview_source(path: Path) -> Path:
    """Return an .xlsx/.xlsm path openpyxl can read, converting legacy .xls.

    HTML rendering and sheet extraction both need the OpenXML workbook.
    .xlsx/.xlsm files are used as-is; .xls is converted through Excel COM.
    """
    if path.name.startswith("~$") or path.name.startswith(".~"):
        raise ValueError(f"Временный файл блокировки Office: {path.name}")
    if path.suffix.casefold() in {".xlsx", ".xlsm"}:
        return path
    if path.suffix.casefold() == ".xls":
        xlsx_path, _ = excel_convert_xls_to_xlsx(path)
        return xlsx_path
    raise ValueError("HTML-просмотр пока поддерживает XLSX/XLSM. Откройте XLS в Excel.")


def excel_cell_color(color: object) -> str | None:
    value = getattr(color, "rgb", None)
    value = str(value) if value is not None else ""
    if not value or len(value) < 6 or (len(value) == 8 and value[:2] == "00"):
        return None
    return f"#{value[-6:]}"


def excel_sheet_html(path: Path, sheet_index: int) -> tuple[str, dict]:
    """Create a read-only HTML sheet with authored geometry.

    The Launcher is a fast navigator, not a replacement for Excel.  The limit is
    intentional: long operational registers are opened in the native program.
    """
    if openpyxl is None:
        raise RuntimeError("Для HTML-просмотра Excel нужен пакет openpyxl")

    preview_source = excel_preview_source(path)

    book = openpyxl.load_workbook(preview_source, read_only=False, data_only=True)
    try:
        sheet = book.worksheets[sheet_index]
        values = sheet
        # Some valid workbooks (notably registers exported by third-party
        # systems) have no stored dimension in one of the loaded views.
        # openpyxl then returns None, which must mean an empty 1x1 sheet here,
        # not an unhandled server exception and a blank Launcher screen.
        sheet_max_row = int(sheet.max_row or 1)
        values_max_row = int(values.max_row or 1)
        sheet_max_column = int(sheet.max_column or 1)
        values_max_column = int(values.max_column or 1)
        max_column = min(max(sheet_max_column, values_max_column), MAX_XLSX_COLS)
        columns = [
            column
            for column in range(1, max_column + 1)
            if not sheet.column_dimensions[openpyxl.utils.get_column_letter(column)].hidden
        ]
        column_pixels = {}
        for column in columns:
            width = sheet.column_dimensions[openpyxl.utils.get_column_letter(column)].width or 8.43
            column_pixels[column] = max(4, min(720, round(width * 7 + 5)))

        merged_start = {}
        merged_skip = set()
        for area in sheet.merged_cells.ranges:
            shown_columns = [column for column in columns if area.min_col <= column <= area.max_col]
            if not shown_columns:
                continue
            start = (area.min_row, shown_columns[0])
            merged_start[start] = (len(shown_columns), area.max_row - area.min_row + 1)
            for row in range(area.min_row, area.max_row + 1):
                for column in shown_columns:
                    if (row, column) != start:
                        merged_skip.add((row, column))

        first_value_row = 0
        last_value_row = 0
        for row in range(1, min(values_max_row, MAX_XLSX_ROWS) + 1):
            if any(values.cell(row, column).value is not None for column in columns):
                if not first_value_row:
                    first_value_row = row
                last_value_row = row
        last_merge_row = max((area.max_row for area in sheet.merged_cells.ranges), default=0)
        rendered_rows = min(max(last_value_row, last_merge_row), MAX_XLSX_ROWS)
        was_limited = max(sheet_max_row, values_max_row) > MAX_XLSX_ROWS

        first_rendered_row = first_value_row or 1
        rows = []
        for row in range(first_rendered_row, rendered_rows + 1):
            if sheet.row_dimensions[row].hidden:
                continue
            cells = []
            for column in columns:
                if (row, column) in merged_skip:
                    continue
                cell = sheet.cell(row, column)
                raw_value = values.cell(row, column).value
                value = "" if raw_value is None else str(raw_value)
                css = []
                # A single neutral grid is supplied by the page stylesheet.
                # Re-emitting four border declarations for every cell made a
                # 2,000-row workbook produce 8–9 MB of HTML and a long white
                # screen before Chrome could paint it.  This is a navigator,
                # so fast first paint has priority over reproducing every
                # individual spreadsheet border colour.
                if cell.font.bold:
                    css.append("font-weight:500")
                if cell.font.italic:
                    css.append("font-style:italic")
                # Defaults are already declared once in the HTML stylesheet.
                # Repeating Calibri 11px on every one of 52,000 cells turns a
                # normal workbook into a multi-megabyte page that opens white.
                if cell.font.sz and round(cell.font.sz) != 11:
                    css.append(f"font-size:{max(6, min(32, cell.font.sz))}px")
                if cell.font.name and cell.font.name.casefold() not in {"calibri", "arial", "segoe ui"}:
                    css.append(f"font-family:{html.escape(cell.font.name, quote=True)}")
                font_color = excel_cell_color(cell.font.color)
                if font_color:
                    css.append(f"color:{font_color}")
                fill = excel_cell_color(cell.fill.fgColor)
                if cell.fill.fill_type == "solid" and fill:
                    css.append(f"background:{fill}")
                if cell.alignment.horizontal in {"left", "center", "right"}:
                    css.append(f"text-align:{cell.alignment.horizontal}")
                if cell.alignment.vertical in {"top", "center", "bottom"}:
                    css.append(f"vertical-align:{cell.alignment.vertical}")
                if cell.alignment.wrap_text is False:
                    css.append("white-space:pre;overflow:hidden")
                colspan, rowspan = merged_start.get((row, column), (1, 1))
                span = (f' colspan="{colspan}"' if colspan > 1 else "") + (f' rowspan="{rowspan}"' if rowspan > 1 else "")
                cells.append(f'<td{span} style="{";".join(css)}">{html.escape(value)}</td>')
            authored_height = sheet.row_dimensions[row].height
            row_style = f' style="height:{max(1, round(authored_height * 1.33))}px"' if authored_height else ""
            rows.append(f"<tr{row_style}><th>{row}</th>{''.join(cells)}</tr>")

        cols = '<col class="row-number">' + "".join(
            f'<col style="width:{column_pixels[column]}px">' for column in columns
        )
        table_width = 38 + sum(column_pixels.values())
        notice = (
            f"Показаны первые {MAX_XLSX_ROWS:,} строк. Полный рабочий файл откройте в Excel."
            if was_limited else "Просмотр без редактирования"
        )
        rendered = {
            "name": sheet.title,
            "rows": max(0, rendered_rows - first_rendered_row + 1),
            "firstRow": first_rendered_row,
            "columns": len(columns),
            "limited": was_limited,
        }
        page = f'''<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>{html.escape(sheet.title)}</title>
<style>
html,body{{margin:0;min-width:max-content;background:#fff;color:#20262d;font:11px "Segoe UI",Arial,sans-serif;overflow:auto}}
#sheet-canvas{{position:relative;transform-origin:0 0}}#sheet{{position:absolute;left:0;top:0;transform-origin:50% 50%}}
.notice{{position:sticky;top:0;z-index:2;padding:6px 10px;border-bottom:1px solid #d3dde4;background:#f7fafc;color:#607080;font-size:11px}}
table{{border-collapse:collapse;table-layout:fixed;width:{table_width}px}}col.row-number{{width:38px}}
th,td{{box-sizing:border-box;border:1px solid #cbd5dc;padding:2px 4px;vertical-align:top;white-space:pre-wrap;overflow-wrap:break-word}}
th{{position:sticky;left:0;z-index:1;background:#f1f5f7;color:#657687;font:10px "Segoe UI",Arial,sans-serif;text-align:right}}td{{overflow:hidden}}
</style></head><body><div id="sheet-canvas"><div id="sheet"><table><colgroup>{cols}</colgroup><tbody>{''.join(rows)}</tbody></table></div></div>
<script>const canvas=document.getElementById('sheet-canvas'),sheet=document.getElementById('sheet');let width=0,height=0,padding=0,scale=1,rotation=0,hand=true,dragging=false,startX=0,startY=0,startLeft=0,startTop=0;function dimensions(){{return Math.abs(rotation%180)===90?{{width:height,height:width}}:{{width,height}}}}function zoom(value){{if(!width){{width=sheet.offsetWidth;height=sheet.offsetHeight}}scale=Math.max(.35,Math.min(3,value));padding=Math.max(innerWidth,innerHeight);const size=dimensions();canvas.style.width=(size.width*scale+padding*2)+'px';canvas.style.height=(size.height*scale+padding*2)+'px';sheet.style.transform='translate('+(padding+size.width*scale/2-width/2)+'px,'+(padding+size.height*scale/2-height/2)+'px) rotate('+rotation+'deg) scale('+scale+')'}}function fit(){{if(!width)zoom(1);const size=dimensions(),value=Math.min(1,(innerWidth-48)/size.width,(innerHeight-48)/size.height);zoom(value);requestAnimationFrame(()=>{{scrollTo(padding,padding);parent.postMessage({{type:'launcher-sheet-fitted',value:scale}},'*')}})}}function cursor(){{document.body.style.cursor=hand?(dragging?'grabbing':'grab'):'default'}}addEventListener('load',()=>{{zoom(1);cursor();requestAnimationFrame(()=>scrollTo(padding,padding))}});addEventListener('wheel',event=>{{if(!event.ctrlKey)return;event.preventDefault();zoom(scale*(event.deltaY<0?1.12:.89))}},{{passive:false}});addEventListener('pointerdown',event=>{{if(!hand||event.button!==0)return;dragging=true;startX=event.clientX;startY=event.clientY;startLeft=scrollX;startTop=scrollY;document.body.setPointerCapture?.(event.pointerId);cursor();event.preventDefault()}});addEventListener('pointermove',event=>{{if(!dragging)return;scrollTo(startLeft-(event.clientX-startX),startTop-(event.clientY-startY))}});addEventListener('pointerup',event=>{{if(!dragging)return;dragging=false;document.body.releasePointerCapture?.(event.pointerId);cursor()}});addEventListener('pointercancel',()=>{{dragging=false;cursor()}});addEventListener('keydown',event=>{{if(event.key==='Escape')parent.postMessage({{type:'launcher-escape'}},'*')}});addEventListener('dblclick',()=>{{parent.postMessage({{type:'launcher-toggle-full-view'}},'*')}});addEventListener('message',event=>{{if(!event.data)return;if(event.data.type==='launcher-sheet-zoom')zoom(event.data.value);if(event.data.type==='launcher-sheet-fit')fit();if(event.data.type==='launcher-sheet-rotate'){{rotation=((Number(event.data.value)||0)%360+360)%360;zoom(scale)}}if(event.data.type==='launcher-sheet-hand'){{hand=Boolean(event.data.value);dragging=false;cursor()}}}});</script></body></html>'''
        return page, rendered
    finally:
        try:
            book.close()
        except Exception:
            pass


def excel_sheet_preview(path: Path, sheet_index: int) -> dict:
    """Build one sheet on demand; opening a book must not wait for every tab."""
    cache_dir = excel_html_cache_dir(path)
    cache_dir.mkdir(parents=True, exist_ok=True)
    # A new cache version makes existing pages harmless without deleting a
    # user's cache, including the viewer controls embedded in this HTML.
    output = cache_dir / f"sheet-{sheet_index + 1}-v10.html"
    metadata_path = cache_dir / f"sheet-{sheet_index + 1}-v10.json"
    metadata = None
    if output.exists() and metadata_path.exists():
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            metadata = None
    if metadata is None:
        page, metadata = excel_sheet_html(path, sheet_index)
        output.write_text(page, encoding="utf-8")
        metadata_path.write_text(json.dumps(metadata, ensure_ascii=False), encoding="utf-8")
    return {
        "index": sheet_index,
        "url": f"/cache/excel/html/{cache_dir.name}/{output.name}",
        **metadata,
    }


def excel_thumbnail_preview(path: Path) -> str:
    """Build a visual card from the exact same HTML sheet as the large view.

    A workbook must have one visual source of truth.  The former thumbnail
    constructed a second, simplified table with its own font and row rules;
    it could therefore disagree with the readable sheet.  The card is now a
    scaled viewport of the cached HTML sheet used by the full viewer.
    """
    if openpyxl is None:
        raise RuntimeError("Для HTML-просмотра Excel нужен пакет openpyxl")
    # The rail already places this page in a scaled, clipped iframe.  Returning
    # the sheet directly means its geometry, fonts and initial position are
    # exactly the same in the card and in the large viewer.
    return excel_sheet_preview(path, 0)["url"]


def excel_workbook_preview(path: Path) -> dict:
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"Excel-файл не найден: {path}")
    if path.name.startswith("~$") or path.name.startswith(".~"):
        raise ValueError(f"Временный файл блокировки Office: {path.name}")
    if not is_excel_file(path):
        raise ValueError(f"Это не Excel-файл: {path}")
    if openpyxl is None:
        raise RuntimeError("Для HTML-просмотра Excel нужен пакет openpyxl")

    preview_source = excel_preview_source(path)
    sheet_names = None
    if zipfile.is_zipfile(preview_source):
        try:
            with zipfile.ZipFile(preview_source, "r") as z:
                tree = ET.fromstring(z.read("xl/workbook.xml"))
                ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                extracted = [el.attrib.get("name", "") for el in tree.findall(".//main:sheet", ns)]
                if extracted and all(extracted):
                    sheet_names = extracted
        except Exception:
            sheet_names = None
    if sheet_names is None:
        book = openpyxl.load_workbook(preview_source, read_only=True, data_only=False)
        try:
            sheet_names = list(book.sheetnames)
        finally:
            book.close()

    sheets = [{"index": index, "name": name} for index, name in enumerate(sheet_names)]
    return {
        "name": path.name,
        "path": str(path),
        "sheets": sheets,
        "maxRows": MAX_XLSX_ROWS,
        "cacheKey": excel_html_cache_dir(path).name,
        "thumbnailUrl": excel_thumbnail_preview(path),
    }


def excel_to_pdf(path: Path) -> tuple[Path, bool]:
    if not path.exists():
        raise FileNotFoundError(f"Excel-файл не найден: {path}")
    if not path.is_file() or not is_excel_file(path):
        raise ValueError(f"Это не Excel-файл: {path}")
    if not EXCEL_CONVERT_SCRIPT.exists():
        raise RuntimeError(f"Скрипт конвертации Excel не найден: {EXCEL_CONVERT_SCRIPT}")

    key = file_cache_key(path, "excel-pdf")
    target_dir = EXCEL_CACHE_DIR / key
    target_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = target_dir / f"{path.stem}.pdf"
    manifest_path = target_dir / "manifest.json"

    cached = False
    if pdf_path.exists() and pdf_path.stat().st_size > 0 and manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            cached = (
                manifest.get("sourcePath") == str(path)
                and manifest.get("cacheKey") == key
                and manifest.get("sourceMtimeNs") == path.stat().st_mtime_ns
                and manifest.get("sourceSize") == path.stat().st_size
            )
        except (OSError, json.JSONDecodeError):
            cached = False

    if cached:
        return pdf_path, True

    if pdf_path.exists():
        pdf_path.unlink()

    process = subprocess.run(
        [
            "powershell",
            "-STA",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(EXCEL_CONVERT_SCRIPT),
            "-InputPath",
            str(path),
            "-OutputPath",
            str(pdf_path),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=EXCEL_CONVERT_TIMEOUT_SECONDS,
        **hidden_process_kwargs(),
    )
    if process.returncode != 0:
        message = process.stderr.strip() or process.stdout.strip() or "Excel не смог экспортировать книгу в PDF"
        raise RuntimeError(message)
    if not pdf_path.exists() or pdf_path.stat().st_size <= 0:
        raise RuntimeError("Excel не создал PDF для preview")

    manifest_path.write_text(
        json.dumps(
            {
                "sourcePath": str(path),
                "sourceName": path.name,
                "sourceMtimeNs": path.stat().st_mtime_ns,
                "sourceSize": path.stat().st_size,
                "cacheKey": key,
                "pdfPath": str(pdf_path),
                "convertedAt": datetime.now().isoformat(timespec="seconds"),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return pdf_path, False


def render_excel(path: Path, dpi: int = DEFAULT_PDF_DPI, first_page_only: bool = False) -> dict:
    if not path.exists():
        return {
            "name": path.name,
            "path": str(path),
            "sourcePath": str(path),
            "sourceName": path.name,
            "sourceType": file_extension(path),
            "dpi": dpi,
            "pages": 0,
            "renderedPages": 0,
            "cacheKey": "",
            "cacheHit": False,
            "cacheHitPages": 0,
            "newRenderedPages": 0,
            "is_missing": True,
            "errors": [{"error": f"Таблица не найдена на диске: {path.name}", "is_missing": True}],
            "items": [],
        }
    pdf_path, convert_cache_hit = excel_to_pdf(path)
    document = render_pdf(pdf_path, dpi=dpi, first_page_only=first_page_only)
    document["sourcePath"] = str(path)
    document["sourceName"] = path.name
    document["sourceType"] = file_extension(path)
    document["convertedPdfPath"] = str(pdf_path)
    document["convertCacheHit"] = convert_cache_hit
    return document


# PyMuPDF (fitz) дважды ронял процесс при параллельном рендеринге из
# нескольких потоков (Application Error, access violation в нативном коде).
# Весь нативный fitz-рендеринг идёт строго по очереди; кэш-возвраты и
# Poppler-путь (отдельные процессы) блокировкой не затрагиваются.
_FITZ_RENDER_LOCK = threading.Lock()


def render_pdf(
    path: Path,
    dpi: int = DEFAULT_PDF_DPI,
    page_timeout_seconds: int = PDF_PAGE_TIMEOUT_SECONDS,
    first_page_only: bool = False,
) -> dict:
    if not path.exists():
        return {
            "name": path.name,
            "path": str(path),
            "sourcePath": str(path),
            "dpi": dpi,
            "pages": 0,
            "renderedPages": 0,
            "cacheKey": "",
            "cacheHit": False,
            "cacheHitPages": 0,
            "newRenderedPages": 0,
            "is_missing": True,
            "errors": [{"error": f"PDF не найден на диске: {path.name}", "is_missing": True}],
            "items": [],
        }
    if not path.is_file() or path.suffix.casefold() != ".pdf":
        raise ValueError(f"Это не PDF-файл: {path}")

    key = pdf_cache_key(path, dpi)
    target_dir = PDF_CACHE_DIR / key
    target_dir.mkdir(parents=True, exist_ok=True)

    png1 = target_dir / "page-1.png"
    cached_manifest = read_pdf_cache_manifest(path, dpi, key, target_dir)
    if cached_manifest:
        cached_items = cached_manifest["items"]
        cached_errors = cached_manifest.get("errors", [])
        cached_total = cached_manifest.get("pages", len(cached_items)) or 0
        if first_page_only and png1.exists() and png1.stat().st_size > 0:
            item1 = cached_items[0] if cached_items else pdf_page_item(path, key, 1, png1)
            return {
                "name": path.name,
                "path": str(path),
                "dpi": dpi,
                "pages": cached_total or 1,
                "renderedPages": 1,
                "cacheKey": key,
                "cacheHit": True,
                "cacheHitPages": 1,
                "newRenderedPages": 0,
                "errors": cached_errors,
                "items": [item1],
            }
        if not first_page_only and cached_total and len(cached_items) >= cached_total:
            return {
                "name": path.name,
                "path": str(path),
                "dpi": dpi,
                "pages": cached_total,
                "renderedPages": len(cached_items),
                "cacheKey": key,
                "cacheHit": True,
                "cacheHitPages": len(cached_items),
                "newRenderedPages": 0,
                "errors": cached_errors,
                "items": cached_items,
            }

    # 1. Быстрый рендеринг через PyMuPDF (fitz) без создания сотен процессов
    _FITZ_RENDER_LOCK.acquire()
    try:
        import fitz
        try:
            doc = fitz.open(str(path))
        except Exception:
            doc = fitz.open(stream=path.read_bytes(), filetype="pdf")
        page_count = len(doc)
        pages = []
        errors = []
        rendered_count = 0
        cache_hit_count = 0

        pages_to_render = 1 if first_page_only else page_count
        for page_idx in range(pages_to_render):
            page_num = page_idx + 1
            png = target_dir / f"page-{page_num}.png"
            if png.exists() and png.stat().st_size > 0:
                cache_hit_count += 1
            else:
                try:
                    pdf_page = doc[page_idx]
                    pix = pdf_page.get_pixmap(dpi=dpi)
                    pix.save(str(png))
                    rendered_count += 1
                except Exception as err:
                    errors.append({"page": page_num, "error": str(err)})
                    continue

            if png.exists() and png.stat().st_size > 0:
                pages.append(pdf_page_item(path, key, page_num, png))

        doc.close()

        if pages:
            write_pdf_cache_manifest(path, dpi, key, target_dir, page_count, pages, errors)
            return {
                "name": path.name,
                "path": str(path),
                "sourcePath": str(path),
                "dpi": dpi,
                "pages": page_count,
                "renderedPages": len(pages),
                "cacheKey": key,
                "cacheHit": cache_hit_count == pages_to_render,
                "cacheHitPages": cache_hit_count,
                "newRenderedPages": rendered_count,
                "errors": errors,
                "items": pages,
            }
    except Exception as fitz_err:
        pass
    finally:
        _FITZ_RENDER_LOCK.release()

    # 2. Запасной рендеринг через Poppler (если PyMuPDF недоступен или выдал сбой)
    pdftoppm = poppler_tool("pdftoppm")
    if not pdftoppm:
        raise RuntimeError("pdftoppm не найден. Нужен Poppler из runtime.")

    page_count = pdf_page_count(path)
    started_at = time.monotonic()

    pages = []
    errors = []
    rendered_count = 0
    cache_hit_count = 0

    pages_to_render = 1 if first_page_only else page_count
    for page in range(1, pages_to_render + 1):
        if time.monotonic() - started_at > PDF_DOCUMENT_TIMEOUT_SECONDS:
            for skipped_page in range(page, pages_to_render + 1):
                errors.append({"page": skipped_page, "error": f"PDF остановлен по лимиту {PDF_DOCUMENT_TIMEOUT_SECONDS} сек. на файл."})
            break

        png = target_dir / f"page-{page}.png"
        if png.exists() and png.stat().st_size > 0:
            cache_hit_count += 1
        else:
            for stale in target_dir.glob(f"page-{page}*.png"):
                stale.unlink()
            prefix = target_dir / f"page-{page}"
            try:
                result = run_poppler(
                    [
                        pdftoppm,
                        "-f",
                        str(page),
                        "-l",
                        str(page),
                        "-singlefile",
                        "-r",
                        str(dpi),
                        "-png",
                        str(path),
                        str(prefix),
                    ],
                        timeout=page_timeout_seconds,
                )
                candidate = target_dir / f"page-{page}.png"
                if result.returncode != 0:
                    raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "pdftoppm не смог отрендерить страницу")
                if not candidate.exists() or candidate.stat().st_size <= 0:
                    raise RuntimeError("pdftoppm не создал PNG страницы")
                rendered_count += 1
            except subprocess.TimeoutExpired:
                errors.append({"page": page, "error": f"Таймаут рендера страницы {page}: {page_timeout_seconds} сек."})
                continue
            except (RuntimeError, OSError) as error:
                errors.append({"page": page, "error": str(error)})
                continue

        if png.exists() and png.stat().st_size > 0:
            pages.append(pdf_page_item(path, key, page, png))

    if not pages:
        result = run_poppler(
            [pdftoppm, "-v"],
            timeout=10,
        )
        tool_version = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"Не удалось отрендерить ни одной страницы PDF. {tool_version}. Ошибки: {errors}")

    write_pdf_cache_manifest(path, dpi, key, target_dir, page_count, pages, errors)
    return {
        "name": path.name,
        "path": str(path),
        "sourcePath": str(path),
        "dpi": dpi,
        "pages": page_count,
        "renderedPages": len(pages),
        "cacheKey": key,
        "cacheHit": cache_hit_count == page_count,
        "cacheHitPages": cache_hit_count,
        "newRenderedPages": rendered_count,
        "errors": errors,
        "items": pages,
    }


def render_pdf_page(path: Path, page: int, dpi: int = DEFAULT_PDF_DPI, page_timeout_seconds: int = PDF_PAGE_TIMEOUT_SECONDS) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"PDF не найден: {path}")
    if not path.is_file() or path.suffix.casefold() != ".pdf":
        raise ValueError(f"Это не PDF-файл: {path}")

    key = pdf_cache_key(path, dpi)
    target_dir = PDF_CACHE_DIR / key
    target_dir.mkdir(parents=True, exist_ok=True)
    png = target_dir / f"page-{page}.png"
    cache_hit = png.exists() and png.stat().st_size > 0

    page_count = 0
    cached_manifest = read_pdf_cache_manifest(path, dpi, key, target_dir)
    if cached_manifest:
        page_count = cached_manifest.get("pages", 0)

    # 1. Если страница уже в кэше — мгновенный возврат без сетевых обращений
    if cache_hit and page_count > 0:
        item = pdf_page_item(path, key, page, png)
        return {
            "name": path.name,
            "path": str(path),
            "dpi": dpi,
            "pages": page_count,
            "page": page,
            "cacheKey": key,
            "cacheHit": True,
            "item": item,
        }

    # 2. Быстрый рендеринг через PyMuPDF (одно открытие документа)
    if not cache_hit:
        _FITZ_RENDER_LOCK.acquire()
        try:
            import fitz
            doc = fitz.open(str(path))
            page_count = len(doc)
            if page < 1 or page > page_count:
                doc.close()
                raise ValueError(f"Страница {page} вне диапазона 1-{page_count}")
            pdf_page = doc[page - 1]
            pix = pdf_page.get_pixmap(dpi=dpi)
            pix.save(str(png))
            doc.close()
            cache_hit = png.exists() and png.stat().st_size > 0
        except ValueError:
            raise
        except Exception as fitz_err:
            pass
        finally:
            _FITZ_RENDER_LOCK.release()

    # 3. Запасной рендеринг через Poppler (только если PyMuPDF не сработал)
    if not cache_hit:
        pdftoppm = poppler_tool("pdftoppm")
        if not pdftoppm:
            raise RuntimeError("Рендеринг страницы не удался (PyMuPDF и Poppler недоступны)")
        if not page_count:
            page_count = pdf_page_count(path)
        if page < 1 or page > page_count:
            raise ValueError(f"Страница {page} вне диапазона 1-{page_count}")
        for stale in target_dir.glob(f"page-{page}*.png"):
            stale.unlink()
        prefix = target_dir / f"page-{page}"
        try:
            result = run_poppler(
                [
                    pdftoppm,
                    "-f", str(page),
                    "-l", str(page),
                    "-singlefile",
                    "-r", str(dpi),
                    "-png",
                    str(path),
                    str(prefix),
                ],
                timeout=page_timeout_seconds,
            )
        except subprocess.TimeoutExpired as error:
            raise RuntimeError(f"Таймаут рендера страницы {page}: {page_timeout_seconds} сек.") from error
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "pdftoppm не смог отрендерить страницу")
        if not png.exists() or png.stat().st_size <= 0:
            raise RuntimeError("pdftoppm не создал PNG страницы")

    if not page_count:
        page_count = page

    item = pdf_page_item(path, key, page, png)
    cached_manifest = read_pdf_cache_manifest(path, dpi, key, target_dir)
    items_by_page = {existing["page"]: existing for existing in (cached_manifest or {}).get("items", [])}
    items_by_page[page] = item
    items = [items_by_page[index] for index in sorted(items_by_page)]
    write_pdf_cache_manifest(path, dpi, key, target_dir, page_count, items, (cached_manifest or {}).get("errors", []))

    return {
        "name": path.name,
        "path": str(path),
        "dpi": dpi,
        "pages": page_count,
        "page": page,
        "cacheKey": key,
        "cacheHit": cache_hit,
        "item": item,
    }



try:
    import winreg
except ImportError:
    winreg = None


def _get_command_from_progid(progid: str) -> str:
    if not winreg or not progid:
        return ""
    keys_to_try = [
        f"{progid}\\shell\\open\\command",
        f"{progid}\\shell\\Open\\command",
        f"{progid}\\shell\\edit\\command",
    ]
    for subkey in keys_to_try:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, subkey) as k:
                val, _ = winreg.QueryValueEx(k, "")
                cleaned = _clean_configured_exe_path(val)
                if cleaned:
                    if Path(cleaned).name.casefold() in FORBIDDEN_AUTODETECT_EXES:
                        return ""
                    return cleaned
        except OSError:
            pass
    return ""


def detect_windows_app_for_ext(ext: str) -> str:
    """Определяет установленную в Windows программу по умолчанию для расширения."""
    if not winreg:
        return ""
    norm_ext = normalize_file_extension(ext)
    if not norm_ext:
        return ""

    candidate = ""
    # 1. UserChoice (Windows 10 / 11 modern default)
    user_choice = f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\{norm_ext}\\UserChoice"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, user_choice) as k:
            progid, _ = winreg.QueryValueEx(k, "ProgId")
            res = _get_command_from_progid(progid)
            if res:
                candidate = res
    except OSError:
        pass

    # 2. HKEY_CLASSES_ROOT default ProgID
    if not candidate:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, norm_ext) as k:
                progid = winreg.QueryValue(k, "")
                res = _get_command_from_progid(progid)
                if res:
                    candidate = res
        except OSError:
            pass

    # 3. OpenWithProgids
    if not candidate:
        openwith = f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\{norm_ext}\\OpenWithProgids"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, openwith) as k:
                i = 0
                while True:
                    try:
                        name, _, _ = winreg.EnumValue(k, i)
                        res = _get_command_from_progid(name)
                        if res:
                            candidate = res
                            break
                        i += 1
                    except (OSError, ValueError, TypeError):
                        break
        except OSError:
            pass

    if candidate:
        exe_file = Path(candidate).name.casefold()
        if exe_file in FORBIDDEN_AUTODETECT_EXES:
            return ""
        if Path(candidate).is_file():
            return candidate

    return ""

class LauncherHandler(BaseHTTPRequestHandler):
    server_version = "FEngineeringLauncherV3/0.1"

    def log_message(self, format: str, *args: object) -> None:
        logs_dir = RUNTIME_DIR / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        with (logs_dir / "server.log").open("a", encoding="utf-8") as log:
            log.write("%s - %s\n" % (self.log_date_time_string(), format % args))

    def send_json(self, status: HTTPStatus, payload: dict) -> None:
        try:
            body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError):
            pass

    def read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0") or "0")
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8-sig"))

    def handle_choose_exe(self, initial_path: str = "") -> None:
        try:
            choose_script = REPO_ROOT / "scripts" / "choose_exe.py"
            choose_env = dict(os.environ)
            choose_env["PYTHONUTF8"] = "1"
            choose_env["PYTHONIOENCODING"] = "utf-8"
            args = [sys.executable, str(choose_script)]
            if initial_path:
                args.append(initial_path)
            creation_flags = 0
            if os.name == "nt":
                creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
            proc = subprocess.run(
                args,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=120,
                env=choose_env,
                creationflags=creation_flags,
            )
            raw = proc.stdout.strip()
            selected = ""
            if raw:
                try:
                    parsed_json = json.loads(raw)
                    if isinstance(parsed_json, str):
                        selected = parsed_json
                except Exception:
                    selected = raw.strip('"')
            self.send_json(HTTPStatus.OK, {"path": selected})
        except subprocess.TimeoutExpired:
            self.send_json(
                HTTPStatus.OK,
                {"path": "", "timedOut": True, "error": "Диалог выбора программы не ответил за 120 секунд."},
            )
        except Exception as error:
            self.send_json(HTTPStatus.OK, {"path": "", "error": f"Ошибка вызова диалога: {error}"})

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/choose-exe":
            self.handle_choose_exe(initial_path="")
            return
        if parsed.path == "/api/health":
            self.send_json(
                HTTPStatus.OK,
                {
                    "status": "ok",
                    "service": "F-Engineering Launcher v3",
                    "version": VERSION,
                    "repoRoot": str(REPO_ROOT),
                    "runtime": str(RUNTIME_DIR),
                    "wordPreview": True,
                    "excelPreview": True,
                },
            )
            return

        if parsed.path == "/api/version":
            info = get_launcher_version_info()
            self.send_json(
                HTTPStatus.OK,
                {"status": "ok", "version": info["version"], "branch": info["branch"]},
            )
            return

        if parsed.path == "/api/updates/check":
            from app.updater import check_for_updates

            res = check_for_updates()
            self.send_json(HTTPStatus.OK, res)
            return

        if parsed.path == "/api/config/apps":
            self.send_json(HTTPStatus.OK, load_native_apps_config())
            return

        if parsed.path == "/api/file/raw":
            try:
                params = parse_qs(parsed.query)
                raw_path = params.get("path", [""])[0].strip()
                if not raw_path:
                    self.send_error(HTTPStatus.BAD_REQUEST, "Missing path")
                    return
                target = Path(raw_path).expanduser().resolve()
                if not target.exists() or not target.is_file():
                    self.send_error(HTTPStatus.NOT_FOUND, "File not found")
                    return
                content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
                body = target.read_bytes()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as error:
                self.send_error(HTTPStatus.BAD_REQUEST, str(error))
            return

        if parsed.path == "/api/dwg/thumbnail":
            try:
                params = parse_qs(parsed.query)
                raw_path = params.get("path", [""])[0].strip()
                if not raw_path:
                    self.send_error(HTTPStatus.BAD_REQUEST, "Missing path")
                    return
                target = Path(raw_path).expanduser().resolve()
                if not target.exists() or not target.is_file():
                    self.send_error(HTTPStatus.NOT_FOUND, "File not found")
                    return
                try:
                    from app.backend.dwg_engine import extract_raw_thumbnail_from_dwg, _generate_placeholder_png
                except ImportError:
                    from dwg_engine import extract_raw_thumbnail_from_dwg, _generate_placeholder_png
                thumb = extract_raw_thumbnail_from_dwg(target)
                body = thumb[0] if thumb else _generate_placeholder_png(target.name, "Model")
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "image/png")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as error:
                self.send_error(HTTPStatus.BAD_REQUEST, str(error))
            return

        if parsed.path == "/api/objects":
            self.send_json(HTTPStatus.OK, {"items": list_object_summaries()})
            return

        if parsed.path.startswith("/api/objects/"):
            object_id = unquote(parsed.path.removeprefix("/api/objects/")).strip()
            manifest = load_manifest(object_id)
            if manifest:
                self.send_json(HTTPStatus.OK, manifest)
            else:
                self.send_json(HTTPStatus.NOT_FOUND, {"error": "Объект не найден в manifest-хранилище"})
            return

        if parsed.path.startswith("/cache/"):
            self.serve_cache(parsed.path)
            return

        self.serve_static(parsed.path)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/objects/import":
            raw_path = ""
            try:
                body = self.read_json()
                raw_path = str(body.get("path", ""))
                manifest = scan_object(raw_path)
                self.send_json(HTTPStatus.OK, manifest)
            except (ValueError, FileNotFoundError, NotADirectoryError, OSError, json.JSONDecodeError) as error:
                append_import_log(
                    {
                        "at": datetime.now().isoformat(timespec="seconds"),
                        "status": "error",
                        "path": raw_path,
                        "error": str(error),
                    }
                )
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            except Exception as error:
                append_import_log(
                    {
                        "at": datetime.now().isoformat(timespec="seconds"),
                        "status": "unexpected-error",
                        "path": raw_path,
                        "error": "%s: %s" % (type(error).__name__, error),
                        "traceback": traceback.format_exc(limit=12),
                    }
                )
                self.send_json(
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                    {"error": "Загрузка не удалась (%s). Подробности записаны в runtime/logs/import.jsonl." % type(error).__name__},
                )
            return

        if parsed.path == "/api/objects/exclude":
            try:
                body = self.read_json()
                object_id = str(body.get("id", "")).strip()
                if not object_id:
                    raise ValueError("Не выбран объект для исключения")
                target = MANIFESTS_DIR / f"{object_id}.json"
                if target.exists():
                    target.unlink()
                self.send_json(HTTPStatus.OK, {"ok": True, "removed": object_id})
            except (ValueError, OSError, json.JSONDecodeError) as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/objects/diff":
            try:
                body = self.read_json()
                object_id = str(body.get("id", "")).strip()
                manifest = load_manifest(object_id)
                if not manifest:
                    raise ValueError("Объект не найден")
                root = Path(manifest.get("rootPath", ""))
                if not root.exists():
                    raise FileNotFoundError(f"Папка объекта не найдена: {root}")
                tree, extension_counts, folder_count, file_count = build_tree(root)
                last_diff = diff_trees(manifest.get("tree"), tree)
                self.send_json(
                    HTTPStatus.OK,
                    {
                        "id": object_id,
                        "scannedAt": datetime.now().isoformat(timespec="seconds"),
                        "statistics": {
                            "folders": folder_count,
                            "files": file_count,
                            "extensions": extension_counts,
                        },
                        "lastDiff": last_diff,
                    },
                )
            except (ValueError, FileNotFoundError, OSError, json.JSONDecodeError) as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/choose-folder":
            # Контракт docs/LOAD_BUTTON_CONTRACT.md: один PowerShell
            # FolderBrowserDialog, строго JSON в ответе.
            try:
                ps_command = (
                    "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false); "
                    "Add-Type -AssemblyName System.Windows.Forms; "
                    "$dlg = New-Object System.Windows.Forms.FolderBrowserDialog; "
                    '$dlg.Description = "Select object folder - FLauncher"; '
                    "$dlg.ShowNewFolderButton = $false; "
                    "$owner = New-Object System.Windows.Forms.Form; "
                    "$owner.TopMost = $true; $owner.ShowInTaskbar = $false; "
                    "$null = $owner.Handle; "
                    "try { if ($dlg.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) "
                    "{ [Console]::WriteLine($dlg.SelectedPath) } } "
                    "finally { $owner.Dispose(); $dlg.Dispose() }"
                )
                # Доводчик: диалог должен открыться главным окном, а не значком в фоне.
                _bring_window_to_front(None, "#32770", 115.0, "Select object folder")
                # ВАЖНО: без hidden_process_kwargs() — флаг STARTF_USESHOWWINDOW
                # с wShowWindow=0 прячет модальное окно FolderBrowserDialog.
                # Диалог выбора папки обязан открываться поверх лаунчера.
                creation_flags = 0
                if os.name == "nt":
                    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
                proc = subprocess.run(
                    [
                        "powershell.exe",
                        "-STA",
                        "-NoProfile",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-Command",
                        ps_command,
                    ],
                    capture_output=True,
                    text=False,
                    timeout=120,
                    creationflags=creation_flags,
                )
                selected = decode_folder_dialog_output(proc.stdout)
                if selected:
                    self.send_json(HTTPStatus.OK, {"path": selected})
                else:
                    self.send_json(HTTPStatus.OK, {"path": ""})
            except subprocess.TimeoutExpired:
                self.send_json(
                    HTTPStatus.OK,
                    {
                        "path": "",
                        "error": "Системный диалог выбора папки Windows не ответил за 120 секунд. Попробуйте ещё раз.",
                    },
                )
            except Exception as error:
                self.send_json(HTTPStatus.OK, {"path": "", "error": f"Системный диалог выбора папки недоступен: {error}"})
            return

        if parsed.path == "/api/config/apps":
            try:
                body = self.read_json()
                if not isinstance(body, dict):
                    raise ValueError("Ожидался JSON-объект")
                save_native_apps_config(body)
                self.send_json(HTTPStatus.OK, {"ok": True, "config": load_native_apps_config()})
            except Exception as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/choose-exe":
            initial_path = ""
            try:
                body = self.read_json()
                initial_path = str(body.get("current", "") or body.get("path", "")).strip()
            except Exception:
                pass
            self.handle_choose_exe(initial_path=initial_path)
            return

        if parsed.path == "/api/config/apps/autodetect":
            try:
                target_exts = [
                    ".pdf", ".dwg", ".dxf", ".xlsx", ".xls", ".xlsm", ".csv", ".ods",
                    ".docx", ".doc", ".rtf", ".odt", ".pptx", ".ppt", ".txt",
                    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp",
                    ".mp4", ".avi", ".mov", ".mkv", ".mp3", ".wav", ".zip", ".rar", ".7z"
                ]
                detected = {}
                for ext in target_exts:
                    app_path = detect_windows_app_for_ext(ext)
                    if app_path:
                        detected[ext] = app_path
                self.send_json(HTTPStatus.OK, {"ok": True, "detected": detected})
            except Exception as error:
                self.send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": str(error)})
            return

        if parsed.path == "/api/pdf/render":
            try:
                body = self.read_json()
                raw_files = body.get("files", [])
                dpi = int(body.get("dpi") or DEFAULT_PDF_DPI)
                if dpi < 72 or dpi > 600:
                    raise ValueError("DPI должен быть в диапазоне 72-600")
                if not isinstance(raw_files, list) or not raw_files:
                    raise ValueError("Не выбраны PDF-файлы для отображения")
                if len(raw_files) > 15:
                    raise ValueError("За один раз пока можно отрендерить не больше 15 PDF")
                first_page_only = bool(body.get("firstPageOnly", False))
                documents = [render_pdf(Path(str(file_path)), dpi=dpi, first_page_only=first_page_only) for file_path in raw_files]
                document_errors = [
                    {"document": document["name"], "path": document["path"], **error}
                    for document in documents
                    for error in document.get("errors", [])
                ]
                self.send_json(
                    HTTPStatus.OK,
                    {
                        "dpi": dpi,
                        "documents": documents,
                        "totalPages": sum(document["pages"] for document in documents),
                        "renderedPages": sum(document["renderedPages"] for document in documents),
                        "errors": document_errors,
                        "renderedAt": datetime.now().isoformat(timespec="seconds"),
                    },
                )
            except (ValueError, FileNotFoundError, RuntimeError, OSError, subprocess.TimeoutExpired, json.JSONDecodeError) as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/pdf/page":
            try:
                body = self.read_json()
                raw_file = str(body.get("file", "")).strip()
                page = int(body.get("page") or 1)
                dpi = int(body.get("dpi") or DEFAULT_PDF_DPI)
                if dpi < 72 or dpi > 600:
                    raise ValueError("DPI должен быть в диапазоне 72-600")
                if not raw_file:
                    raise ValueError("Не выбран PDF-файл для отображения")
                payload_page = render_pdf_page(Path(raw_file), page=page, dpi=dpi)
                self.send_json(HTTPStatus.OK, payload_page)
            except (ValueError, FileNotFoundError, RuntimeError, OSError, subprocess.TimeoutExpired, json.JSONDecodeError) as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/word/render":
            try:
                body = self.read_json()
                raw_files = body.get("files", [])
                dpi = int(body.get("dpi") or DEFAULT_PDF_DPI)
                if dpi < 72 or dpi > 600:
                    raise ValueError("DPI должен быть в диапазоне 72-600")
                if not isinstance(raw_files, list) or not raw_files:
                    raise ValueError("Не выбраны Word-файлы для отображения")
                if len(raw_files) > 10:
                    raise ValueError("За один раз пока можно отрендерить не больше 10 Word-файлов")
                first_page_only = bool(body.get("firstPageOnly", False))
                documents = [render_word(Path(str(file_path)), dpi=dpi, first_page_only=first_page_only) for file_path in raw_files]
                document_errors = [
                    {"document": document["name"], "path": document["path"], **error}
                    for document in documents
                    for error in document.get("errors", [])
                ]
                self.send_json(
                    HTTPStatus.OK,
                    {
                        "dpi": dpi,
                        "documents": documents,
                        "totalPages": sum(document["pages"] for document in documents),
                        "renderedPages": sum(document["renderedPages"] for document in documents),
                        "errors": document_errors,
                        "renderedAt": datetime.now().isoformat(timespec="seconds"),
                    },
                )
            except (ValueError, FileNotFoundError, RuntimeError, OSError, subprocess.TimeoutExpired, json.JSONDecodeError) as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/word/page":
            try:
                body = self.read_json()
                raw_file = str(body.get("file", "")).strip()
                page = int(body.get("page") or 1)
                dpi = int(body.get("dpi") or DEFAULT_PDF_DPI)
                if dpi < 72 or dpi > 600:
                    raise ValueError("DPI должен быть в диапазоне 72-600")
                if not raw_file:
                    raise ValueError("Не выбран Word-файл для отображения")
                pdf_path, convert_cache_hit = word_to_pdf(Path(raw_file))
                payload = render_pdf_page(pdf_path, page=page, dpi=dpi)
                payload["sourcePath"] = raw_file
                payload["sourceType"] = file_extension(Path(raw_file))
                payload["convertedPdfPath"] = str(pdf_path)
                payload["convertCacheHit"] = convert_cache_hit
                self.send_json(HTTPStatus.OK, payload)
            except (ValueError, FileNotFoundError, RuntimeError, OSError, subprocess.TimeoutExpired, json.JSONDecodeError) as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/dwg/model-render":
            try:
                body = self.read_json()
                raw_files = body.get("files", [])
                dpi = int(body.get("dpi") or DEFAULT_PDF_DPI)
                if dpi < 72 or dpi > 600:
                    raise ValueError("DPI должен быть в диапазоне 72-600")
                if not isinstance(raw_files, list) or not raw_files:
                    raise ValueError("Не выбраны DWG-файлы для отображения")
                if len(raw_files) > 1:
                    raise ValueError("Model Space preview пока создаётся по одному DWG-файлу")
                documents = [render_dwg_model(Path(str(file_path)), dpi=dpi) for file_path in raw_files]
                document_errors = [
                    {"document": document["name"], "path": document["path"], **error}
                    for document in documents
                    for error in document.get("errors", [])
                ]
                self.send_json(
                    HTTPStatus.OK,
                    {
                        "dpi": dpi,
                        "documents": documents,
                        "totalPages": sum(document["pages"] for document in documents),
                        "renderedPages": sum(document["renderedPages"] for document in documents),
                        "errors": document_errors,
                        "renderedAt": datetime.now().isoformat(timespec="seconds"),
                    },
                )
            except (ValueError, FileNotFoundError, RuntimeError, OSError, subprocess.TimeoutExpired, json.JSONDecodeError) as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/dwg/model-page":
            try:
                body = self.read_json()
                raw_file = str(body.get("file", "")).strip()
                page = int(body.get("page") or 1)
                dpi = int(body.get("dpi") or DEFAULT_PDF_DPI)
                if dpi < 72 or dpi > 600:
                    raise ValueError("DPI должен быть в диапазоне 72-600")
                if not raw_file:
                    raise ValueError("Не выбран DWG-файл для отображения")
                file_path = Path(raw_file)
                key = file_cache_key(file_path, "dwg-smart-cad-v2")
                cached_png = DWG_CACHE_DIR / key / f"page-{page}.png"
                manifest_file = DWG_CACHE_DIR / key / "manifest.json"
                if cached_png.exists() and cached_png.stat().st_size > 0 and manifest_file.exists():
                    self.send_json(
                        HTTPStatus.OK,
                        {
                            "page": page,
                            "name": f"{file_path.name} · стр. {page}",
                            "url": f"/cache/dwg/{key}/{cached_png.name}",
                            "bytes": cached_png.stat().st_size,
                            "cacheHit": True,
                            "sourcePath": raw_file,
                            "sourceType": "DWG",
                            "convertedPdfPath": "",
                            "convertCacheHit": True,
                        },
                    )
                    return
                pdf_path, convert_cache_hit = dwg_to_model_pdf(file_path)
                payload = render_pdf_page(pdf_path, page=page, dpi=dpi, page_timeout_seconds=DWG_MODEL_PAGE_TIMEOUT_SECONDS)
                payload["sourcePath"] = raw_file
                payload["sourceType"] = "DWG"
                payload["convertedPdfPath"] = str(pdf_path)
                payload["convertCacheHit"] = convert_cache_hit
                self.send_json(HTTPStatus.OK, payload)
            except (ValueError, FileNotFoundError, RuntimeError, OSError, subprocess.TimeoutExpired, json.JSONDecodeError) as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/excel/workbook":
            try:
                body = self.read_json()
                raw_file = str(body.get("file", "")).strip()
                if not raw_file:
                    raise ValueError("Не выбран Excel-файл для отображения")
                self.send_json(HTTPStatus.OK, excel_workbook_preview(Path(raw_file)))
            except Exception as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/excel/sheet":
            try:
                body = self.read_json()
                raw_file = str(body.get("file", "")).strip()
                sheet_index = int(body.get("sheetIndex", 0))
                sheet = excel_sheet_preview(Path(raw_file), sheet_index)
                self.send_json(HTTPStatus.OK, sheet)
            except Exception as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/excel/render":
            try:
                body = self.read_json()
                raw_files = body.get("files", [])
                dpi = int(body.get("dpi") or DEFAULT_PDF_DPI)
                if dpi < 72 or dpi > 600:
                    raise ValueError("DPI должен быть в диапазоне 72-600")
                if not isinstance(raw_files, list) or not raw_files:
                    raise ValueError("Не выбраны Excel-файлы для отображения")
                if len(raw_files) > 10:
                    raise ValueError("За один раз пока можно отрендерить не больше 10 Excel-файлов")
                first_page_only = bool(body.get("firstPageOnly", False))
                documents = [render_excel(Path(str(file_path)), dpi=dpi, first_page_only=first_page_only) for file_path in raw_files]
                document_errors = [
                    {"document": document["name"], "path": document["path"], **error}
                    for document in documents
                    for error in document.get("errors", [])
                ]
                self.send_json(
                    HTTPStatus.OK,
                    {
                        "dpi": dpi,
                        "documents": documents,
                        "totalPages": sum(document["pages"] for document in documents),
                        "renderedPages": sum(document["renderedPages"] for document in documents),
                        "errors": document_errors,
                        "renderedAt": datetime.now().isoformat(timespec="seconds"),
                    },
                )
            except Exception as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/excel/page":
            try:
                body = self.read_json()
                raw_file = str(body.get("file", "")).strip()
                page = int(body.get("page") or 1)
                dpi = int(body.get("dpi") or DEFAULT_PDF_DPI)
                if dpi < 72 or dpi > 600:
                    raise ValueError("DPI должен быть в диапазоне 72-600")
                if not raw_file:
                    raise ValueError("Не выбран Excel-файл для отображения")
                pdf_path, convert_cache_hit = excel_to_pdf(Path(raw_file))
                payload = render_pdf_page(pdf_path, page=page, dpi=dpi)
                payload["sourcePath"] = raw_file
                payload["sourceType"] = file_extension(Path(raw_file))
                payload["convertedPdfPath"] = str(pdf_path)
                payload["convertCacheHit"] = convert_cache_hit
                self.send_json(HTTPStatus.OK, payload)
            except Exception as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/open-file":
            started = time.perf_counter()
            raw_file = ""
            target: Path | None = None
            try:
                body = self.read_json()
                raw_file = str(body.get("path", "")).strip()
                if not raw_file:
                    raise ValueError("Не выбран файл для открытия")
                target = Path(raw_file).expanduser()
                if not target.exists():
                    raise FileNotFoundError(f"Файл не найден: {target}")
                opened_path = str(target)

                # Защита от двойного запуска (двойной клик пользователя или повторный вызов)
                target_key = str(target).casefold()
                now_ts = time.perf_counter()
                with _OPEN_FILE_LOCK:
                    last_time = _LAST_OPENED_TARGETS.get(target_key, 0.0)
                    if (now_ts - last_time) < 1.8:
                        self.send_json(
                            HTTPStatus.OK,
                            {
                                "ok": True,
                                "path": str(target),
                                "openedPath": opened_path,
                                "mode": "deduplicated",
                                "longPathWarning": None,
                            },
                        )
                        return
                    _LAST_OPENED_TARGETS[target_key] = now_ts
                    if len(_LAST_OPENED_TARGETS) > 100:
                        _LAST_OPENED_TARGETS.clear()

                suffix = target.suffix.casefold() if target.is_file() else ""
                # Google Drive shortcuts (.gsheet / .gdoc / .gslides) are virtual
                # reparse points, not real Office documents.  Launching them
                # through a local EXE is unpredictable, so give a clear hint.
                if suffix in {".gsheet", ".gdoc", ".gslides"}:
                    raise ValueError("Это облачный документ Google. Откройте его через браузер (Google Docs / Sheets).")
                action = str(body.get("action", "") or "default").strip().lower()
                if action == "explorer":
                    mode = open_in_explorer(target)
                elif action == "system":
                    mode = launch_system_default(target)
                else:
                    mode = launch_native_file(target)
                append_native_open_log(
                    {
                        "at": datetime.now().isoformat(timespec="seconds"),
                        "status": "requested",
                        "extension": suffix,
                        "sourcePath": str(target),
                        "openedPath": opened_path,
                        "mode": mode,
                        "elapsedMs": round((time.perf_counter() - started) * 1000),
                    }
                )
                self.send_json(
                    HTTPStatus.OK,
                    {
                        "ok": True,
                        "path": str(target),
                        "openedPath": opened_path,
                        "mode": mode,
                        "longPathWarning": (
                            "Путь длиннее 240 символов — Windows может не открыть файл."
                            if len(opened_path) > 240
                            else None
                        ),
                    },
                )
            except NativeOpenError as error:
                append_native_open_log(
                    {
                        "at": datetime.now().isoformat(timespec="seconds"),
                        "status": "error",
                        "error_code": error.error_code,
                        "extension": target.suffix.casefold() if target else Path(raw_file).suffix.casefold(),
                        "sourcePath": str(target) if target else raw_file,
                        "error": str(error),
                        "details": error.details,
                        "elapsedMs": round((time.perf_counter() - started) * 1000),
                    }
                )
                self.send_json(
                    HTTPStatus.BAD_REQUEST,
                    {
                        "ok": False,
                        "error": str(error),
                        "error_code": error.error_code,
                        "details": error.details,
                    },
                )
            except Exception as error:
                append_native_open_log(
                    {
                        "at": datetime.now().isoformat(timespec="seconds"),
                        "status": "error",
                        "error_code": "unknown_error",
                        "extension": target.suffix.casefold() if target else Path(raw_file).suffix.casefold(),
                        "sourcePath": str(target) if target else raw_file,
                        "error": str(error),
                        "elapsedMs": round((time.perf_counter() - started) * 1000),
                    }
                )
                self.send_json(
                    HTTPStatus.BAD_REQUEST,
                    {
                        "ok": False,
                        "error": str(error),
                        "error_code": "unknown_error",
                    },
                )
            return

        if parsed.path == "/api/clipboard":
            try:
                body = self.read_json()
                raw_text = str(body.get("text", "")).strip()
                if not raw_text:
                    raise ValueError("Текст для копирования не передан")
                ok = set_windows_clipboard(raw_text)
                if not ok:
                    raise RuntimeError("Не удалось записать текст в буфер обмена Windows")
                self.send_json(HTTPStatus.OK, {"ok": True, "text": raw_text})
            except Exception as error:
                self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        if parsed.path == "/api/updates/apply":
            # Запуск скачивания и обновления в фоновом потоке, чтобы сразу вернуть ответ браузеру
            import threading
            from app.updater import check_for_updates, download_update_payload, apply_update_and_exit

            def _worker():
                try:
                    info = check_for_updates()
                    if info.get("has_update") and info.get("download_url"):
                        zip_path = download_update_payload(info["download_url"], info.get("sha256", ""))
                        apply_update_and_exit(zip_path)
                except Exception as e:
                    print(f"[Update Error] {e}")

            threading.Thread(target=_worker, daemon=True).start()
            self.send_json(HTTPStatus.OK, {"status": "updating", "message": "Загрузка обновления началась, лаунчер скоро перезапустится."})
            return

        self.send_json(HTTPStatus.NOT_FOUND, {"error": "Маршрут не найден"})

    def serve_cache(self, request_path: str) -> None:
        relative = unquote(request_path).removeprefix("/cache/").lstrip("/")
        target = (RUNTIME_DIR / "cache" / relative).resolve()

        try:
            target.relative_to((RUNTIME_DIR / "cache").resolve())
        except ValueError:
            self.send_error(HTTPStatus.FORBIDDEN, "Forbidden")
            return

        if not target.exists() or not target.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
            return

        content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        body = target.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "private, max-age=31536000, immutable")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def serve_static(self, request_path: str) -> None:
        relative = "index.html" if request_path in ("", "/") else unquote(request_path).lstrip("/")
        target = (FRONTEND_DIR / relative).resolve()

        try:
            target.relative_to(FRONTEND_DIR.resolve())
        except ValueError:
            self.send_error(HTTPStatus.FORBIDDEN, "Forbidden")
            return

        if not target.exists() or not target.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
            return

        content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        if content_type.startswith("text/") or content_type in {"application/javascript", "text/javascript"}:
            content_type = f"{content_type}; charset=utf-8"
        body = target.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        # HTML/JS/CSS must never be served stale: a cached app.js silently
        # reverts the native-open logic to the broken WPS associations.
        # no-store (not no-cache) also wins without a manual ?v= bump.
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.end_headers()
        self.wfile.write(body)


def main() -> None:
    parser = argparse.ArgumentParser(description="Run F-Engineering Launcher v3 local server.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8780)
    args = parser.parse_args()

    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    (RUNTIME_DIR / "cache").mkdir(exist_ok=True)
    PDF_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    MANIFESTS_DIR.mkdir(exist_ok=True)
    (RUNTIME_DIR / "logs").mkdir(exist_ok=True)

    atexit.register(lambda: kill_dwg_daemon("server-exit"))

    try:
        server = ThreadingHTTPServer((args.host, args.port), LauncherHandler)
    except OSError as err:
        if getattr(err, "winerror", None) == 10048 or getattr(err, "errno", None) == 10048 or "10048" in str(err):
            print(f"ERROR: Port {args.port} is already in use by another process: {err}", file=sys.stderr, flush=True)
            sys.exit(48)
        raise
    print(f"F-Engineering Launcher v3: http://{args.host}:{args.port}/", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
