# -*- coding: utf-8 -*-
"""Автообновление лаунчера через GitHub (Спринт 2).

Только стандартная библиотека Python: urllib, json, hashlib, shutil,
zipfile, subprocess, pathlib, time, sys, os. Никаких pip-зависимостей.
"""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import urllib.request
from pathlib import Path

GITHUB_REPO = "f-engineering-spb/f-engineering-launcher"
VERSION_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/dvg-main/version.json"
APP_DIR = Path(__file__).resolve().parent
REPO_ROOT = APP_DIR.parent
CACHE_DIR = REPO_ROOT / "runtime" / "cache"

USER_AGENT = "FEngineering-Launcher"
REQUEST_TIMEOUT = 4
PAYLOAD_NAME = "update_payload.zip"
DEFAULT_VERSION = "3.1.0"


def _parse_version(value: str) -> tuple:
    """Преобразовать '3.1.0' в (3, 1, 0) для корректного сравнения."""
    parts: list = []
    for chunk in str(value or "").strip().lstrip("vV").split("."):
        digits = ""
        for ch in chunk:
            if ch.isdigit():
                digits += ch
            else:
                break
        try:
            parts.append(int(digits) if digits else 0)
        except ValueError:
            parts.append(0)
    return tuple(parts) if parts else (0,)


def _read_local_version() -> str:
    """Прочитать локальную версию через version_info или version.json."""
    try:
        from app.shared.version_info import get_app_version

        version = get_app_version()
        if version:
            return str(version)
    except FileNotFoundError:
        pass
    except Exception:
        pass
    for candidate in (REPO_ROOT / "version.json", APP_DIR / "version.json"):
        try:
            with candidate.open("r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and data.get("version"):
                return str(data["version"])
        except FileNotFoundError:
            continue
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            continue
    return DEFAULT_VERSION


def check_for_updates() -> dict:
    """Проверить наличие обновления на GitHub.

    Возвращает словарь с ключами has_update, current_version,
    latest_version, download_url, changelog, sha256. При ошибке сети
    возвращает has_update=False с полем error.
    """
    current_ver = _read_local_version()
    try:
        request = urllib.request.Request(
            VERSION_URL, headers={"User-Agent": USER_AGENT}
        )
        with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT) as response:
            raw = response.read().decode("utf-8")
        remote = json.loads(raw)
        if not isinstance(remote, dict):
            raise ValueError("Некорректный манифест version.json")
        latest_ver = str(remote.get("version") or current_ver)
        has_update = _parse_version(latest_ver) > _parse_version(current_ver)
        return {
            "has_update": bool(has_update),
            "current_version": current_ver,
            "latest_version": latest_ver,
            "download_url": str(remote.get("download_url") or ""),
            "changelog": str(remote.get("changelog") or ""),
            "sha256": str(remote.get("sha256") or ""),
        }
    except Exception as e:
        return {
            "has_update": False,
            "error": str(e),
            "current_version": current_ver,
            "latest_version": current_ver,
        }


def download_update_payload(download_url: str, expected_sha256: str = "") -> str:
    """Скачать архив обновления в runtime/cache/update_payload.zip.

    При передаче expected_sha256 сверяет контрольную сумму SHA-256,
    при несовпадении удаляет архив и выбрасывает ValueError.
    Возвращает строковый путь к файлу.
    """
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    target = CACHE_DIR / PAYLOAD_NAME
    request = urllib.request.Request(
        download_url, headers={"User-Agent": USER_AGENT}
    )
    hasher = hashlib.sha256()
    with urllib.request.urlopen(request, timeout=30) as response, target.open(
        "wb"
    ) as out:
        while True:
            chunk = response.read(65536)
            if not chunk:
                break
            out.write(chunk)
            hasher.update(chunk)
    if expected_sha256:
        actual = hasher.hexdigest()
        if actual.lower() != expected_sha256.strip().lower():
            try:
                target.unlink()
            except OSError:
                pass
            raise ValueError("Контрольная сумма SHA-256 не совпала")
    return str(target)


def apply_update_and_exit(zip_path: str) -> None:
    """Запустить scripts/update_helper.cmd и завершить текущий процесс.

    Helper работает в отдельном окне, чтобы снять файловые блокировки
    Windows с server.py и папки app/.
    """
    helper = REPO_ROOT / "scripts" / "update_helper.cmd"
    if not helper.exists():
        raise FileNotFoundError(f"Скрипт обновления не найден: {helper}")
    creationflags = 0
    if sys.platform.startswith("win"):
        creationflags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
    subprocess.Popen(
        ["cmd.exe", "/c", str(helper), str(zip_path)],
        creationflags=creationflags,
        cwd=str(REPO_ROOT),
    )
    sys.exit(0)
