# -*- coding: utf-8 -*-
"""Хранение и чтение версии приложения Launcher v3.

Манифест ``version.json`` лежит в корне репозитория (ветка ``dvg-main``).
Модуль ищет его в корне репозитория, затем в папке ``app/``.
"""

from __future__ import annotations

import json
from pathlib import Path

DEFAULT_VERSION = "3.1.0"

DEFAULT_PAYLOAD: dict = {
    "version": "3.1.0",
    "build_date": "2026-09-27",
    "git_branch": "dvg-main",
    "runtime_version": "1.0",
    "update_manifest_url": (
        "https://raw.githubusercontent.com/f-engineering-spb/"
        "f-engineering-launcher/dvg-main/version.json"
    ),
}

_MANIFEST_FILENAME = "version.json"


def _candidate_paths() -> list[Path]:
    """Вернуть пути к version.json в порядке приоритета."""
    here = Path(__file__).resolve()
    # app/shared/version_info.py -> app/ -> repo root
    app_dir = here.parents[1]
    repo_root = here.parents[2]
    return [repo_root / _MANIFEST_FILENAME, app_dir / _MANIFEST_FILENAME]


def get_version_payload() -> dict:
    """Вернуть полный словарь манифеста version.json.

    При отсутствии файла или ошибке чтения возвращается
    копия словаря по умолчанию.
    """
    for path in _candidate_paths():
        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
        except FileNotFoundError:
            continue
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            continue
        if isinstance(data, dict):
            return data
    return dict(DEFAULT_PAYLOAD)


def get_app_version() -> str:
    """Вернуть версию приложения из version.json.

    Возвращает ``DEFAULT_VERSION`` при любой ошибке чтения.
    """
    try:
        payload = get_version_payload()
        version = payload.get("version", DEFAULT_VERSION)
    except FileNotFoundError:
        return DEFAULT_VERSION
    if not version:
        return DEFAULT_VERSION
    return str(version)
