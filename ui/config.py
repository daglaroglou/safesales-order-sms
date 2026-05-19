"""Persistent application settings (export paths, etc.)."""

from __future__ import annotations

import json
import os
from pathlib import Path

from excel.utils import excel_dir

_CONFIG_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "safesales-order-sms"
_CONFIG_FILE = _CONFIG_DIR / "settings.json"


def default_export_directory() -> Path:
    return excel_dir()


def _load_raw() -> dict:
    if not _CONFIG_FILE.is_file():
        return {}
    try:
        data = json.loads(_CONFIG_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def _save_raw(data: dict) -> None:
    _CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    _CONFIG_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def get_export_directory() -> Path:
    raw = (_load_raw().get("export_directory") or "").strip()
    if raw:
        path = Path(raw)
        if path.is_dir():
            return path
    return default_export_directory()


def get_api_key() -> str | None:
    raw = (_load_raw().get("api_key") or "").strip()
    return raw or None


def set_api_key(api_key: str) -> str:
    key = (api_key or "").strip()
    if not key:
        raise ValueError("api_key cannot be empty")
    data = _load_raw()
    data["api_key"] = key
    _save_raw(data)
    return key


def set_export_directory(directory: Path | str) -> Path:
    path = Path(directory).expanduser().resolve()
    path.mkdir(parents=True, exist_ok=True)
    data = _load_raw()
    data["export_directory"] = str(path)
    _save_raw(data)
    return path
