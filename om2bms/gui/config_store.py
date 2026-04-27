from __future__ import annotations

from pathlib import Path


PROJECT_ROOT = Path.cwd()
CONFIG_DIR = PROJECT_ROOT / "config"

DEFAULT_OUTPUT_FILE = CONFIG_DIR / "default_output_dir.txt"
DEFAULT_BMS_FILE = CONFIG_DIR / "default_bms_dir.txt"
DEFAULT_JSON_FILE = CONFIG_DIR / "default_json_dir.txt"


def _ensure_config_dir() -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)


def _load_path(file_path: Path, fallback: str = "") -> str:
    try:
        if not file_path.exists():
            return fallback

        value = file_path.read_text(encoding="utf-8-sig").strip()
        return value or fallback
    except OSError:
        return fallback


def _save_path(file_path: Path, value: str) -> None:
    try:
        _ensure_config_dir()
        file_path.write_text(value, encoding="utf-8")
    except OSError:
        pass


def load_default_output_dir() -> str:
    return _load_path(DEFAULT_OUTPUT_FILE, fallback=str(Path.cwd()))


def save_default_output_dir(path: str) -> None:
    _save_path(DEFAULT_OUTPUT_FILE, path)


def load_default_bms_dir() -> str:
    return _load_path(DEFAULT_BMS_FILE, fallback="")


def save_default_bms_dir(path: str) -> None:
    _save_path(DEFAULT_BMS_FILE, path)


def load_default_json_dir() -> str:
    return _load_path(DEFAULT_JSON_FILE, fallback="")


def save_default_json_dir(path: str) -> None:
    _save_path(DEFAULT_JSON_FILE, path)
