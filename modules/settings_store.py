from __future__ import annotations

import json
from pathlib import Path


SETTINGS_PATH = Path.home() / ".audio_batch_renamer_settings.json"
WORKFLOW_STATE_PATH = Path.home() / ".audio_batch_renamer_state.json"
DEFAULT_SETTINGS = {
    "digits": "3桁",
    "keep_text": True,
    "move_ng": True,
    "export_csv": True,
    "geometry": "1380x820",
}


def load_settings() -> dict:
    if not SETTINGS_PATH.exists():
        return dict(DEFAULT_SETTINGS)

    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return dict(DEFAULT_SETTINGS)

    settings = dict(DEFAULT_SETTINGS)
    settings.update({key: value for key, value in data.items() if key in DEFAULT_SETTINGS})
    return settings


def save_settings(settings: dict) -> None:
    merged = dict(DEFAULT_SETTINGS)
    merged.update(settings)
    SETTINGS_PATH.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")


def load_workflow_state() -> dict:
    if not WORKFLOW_STATE_PATH.exists():
        return {}
    try:
        data = json.loads(WORKFLOW_STATE_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


def save_workflow_state(state: dict) -> None:
    WORKFLOW_STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def clear_workflow_state() -> None:
    WORKFLOW_STATE_PATH.unlink(missing_ok=True)
