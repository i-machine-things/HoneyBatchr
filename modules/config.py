"""Config file path, defaults, load/write helpers."""

import json
from pathlib import Path

CONFIG_FILE = Path.home() / ".honeybatchr" / "config.json"

DEFAULT_CONFIG: dict = {
    "copies": 1,
    "collate": True,
    "grayscale": False,
    "print_as_image": False,
    "bleed_marks": False,
    "duplex": False,
    "auto_rotate": True,
    "auto_center": True,
    "orientation": "Portrait",
    "print_what": "Document and markups",
    "simulate_overprint": False,
    "pages_per_sheet": 2,
    "page_order": "Horizontal",
    "margins": 0.200,
    "margins_enabled": True,
    "print_page_border": True,
    "theme": "Fusion Light",
}


def load_config() -> dict:
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading config: {e}")
    return dict(DEFAULT_CONFIG)


def write_config(data: dict) -> None:
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def update_config_value(key: str, value) -> None:
    """Update a single key without overwriting unrelated settings."""
    try:
        saved: dict = {}
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, encoding="utf-8") as f:
                saved = json.load(f)
        saved[key] = value
        write_config(saved)
    except Exception as e:
        print(f"Error updating config: {e}")
