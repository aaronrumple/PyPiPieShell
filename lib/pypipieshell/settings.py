# -*- coding: utf-8 -*-
"""Per-user settings, REPL history and editor autosave for PyPiPie Shell.

Stored under ``%APPDATA%\\pyRevit\\PyPiPieShell`` so every host (modal window,
modeless window, dockable pane) shares the same settings, history and the
last unsaved buffer.
"""
import io
import json
import os

DEFAULTS = {
    "font_family": "Cascadia Mono, Consolas",
    "font_size": 13.0,
    "word_wrap": False,
    "show_whitespace": False,
    "show_line_numbers": True,
    "highlight_current_line": True,
    "folding": True,
    # auto | group | none   (see execution.TransactionPolicy)
    "transaction_mode": "auto",
    # Esc interrupts a running script (uses a Python trace hook)
    "esc_interrupt": True,
    "complete_as_you_type": True,
    "hover_tips": True,
    "restore_buffer": True,
    "stub_paths": [],
    "last_file": None,
    "launch_mode": "Modeless",
    "find_bar_visible": True,
}

_HISTORY_MAX = 500


def data_dir():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "pyRevit", "PyPiPieShell")
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except OSError:
            pass
    return path


def _path(name):
    return os.path.join(data_dir(), name)


def _read_json(name, default):
    try:
        with io.open(_path(name), "r", encoding="utf-8") as f:
            return json.loads(f.read())
    except Exception:
        return default


def _write_json(name, data):
    try:
        text = json.dumps(data, indent=2, sort_keys=True)
        if not isinstance(text, type(u"")):
            text = text.decode("utf-8")
        with io.open(_path(name), "w", encoding="utf-8") as f:
            f.write(text)
        return True
    except Exception:
        return False


class Settings(object):
    """Dictionary-like settings with defaults and explicit ``save()``."""

    def __init__(self):
        self._data = dict(DEFAULTS)
        stored = _read_json("settings.json", {})
        if isinstance(stored, dict):
            self._data.update(stored)

    def __getitem__(self, key):
        return self._data.get(key, DEFAULTS.get(key))

    def __setitem__(self, key, value):
        self._data[key] = value

    def get(self, key, default=None):
        return self._data.get(key, DEFAULTS.get(key, default))

    def save(self):
        return _write_json("settings.json", self._data)


def load_history():
    data = _read_json("history.json", [])
    return [h for h in data if isinstance(h, type(u"")) or isinstance(h, str)] if isinstance(data, list) else []


def save_history(history):
    return _write_json("history.json", list(history)[-_HISTORY_MAX:])


def read_autosave():
    try:
        with io.open(_path("autosave.py"), "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return None


def write_autosave(text):
    try:
        with io.open(_path("autosave.py"), "w", encoding="utf-8") as f:
            f.write(text or u"")
        return True
    except Exception:
        return False


def read_text_file(path):
    """Read a source file as unicode, tolerating UTF-8 BOM and cp1252."""
    with io.open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", "replace")


def write_text_file(path, text):
    with io.open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)
