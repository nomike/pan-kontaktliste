"""
Read package version for display in GUI (About dialog).
Release-please updates version in pyproject.toml; we read it from there or from metadata when installed.
In a PyInstaller frozen exe, pyproject.toml is bundled so the fallback still works.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path


def get_version() -> str:
    """Return current package version (from pyproject.toml or install metadata)."""
    try:
        from importlib.metadata import version
        return version("pan-kontaktliste")
    except Exception:
        pass
    # In frozen (PyInstaller) exe, bundle root is sys._MEIPASS; pyproject.toml is added there at build
    base = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    path = base / "pyproject.toml"
    if path.exists():
        text = path.read_text(encoding="utf-8")
        m = re.search(r'version\s*=\s*["\']([^"\']+)["\']', text)
        if m:
            return m.group(1)
    return "0.0.0"
