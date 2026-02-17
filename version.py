"""
Read package version for display in GUI (About dialog).
Release-please updates version in pyproject.toml; we read it from there or from metadata when installed.
"""
from __future__ import annotations

import re
from pathlib import Path


def get_version() -> str:
    """Return current package version (from pyproject.toml or install metadata)."""
    try:
        from importlib.metadata import version
        return version("pan-kontaktliste")
    except Exception:
        pass
    path = Path(__file__).resolve().parent / "pyproject.toml"
    if path.exists():
        text = path.read_text(encoding="utf-8")
        m = re.search(r'version\s*=\s*["\']([^"\']+)["\']', text)
        if m:
            return m.group(1)
    return "0.0.0"
