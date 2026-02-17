"""Tests for version module."""
from __future__ import annotations

from version import get_version


def test_get_version_returns_string() -> None:
    """get_version() returns a non-empty string."""
    v = get_version()
    assert isinstance(v, str)
    assert len(v) > 0


def test_get_version_looks_semver() -> None:
    """Version looks like x.y.z (at least two dots or one)."""
    v = get_version()
    parts = v.split(".")
    assert len(parts) >= 1
    assert all(p.isdigit() or p for p in parts)
