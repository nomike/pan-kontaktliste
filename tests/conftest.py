"""Pytest fixtures for PAN Kontaktliste tests."""
from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image


@pytest.fixture
def placeholder_path(tmp_path: Path) -> Path:
    """A minimal placeholder image (1x1 PNG) for tests."""
    png = tmp_path / "placeholder.png"
    Image.new("RGB", (1, 1), color=(0, 0, 0)).save(png)
    return png
