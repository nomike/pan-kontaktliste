"""Tests for render: HTML output, image data URLs, edge cases."""
from __future__ import annotations

import base64
import io
from pathlib import Path

import pytest
from PIL import Image

from render import _image_to_data_url, render_html


def test_render_html_empty_participants(tmp_path: Path, placeholder_path: Path) -> None:
    """Empty participant list produces valid HTML with empty columns."""
    out = tmp_path / "out.html"
    render_html([], out)
    assert out.exists()
    content = out.read_text(encoding="utf-8")
    assert "Teilnehmendenkontaktliste" in content
    assert "columns" in content


def test_render_html_one_participant(tmp_path: Path, placeholder_path: Path) -> None:
    """One participant with image_path produces HTML with data URL."""
    participants = [
        {
            "land": "DE",
            "rufname": "Test",
            "couch": "",
            "image_path": str(placeholder_path),
        },
    ]
    out = tmp_path / "out.html"
    render_html(participants, out)
    assert out.exists()
    content = out.read_text(encoding="utf-8")
    assert "Test" in content
    assert "data:image/png;base64," in content


def test_render_html_missing_image_path(tmp_path: Path) -> None:
    """Missing image file: image_data is empty string (no crash)."""
    participants = [
        {
            "land": "DE",
            "rufname": "NoPic",
            "couch": "",
            "image_path": str(tmp_path / "nonexistent.png"),
        },
    ]
    out = tmp_path / "out.html"
    render_html(participants, out)
    assert out.exists()
    content = out.read_text(encoding="utf-8")
    assert "NoPic" in content


def test_render_html_optional_fields(tmp_path: Path, placeholder_path: Path) -> None:
    """Participant with email, phone, vorname, nachname: all appear when set."""
    participants = [
        {
            "land": "AT",
            "rufname": "Full",
            "couch": "yes",
            "email": "a@b.at",
            "phone": "+43 1",
            "vorname": "Anna",
            "nachname": "Z",
            "image_path": str(placeholder_path),
        },
    ]
    out = tmp_path / "out.html"
    render_html(participants, out)
    content = out.read_text(encoding="utf-8")
    assert "a@b.at" in content
    assert "+43 1" in content
    assert "Anna" in content
    assert "Z" in content
    assert "yes" in content


def test_render_html_escapes_content(tmp_path: Path, placeholder_path: Path) -> None:
    """HTML-unsafe content in participant is escaped."""
    participants = [
        {
            "land": "DE",
            "rufname": "<script>alert(1)</script>",
            "couch": "&",
            "image_path": str(placeholder_path),
        },
    ]
    out = tmp_path / "out.html"
    render_html(participants, out)
    content = out.read_text(encoding="utf-8")
    assert "&lt;script&gt;" in content or "<script>" not in content
    assert "&amp;" in content


def test_image_to_data_url_rotation(tmp_path: Path) -> None:
    """Rotation swaps width and height before thumbnailing."""
    src = tmp_path / "wide.png"
    Image.new("RGB", (200, 100), color=(255, 0, 0)).save(src)

    data_url = _image_to_data_url(src, rotation=90)
    assert data_url.startswith("data:image/png;base64,")
    raw = base64.b64decode(data_url.split(",", 1)[1])
    with Image.open(io.BytesIO(raw)) as thumb:
        assert thumb.width < thumb.height


def test_render_html_applies_image_rotation(tmp_path: Path) -> None:
    """render_html passes image_rotation through to thumbnail generation."""
    src = tmp_path / "wide.png"
    Image.new("RGB", (200, 100), color=(0, 255, 0)).save(src)
    participants = [
        {
            "land": "DE",
            "rufname": "Rotated",
            "couch": "",
            "image_path": str(src),
            "image_rotation": 90,
        },
    ]
    out = tmp_path / "out.html"
    render_html(participants, out)
    content = out.read_text(encoding="utf-8")
    assert "Rotated" in content
    assert "data:image/png;base64," in content
