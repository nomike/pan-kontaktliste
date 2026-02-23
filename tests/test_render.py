"""Tests for render: HTML output, image data URLs, edge cases."""
from __future__ import annotations

from pathlib import Path

import pytest

from render import render_html


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
