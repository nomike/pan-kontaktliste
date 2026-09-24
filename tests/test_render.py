"""Tests for render: PDF output, image data URLs, edge cases."""
from __future__ import annotations

import base64
import io
from pathlib import Path

from PIL import Image

from render import _build_html, _image_to_data_url, render_pdf


def _assert_pdf(path: Path) -> None:
    assert path.exists()
    raw = path.read_bytes()
    assert raw.startswith(b"%PDF")
    assert len(raw) > 100


def test_render_pdf_empty_participants(tmp_path: Path, placeholder_path: Path) -> None:
    """Empty participant list produces a valid PDF."""
    out = tmp_path / "out.pdf"
    render_pdf([], out)
    _assert_pdf(out)


def test_render_pdf_one_participant(tmp_path: Path, placeholder_path: Path) -> None:
    """One participant with image_path produces a PDF; HTML includes data URL."""
    participants = [
        {
            "land": "DE",
            "rufname": "Test",
            "couch": "",
            "image_path": str(placeholder_path),
        },
    ]
    html = _build_html(participants)
    assert "Test" in html
    assert "data:image/png;base64," in html

    out = tmp_path / "out.pdf"
    render_pdf(participants, out)
    _assert_pdf(out)


def test_render_pdf_missing_image_path(tmp_path: Path) -> None:
    """Missing image file: image_data is empty string (no crash)."""
    participants = [
        {
            "land": "DE",
            "rufname": "NoPic",
            "couch": "",
            "image_path": str(tmp_path / "nonexistent.png"),
        },
    ]
    html = _build_html(participants)
    assert "NoPic" in html

    out = tmp_path / "out.pdf"
    render_pdf(participants, out)
    _assert_pdf(out)


def test_render_pdf_optional_fields(tmp_path: Path, placeholder_path: Path) -> None:
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
    html = _build_html(participants)
    assert "a@b.at" in html
    assert "+43 1" in html
    assert "Anna" in html
    assert "Z" in html
    assert "yes" in html

    out = tmp_path / "out.pdf"
    render_pdf(participants, out)
    _assert_pdf(out)


def test_build_html_escapes_content(tmp_path: Path, placeholder_path: Path) -> None:
    """HTML-unsafe content in participant is escaped."""
    participants = [
        {
            "land": "DE",
            "rufname": "<script>alert(1)</script>",
            "couch": "&",
            "image_path": str(placeholder_path),
        },
    ]
    content = _build_html(participants)
    assert "&lt;script&gt;" in content or "<script>" not in content
    assert "&amp;" in content


def test_image_to_data_url_exif_orientation(tmp_path: Path) -> None:
    """EXIF orientation 6 (rotate 90 CW) is applied before thumbnailing."""
    src = tmp_path / "exif.jpg"
    img = Image.new("RGB", (200, 100), color=(0, 0, 255))
    # Pillow writes Orientation via exif; tag 274 = Orientation
    exif = img.getexif()
    exif[274] = 6  # Rotate 90 CW → tall image
    img.save(src, format="JPEG", exif=exif)

    data_url = _image_to_data_url(src)
    assert data_url.startswith("data:image/png;base64,")
    raw = base64.b64decode(data_url.split(",", 1)[1])
    with Image.open(io.BytesIO(raw)) as thumb:
        assert thumb.height >= thumb.width


def test_build_html_meetup_name_title(tmp_path: Path, placeholder_path: Path) -> None:
    """meetup_name is used as title; empty falls back to default."""
    html_default = _build_html([])
    assert "Teilnehmendenkontaktliste" in html_default

    html_named = _build_html([], meetup_name="PAN Testtreffen")
    assert "PAN Testtreffen" in html_named
