"""
Render participant list to HTML (with optional print-to-PDF via browser).
"""
from __future__ import annotations

import base64
import io
import sys
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from PIL import Image

_THUMBNAIL_SIZE = (144, 144)  # 2x display size (72px CSS) for retina


def _base_path() -> Path:
    """Project root (or PyInstaller bundle root)."""
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _image_to_data_url(image_path: str | Path) -> str:
    """Resize image to thumbnail and return a PNG data URL."""
    path = Path(image_path)
    if not path.exists():
        return ""
    try:
        with Image.open(path) as img:
            img.thumbnail(_THUMBNAIL_SIZE, Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="PNG", optimize=True)
            b64 = base64.b64encode(buf.getvalue()).decode("ascii")
            return f"data:image/png;base64,{b64}"
    except Exception:
        return ""


def render_html(
    participants: list[dict],
    output_html_path: Path,
    meetup_name: str = "",
    template_dir: Path | None = None,
) -> None:
    """
    Render participants to a single HTML file with embedded images (data URLs).
    Each participant must have 'image_path' (path to image file).
    Adds 'image_data' (data URL) to each participant for the template.
    meetup_name is used as the HTML page title and h1; if empty, falls back to "Teilnehmendenkontaktliste".
    """
    output_html_path = Path(output_html_path)
    if template_dir is None:
        template_dir = _base_path() / "template"

    env = Environment(
        loader=FileSystemLoader(str(template_dir)),
        autoescape=select_autoescape(["html", "htm", "xml", "j2"]),
    )

    for p in participants:
        p["image_data"] = _image_to_data_url(p["image_path"])

    template = env.get_template("contact_list.html.j2")
    html_content = template.render(
        participants=participants,
        meetup_name=meetup_name.strip(),
    )
    output_html_path.write_text(html_content, encoding="utf-8")
