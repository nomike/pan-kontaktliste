"""
Render participant list to HTML (with optional print-to-PDF via browser).
"""
from __future__ import annotations

import base64
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape


def _image_to_data_url(image_path: str | Path) -> str:
    """Read image file and return a data URL (e.g. data:image/png;base64,...)."""
    path = Path(image_path)
    if not path.exists():
        return ""
    data = path.read_bytes()
    ext = path.suffix.lower()
    mime = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".webp": "image/webp",
    }.get(ext, "image/png")
    b64 = base64.b64encode(data).decode("ascii")
    return f"data:{mime};base64,{b64}"


def render_html(
    participants: list[dict],
    output_html_path: Path,
    template_dir: Path | None = None,
) -> None:
    """
    Render participants to a single HTML file with embedded images (data URLs).
    Each participant must have 'image_path' (path to image file).
    Adds 'image_data' (data URL) to each participant for the template.
    """
    output_html_path = Path(output_html_path)
    if template_dir is None:
        template_dir = Path(__file__).resolve().parent / "template"

    env = Environment(
        loader=FileSystemLoader(str(template_dir)),
        autoescape=select_autoescape(["html", "htm", "xml"]),
    )

    for p in participants:
        p["image_data"] = _image_to_data_url(p["image_path"])

    template = env.get_template("contact_list.html.j2")
    html_content = template.render(participants=participants)
    output_html_path.write_text(html_content, encoding="utf-8")
