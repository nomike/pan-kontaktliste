"""
Read PAN sign-up Excel and build participant list for the contact list PDF.
Respects Teilnehmyliste and per-field consent; extracts images when consented.
"""
from __future__ import annotations

import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.reader.drawings import find_images


# Column names in the spreadsheet (exact match)
CONSENT_LIST = "Teilnehmyliste"
CONSENT_EMAIL = "Teilnehmyliste E-Mail"
CONSENT_PHONE = "Teilnehmyliste Telefonnummer"
CONSENT_NACHNAME = "Teilnehmyliste Nachname"
CONSENT_VORNAME = "Teilnehmerliste Vorname"  # note: "Teilnehmerliste" spelling
CONSENT_BILD = "Teilnehmyliste Bild"

DATA_LAND = "Land"
DATA_RUFNAME = "Rufname/Pseudonym"
DATA_COUCH = "Teilnehmyliste_Couch"
DATA_EMAIL = "E-Mail Adresse"
DATA_PHONE = "Telefonnummer (mit Ländercode!)"
DATA_FAMILIENNAME = "Familiename"
DATA_VORNAME = "Vorname"
DATA_BILD = "Bild"


def _truthy(value: Any) -> bool:
    """Normalize Excel booleans and strings to bool."""
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    s = str(value).strip().lower()
    return s in ("true", "1", "yes", "ja", "x")


def _anchor_row(anchor) -> int | None:
    """Get 0-based row from an openpyxl anchor (OneCellAnchor or TwoCellAnchor)."""
    if anchor is None:
        return None
    if hasattr(anchor, "_from") and anchor._from is not None:
        return getattr(anchor._from, "row", None)
    return None


def _extract_images_by_row(xlsx_path: str | Path) -> dict[int, str]:
    """
    Open xlsx as zip, find drawing for first sheet, extract images and map by Excel row (1-based).
    Returns dict: excel_row -> path to extracted image file (temp file).
    """
    path = Path(xlsx_path)
    if not path.exists():
        return {}

    row_to_path: dict[int, str] = {}
    try:
        with zipfile.ZipFile(path, "r") as archive:
            names = archive.namelist()
            drawing_paths = [n for n in names if n.startswith("xl/drawings/") and n.endswith(".xml") and "_rels" not in n]
            if not drawing_paths:
                return {}

            # First sheet typically has drawing1.xml
            drawing_path = "xl/drawings/drawing1.xml"
            if drawing_path not in names:
                drawing_path = drawing_paths[0]

            charts, images = find_images(archive, drawing_path)
    except Exception:
        return {}

    if not images:
        return {}

    # Map anchor row (0-based) -> list of images; Excel data row 2 = 0-based row 1
    by_row: dict[int, list] = {}
    for img in images:
        row_0 = _anchor_row(img.anchor)
        if row_0 is not None:
            by_row.setdefault(row_0, []).append(img)

    # Save each image to a temp file and map Excel row (1-based) to path
    for row_0, img_list in by_row.items():
        excel_row = row_0 + 1
        img = img_list[0]
        try:
            data = img._data()
            ext = (img.format or "png").lower()
            if ext not in ("png", "jpeg", "jpg", "gif"):
                ext = "png"
            fd, tmp_path = tempfile.mkstemp(suffix=f".{ext}", prefix="pan_contact_")
            os.close(fd)
            with open(tmp_path, "wb") as f:
                f.write(data)
            row_to_path[excel_row] = tmp_path
        except Exception:
            continue

    return row_to_path


def load_participants(
    xlsx_path: str | Path,
    placeholder_image_path: str | Path,
    image_output_dir: Path | None = None,
) -> list[dict[str, Any]]:
    """
    Load workbook, filter by Teilnehmyliste, apply per-field consent, resolve image or placeholder.
    If image_output_dir is given, extracted/placeholder images are copied there (for LaTeX build).
    Returns list of participant dicts with keys: land, rufname, couch, email?, phone?, nachname?,
    vorname?, image_path (always set).
    """
    xlsx_path = Path(xlsx_path)
    placeholder_image_path = Path(placeholder_image_path)
    if image_output_dir is None:
        image_output_dir = Path(tempfile.mkdtemp(prefix="pan_contact_images_"))
    image_output_dir = Path(image_output_dir)
    image_output_dir.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.load_workbook(xlsx_path, read_only=False, data_only=True)
    sh = wb.active
    if sh is None:
        wb.close()
        return []

    # Header row 1
    headers: list[str | None] = []
    for c in sh[1]:
        headers.append(c.value)
    col_index: dict[str, int] = {}
    for i, h in enumerate(headers):
        if h is not None and str(h).strip():
            col_index[str(h).strip()] = i

    def get_col(name: str) -> int | None:
        return col_index.get(name)

    def cell(row_idx: int, col_key: str) -> Any:
        j = get_col(col_key)
        if j is None:
            return None
        row = sh[row_idx]
        if j >= len(row):
            return None
        return row[j].value

    # Extract images by row (Excel row number = 2, 3, ...)
    row_to_image_path = _extract_images_by_row(xlsx_path)
    placeholder_path = placeholder_image_path.resolve()

    def _str(v: Any) -> str:
        if v is None:
            return ""
        s = str(v).strip()
        return s if s else ""

    participants: list[dict[str, Any]] = []
    for i, row_idx in enumerate(range(2, sh.max_row + 1)):
        if not _truthy(cell(row_idx, CONSENT_LIST)):
            continue

        email_ok = _truthy(cell(row_idx, CONSENT_EMAIL))
        phone_ok = _truthy(cell(row_idx, CONSENT_PHONE))
        nachname_ok = _truthy(cell(row_idx, CONSENT_NACHNAME))
        vorname_ok = _truthy(cell(row_idx, CONSENT_VORNAME))
        bild_ok = _truthy(cell(row_idx, CONSENT_BILD))

        land = _str(cell(row_idx, DATA_LAND))
        rufname = _str(cell(row_idx, DATA_RUFNAME))
        couch = _str(cell(row_idx, DATA_COUCH))

        if bild_ok and row_idx in row_to_image_path:
            src = Path(row_to_image_path[row_idx])
            ext = src.suffix or ".png"
            dest = image_output_dir / f"teilnehmer_{len(participants)}{ext}"
            try:
                shutil.copy2(src, dest)
                image_path = str(dest)
            except Exception:
                image_path = str(placeholder_path)
            try:
                os.unlink(src)
            except Exception:
                pass
        else:
            dest = image_output_dir / f"teilnehmer_{len(participants)}{placeholder_path.suffix}"
            try:
                shutil.copy2(placeholder_path, dest)
                image_path = str(dest)
            except Exception:
                image_path = str(placeholder_path)

        p: dict[str, Any] = {
            "land": land,
            "rufname": rufname,
            "couch": couch,
            "image_path": image_path,
        }
        if email_ok:
            p["email"] = _str(cell(row_idx, DATA_EMAIL))
        if phone_ok:
            p["phone"] = _str(cell(row_idx, DATA_PHONE))
        if nachname_ok:
            p["nachname"] = _str(cell(row_idx, DATA_FAMILIENNAME))
        if vorname_ok:
            p["vorname"] = _str(cell(row_idx, DATA_VORNAME))

        participants.append(p)

    wb.close()
    return participants
