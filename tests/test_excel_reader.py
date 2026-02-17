"""Tests for excel_reader: consent filtering, edge cases."""
from __future__ import annotations

from pathlib import Path

import pytest

from excel_reader import load_participants
from tests.conftest import build_sample_xlsx


def test_load_participants_empty_sheet(tmp_path: Path, placeholder_path: Path) -> None:
    """Only headers, no data rows: result is empty list."""
    xlsx = build_sample_xlsx(tmp_path, [])
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(xlsx, placeholder_path, image_output_dir=out_dir)
    assert result == []


def test_load_participants_no_consent(tmp_path: Path, placeholder_path: Path) -> None:
    """Row with Teilnehmyliste=false is excluded."""
    xlsx = build_sample_xlsx(tmp_path, [
        {
            "Teilnehmyliste": False,
            "Land": "DE",
            "Rufname/Pseudonym": "Test",
            "Teilnehmyliste_Couch": "",
        },
    ])
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(xlsx, placeholder_path, image_output_dir=out_dir)
    assert len(result) == 0


def test_load_participants_consent_true_string(tmp_path: Path, placeholder_path: Path) -> None:
    """Teilnehmyliste='True' (string) is treated as truthy."""
    xlsx = build_sample_xlsx(tmp_path, [
        {
            "Teilnehmyliste": "True",
            "Land": "DE",
            "Rufname/Pseudonym": "Test",
            "Teilnehmyliste_Couch": "",
        },
    ])
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(xlsx, placeholder_path, image_output_dir=out_dir)
    assert len(result) == 1
    assert result[0]["land"] == "DE"
    assert result[0]["rufname"] == "Test"


def test_load_participants_partial_consent(tmp_path: Path, placeholder_path: Path) -> None:
    """Only consented fields appear in participant dict."""
    xlsx = build_sample_xlsx(tmp_path, [
        {
            "Teilnehmyliste": True,
            "Teilnehmyliste E-Mail": True,
            "Teilnehmyliste Telefonnummer": False,
            "Teilnehmyliste Nachname": False,
            "Teilnehmerliste Vorname": False,
            "Teilnehmyliste Bild": False,
            "Land": "CH",
            "Rufname/Pseudonym": "Alias",
            "Teilnehmyliste_Couch": "",
            "E-Mail Adresse": "a@b.ch",
            "Telefonnummer (mit Ländercode!)": "+41",
            "Familiename": "Secret",
            "Vorname": "Alice",
        },
    ])
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(xlsx, placeholder_path, image_output_dir=out_dir)
    assert len(result) == 1
    p = result[0]
    assert p["land"] == "CH"
    assert p["rufname"] == "Alias"
    assert p["email"] == "a@b.ch"
    assert "phone" not in p
    assert "nachname" not in p
    assert "vorname" not in p
    assert "image_path" in p


def test_load_participants_unicode(tmp_path: Path, placeholder_path: Path) -> None:
    """Unicode in fields is preserved."""
    xlsx = build_sample_xlsx(tmp_path, [
        {
            "Teilnehmyliste": 1,
            "Land": "DE",
            "Rufname/Pseudonym": "Müller-Größer",
            "Teilnehmyliste_Couch": "Café",
        },
    ])
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(xlsx, placeholder_path, image_output_dir=out_dir)
    assert len(result) == 1
    assert result[0]["rufname"] == "Müller-Größer"
    assert result[0]["couch"] == "Café"


def test_load_participants_empty_strings(tmp_path: Path, placeholder_path: Path) -> None:
    """None and empty strings become empty string in output."""
    xlsx = build_sample_xlsx(tmp_path, [
        {
            "Teilnehmyliste": "x",
            "Land": "",
            "Rufname/Pseudonym": None,
            "Teilnehmyliste_Couch": "  ",
        },
    ])
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(xlsx, placeholder_path, image_output_dir=out_dir)
    assert len(result) == 1
    assert result[0]["land"] == ""
    assert result[0]["rufname"] == ""
    assert result[0]["couch"] == ""


def test_load_participants_nonexistent_xlsx(placeholder_path: Path, tmp_path: Path) -> None:
    """Missing xlsx raises."""
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    with pytest.raises(Exception):
        load_participants(tmp_path / "does_not_exist.xlsx", placeholder_path, image_output_dir=out_dir)


def test_load_participants_multiple(tmp_path: Path, placeholder_path: Path) -> None:
    """Multiple rows with consent yield multiple participants."""
    xlsx = build_sample_xlsx(tmp_path, [
        {"Teilnehmyliste": True, "Land": "DE", "Rufname/Pseudonym": "A", "Teilnehmyliste_Couch": ""},
        {"Teilnehmyliste": False, "Land": "AT", "Rufname/Pseudonym": "B", "Teilnehmyliste_Couch": ""},
        {"Teilnehmyliste": "1", "Land": "CH", "Rufname/Pseudonym": "C", "Teilnehmyliste_Couch": ""},
    ])
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(xlsx, placeholder_path, image_output_dir=out_dir)
    assert len(result) == 2
    assert result[0]["rufname"] == "A"
    assert result[1]["rufname"] == "C"
