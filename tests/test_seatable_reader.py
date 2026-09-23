"""Tests for seatable_reader (mocked SeaTable SDK)."""
from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from seatable_reader import (
    CONSENT_BILD,
    CONSENT_EMAIL,
    CONSENT_LIST,
    CONSENT_NACHNAME,
    CONSENT_PHONE,
    CONSENT_VORNAME,
    DATA_BILD,
    DATA_COUCH,
    DATA_EMAIL,
    DATA_FAMILIENNAME,
    DATA_LAND,
    DATA_PHONE,
    DATA_RUFNAME,
    DATA_VORNAME,
    BaseInfo,
    SeaTableAuthError,
    SeaTableSession,
    _image_urls_from_cell,
    _truthy,
    list_bases,
    load_participants,
    login,
)


def test_truthy_variants() -> None:
    assert _truthy(True) is True
    assert _truthy("Ja") is True
    assert _truthy("x") is True
    assert _truthy(1) is True
    assert _truthy(False) is False
    assert _truthy("") is False
    assert _truthy(None) is False


def test_image_urls_from_cell() -> None:
    assert _image_urls_from_cell(None) == []
    assert _image_urls_from_cell("https://example/img.jpg") == ["https://example/img.jpg"]
    assert _image_urls_from_cell(["https://a", {"url": "https://b"}]) == [
        "https://a",
        "https://b",
    ]


def test_login_success() -> None:
    account = MagicMock()
    account.token = "tok"

    with patch("seatable_reader.Account", return_value=account) as AccountCls:
        session = login("user@example.com", "secret", "https://cloud.seatable.io")

    AccountCls.assert_called_once_with(
        "user@example.com", "secret", "https://cloud.seatable.io"
    )
    account.auth.assert_called_once()
    assert isinstance(session, SeaTableSession)
    assert session.account is account


def test_login_failure() -> None:
    account = MagicMock()
    account.auth.side_effect = ConnectionError(400, "bad credentials")
    with patch("seatable_reader.Account", return_value=account):
        with pytest.raises(SeaTableAuthError):
            login("user", "wrong")


def test_list_bases_parses_workspaces() -> None:
    account = MagicMock()
    account.list_workspaces.return_value = {
        "workspace_list": [
            {
                "id": 10,
                "name": "Personal",
                "table_list": [
                    {"name": "PAN Winter", "workspace_id": 10},
                    {"name": "PAN Herbst", "workspace_id": 10},
                ],
            },
            {
                "id": 20,
                "name": "Group",
                "table_list": [{"name": "Other", "workspace_id": 20}],
            },
        ]
    }
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")
    bases = list_bases(session)
    assert [b.name for b in bases] == ["Other", "PAN Herbst", "PAN Winter"]
    assert all(isinstance(b, BaseInfo) for b in bases)


def _row(**overrides: object) -> dict:
    base = {
        CONSENT_LIST: True,
        CONSENT_EMAIL: False,
        CONSENT_PHONE: False,
        CONSENT_NACHNAME: False,
        CONSENT_VORNAME: False,
        CONSENT_BILD: False,
        DATA_LAND: "DE",
        DATA_RUFNAME: "Test",
        DATA_COUCH: "",
        DATA_EMAIL: "a@b.de",
        DATA_PHONE: "+49",
        DATA_FAMILIENNAME: "Secret",
        DATA_VORNAME: "Alice",
        DATA_BILD: [],
    }
    base.update(overrides)
    return base


def test_load_participants_consent_and_partial_fields(
    tmp_path: Path, placeholder_path: Path
) -> None:
    base = MagicMock()
    base.get_metadata.return_value = {
        "tables": [
            {
                "name": "Anmeldung",
                "columns": [{"name": CONSENT_LIST}, {"name": DATA_RUFNAME}],
            }
        ]
    }
    base.list_views.return_value = [{"name": "Default View"}]
    base.list_rows.return_value = [
        _row(**{CONSENT_LIST: False, DATA_RUFNAME: "Skip"}),
        _row(
            **{
                CONSENT_EMAIL: True,
                CONSENT_BILD: True,
                DATA_BILD: [
                    "https://cloud.seatable.io/workspace/1/asset/"
                    "11111111-1111-1111-1111-111111111111/images/2026-01/a.jpg"
                ],
                DATA_RUFNAME: "Keep",
            }
        ),
    ]
    base.dtable_uuid = "11111111-1111-1111-1111-111111111111"

    def fake_download(url: str, save_path: str) -> None:
        Path(save_path).write_bytes(b"fake-image")

    base.download_file.side_effect = fake_download

    account = MagicMock()
    account.get_base.return_value = base
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(
        session, 10, "PAN Winter", placeholder_path, image_output_dir=out_dir
    )
    assert len(result) == 1
    p = result[0]
    assert p["rufname"] == "Keep"
    assert p["email"] == "a@b.de"
    assert "phone" not in p
    assert "nachname" not in p
    assert Path(p["image_path"]).is_file()
    assert Path(p["image_path"]).read_bytes() == b"fake-image"


def test_load_participants_placeholder_without_image_consent(
    tmp_path: Path, placeholder_path: Path
) -> None:
    base = MagicMock()
    base.get_metadata.return_value = {
        "tables": [{"name": "T", "columns": [{"name": CONSENT_LIST}]}]
    }
    base.list_views.return_value = []
    base.list_rows.return_value = [_row(**{CONSENT_BILD: False, DATA_RUFNAME: "NoPic"})]
    account = MagicMock()
    account.get_base.return_value = base
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(
        session, 1, "Base", placeholder_path, image_output_dir=out_dir
    )
    assert len(result) == 1
    assert Path(result[0]["image_path"]).read_bytes() == placeholder_path.read_bytes()
