"""Tests for seatable_reader (mocked SeaTable SDK)."""
from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch
import threading

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
    SeaTableCancelled,
    SeaTableError,
    SeaTableSession,
    _cache_identity_from_url,
    _coerce_view_list,
    _image_urls_from_cell,
    _pick_view_name,
    _truthy,
    cache_path_for_url,
    choose_default_view,
    list_bases,
    list_view_names,
    load_participants,
    login,
    prune_image_cache,
    wipe_image_cache,
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


def test_format_connection_error_context() -> None:
    from seatable_reader import _format_connection_error

    msg = _format_connection_error(
        ConnectionError(429, "rate limited"),
        context="Bases konnten nicht geladen werden",
    )
    assert msg.startswith("Bases konnten nicht geladen werden")
    assert "429" in msg
    assert "Anmeldung" not in msg


def test_list_bases_accepts_workspace_id_alias() -> None:
    account = MagicMock()
    account.list_workspaces.return_value = {
        "workspace_list": [
            {
                "workspace_id": 99,
                "workspace_name": "Alias WS",
                "table_list": [{"name": "Meetup", "workspace_id": 99}],
            }
        ]
    }
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")
    bases = list_bases(session)
    assert len(bases) == 1
    assert bases[0].workspace_id == 99
    assert bases[0].name == "Meetup"


def test_load_participants_reports_progress_and_honours_cancel(
    tmp_path: Path, placeholder_path: Path
) -> None:
    base = MagicMock()
    base.get_metadata.return_value = {
        "tables": [{"name": "T", "columns": [{"name": CONSENT_LIST}]}]
    }
    base.list_views.return_value = []
    base.list_rows.return_value = [
        _row(
            **{
                CONSENT_BILD: True,
                DATA_BILD: [
                    "https://cloud.seatable.io/workspace/1/asset/"
                    "11111111-1111-1111-1111-111111111111/images/2026-01/a.jpg"
                ],
                DATA_RUFNAME: "A",
            }
        ),
        _row(
            **{
                CONSENT_BILD: True,
                DATA_BILD: [
                    "https://cloud.seatable.io/workspace/1/asset/"
                    "11111111-1111-1111-1111-111111111111/images/2026-01/b.jpg"
                ],
                DATA_RUFNAME: "B",
            }
        ),
    ]
    base.dtable_uuid = "11111111-1111-1111-1111-111111111111"
    base.download_file.side_effect = lambda url, path: Path(path).write_bytes(b"img")

    account = MagicMock()
    account.get_base.return_value = base
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")

    messages: list[str] = []
    cancel = threading.Event()

    def progress(msg: str, current: int, total: int) -> None:
        messages.append(msg)
        if "Bild 1/" in msg:
            cancel.set()

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    with pytest.raises(SeaTableCancelled):
        load_participants(
            session,
            1,
            "Base",
            placeholder_path,
            image_output_dir=out_dir,
            image_cache_dir=cache_dir,
            progress=progress,
            cancel_event=cancel,
        )
    assert any("Bild 1/" in m for m in messages)


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


def test_failed_png_download_keeps_intact_placeholder(
    tmp_path: Path, placeholder_path: Path
) -> None:
    """A failed .png download must not leave a truncated file as the participant image."""
    base = MagicMock()
    base.get_metadata.return_value = {
        "tables": [{"name": "T", "columns": [{"name": CONSENT_LIST}]}]
    }
    base.list_views.return_value = []
    base.list_rows.return_value = [
        _row(
            **{
                CONSENT_BILD: True,
                DATA_BILD: [
                    "https://cloud.seatable.io/workspace/1/asset/"
                    "11111111-1111-1111-1111-111111111111/images/2026-01/photo.png"
                ],
                DATA_RUFNAME: "Broken",
            }
        ),
    ]
    base.dtable_uuid = "11111111-1111-1111-1111-111111111111"

    def empty_download(url: str, save_path: str) -> None:
        # Simulate a failed transfer that creates an empty/truncated file.
        Path(save_path).write_bytes(b"")

    base.download_file.side_effect = empty_download
    base.get_file_download_link.side_effect = Exception("no link")

    account = MagicMock()
    account.get_base.return_value = base
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    result = load_participants(
        session,
        1,
        "Base",
        placeholder_path,
        image_output_dir=out_dir,
        image_cache_dir=cache_dir,
    )
    assert len(result) == 1
    image = Path(result[0]["image_path"])
    assert image.is_file()
    assert image.read_bytes() == placeholder_path.read_bytes()
    assert image.stat().st_size > 0
    assert not (out_dir / "teilnehmer_0.png.partial").exists()
    assert not list(cache_dir.glob("*.partial"))


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
    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    result = load_participants(
        session,
        10,
        "PAN Winter",
        placeholder_path,
        image_output_dir=out_dir,
        image_cache_dir=cache_dir,
    )
    assert len(result) == 1
    p = result[0]
    assert p["rufname"] == "Keep"
    assert p["email"] == "a@b.de"
    assert "phone" not in p
    assert "nachname" not in p
    assert Path(p["image_path"]).is_file()
    assert Path(p["image_path"]).read_bytes() == b"fake-image"
    assert any(cache_dir.iterdir())


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
    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    result = load_participants(
        session,
        1,
        "Base",
        placeholder_path,
        image_output_dir=out_dir,
        image_cache_dir=cache_dir,
    )
    assert len(result) == 1
    assert Path(result[0]["image_path"]).read_bytes() == placeholder_path.read_bytes()


def test_load_participants_propagates_view_list_error(
    tmp_path: Path, placeholder_path: Path
) -> None:
    base = MagicMock()
    base.get_metadata.return_value = {
        "tables": [{"name": "T", "columns": [{"name": CONSENT_LIST}]}]
    }
    base.list_views.side_effect = Exception("views down")
    account = MagicMock()
    account.get_base.return_value = base
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    with pytest.raises(SeaTableError, match="views down"):
        load_participants(
            session,
            1,
            "Base",
            placeholder_path,
            image_output_dir=out_dir,
            image_cache_dir=cache_dir,
        )


def test_choose_default_view_preferences() -> None:
    assert choose_default_view([]) is None
    assert choose_default_view(["Alpha", "Beta"]) == "Alpha"
    assert choose_default_view(["Alpha", "Default View", "Beta"]) == "Default View"
    assert choose_default_view(["Alpha", "default view", "Kontaktliste"]) == "Kontaktliste"
    assert choose_default_view(["Default View", "kontaktliste"]) == "kontaktliste"


def test_pick_view_name_uses_exact_defaults() -> None:
    base = MagicMock()
    base.list_views.return_value = {
        "views": [
            {"name": "All"},
            {"name": "Default View"},
            {"name": "My Kontaktliste Filter"},
        ]
    }
    assert _pick_view_name(base, "T") == "Default View"

    base.list_views.return_value = {
        "views": [
            {"name": "All"},
            {"name": "Kontaktliste"},
            {"name": "Default View"},
        ]
    }
    assert _pick_view_name(base, "T") == "Kontaktliste"

    base.list_views.return_value = {"views": [{"name": "Only"}]}
    assert _pick_view_name(base, "T") == "Only"

    assert _pick_view_name(base, "T", view_name="Custom") == "Custom"


def test_coerce_view_list_unwraps_sdk_dict() -> None:
    assert _coerce_view_list(None) == []
    assert _coerce_view_list([{"name": "A"}]) == [{"name": "A"}]
    wrapped = {
        "views": [
            {"name": "Default View"},
            {"name": "Kontaktliste"},
        ]
    }
    assert _coerce_view_list(wrapped) == wrapped["views"]
    # Iterating the dict itself would yield only the key "views".
    assert list(wrapped) == ["views"]


def test_list_view_names_resolves_table() -> None:
    base = MagicMock()
    base.get_metadata.return_value = {
        "tables": [
            {"name": "Other", "columns": [{"name": "X"}]},
            {"name": "Anmeldung", "columns": [{"name": CONSENT_LIST}]},
        ]
    }
    base.list_views.return_value = {
        "views": [
            {"name": "Default View"},
            {"name": "Kontaktliste"},
        ]
    }
    account = MagicMock()
    account.get_base.return_value = base
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")

    names = list_view_names(session, 1, "Base")
    assert names == ["Default View", "Kontaktliste"]
    base.list_views.assert_called_once_with("Anmeldung")


def test_cache_identity_includes_dtable_uuid() -> None:
    url_a = (
        "https://cloud.seatable.io/workspace/1/asset/"
        "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa/images/2026-01/photo.jpg"
    )
    url_b = (
        "https://cloud.seatable.io/workspace/1/asset/"
        "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb/images/2026-01/photo.jpg"
    )
    assert _cache_identity_from_url(url_a) == (
        "cloud.seatable.io/aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa/images/2026-01/photo.jpg"
    )
    assert _cache_identity_from_url(url_a) != _cache_identity_from_url(url_b)
    path_a = cache_path_for_url(url_a, Path("/tmp/cache"))
    path_b = cache_path_for_url(url_b, Path("/tmp/cache"))
    assert path_a.name != path_b.name
    assert path_a.suffix == ".jpg"

    relative = "/images/2026-01/photo.jpg"
    id_a = _cache_identity_from_url(
        relative,
        dtable_uuid="aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        server_url="https://cloud.seatable.io",
    )
    id_b = _cache_identity_from_url(
        relative,
        dtable_uuid="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
        server_url="https://cloud.seatable.io",
    )
    assert id_a != id_b
    assert "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa" in id_a


def test_image_cache_hit_skips_download(
    tmp_path: Path, placeholder_path: Path
) -> None:
    url = (
        "https://cloud.seatable.io/workspace/1/asset/"
        "11111111-1111-1111-1111-111111111111/images/2026-01/a.jpg"
    )
    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    cached = cache_path_for_url(
        url,
        cache_dir,
        server_url="https://cloud.seatable.io",
    )
    cached.write_bytes(b"cached-bytes")

    base = MagicMock()
    base.get_metadata.return_value = {
        "tables": [{"name": "T", "columns": [{"name": CONSENT_LIST}]}]
    }
    base.list_views.return_value = []
    base.list_rows.return_value = [
        _row(**{CONSENT_BILD: True, DATA_BILD: [url], DATA_RUFNAME: "Cached"})
    ]
    account = MagicMock()
    account.get_base.return_value = base
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(
        session,
        1,
        "Base",
        placeholder_path,
        image_output_dir=out_dir,
        image_cache_dir=cache_dir,
    )
    assert Path(result[0]["image_path"]).read_bytes() == b"cached-bytes"
    base.download_file.assert_not_called()
    base.get_file_download_link.assert_not_called()


def test_image_cache_expired_redownloads(
    tmp_path: Path, placeholder_path: Path
) -> None:
    import os
    import time

    import seatable_reader as sr

    url = (
        "https://cloud.seatable.io/workspace/1/asset/"
        "11111111-1111-1111-1111-111111111111/images/2026-01/a.jpg"
    )
    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    cached = cache_path_for_url(
        url,
        cache_dir,
        server_url="https://cloud.seatable.io",
    )
    cached.write_bytes(b"stale")
    old = time.time() - (sr.IMAGE_CACHE_TTL_S + 60)
    os.utime(cached, (old, old))

    base = MagicMock()
    base.get_metadata.return_value = {
        "tables": [{"name": "T", "columns": [{"name": CONSENT_LIST}]}]
    }
    base.list_views.return_value = []
    base.list_rows.return_value = [
        _row(**{CONSENT_BILD: True, DATA_BILD: [url], DATA_RUFNAME: "Stale"})
    ]

    def fake_download(u: str, path: str) -> None:
        Path(path).write_bytes(b"fresh")

    base.download_file.side_effect = fake_download
    account = MagicMock()
    account.get_base.return_value = base
    session = SeaTableSession(account=account, server_url="https://cloud.seatable.io")

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    result = load_participants(
        session,
        1,
        "Base",
        placeholder_path,
        image_output_dir=out_dir,
        image_cache_dir=cache_dir,
    )
    assert Path(result[0]["image_path"]).read_bytes() == b"fresh"
    assert cached.read_bytes() == b"fresh"
    base.download_file.assert_called_once()


def test_prune_and_wipe_image_cache(tmp_path: Path) -> None:
    import os
    import time

    import seatable_reader as sr

    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    fresh = cache_dir / "fresh.jpg"
    stale = cache_dir / "stale.jpg"
    fresh_partial = cache_dir / "x.jpg.abc123.partial"
    abandoned_partial = cache_dir / "y.jpg.def456.partial"
    fresh.write_bytes(b"ok")
    stale.write_bytes(b"old")
    fresh_partial.write_bytes(b"tmp")
    abandoned_partial.write_bytes(b"tmp")
    old = time.time() - (sr.IMAGE_CACHE_TTL_S + 10)
    os.utime(stale, (old, old))
    os.utime(abandoned_partial, (time.time() - (sr.PARTIAL_MAX_AGE_S + 10),) * 2)

    removed = prune_image_cache(cache_dir)
    assert removed == 2
    assert fresh.exists()
    assert fresh_partial.exists()
    assert not stale.exists()
    assert not abandoned_partial.exists()

    assert wipe_image_cache(cache_dir) == 2
    assert not fresh.exists()
    assert not fresh_partial.exists()


def test_wipe_image_cache_raises_on_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    cache_dir = tmp_path / "cache"
    cache_dir.mkdir()
    kept = cache_dir / "kept.jpg"
    kept.write_bytes(b"x")

    real_unlink = Path.unlink

    def flaky_unlink(self: Path, *args: object, **kwargs: object) -> None:
        if self.name == "kept.jpg":
            raise OSError("permission denied")
        return real_unlink(self, *args, **kwargs)

    monkeypatch.setattr(Path, "unlink", flaky_unlink)
    with pytest.raises(OSError, match="nicht vollständig"):
        wipe_image_cache(cache_dir)
