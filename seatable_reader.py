"""
Load PAN meetup participants from SeaTable Cloud.

Authenticates with account credentials (in-memory only), lists bases, fetches
rows and images via the official seatable-api SDK.
"""
from __future__ import annotations

import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

import requests
from seatable_api import Account

DEFAULT_SERVER_URL = "https://cloud.seatable.io"

# Column names in the SeaTable base (exact match)
CONSENT_LIST = "Teilnehmyliste"
CONSENT_EMAIL = "Teilnehmyliste E-Mail"
CONSENT_PHONE = "Teilnehmyliste Telefonnummer"
CONSENT_NACHNAME = "Teilnehmyliste Nachname"
CONSENT_VORNAME = "Teilnehmerliste Vorname"  # note: "Teilnehmerliste" spelling
CONSENT_BILD = "Teilnehmyliste Bild"

DATA_LAND = "Land"
DATA_PLZ = "PLZ"
DATA_ORT = "Ort"
DATA_RUFNAME = "Rufname/Pseudonym"
DATA_COUCH = "Teilnehmyliste_Couch"
DATA_EMAIL = "E-Mail Adresse"
DATA_PHONE = "Telefonnummer (mit Ländercode!)"
DATA_FAMILIENNAME = "Familiename"
DATA_VORNAME = "Vorname"
DATA_BILD = "Bild"


@dataclass(frozen=True)
class BaseInfo:
    """A SeaTable base (meetup) reachable by the logged-in account."""

    workspace_id: int
    name: str
    workspace_name: str = ""

    @property
    def label(self) -> str:
        if self.workspace_name:
            return f"{self.name}  ({self.workspace_name})"
        return self.name


@dataclass
class SeaTableSession:
    """Authenticated SeaTable account; credentials stay in memory only."""

    account: Account
    server_url: str


class SeaTableAuthError(Exception):
    """Login or token failure."""


class SeaTableError(Exception):
    """General SeaTable API / data error."""


def _truthy(value: Any) -> bool:
    """Normalize SeaTable / Excel-like booleans and strings to bool."""
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    s = str(value).strip().lower()
    return s in ("true", "1", "yes", "ja", "x")


def _str(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    return s if s else ""


def _format_connection_error(
    exc: ConnectionError,
    *,
    context: str = "SeaTable-Anfrage fehlgeschlagen",
) -> str:
    args = getattr(exc, "args", ())
    if len(args) >= 2:
        status, body = args[0], args[1]
        return f"{context} (HTTP {status}): {body}"
    detail = str(exc).strip()
    return f"{context}: {detail}" if detail else context


def login(
    username: str,
    password: str,
    server_url: str = DEFAULT_SERVER_URL,
) -> SeaTableSession:
    """
    Authenticate against SeaTable Cloud (or another server).
    Raises SeaTableAuthError on failure.
    """
    username = username.strip()
    server_url = (server_url or DEFAULT_SERVER_URL).strip().rstrip("/")
    if not username or not password:
        raise SeaTableAuthError("Benutzername und Passwort sind erforderlich.")
    account = Account(username, password, server_url)
    try:
        account.auth()
    except ConnectionError as e:
        raise SeaTableAuthError(
            _format_connection_error(e, context="Anmeldung fehlgeschlagen")
        ) from e
    except Exception as e:
        raise SeaTableAuthError(str(e) or "Anmeldung fehlgeschlagen.") from e
    if not account.token:
        raise SeaTableAuthError("Anmeldung fehlgeschlagen (kein Token).")
    return SeaTableSession(account=account, server_url=server_url)


def list_bases(session: SeaTableSession) -> list[BaseInfo]:
    """
    List all bases visible to the account (personal + group workspaces).
    """
    try:
        data = session.account.list_workspaces()
    except ConnectionError as e:
        raise SeaTableError(
            _format_connection_error(e, context="Bases konnten nicht geladen werden")
        ) from e
    except Exception as e:
        raise SeaTableError(str(e) or "Bases konnten nicht geladen werden.") from e

    bases: list[BaseInfo] = []
    seen: set[tuple[int, str]] = set()

    def _add(workspace_id: Any, name: Any, workspace_name: str = "") -> None:
        if workspace_id is None or not name:
            return
        try:
            wid = int(workspace_id)
        except (TypeError, ValueError):
            return
        base_name = str(name).strip()
        if not base_name:
            return
        key = (wid, base_name)
        if key in seen:
            return
        seen.add(key)
        bases.append(
            BaseInfo(
                workspace_id=wid,
                name=base_name,
                workspace_name=workspace_name.strip(),
            )
        )

    workspace_list = (data or {}).get("workspace_list") or []
    for ws in workspace_list:
        if not isinstance(ws, dict):
            continue
        wid = ws.get("id")
        ws_name = _str(ws.get("name") or ws.get("owner_name") or "")
        for key in ("table_list", "shared_table_list"):
            for entry in ws.get(key) or []:
                if isinstance(entry, dict):
                    _add(entry.get("workspace_id", wid), entry.get("name"), ws_name)

    # Shared / starred lists may appear at the top level
    for key in ("shared_table_list", "starred_dtable_list"):
        for entry in (data or {}).get(key) or []:
            if isinstance(entry, dict):
                _add(
                    entry.get("workspace_id"),
                    entry.get("name"),
                    _str(entry.get("workspace_name") or ""),
                )

    bases.sort(key=lambda b: b.name.lower())
    return bases


def _open_base(session: SeaTableSession, workspace_id: int, base_name: str):
    try:
        return session.account.get_base(workspace_id, base_name)
    except ConnectionError as e:
        raise SeaTableError(
            _format_connection_error(
                e, context=f"Base „{base_name}“ konnte nicht geöffnet werden"
            )
        ) from e
    except Exception as e:
        raise SeaTableError(str(e) or f"Base „{base_name}“ konnte nicht geöffnet werden.") from e


def _pick_table_name(base, table_name: str | None = None) -> str:
    if table_name:
        return table_name
    try:
        meta = base.get_metadata()
    except Exception as e:
        raise SeaTableError(str(e) or "Metadaten der Base konnten nicht geladen werden.") from e
    tables = (meta or {}).get("tables") or []
    if not tables:
        raise SeaTableError("Die Base enthält keine Tabellen.")

    for table in tables:
        colnames = {c.get("name") for c in (table.get("columns") or [])}
        if CONSENT_LIST in colnames:
            return str(table["name"])
    return str(tables[0]["name"])


def _pick_view_name(base, table_name: str, view_name: str | None = None) -> str | None:
    if view_name:
        return view_name
    try:
        views = base.list_views(table_name) or []
    except Exception:
        return None
    for view in views:
        name = _str(view.get("name") if isinstance(view, dict) else view)
        if "kontaktliste" in name.lower():
            return name
    return None


def _image_urls_from_cell(value: Any) -> list[str]:
    """Normalize an image column cell to a list of absolute or asset URLs."""
    if value is None:
        return []
    if isinstance(value, str):
        s = value.strip()
        return [s] if s else []
    if not isinstance(value, (list, tuple)):
        return []
    urls: list[str] = []
    for item in value:
        if isinstance(item, str):
            s = item.strip()
            if s:
                urls.append(s)
        elif isinstance(item, dict):
            u = item.get("url")
            if u and str(u).strip():
                urls.append(str(u).strip())
    return urls


def _suffix_from_url(url: str) -> str:
    path = unquote(urlparse(url).path)
    ext = Path(path).suffix.lower()
    if ext in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"):
        return ext
    return ".jpg"


def _download_image(base, url: str, dest: Path) -> bool:
    """Download one SeaTable image URL to dest. Returns True on success."""
    try:
        base.download_file(url, str(dest))
        return dest.is_file() and dest.stat().st_size > 0
    except Exception:
        # Fallback: path-based download link (URLs that include /images/...)
        try:
            marker = "/images/"
            idx = url.find(marker)
            if idx < 0:
                marker = "/files/"
                idx = url.find(marker)
            if idx < 0:
                return False
            path = unquote(url[idx:])
            link = base.get_file_download_link(path)
            if not link:
                return False
            resp = requests.get(link, timeout=60)
            if resp.status_code != 200 or not resp.content:
                return False
            dest.write_bytes(resp.content)
            return True
        except Exception:
            return False


def load_participants(
    session: SeaTableSession,
    workspace_id: int,
    base_name: str,
    placeholder_image_path: str | Path,
    image_output_dir: Path | None = None,
    table_name: str | None = None,
    view_name: str | None = None,
) -> list[dict[str, Any]]:
    """
    Load rows from a SeaTable base, filter by Teilnehmyliste, apply per-field
    consent, and download images when consented.

    Returns the same participant dict shape as the former Excel loader:
    land, plz, ort, rufname, couch, image_path, and optional email/phone/
    nachname/vorname.
    """
    placeholder_image_path = Path(placeholder_image_path)
    if image_output_dir is None:
        image_output_dir = Path(tempfile.mkdtemp(prefix="pan_contact_images_"))
    image_output_dir = Path(image_output_dir)
    image_output_dir.mkdir(parents=True, exist_ok=True)
    placeholder_path = placeholder_image_path.resolve()

    base = _open_base(session, workspace_id, base_name)
    resolved_table = _pick_table_name(base, table_name)
    resolved_view = _pick_view_name(base, resolved_table, view_name)

    try:
        if resolved_view:
            rows = base.list_rows(resolved_table, view_name=resolved_view) or []
        else:
            rows = base.list_rows(resolved_table) or []
    except Exception as e:
        raise SeaTableError(str(e) or "Zeilen konnten nicht geladen werden.") from e

    participants: list[dict[str, Any]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        if not _truthy(row.get(CONSENT_LIST)):
            continue

        email_ok = _truthy(row.get(CONSENT_EMAIL))
        phone_ok = _truthy(row.get(CONSENT_PHONE))
        nachname_ok = _truthy(row.get(CONSENT_NACHNAME))
        vorname_ok = _truthy(row.get(CONSENT_VORNAME))
        bild_ok = _truthy(row.get(CONSENT_BILD))

        image_path = str(placeholder_path)
        if bild_ok:
            urls = _image_urls_from_cell(row.get(DATA_BILD))
            if urls:
                ext = _suffix_from_url(urls[0])
                dest = image_output_dir / f"teilnehmer_{len(participants)}{ext}"
                if _download_image(base, urls[0], dest):
                    image_path = str(dest)
                else:
                    dest = image_output_dir / f"teilnehmer_{len(participants)}{placeholder_path.suffix}"
                    try:
                        shutil.copy2(placeholder_path, dest)
                        image_path = str(dest)
                    except Exception:
                        image_path = str(placeholder_path)
            else:
                dest = image_output_dir / f"teilnehmer_{len(participants)}{placeholder_path.suffix}"
                try:
                    shutil.copy2(placeholder_path, dest)
                    image_path = str(dest)
                except Exception:
                    image_path = str(placeholder_path)
        else:
            dest = image_output_dir / f"teilnehmer_{len(participants)}{placeholder_path.suffix}"
            try:
                shutil.copy2(placeholder_path, dest)
                image_path = str(dest)
            except Exception:
                image_path = str(placeholder_path)

        p: dict[str, Any] = {
            "land": _str(row.get(DATA_LAND)),
            "plz": _str(row.get(DATA_PLZ)),
            "ort": _str(row.get(DATA_ORT)),
            "rufname": _str(row.get(DATA_RUFNAME)),
            "couch": _str(row.get(DATA_COUCH)),
            "image_path": image_path,
        }
        if email_ok:
            p["email"] = _str(row.get(DATA_EMAIL))
        if phone_ok:
            p["phone"] = _str(row.get(DATA_PHONE))
        if nachname_ok:
            p["nachname"] = _str(row.get(DATA_FAMILIENNAME))
        if vorname_ok:
            p["vorname"] = _str(row.get(DATA_VORNAME))
        participants.append(p)

    return participants
