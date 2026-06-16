"""Google Drive uploader for generated offer files (OAuth flow).

Uses an OAuth refresh token captured by /auth/drive/login one-time
sign-in (process@encon.in). The refresh token lives in vlph.db's
oauth_tokens table so it survives redeploys.

Routing:
  Vertical / Horizontal Ladle Preheater  -> GOOGLE_DRIVE_FOLDER_LADLE_ID
  Tundish                                 -> GOOGLE_DRIVE_FOLDER_TUNDISH_ID
  Other product types                     -> GOOGLE_DRIVE_FOLDER_DEFAULT_ID

Env vars required for OAuth:
  GOOGLE_OAUTH_CLIENT_ID
  GOOGLE_OAUTH_CLIENT_SECRET
  GOOGLE_OAUTH_REDIRECT_URI    (e.g. https://automation.encon.co.in/auth/drive/callback)

Failures are logged and swallowed — Drive being unreachable should
never break offer generation.
"""

from __future__ import annotations

import os
import sqlite3
import threading
import traceback
from typing import Optional


_SCOPES = ["https://www.googleapis.com/auth/drive.file"]
_DB_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "vlph.db"
)
_TOKEN_KEY = "drive_refresh_token"

# Cached Drive service. Cleared whenever the stored refresh token changes.
_service = None


def _ensure_token_table():
    conn = sqlite3.connect(_DB_PATH)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS oauth_tokens "
        "(key TEXT PRIMARY KEY, value TEXT)"
    )
    conn.commit()
    conn.close()


def get_refresh_token() -> Optional[str]:
    # Prefer an env-var token (survives Railway redeploys, which wipe vlph.db).
    # Fall back to the vlph.db oauth_tokens table for local dev.
    env_token = os.environ.get("GOOGLE_DRIVE_REFRESH_TOKEN", "").strip()
    if env_token:
        return env_token
    _ensure_token_table()
    conn = sqlite3.connect(_DB_PATH)
    row = conn.execute(
        "SELECT value FROM oauth_tokens WHERE key=?", (_TOKEN_KEY,)
    ).fetchone()
    conn.close()
    return row[0] if row else None


def save_refresh_token(token: str) -> None:
    """Persist the refresh token. Drops the cached Drive service so the
    next upload picks up the new credentials."""
    global _service
    _ensure_token_table()
    conn = sqlite3.connect(_DB_PATH)
    conn.execute(
        "INSERT INTO oauth_tokens (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (_TOKEN_KEY, token),
    )
    conn.commit()
    conn.close()
    _service = None


def is_authorized() -> bool:
    """Quick check used by the UI to show 'Connect Drive' vs 'Connected'."""
    return bool(get_refresh_token())


def _get_service():
    """Build a Google Drive v3 service from the saved refresh token.
    Returns None when not yet authorised or env vars are missing."""
    global _service
    if _service is not None:
        return _service

    refresh_token = get_refresh_token()
    if not refresh_token:
        return None

    client_id = os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "").strip()
    client_secret = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "").strip()
    if not client_id or not client_secret:
        print("WARN: drive_uploader missing GOOGLE_OAUTH_CLIENT_ID / SECRET env vars")
        return None

    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build

        creds = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret,
            scopes=_SCOPES,
        )
        _service = build("drive", "v3", credentials=creds, cache_discovery=False)
        return _service
    except Exception as e:
        print(f"WARN: drive_uploader auth failed: {e}")
        traceback.print_exc()
        return None


def _folder_id_for_product(product_type: str) -> Optional[str]:
    """Map a product_type string to the Drive folder ID it belongs in.
    Env vars take precedence over the hard-coded fallback below."""
    pt = (product_type or "").lower()
    if "combined" in pt:
        # Dedicated "Combined Offers" folder. Env override wins; otherwise the
        # uploader auto-creates/finds it by name (see _ensure_combined_folder).
        return os.environ.get("GOOGLE_DRIVE_FOLDER_COMBINED_ID", "").strip() or None
    if "vertical" in pt or "horizontal" in pt or "ladle" in pt:
        return os.environ.get("GOOGLE_DRIVE_FOLDER_LADLE_ID", "").strip() or None
    if "tundish" in pt:
        return os.environ.get("GOOGLE_DRIVE_FOLDER_TUNDISH_ID", "").strip() or None
    if "recup" in pt:
        # User-supplied recuperator offers folder
        # (OFFERS > 2026 > Recuperator). Env var override available
        # via GOOGLE_DRIVE_FOLDER_RECUP_ID.
        return (os.environ.get("GOOGLE_DRIVE_FOLDER_RECUP_ID", "").strip()
                or "1i4wbS9JpSJfC5fJLSgPEv-z1_Dm9Jr2v")
    return os.environ.get("GOOGLE_DRIVE_FOLDER_DEFAULT_ID", "").strip() or None


def drive_status(product_type: str = "") -> dict:
    """Whether a generated file for `product_type` will actually reach Drive.
    Used to surface live feedback on the result screens. Upload itself runs
    async, but if we're authorised AND a folder is configured it succeeds in
    the background; otherwise it is skipped for the returned reason."""
    if not is_authorized():
        return {"ok": False, "reason": "not_connected",
                "msg": "Drive not connected — files not uploaded"}
    # Combined offers get a dedicated 'Combined Offers' folder that is
    # auto-created if absent, so a folder is always available once authorised.
    if "combined" in (product_type or "").lower():
        return {"ok": True, "reason": "ok", "msg": "Saved to Google Drive (Combined Offers)"}
    if not _folder_id_for_product(product_type):
        return {"ok": False, "reason": "no_folder",
                "msg": "Drive folder not configured — files not uploaded"}
    return {"ok": True, "reason": "ok", "msg": "Saved to Google Drive"}


_COMBINED_FOLDER_NAME = "Combined Offers"
_combined_folder_id = None   # cached app-created folder id


def _ensure_combined_folder(service) -> Optional[str]:
    """Find (or create) the 'Combined Offers' folder and return its id.
    Cached for the process; survives redeploys because the search finds the
    folder this app created last time (drive.file scope sees app-created files)."""
    global _combined_folder_id
    if _combined_folder_id:
        return _combined_folder_id
    try:
        q = ("mimeType='application/vnd.google-apps.folder' and trashed=false "
             f"and name='{_COMBINED_FOLDER_NAME}'")
        res = service.files().list(q=q, spaces="drive", pageSize=1,
                                   fields="files(id,name)").execute()
        files = res.get("files", [])
        if files:
            _combined_folder_id = files[0]["id"]
        else:
            meta = {"name": _COMBINED_FOLDER_NAME,
                    "mimeType": "application/vnd.google-apps.folder"}
            created = service.files().create(body=meta, fields="id").execute()
            _combined_folder_id = created.get("id")
            print(f"INFO: created Drive folder '{_COMBINED_FOLDER_NAME}' id={_combined_folder_id}")
        return _combined_folder_id
    except Exception as e:
        print(f"WARN: could not ensure '{_COMBINED_FOLDER_NAME}' folder: {e}")
        return None


def upload_offer(local_path: str, filename: str, product_type: str) -> Optional[str]:
    """Upload one offer file to the appropriate Drive folder.

    Returns the Drive web view link on success, or None if the upload
    was skipped (no auth / no folder configured) or failed."""
    if not os.path.exists(local_path):
        print(f"WARN: drive upload skipped - file not found: {local_path}")
        return None

    service = _get_service()
    if service is None:
        print(f"WARN: drive upload skipped - service unavailable (not authorised or missing client creds)")
        return None

    folder_id = _folder_id_for_product(product_type)
    if not folder_id and "combined" in (product_type or "").lower():
        folder_id = _ensure_combined_folder(service)   # auto-create the dedicated folder
    if not folder_id:
        print(f"WARN: drive upload skipped - no folder configured for product_type={product_type!r}")
        return None

    print(f"INFO: drive upload starting - filename={filename!r} product_type={product_type!r} folder_id={folder_id[:8]}...")

    try:
        from googleapiclient.http import MediaFileUpload
        ext = os.path.splitext(filename)[1].lower()
        mime = {
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".pdf":  "application/pdf",
            ".json": "application/json",
        }.get(ext, "application/octet-stream")

        media = MediaFileUpload(local_path, mimetype=mime, resumable=False)
        meta = {"name": filename, "parents": [folder_id]}
        created = service.files().create(
            body=meta, media_body=media, fields="id, webViewLink",
            supportsAllDrives=True,
        ).execute()
        return created.get("webViewLink")
    except Exception as e:
        print(f"WARN: drive upload failed for {filename}: {e}")
        traceback.print_exc()
        return None


def upload_offer_async(local_path: str, filename: str, product_type: str) -> None:
    """Fire-and-forget upload on a background thread."""
    t = threading.Thread(
        target=upload_offer,
        args=(local_path, filename, product_type),
        daemon=True,
    )
    t.start()
