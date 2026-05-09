"""Google Drive uploader for generated offer files.

Uploads each docx + pdf into a product-specific subfolder of the
ENCON Offers Drive (owner: process@encon.in).

Auth is via a Service Account JSON pasted into Railway env var
GOOGLE_SERVICE_ACCOUNT_JSON. The target subfolder for VLPH/HLPH
offers is set via GOOGLE_DRIVE_FOLDER_LADLE_ID. Other product
types fall back to GOOGLE_DRIVE_FOLDER_DEFAULT_ID (the parent
folder) if set; otherwise upload is skipped silently.

Failures are logged and swallowed — Drive being unreachable
should never break offer generation.
"""

from __future__ import annotations

import json
import os
import threading
import traceback
from typing import Optional


_SCOPES = ["https://www.googleapis.com/auth/drive.file"]

# Cached Drive service so we don't rebuild credentials on every upload.
_service = None
_service_error: Optional[str] = None


def _get_service():
    """Build a Google Drive v3 service from GOOGLE_SERVICE_ACCOUNT_JSON.
    Returns None when credentials aren't configured or fail to load."""
    global _service, _service_error
    if _service is not None or _service_error is not None:
        return _service

    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        _service_error = "GOOGLE_SERVICE_ACCOUNT_JSON not set"
        return None

    try:
        info = json.loads(raw)
        from google.oauth2 import service_account
        from googleapiclient.discovery import build

        creds = service_account.Credentials.from_service_account_info(
            info, scopes=_SCOPES,
        )
        _service = build("drive", "v3", credentials=creds, cache_discovery=False)
        return _service
    except Exception as e:
        _service_error = f"Drive auth failed: {e}"
        print(f"WARN: drive_uploader auth: {_service_error}")
        return None


def _folder_id_for_product(product_type: str) -> Optional[str]:
    """Map a product_type string to the Drive folder ID it belongs in.
    Returns None if no folder is configured for this kind of offer
    (caller should skip the upload in that case)."""
    pt = (product_type or "").lower()
    if "vertical" in pt or "horizontal" in pt or "ladle" in pt:
        return os.environ.get("GOOGLE_DRIVE_FOLDER_LADLE_ID", "").strip() or None
    if "tundish" in pt:
        return os.environ.get("GOOGLE_DRIVE_FOLDER_TUNDISH_ID", "").strip() or None
    # Fallback for BTF / SNSF BRF / Regen / unknown — use the parent folder
    # if configured. Otherwise skip the upload.
    return os.environ.get("GOOGLE_DRIVE_FOLDER_DEFAULT_ID", "").strip() or None


def upload_offer(local_path: str, filename: str, product_type: str) -> Optional[str]:
    """Upload one offer file to the appropriate Drive folder.

    Returns the Drive web view link on success, or None if the upload
    was skipped (no credentials / no folder configured) or failed."""
    if not os.path.exists(local_path):
        return None

    folder_id = _folder_id_for_product(product_type)
    if not folder_id:
        return None

    service = _get_service()
    if service is None:
        return None

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
    """Fire-and-forget upload on a background thread. Caller doesn't
    wait for the result; the API response stays fast even when Drive
    is slow."""
    t = threading.Thread(
        target=upload_offer,
        args=(local_path, filename, product_type),
        daemon=True,
    )
    t.start()
