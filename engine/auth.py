"""
Portal authentication — two roles (admin / user), credentials from env vars.

Sessions are stateless signed tokens stored in an HTTP-only cookie (HMAC-SHA256
over the payload, so no DB/session store is needed and it survives redeploys).

Environment variables (set these in Railway; the dev fallbacks below are only
for local use and should NOT be relied on in production):

    AUTH_SECRET       signing key for session cookies   (default: dev key)
    ADMIN_USERNAME    admin login name                  (default: "admin")
    ADMIN_PASSWORD    admin password                    (default: "admin")
    USER_USERNAME     non-admin login name              (default: "user")
    USER_PASSWORD     non-admin password                (default: "user")
"""

import os
import json
import time
import hmac
import base64
import hashlib

SESSION_COOKIE = "encon_session"
SESSION_MAX_AGE = 60 * 60 * 12  # 12 hours


def _secret() -> bytes:
    return (os.environ.get("AUTH_SECRET") or "dev-insecure-secret-change-me").encode()


def _users() -> dict:
    """username -> {password, role}. Read fresh each call so env changes apply."""
    return {
        (os.environ.get("ADMIN_USERNAME") or "admin"): {
            "password": os.environ.get("ADMIN_PASSWORD") or "admin",
            "role": "admin",
        },
        (os.environ.get("USER_USERNAME") or "user"): {
            "password": os.environ.get("USER_PASSWORD") or "user",
            "role": "user",
        },
    }


def verify_credentials(username: str, password: str):
    """Return the role ('admin'/'user') for valid credentials, else None."""
    user = _users().get((username or "").strip())
    if user and hmac.compare_digest(user["password"], password or ""):
        return user["role"]
    return None


def make_token(username: str, role: str) -> str:
    payload = {"u": username, "r": role, "exp": int(time.time()) + SESSION_MAX_AGE}
    raw = base64.urlsafe_b64encode(json.dumps(payload).encode()).decode()
    sig = hmac.new(_secret(), raw.encode(), hashlib.sha256).hexdigest()
    return f"{raw}.{sig}"


def verify_token(token: str):
    """Return the payload dict for a valid, unexpired token, else None."""
    if not token or "." not in token:
        return None
    raw, sig = token.rsplit(".", 1)
    expected = hmac.new(_secret(), raw.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(sig, expected):
        return None
    try:
        payload = json.loads(base64.urlsafe_b64decode(raw.encode()))
    except Exception:
        return None
    if int(payload.get("exp", 0)) < time.time():
        return None
    return payload
