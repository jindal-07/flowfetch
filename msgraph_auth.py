"""
Microsoft Graph OAuth (delegated, auth-code flow) using MSAL.

Permissions requested: Files.Read.All, User.Read, offline_access.
(MSAL adds `openid` and `profile` automatically.)

Token caches are per-session and held in memory only. They're lost on
container restart, requiring users to re-sign-in. That's intentional --
no token persistence on disk, no secrets in the git repo.
"""

from __future__ import annotations

import os
import secrets
from typing import Optional

import msal


# ---------------------------------------------------------------------------
# Configuration (all via environment variables)
# ---------------------------------------------------------------------------

CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "").strip()
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "").strip()
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "common").strip() or "common"
REDIRECT_URI = os.environ.get("AZURE_REDIRECT_URI", "").strip()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Only the resource scopes here. MSAL injects `openid profile offline_access`.
SCOPES = ["https://graph.microsoft.com/Files.Read.All",
          "https://graph.microsoft.com/User.Read"]


def _looks_like_placeholder(v: str) -> bool:
    return bool(v) and (v.startswith("<") or v.endswith(">") or v.lower() == "your-tenant-id")


def is_configured() -> bool:
    if not (CLIENT_ID and CLIENT_SECRET and REDIRECT_URI):
        return False
    if any(_looks_like_placeholder(x) for x in (CLIENT_ID, CLIENT_SECRET, TENANT_ID, REDIRECT_URI)):
        return False
    return True


# ---------------------------------------------------------------------------
# Per-session token caches (in-memory)
# ---------------------------------------------------------------------------

# session_id -> {"cache": SerializableTokenCache, "account": dict, "email": str}
_SESSIONS: dict[str, dict] = {}


def new_session_id() -> str:
    return secrets.token_urlsafe(24)


def get_session(session_id: Optional[str]) -> Optional[dict]:
    if not session_id:
        return None
    return _SESSIONS.get(session_id)


def clear_session(session_id: Optional[str]) -> None:
    if session_id and session_id in _SESSIONS:
        _SESSIONS.pop(session_id, None)


# ---------------------------------------------------------------------------
# MSAL app helpers
# ---------------------------------------------------------------------------

def _build_msal_app(cache: Optional[msal.SerializableTokenCache] = None) -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache,
    )


def build_auth_url(state: str) -> str:
    """Return the Microsoft sign-in URL to redirect the user to."""
    app = _build_msal_app()
    return app.get_authorization_request_url(
        SCOPES,
        state=state,
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
    )


def acquire_token_from_code(code: str, session_id: str) -> dict:
    """
    Exchange an auth code for tokens, store the cache + account in our
    per-session dict, and return the raw token response.
    """
    cache = msal.SerializableTokenCache()
    app = _build_msal_app(cache)
    result = app.acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    if "error" in result:
        raise RuntimeError(
            f"Token exchange failed: {result.get('error')}: {result.get('error_description')}"
        )

    accounts = app.get_accounts()
    account = accounts[0] if accounts else None
    claims = result.get("id_token_claims") or {}
    email = (
        claims.get("preferred_username")
        or claims.get("email")
        or (account or {}).get("username")
        or ""
    )

    _SESSIONS[session_id] = {
        "cache": cache,
        "account": account,
        "email": email,
    }
    return result


def get_access_token(session_id: str) -> Optional[str]:
    """
    Return a valid access token for the session, refreshing silently if needed.
    Returns None if the session has no cached account.
    """
    sess = _SESSIONS.get(session_id)
    if not sess:
        return None
    cache = sess["cache"]
    account = sess["account"]
    if not account:
        return None

    app = _build_msal_app(cache)
    result = app.acquire_token_silent(SCOPES, account=account)
    if not result or "access_token" not in result:
        return None
    return result["access_token"]


def session_email(session_id: str) -> Optional[str]:
    sess = _SESSIONS.get(session_id)
    return sess.get("email") if sess else None
