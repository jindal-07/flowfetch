"""
Microsoft Graph app-only auth (client-credentials flow) using MSAL.

Uses the Application permissions admin-consented on the app registration
(Files.Read.All). There is no user sign-in: the server holds one token and
can read every file the registration is granted, so the deployment must be
protected (see FLOWFETCH_PASSWORD in app.py).
"""

from __future__ import annotations

import os
from typing import Optional

import msal


# ---------------------------------------------------------------------------
# Configuration (all via environment variables)
# ---------------------------------------------------------------------------

CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "").strip()
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "").strip()
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "").strip()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# App-only tokens always use `.default`: whatever Application permissions
# have been admin-consented on the registration.
SCOPES = ["https://graph.microsoft.com/.default"]


def _looks_like_placeholder(v: str) -> bool:
    return bool(v) and (v.startswith("<") or v.endswith(">") or v.lower() == "your-tenant-id")


def is_configured() -> bool:
    required = (CLIENT_ID, CLIENT_SECRET, TENANT_ID)
    # Client credentials needs a real tenant; the multi-tenant aliases can't issue app tokens.
    if not all(required) or TENANT_ID.lower() in ("common", "organizations", "consumers"):
        return False
    return not any(_looks_like_placeholder(x) for x in required)


# ---------------------------------------------------------------------------
# Token
# ---------------------------------------------------------------------------

_APP_CLIENT: Optional[msal.ConfidentialClientApplication] = None


def get_access_token() -> Optional[str]:
    """App-only Graph token, or None when not configured. MSAL caches it
    in-process and only contacts Entra ID again when it is close to expiry."""
    global _APP_CLIENT
    if not is_configured():
        return None
    try:
        if _APP_CLIENT is None:
            _APP_CLIENT = msal.ConfidentialClientApplication(
                CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET,
            )
        result = _APP_CLIENT.acquire_token_for_client(scopes=SCOPES)
    except Exception as e:  # bad tenant ID, network failure, etc.
        raise RuntimeError(f"App-only token request failed: {e}") from e
    if "access_token" not in result:
        raise RuntimeError(
            f"App-only token request failed: {result.get('error')}: {result.get('error_description')}"
        )
    return result["access_token"]
