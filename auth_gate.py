"""
Password-only login gate for the whole app.

Enabled when FLOWFETCH_PASSWORD is set. A correct password sets a signed,
HttpOnly cookie: with "Remember me" it lasts REMEMBER_DAYS, otherwise it ends
with the browser session. The signing key is derived from the password, so
changing FLOWFETCH_PASSWORD logs every device out.
"""

from __future__ import annotations

import hashlib
import hmac
import html
import os
import time
from typing import Optional
from urllib.parse import quote

from fastapi import APIRouter, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse

PASSWORD = os.environ.get("FLOWFETCH_PASSWORD", "").strip()
ENABLED = bool(PASSWORD)

COOKIE_NAME = "ff_auth"
REMEMBER_DAYS = int(os.environ.get("FLOWFETCH_REMEMBER_DAYS", "30") or 30)
SESSION_HOURS = 12  # lifetime of a non-remembered login, even if the browser stays open

# Paths reachable without logging in.
PUBLIC_PATHS = {"/login", "/healthz"}

# Brute-force brake: per client IP, at most MAX_FAILS wrong passwords per window.
MAX_FAILS = 8
FAIL_WINDOW = 15 * 60
_fails: dict[str, list[float]] = {}

_KEY = hashlib.sha256(b"flowfetch-gate:" + PASSWORD.encode()).digest()


# ---------------------------------------------------------------------------
# Token helpers
# ---------------------------------------------------------------------------

def _sign(expiry: int) -> str:
    sig = hmac.new(_KEY, str(expiry).encode(), hashlib.sha256).hexdigest()
    return f"{expiry}.{sig}"


def _valid(token: str) -> bool:
    try:
        expiry_s, _ = token.split(".", 1)
        expiry = int(expiry_s)
    except (ValueError, AttributeError):
        return False
    return expiry > time.time() and hmac.compare_digest(token, _sign(expiry))


def _cookie_from_scope(scope) -> Optional[str]:
    for name, value in scope.get("headers") or []:
        if name == b"cookie":
            for part in value.decode("latin-1").split(";"):
                k, _, v = part.strip().partition("=")
                if k == COOKIE_NAME:
                    return v
    return None


def _client_ip(request: Request) -> str:
    fwd = request.headers.get("x-forwarded-for", "")
    return fwd.split(",")[0].strip() or (request.client.host if request.client else "?")


def _is_https(request: Request) -> bool:
    return request.headers.get("x-forwarded-proto", request.url.scheme) == "https"


def _safe_next(target: str) -> str:
    # Only allow same-site relative paths, never "//host" or absolute URLs.
    return target if target.startswith("/") and not target.startswith("//") else "/"


# ---------------------------------------------------------------------------
# ASGI middleware
# ---------------------------------------------------------------------------

class PasswordGate:
    """Pure ASGI, so SSE and ZIP streaming responses pass through untouched."""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http" or scope.get("path") in PUBLIC_PATHS:
            return await self.app(scope, receive, send)
        token = _cookie_from_scope(scope)
        if token and _valid(token):
            return await self.app(scope, receive, send)

        path = scope.get("path", "/")
        accept = dict(scope.get("headers") or []).get(b"accept", b"").decode("latin-1")
        if scope.get("method") == "GET" and "text/html" in accept:
            qs = scope.get("query_string", b"").decode("latin-1")
            target = path + (f"?{qs}" if qs else "")
            headers = [(b"location", f"/login?next={quote(target)}".encode())]
            await send({"type": "http.response.start", "status": 303, "headers": headers})
            await send({"type": "http.response.body", "body": b""})
            return
        await send({"type": "http.response.start", "status": 401,
                    "headers": [(b"content-type", b"application/json")]})
        await send({"type": "http.response.body",
                    "body": b'{"detail":"Not logged in. Reload the page to log in."}'})


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

router = APIRouter()


@router.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, next: str = "/"):
    if not ENABLED:
        return RedirectResponse("/", status_code=303)
    token = request.cookies.get(COOKIE_NAME)
    if token and _valid(token):
        return RedirectResponse(_safe_next(next), status_code=303)
    return HTMLResponse(_render(next=_safe_next(next)))


@router.post("/login")
async def login_submit(request: Request, password: str = Form(""),
                       remember: Optional[str] = Form(None), next: str = Form("/")):
    if not ENABLED:
        return RedirectResponse("/", status_code=303)
    target = _safe_next(next)
    ip = _client_ip(request)
    now = time.time()
    recent = [t for t in _fails.get(ip, []) if now - t < FAIL_WINDOW]
    _fails[ip] = recent
    if len(recent) >= MAX_FAILS:
        return HTMLResponse(_render(next=target, error="Too many attempts. Try again in a few minutes."),
                            status_code=429)

    if not hmac.compare_digest(password.encode(), PASSWORD.encode()):
        recent.append(now)
        return HTMLResponse(_render(next=target, error="Incorrect password."), status_code=401)

    _fails.pop(ip, None)
    remembered = remember is not None
    lifetime = REMEMBER_DAYS * 86400 if remembered else SESSION_HOURS * 3600
    resp = RedirectResponse(target, status_code=303)
    resp.set_cookie(
        COOKIE_NAME, _sign(int(now + lifetime)),
        max_age=lifetime if remembered else None,  # None = ends with the browser session
        httponly=True, samesite="lax", secure=_is_https(request), path="/",
    )
    return resp


@router.post("/logout")
async def logout():
    resp = RedirectResponse("/login", status_code=303)
    resp.delete_cookie(COOKIE_NAME, path="/")
    return resp


@router.get("/api/gate")
async def gate_status():
    return {"enabled": ENABLED}


# ---------------------------------------------------------------------------
# Login page (matches the site's neo-brutalist theme)
# ---------------------------------------------------------------------------

def _render(next: str, error: str = "") -> str:
    err = f'<p class="err" role="alert">{html.escape(error)}</p>' if error else ""
    return f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FlowFetch — Log in</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Archivo+Black&family=Space+Grotesk:wght@500;600;700&display=swap" rel="stylesheet">
<style>
  :root {{ --bg:#FAF3E7; --surface:#FFFDF7; --ink:#16130F; --muted:#6B6459;
          --orange:#FF9950; --yellow:#FFD23F; --red:#FF6B6B; }}
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:'Space Grotesk',system-ui,sans-serif; background:var(--bg); color:var(--ink);
         min-height:100vh; display:grid; place-items:center; padding:24px; }}
  .card {{ width:100%; max-width:400px; background:var(--surface); border:3px solid var(--ink);
          border-radius:14px; box-shadow:8px 8px 0 var(--ink); padding:32px 28px;
          animation:up .35s cubic-bezier(.2,0,0,1) both; }}
  @keyframes up {{ from {{ opacity:0; transform:translateY(16px); }} to {{ opacity:1; transform:none; }} }}
  .brand {{ display:flex; align-items:center; gap:12px; margin-bottom:26px; }}
  .logo {{ width:46px; height:46px; border-radius:11px; background:var(--orange); border:3px solid var(--ink);
          box-shadow:3px 3px 0 var(--ink); display:grid; place-items:center; }}
  .wordmark {{ font-family:'Archivo Black',sans-serif; font-size:1.3rem; }}
  .wordmark span {{ color:var(--orange); }}
  h1 {{ font-family:'Archivo Black',sans-serif; font-weight:400; font-size:1.5rem; margin-bottom:6px; }}
  .sub {{ color:var(--muted); font-size:.92rem; margin-bottom:22px; }}
  label.f {{ display:block; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.09em; margin-bottom:9px; }}
  input[type=password] {{ width:100%; padding:14px 16px; font:inherit; font-size:1rem; background:var(--surface);
          border:2.5px solid var(--ink); border-radius:10px; box-shadow:3px 3px 0 var(--ink); outline:none;
          transition:transform .15s, box-shadow .15s; }}
  input[type=password]:focus {{ transform:translate(-2px,-2px); box-shadow:5px 5px 0 var(--ink); background:#FFFBEA; }}
  .remember {{ display:flex; align-items:center; gap:10px; margin:18px 0 22px; font-weight:600; font-size:.92rem; cursor:pointer; user-select:none; }}
  .remember input {{ appearance:none; width:22px; height:22px; border:2.5px solid var(--ink); border-radius:6px;
          background:var(--surface); display:grid; place-items:center; cursor:pointer; flex-shrink:0; }}
  .remember input:checked {{ background:var(--yellow); }}
  .remember input:checked::after {{ content:""; width:9px; height:5px; border-left:3px solid var(--ink);
          border-bottom:3px solid var(--ink); transform:translateY(-2px) rotate(-45deg); }}
  .remember small {{ color:var(--muted); font-weight:500; }}
  button {{ width:100%; padding:15px; font:inherit; font-weight:700; font-size:.95rem; text-transform:uppercase;
          letter-spacing:.04em; background:var(--orange); color:var(--ink); border:3px solid var(--ink);
          border-radius:10px; box-shadow:5px 5px 0 var(--ink); cursor:pointer; transition:transform .15s, box-shadow .15s; }}
  button:hover {{ transform:translate(-2px,-2px); box-shadow:8px 8px 0 var(--ink); }}
  button:active {{ transform:translate(5px,5px); box-shadow:none; }}
  .err {{ background:var(--red); border:2.5px solid var(--ink); border-radius:10px; padding:10px 14px;
          font-weight:600; font-size:.9rem; margin-bottom:18px; box-shadow:3px 3px 0 var(--ink); }}
  :focus-visible {{ outline:3px solid var(--ink); outline-offset:2px; }}
</style></head>
<body>
  <main class="card">
    <div class="brand">
      <div class="logo"><svg viewBox="0 0 28 28" width="24" height="24" fill="none" aria-hidden="true">
        <path d="M4.5 7.5C7.5 4.8 10.5 4.8 14 7.5c3.5 2.7 6.5 2.7 9.5 0" stroke="#16130F" stroke-width="2.4" stroke-linecap="round"/>
        <path d="M7.5 13.2c2.2-2 4.3-2 6.5 0 2.2 2 4.3 2 6.5 0" stroke="#16130F" stroke-width="2.4" stroke-linecap="round"/>
        <path d="M10 18.5 14 23l4-4.5" stroke="#16130F" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/></svg></div>
      <div class="wordmark">Flow<span>Fetch</span></div>
    </div>
    <h1>Welcome back</h1>
    <p class="sub">Enter the team password to continue.</p>
    {err}
    <form method="post" action="/login">
      <input type="hidden" name="next" value="{html.escape(next)}">
      <label class="f" for="pw">Password</label>
      <input type="password" id="pw" name="password" autocomplete="current-password" required autofocus>
      <label class="remember"><input type="checkbox" name="remember" checked>
        <span>Remember me <small>· {REMEMBER_DAYS} days on this device</small></span></label>
      <button type="submit">Log in</button>
    </form>
  </main>
</body></html>"""
