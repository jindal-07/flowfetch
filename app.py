import asyncio
import csv as csv_mod
import io
import json
import math
import os
import sys
import time
import uuid
import warnings
from datetime import datetime
from pathlib import Path
from typing import Optional

# Playwright needs the Proactor loop on Windows to spawn the browser subprocess.
# Without this, uvicorn's default Selector loop raises NotImplementedError.
if sys.platform == "win32":
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", DeprecationWarning)
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# Load .env early so msgraph_auth (which reads env vars at import time) sees them.
try:
    from dotenv import load_dotenv
    load_dotenv(dotenv_path=Path(__file__).parent / ".env", override=False)
except ImportError:
    pass

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from starlette.middleware.sessions import SessionMiddleware
from stream_zip import ZIP_64, async_stream_zip

import fb_video_downloader as downloader
import sharepoint_downloader as sp_downloader
import sharepoint_graph as sp_graph
import msgraph_auth
import google_drive as gd

app = FastAPI(title="FlowFetch")

# Signed-cookie sessions.
import secrets as _secrets
SESSION_SECRET = os.environ.get("SESSION_SECRET") or _secrets.token_urlsafe(32)

# On Hugging Face Spaces the app is embedded inside an <iframe>, so the cookie
# must be SameSite=None + Secure to be sent. Locally we keep SameSite=Lax.
_ON_HF = bool(os.environ.get("SPACE_ID") or os.environ.get("SPACE_HOST"))
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    https_only=_ON_HF,
    same_site="none" if _ON_HF else "lax",
)

class BasicAuthGate:
    """HTTP Basic auth over the whole app when FLOWFETCH_PASSWORD is set.
    Pure ASGI so SSE and ZIP streaming responses pass through untouched."""

    def __init__(self, app, password: str, username: str):
        import base64
        self.app = app
        self.expected = "Basic " + base64.b64encode(f"{username}:{password}".encode()).decode()

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http" or scope.get("path") == "/healthz":
            return await self.app(scope, receive, send)
        auth = dict(scope.get("headers") or []).get(b"authorization", b"").decode("latin-1")
        if _secrets.compare_digest(auth, self.expected):
            return await self.app(scope, receive, send)
        await send({
            "type": "http.response.start",
            "status": 401,
            "headers": [(b"www-authenticate", b'Basic realm="FlowFetch"'),
                        (b"content-type", b"text/plain")],
        })
        await send({"type": "http.response.body", "body": b"Authentication required."})


_GATE_PASSWORD = os.environ.get("FLOWFETCH_PASSWORD", "").strip()
if _GATE_PASSWORD:
    app.add_middleware(BasicAuthGate, password=_GATE_PASSWORD,
                       username=os.environ.get("FLOWFETCH_USERNAME", "flowfetch").strip() or "flowfetch")
elif msgraph_auth.APP_ONLY:
    print("[auth] WARNING: AZURE_AUTH_MODE=app without FLOWFETCH_PASSWORD - anyone with the "
          "URL can read every SharePoint file the app registration can access.")

BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"
AUTH_DIR = BASE_DIR / "auth"
SP_STATE_FILE = AUTH_DIR / "sharepoint_state.json"

STATIC_DIR.mkdir(exist_ok=True)
AUTH_DIR.mkdir(exist_ok=True)


# ---------------------------------------------------------------------------
# Batching math
# ---------------------------------------------------------------------------

MAX_PER_BATCH = 100


def equal_split_sizes(total: int, max_per_batch: int = MAX_PER_BATCH) -> list[int]:
    """
    Split `total` items into ceil(total/max) equal-sized batches.
    e.g. total=250, max=100 -> [84, 83, 83]
    e.g. total=80,  max=100 -> [80]
    """
    if total <= 0:
        return []
    if total <= max_per_batch:
        return [total]
    n = math.ceil(total / max_per_batch)
    base, rem = divmod(total, n)
    return [base + 1 if i < rem else base for i in range(n)]


def make_batches(sizes: list[int]) -> list[dict]:
    out: list[dict] = []
    cursor = 0
    for i, sz in enumerate(sizes, start=1):
        out.append({
            "n": i,
            "count": sz,
            "start": cursor,
            "end": cursor + sz,
            "status": "queued",  # queued | streaming | complete | cancelled | error
        })
        cursor += sz
    return out


# ---------------------------------------------------------------------------
# Job store (in-memory, TTL-evicted)
# ---------------------------------------------------------------------------

class JobStore:
    def __init__(self, ttl: int = 3600):
        self.ttl = ttl
        self._jobs: dict[str, dict] = {}

    def __contains__(self, jid: str) -> bool:
        return jid in self._jobs

    def __getitem__(self, jid: str) -> dict:
        return self._jobs[jid]

    def get(self, jid: str) -> Optional[dict]:
        return self._jobs.get(jid)

    def create(self, kind: str, **fields) -> str:
        self._evict_stale()
        jid = uuid.uuid4().hex
        self._jobs[jid] = {
            "id": jid,
            "kind": kind,
            "created": time.time(),
            "status": "running",
            "events": [],
            "all_results": [],
            "summary": {"success": 0, "failed": 0, "skipped": 0, "cancelled": 0},
            "cancelled": False,
            "task": None,
            **fields,
        }
        return jid

    def _evict_stale(self):
        now = time.time()
        stale = [jid for jid, j in self._jobs.items()
                 if now - j.get("created", now) > self.ttl]
        for jid in stale:
            self._jobs.pop(jid, None)


jobs = JobStore()


def _append_event(job_id: str, index: int, total: int, message: str,
                  status: str, filename: Optional[str] = None,
                  extra: Optional[dict] = None):
    job = jobs.get(job_id)
    if not job:
        return
    event = {
        "index": index,
        "total": total,
        "message": message,
        "status": status,
        "filename": filename,
    }
    if extra:
        event.update(extra)
    job["events"].append(event)


def _maybe_finalize_job(job_id: str):
    job = jobs.get(job_id)
    if not job or job.get("kind") != "csv":
        return
    if job["status"] in ("completed", "cancelled", "error"):
        return  # already finalized — don't double-count or re-emit
    all_done = job.get("next_batch", 1) > len(job.get("batches", []))
    # A cancelled job finalizes as soon as the in-flight batch finishes, so the
    # remaining batches never run and the SSE stream can terminate immediately.
    if not (all_done or job.get("cancelled")):
        return
    summary = {"success": 0, "failed": 0, "skipped": 0, "cancelled": 0}
    for r in job["all_results"]:
        st = r.get("status")
        if st in summary:
            summary[st] += 1
    job["summary"] = summary
    job["status"] = "cancelled" if job["cancelled"] else "completed"
    total = job.get("total", 0)
    # On a stop, reflect what was actually processed rather than snapping to 100%.
    processed = len(job["all_results"]) if job["cancelled"] else total
    job["events"].append({
        "index": processed,
        "total": total,
        "message": "All batches saved." if not job["cancelled"] else "Stopped.",
        "status": "done",
        "filename": None,
    })


# ---------------------------------------------------------------------------
# Index page
# ---------------------------------------------------------------------------

@app.get("/healthz")
async def healthz():
    return {"ok": True}


@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = STATIC_DIR / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


# ---------------------------------------------------------------------------
# Facebook single (stream directly to browser, no disk)
# ---------------------------------------------------------------------------

# Resolving a Facebook video means launching a headless browser (~10-15s). The
# UI checks accessibility first (via /api/check) and then downloads, so we cache
# the freshly-resolved CDN URL briefly to avoid resolving the same link twice.
_FB_SRC_CACHE: dict[str, tuple[str, float]] = {}
_FB_SRC_TTL = 120


def _fb_cache_put(url: str, src: str):
    now = time.time()
    _FB_SRC_CACHE[url] = (src, now)
    for k in [k for k, (_, t) in _FB_SRC_CACHE.items() if now - t > _FB_SRC_TTL]:
        _FB_SRC_CACHE.pop(k, None)


def _fb_cache_get(url: str) -> Optional[str]:
    hit = _FB_SRC_CACHE.get(url)
    if not hit:
        return None
    src, t = hit
    if time.time() - t > _FB_SRC_TTL:
        _FB_SRC_CACHE.pop(url, None)
        return None
    return src


@app.get("/api/fb/single.bin")
async def fb_single_bin(url: str):
    url = url.strip()
    if not url.startswith("http"):
        raise HTTPException(400, "Invalid URL.")

    filename = downloader.filename_for(url, 1)
    # Reuse the src from the accessibility pre-check when it's still fresh;
    # otherwise resolve now (and confirm the link is reachable) before streaming.
    src = _fb_cache_get(url)
    if src is None:
        try:
            src, reason = await downloader.resolve_video_with_reason(url)
        except Exception as e:
            raise HTTPException(500, f"Failed to resolve video: {e}")
        if not src:
            raise HTTPException(404, reason)

    return StreamingResponse(
        downloader.iter_remote_bytes(src),
        media_type="video/mp4",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Cache-Control": "no-store",
        },
    )


# ---------------------------------------------------------------------------
# Accessibility pre-check (used by the UI before a single download)
# ---------------------------------------------------------------------------

@app.get("/api/check")
async def check_link(request: Request, source: str, url: str):
    """
    Verify a single link is reachable/downloadable BEFORE the browser starts a
    download. Always returns 200 with {"ok": bool, "reason": str, "filename": str}
    so the UI can show a clear message instead of a broken/empty download.
    """
    url = url.strip()
    if not url.startswith("http"):
        return {"ok": False, "reason": "Enter a valid URL (it must start with http)."}
    try:
        if source == "facebook":
            src, reason = await downloader.resolve_video_with_reason(url)
            if not src:
                return {"ok": False, "reason": reason}
            _fb_cache_put(url, src)
            return {"ok": True, "reason": "", "filename": downloader.filename_for(url, 1)}
        if source == "sharepoint":
            method = _sp_auth_check(request)
            if method == "graph":
                name, _ = await sp_graph.resolve_single_stream(url, _sp_token_getter(request))
            else:
                name, _ = await sp_downloader.resolve_single_stream(url, SP_STATE_FILE)
            return {"ok": True, "reason": "", "filename": name}
        if source == "gdrive":
            name, _ = await gd.resolve_single_stream(url)
            return {"ok": True, "reason": "", "filename": name}
        return {"ok": False, "reason": "Unknown source."}
    except HTTPException as e:
        return {"ok": False, "reason": e.detail}
    except PermissionError as e:
        return {"ok": False, "reason": str(e)}
    except FileNotFoundError as e:
        return {"ok": False, "reason": str(e)}
    except Exception as e:
        return {"ok": False, "reason": f"This link isn't accessible: {e}"}


# ---------------------------------------------------------------------------
# SharePoint single (stream directly to browser, no disk)
# ---------------------------------------------------------------------------

@app.get("/api/sharepoint/single.bin")
async def sp_single_bin(request: Request, url: str):
    url = url.strip()
    if not url.startswith("http"):
        raise HTTPException(400, "Invalid URL.")

    method = _sp_auth_check(request)
    # Resolve (and thereby confirm the share is accessible) before streaming.
    try:
        if method == "graph":
            filename, byte_iter = await sp_graph.resolve_single_stream(
                url, _sp_token_getter(request)
            )
        else:
            filename, byte_iter = await sp_downloader.resolve_single_stream(
                url, SP_STATE_FILE
            )
    except PermissionError as e:
        raise HTTPException(403, str(e))
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    except Exception as e:
        raise HTTPException(500, f"Failed to resolve file: {e}")

    return StreamingResponse(
        byte_iter,
        media_type="application/octet-stream",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Cache-Control": "no-store",
        },
    )


# ---------------------------------------------------------------------------
# CSV: plan + batched ZIP streaming
# ---------------------------------------------------------------------------

@app.post("/api/csv/plan")
async def csv_plan(
    request: Request,
    source: str = Form(...),
    file: UploadFile = File(...),
):
    """
    Parse a CSV, compute equal-split batches, and register a job. The client
    then triggers /api/jobs/{id}/batch/{n}.zip for each batch in order.
    """
    if source not in ("facebook", "sharepoint", "gdrive"):
        raise HTTPException(400, "source must be 'facebook', 'sharepoint', or 'gdrive'.")
    if not file.filename or not file.filename.lower().endswith(".csv"):
        raise HTTPException(400, "Please upload a .csv file.")

    content = await file.read()
    sp_method = None
    sid = None

    if source == "facebook":
        try:
            links = downloader.parse_links_from_csv_bytes(content)
        except Exception as e:
            raise HTTPException(400, f"Could not parse CSV: {e}")
        if not links:
            raise HTTPException(400, "No valid Facebook links found in the CSV.")
    elif source == "sharepoint":
        sp_method = _sp_auth_check(request)
        try:
            links = sp_graph.parse_links_from_csv_bytes(content)
        except Exception as e:
            raise HTTPException(400, f"Could not parse CSV: {e}")
        if not links:
            raise HTTPException(400, "No SharePoint links found in the CSV.")
        sid = _session_id(request) if sp_method == "graph" else None
    else:  # gdrive
        try:
            links = gd.parse_links_from_csv_bytes(content)
        except Exception as e:
            raise HTTPException(400, f"Could not parse CSV: {e}")
        if not links:
            raise HTTPException(400, "No Google Drive links found in the CSV.")

    sizes = equal_split_sizes(len(links), MAX_PER_BATCH)
    batches = make_batches(sizes)

    job_id = jobs.create(
        "csv",
        source=source,
        sp_method=sp_method,
        sp_sid=sid,
        links=links,
        total=len(links),
        batches=batches,
        next_batch=1,
    )
    return {
        "job_id": job_id,
        "total": len(links),
        "batches": [{"n": b["n"], "count": b["count"]} for b in batches],
    }


def _build_batch_generator(job: dict, batch: dict):
    """Pick the right iter_batch_files variant based on job source/method."""
    on_progress = (lambda jid: (lambda i, t, m, s, f=None:
                                _append_event(jid, i, t, m, s, f)))(job["id"])
    cancel = (lambda jid: (lambda: jobs[jid]["cancelled"]))(job["id"])

    source = job["source"]
    total = job["total"]
    start, end = batch["start"], batch["end"]

    if source == "facebook":
        return downloader.iter_batch_files(
            links=job["links"][start:end],
            on_progress=on_progress,
            should_cancel=cancel,
            batch_offset=start,
            overall_total=total,
        )
    if source == "sharepoint":
        links_slice = job["links"][start:end]
        if job["sp_method"] == "graph":
            return sp_graph.iter_batch_files(
                links_slice, _sp_token_getter_for_sid(job.get("sp_sid")),
                on_progress, cancel,
                batch_offset=start, overall_total=total,
            )
        return sp_downloader.iter_batch_files(
            links_slice, SP_STATE_FILE,
            on_progress, cancel,
            batch_offset=start, overall_total=total,
        )
    if source == "gdrive":
        links_slice = job["links"][start:end]
        return gd.iter_batch_files(
            links_slice, on_progress, cancel,
            batch_offset=start, overall_total=total,
        )
    if source == "sp_folder_csv":
        items_slice = job["manifest"][start:end]
        if job["sp_method"] == "graph":
            return sp_graph.iter_batch_files_from_manifest(
                items_slice, on_progress, cancel,
                batch_offset=start, overall_total=total,
            )
        return sp_downloader.iter_batch_files_from_manifest(
            items_slice, SP_STATE_FILE,
            on_progress, cancel,
            batch_offset=start, overall_total=total,
        )
    raise HTTPException(400, f"Unknown job source: {source}")


@app.get("/api/jobs/{job_id}/batch/{n}.zip")
async def stream_batch(job_id: str, n: int):
    """
    Stream batch N as a ZIP directly to the client. The server never persists
    files; fetched bytes flow straight into the zip stream and into the HTTP
    response body. On completion the server advances `next_batch` and emits a
    `batch_complete` SSE event so the UI knows to trigger batch N+1.
    """
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found.")
    if job.get("kind") != "csv":
        raise HTTPException(400, "Job is not a CSV/batch job.")
    if n < 1 or n > len(job.get("batches", [])):
        raise HTTPException(404, "Batch not found.")
    if n != job.get("next_batch"):
        raise HTTPException(
            409,
            f"Batch {n} is not next. Expected batch {job.get('next_batch')}.",
        )

    batch = job["batches"][n - 1]
    batch["status"] = "streaming"
    gen = _build_batch_generator(job, batch)

    modified_at = datetime.now()
    perms = 0o600

    async def member_iter():
        async for name, byte_iter, result in gen:
            if name is None or byte_iter is None:
                # Resolution failed before streaming started.
                job["all_results"].append(result)
                continue

            aiter = byte_iter.__aiter__()
            # Peek the first chunk so pre-transfer failures (expired URL, HTTP
            # error, refused connection — the common case) surface BEFORE this
            # member is committed to the zip. Those become clean 'failed'
            # results with no zip entry, and the batch keeps going.
            try:
                first_chunk = await aiter.__anext__()
                empty = False
            except StopAsyncIteration:
                first_chunk, empty = None, True
            except Exception:
                # The source generator already marked result 'failed' and
                # reported it; just record and move to the next file.
                if result.get("status") not in ("failed", "skipped"):
                    result["status"] = "failed"
                    result["file"] = None
                job["all_results"].append(result)
                continue

            async def _safe_member(head=first_chunk, is_empty=empty,
                                   rest=aiter, res=result):
                # A mid-transfer drop must NOT abort the whole batch zip:
                # swallow it and end this member early so stream-zip stays
                # valid and moves on. The file is still marked 'failed' (and
                # exported in failures.csv for retry).
                try:
                    if not is_empty and head is not None:
                        yield head
                    async for c in rest:
                        yield c
                except Exception:
                    if res.get("status") not in ("failed", "skipped"):
                        res["status"] = "failed"
                    return

            yield name, modified_at, perms, ZIP_64, _safe_member()
            # Control resumes after stream-zip finishes writing this member.
            job["all_results"].append(result)

    async def response_body():
        try:
            async for chunk in async_stream_zip(member_iter()):
                yield chunk
            batch["status"] = "complete" if not job["cancelled"] else "cancelled"
            job["next_batch"] = n + 1
            _append_event(
                job_id, batch["end"], job["total"],
                f"Batch {n} of {len(job['batches'])} saved.",
                "done",
                extra={"event": "batch_complete", "n": n},
            )
            _maybe_finalize_job(job_id)
        except asyncio.CancelledError:
            batch["status"] = "aborted"
            raise
        except Exception as e:
            batch["status"] = "error"
            _append_event(
                job_id, 0, job["total"],
                f"Batch {n} error: {e}", "error",
            )
            raise

    filename = f"flowfetch-batch-{n:02d}.zip"
    return StreamingResponse(
        response_body(),
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Cache-Control": "no-store",
            "X-Accel-Buffering": "no",
        },
    )


# ---------------------------------------------------------------------------
# SharePoint folder: crawl -> manifest CSV (no media downloaded server-side)
# ---------------------------------------------------------------------------

def _manifest_filename(folder_name: str, job_id: str) -> str:
    """Slugify the folder name into a filesystem-safe CSV name."""
    import re
    name = (folder_name or "").strip()
    # Strip path-illegal characters and collapse whitespace.
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", name).strip()
    name = re.sub(r"\s+", " ", name)
    if not name:
        name = f"manifest_{job_id[:8]}"
    return f"{name}.csv"


async def _crawl_task(job_id: str, url: str, method: str):
    job = jobs.get(job_id)
    if not job:
        return

    def on_progress(i, t, m, s, f=None):
        _append_event(job_id, i, t, m, s, f)

    cancel = lambda: job["cancelled"]
    try:
        if method == "graph":
            token_getter = _sp_token_getter_for_sid(job.get("sp_sid"))
            result = await sp_graph.crawl_folder(
                url, token_getter, on_progress, True, cancel
            )
        else:
            result = await sp_downloader.crawl_folder(
                url, SP_STATE_FILE, on_progress, True, cancel
            )

        items = result["items"]
        folder_name = result.get("folder_name") or ""

        if job["cancelled"]:
            job["status"] = "cancelled"
            _append_event(job_id, 0, len(items), "Cancelled.", "done")
            return

        manifest_bytes = sp_graph.items_to_csv_bytes(items)
        fname = _manifest_filename(folder_name, job_id)
        job["manifest_csv"] = manifest_bytes
        job["manifest_count"] = len(items)
        job["manifest_filename"] = fname
        job["folder_name"] = folder_name
        job["status"] = "manifest_ready"

        manifest_url = f"/api/jobs/{job_id}/manifest.csv"
        nice_name = folder_name or "folder"
        _append_event(
            job_id, len(items), len(items),
            f"Manifest ready for '{nice_name}': {len(items)} file(s). Saving as {fname}...",
            "done",
            filename=fname,
            extra={
                "event": "manifest_ready",
                "url": manifest_url,
                "filename": fname,
                "count": len(items),
            },
        )
    except Exception as e:
        job["status"] = "error"
        _append_event(job_id, 0, 0, str(e), "error")


@app.post("/api/sharepoint/folder/start")
async def sp_folder_start(request: Request, url: str = Form(...)):
    url = url.strip()
    if not url.startswith("http"):
        raise HTTPException(400, "Invalid URL.")
    if not sp_graph.is_folder_share(url):
        raise HTTPException(400, "Link doesn't look like a folder share (missing ':f:').")

    method = _sp_auth_check(request)
    sid = _session_id(request) if method == "graph" else None

    job_id = jobs.create(
        "sp_folder",
        source="sp_folder",
        sp_method=method,
        sp_sid=sid,
        total=0,
        next_batch=1,
        batches=[],
        manifest_csv=None,
        manifest_count=0,
    )
    job = jobs[job_id]
    job["task"] = asyncio.create_task(_crawl_task(job_id, url, method))
    return {"job_id": job_id}


@app.get("/api/jobs/{job_id}/manifest.csv")
async def get_manifest_csv(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found.")
    csv_bytes = job.get("manifest_csv")
    if not csv_bytes:
        raise HTTPException(404, "Manifest not yet ready.")
    fname = job.get("manifest_filename") or _manifest_filename(
        job.get("folder_name") or "", job_id
    )
    # RFC 5987 form for non-ASCII folder names (Hindi/regional Indian
    # languages, accents, etc.). HTTP headers are latin-1, so the plain
    # `filename=` fallback must be ASCII-only; browsers that understand
    # `filename*=` use the real UTF-8 name.
    from urllib.parse import quote
    encoded = quote(fname)
    ascii_fallback = fname.encode("ascii", "ignore").decode().strip() or ""
    if ascii_fallback in ("", ".csv"):
        ascii_fallback = f"manifest_{job_id[:8]}.csv"
    return StreamingResponse(
        iter([csv_bytes]),
        media_type="text/csv; charset=utf-8",
        headers={
            "Content-Disposition": f"attachment; filename=\"{ascii_fallback}\"; filename*=UTF-8''{encoded}",
            "Cache-Control": "no-store",
        },
    )


# ---------------------------------------------------------------------------
# Job control: SSE stream, cancel, status, failures CSV
# ---------------------------------------------------------------------------

@app.get("/api/jobs/{job_id}/stream")
async def stream_job(job_id: str):
    if job_id not in jobs:
        raise HTTPException(404, "Job not found.")

    async def event_generator():
        sent = 0
        while True:
            job = jobs.get(job_id)
            if not job:
                break
            while sent < len(job["events"]):
                event = job["events"][sent]
                yield f"data: {json.dumps(event)}\n\n"
                sent += 1
            if job["status"] in ("completed", "error", "cancelled", "manifest_ready"):
                payload = {
                    "status": job["status"],
                    "done": True,
                    "summary": job.get("summary"),
                }
                yield f"data: {json.dumps(payload)}\n\n"
                break
            await asyncio.sleep(0.5)

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found.")
    if job["status"] != "running":
        return {"ok": True, "status": job["status"]}
    job["cancelled"] = True
    return {"ok": True, "status": "cancelling"}


@app.get("/api/jobs/{job_id}")
async def get_job(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found.")
    return {k: v for k, v in job.items() if k not in ("task", "manifest_csv")}


@app.get("/api/jobs/{job_id}/failures.csv")
async def get_job_failures_csv(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found.")
    failures = [r for r in job.get("all_results", []) if r.get("status") == "failed"]

    output = io.StringIO()
    writer = csv_mod.writer(output)
    writer.writerow(["Links"])
    for r in failures:
        if r.get("url"):
            writer.writerow([r["url"]])

    return StreamingResponse(
        iter([output.getvalue().encode("utf-8-sig")]),
        media_type="text/csv; charset=utf-8",
        headers={
            "Content-Disposition": f'attachment; filename="failures_{job_id[:8]}.csv"'
        },
    )


# ---------------------------------------------------------------------------
# Google Drive: public-link downloads (no auth)
# ---------------------------------------------------------------------------

@app.get("/api/gdrive/single.bin")
async def gd_single_bin(url: str):
    url = url.strip()
    if not url.startswith("http"):
        raise HTTPException(400, "Invalid URL.")
    # Confirm the file exists and is publicly shared before streaming.
    try:
        filename, byte_iter = await gd.resolve_single_stream(url)
    except PermissionError as e:
        raise HTTPException(403, str(e))
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    except Exception as e:
        raise HTTPException(500, f"Failed to resolve file: {e}")
    return StreamingResponse(
        byte_iter,
        media_type="application/octet-stream",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Cache-Control": "no-store",
        },
    )


async def _gd_crawl_task(job_id: str, url: str):
    job = jobs.get(job_id)
    if not job:
        return

    def on_progress(i, t, m, s, f=None):
        _append_event(job_id, i, t, m, s, f)

    cancel = lambda: job["cancelled"]
    try:
        result = await gd.crawl_folder(url, on_progress, True, cancel)

        items = result["items"]
        folder_name = result.get("folder_name") or ""

        if job["cancelled"]:
            job["status"] = "cancelled"
            _append_event(job_id, 0, len(items), "Cancelled.", "done")
            return

        manifest_bytes = gd.items_to_csv_bytes(items)
        fname = _manifest_filename(folder_name, job_id)
        job["manifest_csv"] = manifest_bytes
        job["manifest_count"] = len(items)
        job["manifest_filename"] = fname
        job["folder_name"] = folder_name
        job["status"] = "manifest_ready"

        manifest_url = f"/api/jobs/{job_id}/manifest.csv"
        nice_name = folder_name or "folder"
        _append_event(
            job_id, len(items), len(items),
            f"Manifest ready for '{nice_name}': {len(items)} file(s). Saving as {fname}...",
            "done",
            filename=fname,
            extra={
                "event": "manifest_ready",
                "url": manifest_url,
                "filename": fname,
                "count": len(items),
            },
        )
    except Exception as e:
        job["status"] = "error"
        _append_event(job_id, 0, 0, str(e), "error")


@app.post("/api/gdrive/folder/start")
async def gd_folder_start(url: str = Form(...)):
    url = url.strip()
    if not url.startswith("http"):
        raise HTTPException(400, "Invalid URL.")
    if not gd.is_folder_url(url):
        raise HTTPException(400, "Link doesn't look like a Google Drive folder share.")

    job_id = jobs.create(
        "gd_folder",
        source="gd_folder",
        total=0,
        next_batch=1,
        batches=[],
        manifest_csv=None,
        manifest_count=0,
    )
    job = jobs[job_id]
    job["task"] = asyncio.create_task(_gd_crawl_task(job_id, url))
    return {"job_id": job_id}


# ---------------------------------------------------------------------------
# Microsoft Graph OAuth (delegated, auth-code flow)
# ---------------------------------------------------------------------------

def _session_id(request: Request) -> Optional[str]:
    return request.session.get("sid")


def _ensure_session_id(request: Request) -> str:
    sid = request.session.get("sid")
    if not sid:
        sid = msgraph_auth.new_session_id()
        request.session["sid"] = sid
    return sid


def _graph_token(sid: Optional[str]) -> Optional[str]:
    """Delegated or app-only token; app-only failures surface as 'not signed in'."""
    try:
        return msgraph_auth.get_access_token(sid)
    except RuntimeError as e:
        print(f"[auth] {e}")
        return None


@app.get("/auth/status")
async def auth_status(request: Request):
    sid = _session_id(request)
    token = await asyncio.to_thread(_graph_token, sid)
    return {
        "configured": msgraph_auth.is_configured(),
        "mode": msgraph_auth.AUTH_MODE,
        "signed_in": bool(token),
        "email": msgraph_auth.session_email(sid),
    }


@app.get("/auth/debug")
async def auth_debug():
    """Reveal which env vars are missing or look like placeholders."""
    def status(v: str) -> str:
        if not v:
            return "MISSING"
        if v.startswith("<") or v.endswith(">"):
            return "PLACEHOLDER"
        return "OK"
    info = {
        "AZURE_AUTH_MODE": msgraph_auth.AUTH_MODE,
        "AZURE_CLIENT_ID": status(msgraph_auth.CLIENT_ID),
        "AZURE_CLIENT_SECRET": status(msgraph_auth.CLIENT_SECRET),
        "AZURE_TENANT_ID": status(msgraph_auth.TENANT_ID if msgraph_auth.TENANT_ID != "common" else ""),
        "is_configured": msgraph_auth.is_configured(),
    }
    if msgraph_auth.APP_ONLY:
        try:
            info["app_token"] = "OK" if await asyncio.to_thread(msgraph_auth.get_app_token) else "NOT CONFIGURED"
        except RuntimeError as e:
            info["app_token"] = f"ERROR: {e}"
    else:
        info["AZURE_REDIRECT_URI"] = status(msgraph_auth.REDIRECT_URI)
        info["redirect_uri_value"] = msgraph_auth.REDIRECT_URI
    return info


@app.get("/auth/login")
async def auth_login(request: Request):
    if msgraph_auth.APP_ONLY:
        return RedirectResponse("/")
    if not msgraph_auth.is_configured():
        raise HTTPException(500, "Microsoft Graph auth is not configured on this server.")
    state = msgraph_auth.new_session_id()
    request.session["oauth_state"] = state
    _ensure_session_id(request)
    return RedirectResponse(msgraph_auth.build_auth_url(state))


@app.get("/auth/callback")
async def auth_callback(request: Request, code: Optional[str] = None,
                        state: Optional[str] = None, error: Optional[str] = None):
    if error:
        raise HTTPException(400, f"OAuth error: {error}")
    if not code or not state:
        raise HTTPException(400, "Missing code or state.")
    expected_state = request.session.get("oauth_state")
    if state != expected_state:
        raise HTTPException(400, "State mismatch (possible CSRF).")

    sid = _ensure_session_id(request)
    try:
        msgraph_auth.acquire_token_from_code(code, sid)
    except Exception as e:
        raise HTTPException(400, f"Token exchange failed: {e}")

    request.session.pop("oauth_state", None)
    return HTMLResponse("""
<!DOCTYPE html>
<html><head><title>Signed in</title>
<style>body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5ef;color:#1f1f14;
display:grid;place-items:center;height:100vh;margin:0;text-align:center}
.box{background:#ffffff;border:1px solid #d8d8c0;border-radius:18px;padding:32px 40px;
box-shadow:0 4px 24px rgba(31,31,20,0.08)}
.dot{width:38px;height:38px;border-radius:50%;background:#b47818;margin:0 auto 14px;
display:grid;place-items:center;color:#fcf5e9;font-weight:700}
h1{font-size:1.1rem;margin:0 0 6px;color:#1f1f14}p{margin:0;color:#5e5e3b;font-size:.9rem}</style></head>
<body><div class="box"><div class="dot">&check;</div><h1>Signed in successfully</h1><p>You can close this window.</p></div>
<script>
  try {
    if (window.opener && !window.opener.closed) {
      window.opener.postMessage({ type: "flowfetch:auth-success" }, "*");
    }
  } catch (e) {}
  setTimeout(() => { try { window.close(); } catch (e) {} }, 800);
  setTimeout(() => { if (!window.closed) location.replace("/"); }, 2000);
</script></body></html>
""")


@app.post("/auth/logout")
async def auth_logout(request: Request):
    sid = _session_id(request)
    msgraph_auth.clear_session(sid)
    request.session.clear()
    return {"ok": True}


# ---------------------------------------------------------------------------
# SharePoint auth helpers
# ---------------------------------------------------------------------------

def _sp_token_getter_for_sid(sid: Optional[str]):
    def _get():
        return msgraph_auth.get_access_token(sid)
    return _get


def _sp_token_getter(request: Request):
    return _sp_token_getter_for_sid(_session_id(request))


def _sp_auth_check(request: Request) -> str:
    """Return 'graph' or 'cookie' depending on which auth is available."""
    sid = _session_id(request)
    if _graph_token(sid):
        return "graph"
    if sp_downloader.auth_state_exists(SP_STATE_FILE):
        return "cookie"
    raise HTTPException(400, "Not signed in. Use 'Sign in with Microsoft' (or upload a state.json) first.")


def _is_local_host(request: Request) -> bool:
    host = (request.url.hostname or "").lower()
    return host in ("localhost", "127.0.0.1", "0.0.0.0", "::1")


@app.get("/api/sharepoint/auth-status")
async def sp_auth_status(request: Request):
    sid = _session_id(request)
    graph_token = await asyncio.to_thread(_graph_token, sid)
    cookie_ok = sp_downloader.auth_state_exists(SP_STATE_FILE)
    if graph_token:
        method = "app" if msgraph_auth.APP_ONLY else "graph"
    else:
        method = "cookie" if cookie_ok else None
    return {
        "logged_in": bool(graph_token or cookie_ok),
        "method": method,
        "email": msgraph_auth.session_email(sid),
        "graph_configured": msgraph_auth.is_configured(),
        "is_local": _is_local_host(request),
    }


@app.post("/api/sharepoint/auth-upload")
async def sp_auth_upload(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".json"):
        raise HTTPException(400, "Please upload a .json file.")
    content = await file.read()
    try:
        json.loads(content)
    except Exception as e:
        raise HTTPException(400, f"Not a valid JSON file: {e}")
    SP_STATE_FILE.write_bytes(content)
    return {"ok": True, "size": len(content)}


@app.post("/api/sharepoint/auth-login")
async def sp_auth_login():
    """Headed-browser SharePoint login. Local only."""
    try:
        await sp_downloader.login_and_save_state(state_file=SP_STATE_FILE)
    except Exception as e:
        raise HTTPException(500, f"Login failed: {e}")
    return {"ok": True}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run("app:app", host=host, port=port, reload=False, loop="asyncio")
