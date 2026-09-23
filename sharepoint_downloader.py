"""
SharePoint / OneDrive video downloader.

Auth strategy: a one-time interactive Playwright login saves cookies to
`auth/sharepoint_state.json`. All subsequent operations reuse that state to
make authenticated calls to the SharePoint REST API.

URL flow:
  - Single share link / CSV  -> resolve each share -> download the file
  - Folder share link        -> traverse children recursively -> list of files
                                -> write a CSV manifest -> download each file
"""

from __future__ import annotations

import asyncio
import base64
import csv
import io
import os
import random
import re
import time
from pathlib import Path
from typing import Any, AsyncIterator, Callable, Optional
from urllib.parse import urlparse

import pandas as pd
import requests
from playwright.async_api import async_playwright


# ---------------------------------------------------------------------------
# Paths / constants
# ---------------------------------------------------------------------------

DEFAULT_AUTH_DIR = Path("auth")
DEFAULT_STATE_FILE = DEFAULT_AUTH_DIR / "sharepoint_state.json"

DESKTOP_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

VIDEO_EXTENSIONS = {".mp4", ".mov", ".m4v", ".webm", ".mkv", ".avi", ".wmv", ".flv"}
MEDIA_EXTENSIONS = VIDEO_EXTENSIONS | {".jpg", ".jpeg", ".png", ".gif", ".heic", ".webp", ".mp3", ".wav"}


# ---------------------------------------------------------------------------
# URL helpers
# ---------------------------------------------------------------------------

def encode_share_url(url: str) -> str:
    """Encode a SharePoint sharing URL into the `u!...` share-id format."""
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
    return "u!" + encoded


def tenant_from_url(url: str) -> str:
    """e.g. https://ruskmedia.sharepoint.com/... -> https://ruskmedia.sharepoint.com"""
    p = urlparse(url)
    return f"{p.scheme}://{p.netloc}"


def is_folder_share(url: str) -> bool:
    """SharePoint folder shares use the `:f:` token; files use `:b:`, videos `:v:` etc."""
    return ":f:" in url or "/Forms/AllItems.aspx" in url


def looks_like_media(name: str) -> bool:
    ext = os.path.splitext(name)[1].lower()
    return ext in MEDIA_EXTENSIONS


def is_video(name: str) -> bool:
    return os.path.splitext(name)[1].lower() in VIDEO_EXTENSIONS


# ---------------------------------------------------------------------------
# Auth: one-time interactive login
# ---------------------------------------------------------------------------

async def login_and_save_state(
    state_file: Path = DEFAULT_STATE_FILE,
    tenant_url: str = "https://ruskmedia.sharepoint.com",
    timeout_s: int = 300,
) -> Path:
    """
    Opens a headed browser, lets the user sign in to Microsoft / SharePoint,
    waits for them to land on the tenant homepage, then saves cookies.

    Must be run on a machine with a display (your laptop), NOT on a headless
    server like a Hugging Face Space. You can copy the resulting JSON to the
    server.
    """
    state_file = Path(state_file)
    state_file.parent.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(user_agent=DESKTOP_UA)
        page = await context.new_page()
        await page.goto(tenant_url)

        print(f"\nPlease sign in to {tenant_url} in the opened browser.")
        print(f"Waiting up to {timeout_s}s for the page to redirect to your tenant...")

        # Wait until URL contains the tenant host (post-login).
        host = urlparse(tenant_url).netloc
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            if host in page.url and "login.microsoftonline.com" not in page.url:
                # Give the page a couple seconds to set all cookies.
                await page.wait_for_timeout(3000)
                break
            await page.wait_for_timeout(1000)

        await context.storage_state(path=str(state_file))
        await browser.close()

    return state_file


def auth_state_exists(state_file: Path = DEFAULT_STATE_FILE) -> bool:
    p = Path(state_file)
    return p.exists() and p.stat().st_size > 100


# ---------------------------------------------------------------------------
# Authenticated requests session (built from Playwright storage_state)
# ---------------------------------------------------------------------------

def _session_from_state(state_file: Path, tenant_host: str) -> requests.Session:
    """Build a requests.Session preloaded with cookies from storage_state.json."""
    import json
    s = requests.Session()
    s.headers.update({
        "User-Agent": DESKTOP_UA,
        "Accept": "application/json;odata=verbose",
    })
    with open(state_file, "r", encoding="utf-8") as f:
        state = json.load(f)
    for c in state.get("cookies", []):
        # We attach all cookies but they'll only be sent to matching domains.
        s.cookies.set(
            name=c["name"],
            value=c["value"],
            domain=c.get("domain"),
            path=c.get("path", "/"),
        )
    return s


# ---------------------------------------------------------------------------
# SharePoint REST: resolve share, list children, get download URL
# ---------------------------------------------------------------------------

def _resolve_share(session: requests.Session, tenant: str, share_url: str) -> dict:
    """
    Given a sharing URL, return the driveItem JSON describing it.
    Uses SharePoint's v2.0 shares endpoint (no Graph token needed when called
    with browser cookies for the same tenant).
    """
    share_id = encode_share_url(share_url)
    api = f"{tenant}/_api/v2.0/shares/{share_id}/driveItem"
    r = session.get(api, headers={"Accept": "application/json"})
    r.raise_for_status()
    return r.json()


def _list_children(session: requests.Session, drive_id: str, item_id: str) -> list[dict]:
    """List children of a folder driveItem. Handles pagination via @odata.nextLink."""
    base = f"https://{urlparse(session.headers.get('Referer', '')).netloc}"
    # Use the URL embedded in the parent item if present; safer to construct here:
    # SharePoint v2 returns nextLink as full URL.
    children = []
    url = f"/_api/v2.0/drives/{drive_id}/items/{item_id}/children?$top=200"
    next_url: Optional[str] = url
    while next_url:
        if next_url.startswith("/"):
            # We need the tenant host. Use first cookie domain to figure it out.
            tenant_host = None
            for c in session.cookies:
                if c.domain and "sharepoint.com" in c.domain:
                    tenant_host = c.domain.lstrip(".")
                    break
            full = f"https://{tenant_host}{next_url}"
        else:
            full = next_url
        r = session.get(full, headers={"Accept": "application/json"})
        r.raise_for_status()
        data = r.json()
        children.extend(data.get("value", []))
        next_url = data.get("@odata.nextLink")
    return children


def _walk_folder(session: requests.Session, drive_id: str, item_id: str,
                 path_prefix: str = "") -> list[dict]:
    """
    Recursively walk a SharePoint folder and return a flat list of file items.
    Each entry: {name, path, size, download_url, web_url, mime}
    """
    out: list[dict] = []
    try:
        kids = _list_children(session, drive_id, item_id)
    except Exception as e:
        print(f"   List error at {path_prefix}: {e}")
        return out

    for child in kids:
        name = child.get("name", "")
        is_folder = "folder" in child
        rel_path = f"{path_prefix}/{name}".lstrip("/")

        if is_folder:
            out.extend(_walk_folder(session, drive_id, child["id"], rel_path))
        else:
            dl_url = child.get("@content.downloadUrl") or child.get("@microsoft.graph.downloadUrl")
            out.append({
                "name": name,
                "path": rel_path,
                "size": child.get("size", 0),
                "download_url": dl_url or "",
                "web_url": child.get("webUrl", ""),
                "mime": (child.get("file") or {}).get("mimeType", ""),
            })
    return out


def extract_media_from_folder(
    folder_share_url: str,
    state_file: Path = DEFAULT_STATE_FILE,
    media_only: bool = True,
) -> list[dict]:
    """
    Top-level: resolve a folder share, traverse recursively, return media files.
    """
    tenant = tenant_from_url(folder_share_url)
    session = _session_from_state(state_file, urlparse(tenant).netloc)

    root = _resolve_share(session, tenant, folder_share_url)
    drive_id = root.get("parentReference", {}).get("driveId") or (root.get("remoteItem") or {}).get("parentReference", {}).get("driveId")
    item_id = root.get("id") or (root.get("remoteItem") or {}).get("id")

    if not drive_id or not item_id:
        raise RuntimeError(
            "Could not resolve folder share. Is the link a folder (`:f:`) and is your login valid?"
        )

    items = _walk_folder(session, drive_id, item_id, path_prefix=root.get("name", ""))
    if media_only:
        items = [i for i in items if looks_like_media(i["name"])]
    return items


def resolve_single_file(share_url: str, state_file: Path = DEFAULT_STATE_FILE) -> dict:
    """Resolve a single file share link -> download metadata."""
    tenant = tenant_from_url(share_url)
    session = _session_from_state(state_file, urlparse(tenant).netloc)
    item = _resolve_share(session, tenant, share_url)
    return {
        "name": item.get("name", "file"),
        "size": item.get("size", 0),
        "download_url": item.get("@content.downloadUrl") or item.get("@microsoft.graph.downloadUrl") or "",
        "web_url": item.get("webUrl", ""),
        "mime": (item.get("file") or {}).get("mimeType", ""),
    }


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

def _safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "file"


def download_file(download_url: str, out_path: Path, state_file: Path = DEFAULT_STATE_FILE) -> bool:
    """Stream a file from SharePoint to disk using cookies from the saved state."""
    if not download_url:
        print("   No download URL on item.")
        return False

    host = urlparse(download_url).netloc
    session = _session_from_state(state_file, host)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        with session.get(download_url, stream=True, timeout=120) as r:
            if r.status_code != 200:
                print(f"   HTTP {r.status_code} for {download_url[:80]}...")
                return False
            with open(out_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        f.write(chunk)
        return True
    except Exception as e:
        print(f"   Network error: {e}")
        return False


# ---------------------------------------------------------------------------
# CSV helpers
# ---------------------------------------------------------------------------

def read_csv_flexible(content: bytes) -> pd.DataFrame:
    """
    Read CSV bytes regardless of source encoding. CSVs saved from Excel with
    Hindi/regional Indian language text are often UTF-16 ("Unicode Text") or
    a legacy codepage rather than plain UTF-8.
    """
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            return pd.read_csv(io.BytesIO(content), encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return pd.read_csv(io.BytesIO(content), encoding="utf-8", encoding_errors="replace")


def parse_links_from_csv_bytes(content: bytes) -> list[str]:
    """Same shape as sharepoint_graph: recognise the manifest CSV too."""
    df = read_csv_flexible(content)
    for col in ("Links", "Link", "web_url", "download_url"):
        if col in df.columns:
            links = df[col].dropna().tolist()
            break
    else:
        links = df[df.columns[0]].dropna().tolist()
    return [str(u).strip() for u in links if str(u).strip().startswith("http")]


def _title_only(name: str) -> str:
    """File name without its extension, e.g. 'मेरी फ़ाइल.mp4' -> 'मेरी फ़ाइल'."""
    return os.path.splitext(name or "")[0]


def items_to_csv_bytes(items: list[dict]) -> bytes:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["name", "path", "size_mb", "mime", "download_url", "web_url"])
    for it in items:
        w.writerow([
            _title_only(it.get("name", "")),
            it.get("path", ""),
            round((it.get("size") or 0) / (1024 * 1024), 2),
            it.get("mime", ""),
            it.get("download_url", ""),
            it.get("web_url", ""),
        ])
    # UTF-8 with BOM so Excel renders Hindi/regional-language names correctly.
    return buf.getvalue().encode("utf-8-sig")


# ---------------------------------------------------------------------------
# Job runners (mirror the FB downloader interface)
# ---------------------------------------------------------------------------

OnProgress = Callable[..., None]


def _emit(on_progress: Optional[OnProgress], index: int, total: int,
          message: str, status: str, filename: Optional[str] = None) -> None:
    if on_progress:
        on_progress(index, total, message, status, filename)


async def process_folder(
    folder_url: str,
    output_folder: str = "downloads_sharepoint",
    state_file: Path = DEFAULT_STATE_FILE,
    on_progress: Optional[OnProgress] = None,
    media_only: bool = True,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> list[dict]:
    """
    1. Extract every media file (recursively) under a SharePoint folder share.
    2. Write a CSV manifest into the output folder.
    3. Download each file (sequential, with rate-limit-friendly delays).

    Returns a list of result dicts: {url, status, file}.
    """
    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not auth_state_exists(state_file):
        raise RuntimeError(
            "Not logged in to SharePoint. Run the login step first to save auth state."
        )

    _emit(on_progress, 0, 0, "Extracting media list from folder...", "info")

    # _walk_folder is sync (requests) so run in a worker thread to keep loop responsive.
    items = await asyncio.to_thread(extract_media_from_folder, folder_url, state_file, media_only)

    manifest = out_dir / "manifest.csv"
    manifest.write_bytes(items_to_csv_bytes(items))
    _emit(on_progress, 0, len(items),
          f"Found {len(items)} media files. Manifest saved to {manifest.name}.",
          "info", manifest.name)

    results: list[dict] = []
    total = len(items)
    for index, item in enumerate(items, start=1):
        if should_cancel and should_cancel():
            _emit(on_progress, index - 1, total, "Cancelled by user.", "cancelled")
            break

        name = _safe_filename(item["name"])
        # Preserve subfolder structure
        rel = Path(item.get("path") or name)
        out_file = out_dir / rel
        out_file = out_file.with_name(_safe_filename(out_file.name))

        if out_file.exists() and out_file.stat().st_size > 0:
            _emit(on_progress, index, total, f"Already downloaded: {rel}", "skipped", str(rel))
            results.append({"url": item.get("web_url", ""), "status": "skipped", "file": str(rel)})
            continue

        _emit(on_progress, index, total, f"Downloading: {rel}", "processing")
        ok = await asyncio.to_thread(download_file, item["download_url"], out_file, state_file)

        if ok:
            _emit(on_progress, index, total, f"Saved: {rel}", "success", str(rel))
            results.append({"url": item.get("web_url", ""), "status": "success", "file": str(rel)})
        else:
            _emit(on_progress, index, total, f"Failed: {rel}", "failed")
            results.append({"url": item.get("web_url", ""), "status": "failed", "file": None})

        # Rate-limit friendly: SharePoint is more tolerant than FB but still throttles.
        delay = random.uniform(0.5, 1.5)
        await asyncio.sleep(delay)

    _emit(on_progress, total, total, "All downloads complete!", "done")
    return results


async def process_links(
    links: list[str],
    output_folder: str = "downloads_sharepoint",
    state_file: Path = DEFAULT_STATE_FILE,
    on_progress: Optional[OnProgress] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> list[dict]:
    """Process a list of single-file SharePoint share links."""
    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not auth_state_exists(state_file):
        raise RuntimeError("Not logged in to SharePoint. Run the login step first.")

    total = len(links)
    _emit(on_progress, 0, total, f"Found {total} links to process.", "info")

    results: list[dict] = []
    for index, url in enumerate(links, start=1):
        if should_cancel and should_cancel():
            _emit(on_progress, index - 1, total, "Cancelled by user.", "cancelled")
            break

        try:
            meta = await asyncio.to_thread(resolve_single_file, url, state_file)
        except Exception as e:
            _emit(on_progress, index, total, f"Resolve failed: {e}", "failed")
            results.append({"url": url, "status": "failed", "file": None})
            continue

        name = _safe_filename(meta["name"])
        out_file = out_dir / name

        if out_file.exists() and out_file.stat().st_size > 0:
            _emit(on_progress, index, total, f"Already downloaded: {name}", "skipped", name)
            results.append({"url": url, "status": "skipped", "file": name})
            continue

        _emit(on_progress, index, total, f"Downloading: {name}", "processing")
        ok = await asyncio.to_thread(download_file, meta["download_url"], out_file, state_file)

        if ok:
            _emit(on_progress, index, total, f"Saved: {name}", "success", name)
            results.append({"url": url, "status": "success", "file": name})
        else:
            _emit(on_progress, index, total, f"Failed: {name}", "failed")
            results.append({"url": url, "status": "failed", "file": None})

        delay = random.uniform(0.5, 1.5)
        await asyncio.sleep(delay)

    _emit(on_progress, total, total, "All downloads complete!", "done")
    return results


async def process_single(
    url: str,
    output_folder: str = "downloads_sharepoint",
    state_file: Path = DEFAULT_STATE_FILE,
    on_progress: Optional[OnProgress] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> dict:
    """Download a single SharePoint share link."""
    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not auth_state_exists(state_file):
        raise RuntimeError("Not logged in to SharePoint. Run the login step first.")

    _emit(on_progress, 1, 1, f"Resolving: {url}", "processing")
    meta = await asyncio.to_thread(resolve_single_file, url, state_file)

    name = _safe_filename(meta["name"])
    out_file = out_dir / name

    _emit(on_progress, 1, 1, f"Downloading: {name}", "processing")
    ok = await asyncio.to_thread(download_file, meta["download_url"], out_file, state_file)

    if ok:
        _emit(on_progress, 1, 1, f"Saved: {name}", "success", name)
        _emit(on_progress, 1, 1, "Download complete!", "done")
        return {"url": url, "status": "success", "file": name}

    _emit(on_progress, 1, 1, "Failed to download.", "failed")
    return {"url": url, "status": "failed", "file": None}


# ---------------------------------------------------------------------------
# Streaming (cookie-auth) helpers for the direct-to-browser pipeline
# ---------------------------------------------------------------------------

def _iter_remote_bytes_sync(download_url: str, state_file: Path, chunk_size: int = 1024 * 1024):
    """Sync generator yielding chunks via the cookie session."""
    if not download_url:
        raise RuntimeError("Empty download URL.")
    host = urlparse(download_url).netloc
    session = _session_from_state(state_file, host)
    with session.get(download_url, stream=True, timeout=300) as r:
        if r.status_code != 200:
            raise RuntimeError(f"HTTP {r.status_code} on download.")
        for chunk in r.iter_content(chunk_size):
            if chunk:
                yield chunk


async def iter_remote_bytes(download_url: str, state_file: Path, chunk_size: int = 1024 * 1024) -> AsyncIterator[bytes]:
    """
    Async iterator over file bytes via the legacy cookie session.
    `requests` is sync, so we run the underlying generator in a worker thread
    and ferry chunks through an asyncio.Queue.
    """
    loop = asyncio.get_running_loop()
    queue: asyncio.Queue = asyncio.Queue(maxsize=4)
    _DONE = object()

    def _producer():
        try:
            for chunk in _iter_remote_bytes_sync(download_url, state_file, chunk_size):
                asyncio.run_coroutine_threadsafe(queue.put(chunk), loop).result()
            asyncio.run_coroutine_threadsafe(queue.put(_DONE), loop).result()
        except Exception as e:
            asyncio.run_coroutine_threadsafe(queue.put(e), loop).result()

    import threading
    t = threading.Thread(target=_producer, daemon=True)
    t.start()

    while True:
        item = await queue.get()
        if item is _DONE:
            return
        if isinstance(item, Exception):
            raise item
        yield item


def _resolve_folder_name(folder_url: str, state_file: Path) -> str:
    """Resolve a share URL just to read its root folder name."""
    tenant = tenant_from_url(folder_url)
    session = _session_from_state(state_file, urlparse(tenant).netloc)
    try:
        root = _resolve_share(session, tenant, folder_url)
    except Exception:
        return ""
    remote = root.get("remoteItem") or {}
    return (root.get("name") or remote.get("name") or "").strip()


async def crawl_folder(
    folder_url: str,
    state_file: Path,
    on_progress: Optional[OnProgress] = None,
    media_only: bool = True,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> dict:
    """Walk a folder share via cookie auth. Returns {"folder_name", "items"}."""
    if not auth_state_exists(state_file):
        raise RuntimeError("Not logged in to SharePoint. Sign in first.")
    _emit(on_progress, 0, 0, "Crawling SharePoint folder...", "info")
    folder_name = await asyncio.to_thread(_resolve_folder_name, folder_url, state_file)
    items = await asyncio.to_thread(extract_media_from_folder, folder_url, state_file, media_only)
    _emit(on_progress, 0, len(items),
          f"Found {len(items)} media file(s) in '{folder_name or 'folder'}'.", "info")
    if should_cancel and should_cancel():
        _emit(on_progress, 0, len(items), "Cancelled by user.", "cancelled")
    return {"folder_name": folder_name, "items": items}


async def iter_batch_files(
    links: list[str],
    state_file: Path,
    on_progress: Optional[OnProgress] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
    batch_offset: int = 0,
    overall_total: int | None = None,
):
    """For each share link, yield (filename, async byte iter, result dict)."""
    if not auth_state_exists(state_file):
        raise RuntimeError("Not logged in to SharePoint.")

    total = overall_total or len(links)
    if on_progress:
        on_progress(batch_offset, total, f"Starting batch of {len(links)} link(s)...", "info")

    for local_idx, url in enumerate(links, start=1):
        idx_overall = batch_offset + local_idx
        if should_cancel and should_cancel():
            if on_progress:
                on_progress(idx_overall - 1, total, "Cancelled by user.", "cancelled")
            return

        if on_progress:
            on_progress(idx_overall, total, f"Resolving: {url}", "processing")

        try:
            meta = await asyncio.to_thread(resolve_single_file, url, state_file)
        except Exception as e:
            if on_progress:
                on_progress(idx_overall, total, f"Resolve failed: {e}", "failed")
            yield None, None, {"url": url, "status": "failed", "file": None}
            continue

        name = _safe_filename(meta.get("name") or f"file_{idx_overall}")
        download_url = meta.get("download_url") or ""

        if not download_url:
            if on_progress:
                on_progress(idx_overall, total, "No download URL on item.", "failed")
            yield None, None, {"url": url, "status": "failed", "file": None}
            continue

        if on_progress:
            on_progress(idx_overall, total, f"Streaming: {name}", "processing")

        result = {"url": url, "status": "pending", "file": name}

        async def _stream_with_status(dl=download_url, res=result, idx=idx_overall, nm=name):
            try:
                async for chunk in iter_remote_bytes(dl, state_file):
                    yield chunk
                res["status"] = "success"
                if on_progress:
                    on_progress(idx, total, f"Saved: {nm}", "success", nm)
            except Exception as e:
                res["status"] = "failed"
                res["file"] = None
                if on_progress:
                    on_progress(idx, total, f"Stream failed: {e}", "failed")
                raise

        yield name, _stream_with_status(), result

        if local_idx < len(links):
            await asyncio.sleep(random.uniform(0.3, 0.8))


async def iter_batch_files_from_manifest(
    items: list[dict],
    state_file: Path,
    on_progress: Optional[OnProgress] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
    batch_offset: int = 0,
    overall_total: int | None = None,
):
    """Like iter_batch_files but consumes a pre-resolved manifest."""
    total = overall_total or len(items)
    if on_progress:
        on_progress(batch_offset, total, f"Starting batch of {len(items)} file(s)...", "info")

    for local_idx, item in enumerate(items, start=1):
        idx_overall = batch_offset + local_idx
        if should_cancel and should_cancel():
            if on_progress:
                on_progress(idx_overall - 1, total, "Cancelled by user.", "cancelled")
            return

        rel_path = item.get("path") or item.get("name") or f"file_{idx_overall}"
        arcname = _safe_filename(Path(rel_path).name)
        download_url = item.get("download_url") or ""

        if not download_url:
            if on_progress:
                on_progress(idx_overall, total, f"No download URL: {arcname}", "failed")
            yield None, None, {"url": item.get("web_url", ""), "status": "failed", "file": None}
            continue

        if on_progress:
            on_progress(idx_overall, total, f"Streaming: {arcname}", "processing")

        result = {"url": item.get("web_url", ""), "status": "pending", "file": arcname}

        async def _stream_with_status(dl=download_url, res=result, idx=idx_overall, nm=arcname):
            try:
                async for chunk in iter_remote_bytes(dl, state_file):
                    yield chunk
                res["status"] = "success"
                if on_progress:
                    on_progress(idx, total, f"Saved: {nm}", "success", nm)
            except Exception as e:
                res["status"] = "failed"
                res["file"] = None
                if on_progress:
                    on_progress(idx, total, f"Stream failed: {e}", "failed")
                raise

        yield arcname, _stream_with_status(), result

        if local_idx < len(items):
            await asyncio.sleep(random.uniform(0.2, 0.5))


async def resolve_single_stream(
    url: str,
    state_file: Path,
) -> tuple[str, AsyncIterator[bytes]]:
    """Resolve one share link via cookie auth and return (filename, byte iter)."""
    if not auth_state_exists(state_file):
        raise RuntimeError("Not logged in to SharePoint.")
    meta = await asyncio.to_thread(resolve_single_file, url, state_file)
    name = _safe_filename(meta.get("name") or "sharepoint_file")
    dl = meta.get("download_url") or ""
    if not dl:
        raise RuntimeError("No download URL on item.")
    return name, iter_remote_bytes(dl, state_file)


# ---------------------------------------------------------------------------
# CLI: standalone login
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    tenant = sys.argv[1] if len(sys.argv) > 1 else "https://ruskmedia.sharepoint.com"
    print(f"Launching browser for SharePoint login at {tenant}")
    asyncio.run(login_and_save_state(tenant_url=tenant))
    print(f"Saved auth state to {DEFAULT_STATE_FILE}")
