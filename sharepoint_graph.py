"""
SharePoint downloader using Microsoft Graph (delegated auth).

Same interface as the legacy sharepoint_downloader (process_single, process_links,
process_folder) but every API call goes through https://graph.microsoft.com with
a bearer token, instead of cookies scraped from a browser session.

Permissions required: Files.Read.All (delegated) + offline_access.
"""

from __future__ import annotations

import asyncio
import base64
import csv
import io
import os
import random
import re
from pathlib import Path
from typing import AsyncIterator, Callable, Optional

import httpx
import pandas as pd
import requests


GRAPH = "https://graph.microsoft.com/v1.0"

VIDEO_EXT = {".mp4", ".mov", ".m4v", ".webm", ".mkv", ".avi", ".wmv", ".flv"}
MEDIA_EXT = VIDEO_EXT | {".jpg", ".jpeg", ".png", ".gif", ".heic", ".webp", ".mp3", ".wav"}


# ---------------------------------------------------------------------------
# URL helpers
# ---------------------------------------------------------------------------

def encode_share_url(url: str) -> str:
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
    return "u!" + encoded


def is_folder_share(url: str) -> bool:
    return ":f:" in url or "/Forms/AllItems.aspx" in url


def looks_like_media(name: str) -> bool:
    return os.path.splitext(name)[1].lower() in MEDIA_EXT


def _safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "file"


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
    """
    Accept any of: a `Links`/`Link` column, the manifest CSV that we
    auto-generate (which has `web_url` and `download_url`), or a CSV whose
    first column is just URLs. Prefer `web_url` over `download_url` because
    download URLs expire after ~1h.
    """
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
# Graph REST helpers
# ---------------------------------------------------------------------------

def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def _graph_get(url: str, token: str) -> dict:
    r = requests.get(url, headers=_headers(token), timeout=60)
    if r.status_code == 401:
        raise PermissionError("Microsoft Graph returned 401. Token expired or insufficient permissions.")
    r.raise_for_status()
    return r.json()


def resolve_share(share_url: str, token: str) -> dict:
    """Resolve a sharing URL to its driveItem JSON."""
    share_id = encode_share_url(share_url)
    return _graph_get(f"{GRAPH}/shares/{share_id}/driveItem", token)


def list_children(drive_id: str, item_id: str, token: str) -> list[dict]:
    url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/children?$top=200"
    out: list[dict] = []
    while url:
        data = _graph_get(url, token)
        out.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return out


def _walk_folder(drive_id: str, item_id: str, token: str, path_prefix: str = "") -> list[dict]:
    """Recursively collect file items below a folder driveItem."""
    out: list[dict] = []
    try:
        kids = list_children(drive_id, item_id, token)
    except Exception as e:
        print(f"   List error at {path_prefix}: {e}")
        return out

    for child in kids:
        name = child.get("name", "")
        rel_path = f"{path_prefix}/{name}".lstrip("/")
        if "folder" in child:
            out.extend(_walk_folder(drive_id, child["id"], token, rel_path))
        else:
            dl_url = child.get("@microsoft.graph.downloadUrl") or child.get("@content.downloadUrl")
            out.append({
                "name": name,
                "path": rel_path,
                "size": child.get("size", 0),
                "download_url": dl_url or "",
                "web_url": child.get("webUrl", ""),
                "mime": (child.get("file") or {}).get("mimeType", ""),
            })
    return out


def extract_media_from_folder(folder_url: str, token: str, media_only: bool = True) -> list[dict]:
    root = resolve_share(folder_url, token)
    # For shared folders, the driveItem may be `remoteItem` (link to another drive).
    remote = root.get("remoteItem") or {}
    parent_ref = (remote.get("parentReference") or root.get("parentReference") or {})
    drive_id = parent_ref.get("driveId")
    item_id = remote.get("id") or root.get("id")

    if not drive_id or not item_id:
        raise RuntimeError(
            "Could not resolve folder share. Is the link a folder (':f:') and is your token valid?"
        )

    root_name = root.get("name") or remote.get("name") or ""
    items = _walk_folder(drive_id, item_id, token, path_prefix=root_name)
    if media_only:
        items = [i for i in items if looks_like_media(i["name"])]
    return items


def resolve_single_file(share_url: str, token: str) -> dict:
    item = resolve_share(share_url, token)
    remote = item.get("remoteItem") or {}
    dl = (
        item.get("@microsoft.graph.downloadUrl")
        or item.get("@content.downloadUrl")
        or remote.get("@microsoft.graph.downloadUrl")
        or remote.get("@content.downloadUrl")
        or ""
    )
    return {
        "name": item.get("name") or remote.get("name") or "file",
        "size": item.get("size") or remote.get("size") or 0,
        "download_url": dl,
        "web_url": item.get("webUrl") or remote.get("webUrl") or "",
        "mime": (item.get("file") or remote.get("file") or {}).get("mimeType", ""),
    }


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

def download_file(download_url: str, out_path: Path) -> bool:
    """
    Graph `@microsoft.graph.downloadUrl` is a pre-signed URL that requires
    NO auth header. We just stream it.
    """
    if not download_url:
        print("   No download URL on item.")
        return False

    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with requests.get(download_url, stream=True, timeout=120) as r:
            if r.status_code != 200:
                print(f"   HTTP {r.status_code} on download.")
                return False
            with open(out_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        f.write(chunk)
        return True
    except Exception as e:
        print(f"   Network error: {e}")
        return False


async def iter_remote_bytes(download_url: str, chunk_size: int = 1024 * 1024) -> AsyncIterator[bytes]:
    """
    Stream a SharePoint pre-signed download URL without touching disk.
    The @microsoft.graph.downloadUrl is short-lived (~hour) but doesn't need
    an Authorization header.
    """
    if not download_url:
        raise RuntimeError("Empty download URL.")
    async with httpx.AsyncClient(timeout=httpx.Timeout(300.0, connect=30.0), follow_redirects=True) as client:
        async with client.stream("GET", download_url) as r:
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code} on download.")
            async for chunk in r.aiter_bytes(chunk_size):
                if chunk:
                    yield chunk


# ---------------------------------------------------------------------------
# Job runners (sequential, rate-limit friendly)
# ---------------------------------------------------------------------------

OnProgress = Callable[..., None]
TokenGetter = Callable[[], Optional[str]]


def _emit(on_progress: Optional[OnProgress], i: int, n: int, msg: str, status: str, filename: Optional[str] = None):
    if on_progress:
        on_progress(i, n, msg, status, filename)


async def process_folder(
    folder_url: str,
    output_folder: str,
    token_getter: TokenGetter,
    on_progress: Optional[OnProgress] = None,
    media_only: bool = True,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> list[dict]:
    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)

    token = token_getter()
    if not token:
        raise RuntimeError("Not signed in to Microsoft. Sign in first.")

    _emit(on_progress, 0, 0, "Extracting media list from folder...", "info")
    items = await asyncio.to_thread(extract_media_from_folder, folder_url, token, media_only)

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

        rel = Path(item.get("path") or item["name"])
        out_file = out_dir / rel
        out_file = out_file.with_name(_safe_filename(out_file.name))

        if out_file.exists() and out_file.stat().st_size > 0:
            _emit(on_progress, index, total, f"Already downloaded: {rel}", "skipped", str(rel))
            results.append({"url": item.get("web_url", ""), "status": "skipped", "file": str(rel)})
            continue

        _emit(on_progress, index, total, f"Downloading: {rel}", "processing")
        ok = await asyncio.to_thread(download_file, item["download_url"], out_file)

        if ok:
            _emit(on_progress, index, total, f"Saved: {rel}", "success", str(rel))
            results.append({"url": item.get("web_url", ""), "status": "success", "file": str(rel)})
        else:
            _emit(on_progress, index, total, f"Failed: {rel}", "failed")
            results.append({"url": item.get("web_url", ""), "status": "failed", "file": None})

        await asyncio.sleep(random.uniform(0.3, 0.9))

    _emit(on_progress, total, total, "All downloads complete!", "done")
    return results


async def process_links(
    links: list[str],
    output_folder: str,
    token_getter: TokenGetter,
    on_progress: Optional[OnProgress] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> list[dict]:
    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)

    total = len(links)
    _emit(on_progress, 0, total, f"Found {total} links to process.", "info")

    results: list[dict] = []
    for index, url in enumerate(links, start=1):
        if should_cancel and should_cancel():
            _emit(on_progress, index - 1, total, "Cancelled by user.", "cancelled")
            break

        token = token_getter()
        if not token:
            raise RuntimeError("Session expired. Sign in again.")

        try:
            meta = await asyncio.to_thread(resolve_single_file, url, token)
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
        ok = await asyncio.to_thread(download_file, meta["download_url"], out_file)

        if ok:
            _emit(on_progress, index, total, f"Saved: {name}", "success", name)
            results.append({"url": url, "status": "success", "file": name})
        else:
            _emit(on_progress, index, total, f"Failed: {name}", "failed")
            results.append({"url": url, "status": "failed", "file": None})

        await asyncio.sleep(random.uniform(0.3, 0.9))

    _emit(on_progress, total, total, "All downloads complete!", "done")
    return results


async def process_single(
    url: str,
    output_folder: str,
    token_getter: TokenGetter,
    on_progress: Optional[OnProgress] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> dict:
    out_dir = Path(output_folder)
    out_dir.mkdir(parents=True, exist_ok=True)

    token = token_getter()
    if not token:
        raise RuntimeError("Not signed in to Microsoft. Sign in first.")

    _emit(on_progress, 1, 1, f"Resolving: {url}", "processing")
    meta = await asyncio.to_thread(resolve_single_file, url, token)

    name = _safe_filename(meta["name"])
    out_file = out_dir / name

    _emit(on_progress, 1, 1, f"Downloading: {name}", "processing")
    ok = await asyncio.to_thread(download_file, meta["download_url"], out_file)

    if ok:
        _emit(on_progress, 1, 1, f"Saved: {name}", "success", name)
        _emit(on_progress, 1, 1, "Download complete!", "done")
        return {"url": url, "status": "success", "file": name}

    _emit(on_progress, 1, 1, "Failed to download.", "failed")
    return {"url": url, "status": "failed", "file": None}


# ---------------------------------------------------------------------------
# Manifest-only crawl + batch streaming (new direct-to-browser path)
# ---------------------------------------------------------------------------

async def crawl_folder(
    folder_url: str,
    token_getter: TokenGetter,
    on_progress: Optional[OnProgress] = None,
    media_only: bool = True,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> dict:
    """
    Walk a folder share. Returns {"folder_name": str, "items": list[dict]}.
    Items contain pre-signed download URLs (valid ~1 hour). Nothing written
    to disk.
    """
    token = token_getter()
    if not token:
        raise RuntimeError("Not signed in to Microsoft. Sign in first.")

    _emit(on_progress, 0, 0, "Crawling SharePoint folder...", "info")

    # Resolve the share once up-front so we can name the manifest after the
    # actual folder (e.g. "Campaign Q3 2025.csv" rather than a job-id stub).
    root = await asyncio.to_thread(resolve_share, folder_url, token)
    remote = root.get("remoteItem") or {}
    folder_name = (root.get("name") or remote.get("name") or "").strip()

    items = await asyncio.to_thread(extract_media_from_folder, folder_url, token, media_only)
    _emit(on_progress, 0, len(items),
          f"Found {len(items)} media file(s) in '{folder_name or 'folder'}'.", "info")
    if should_cancel and should_cancel():
        _emit(on_progress, 0, len(items), "Cancelled by user.", "cancelled")
    return {"folder_name": folder_name, "items": items}


async def iter_batch_files(
    links: list[str],
    token_getter: TokenGetter,
    on_progress: Optional[OnProgress] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
    batch_offset: int = 0,
    overall_total: int | None = None,
):
    """
    For a list of SharePoint share URLs, yield (filename, async_byte_iter, result_dict)
    triples. The consumer (zip stream) iterates bytes; we update result_dict
    once iteration finishes and call on_progress.
    """
    total = overall_total or len(links)
    if on_progress:
        on_progress(batch_offset, total, f"Starting batch of {len(links)} link(s)...", "info")

    for local_idx, url in enumerate(links, start=1):
        idx_overall = batch_offset + local_idx
        if should_cancel and should_cancel():
            if on_progress:
                on_progress(idx_overall - 1, total, "Cancelled by user.", "cancelled")
            return

        token = token_getter()
        if not token:
            if on_progress:
                on_progress(idx_overall, total, "Session expired. Sign in again.", "failed")
            yield None, None, {"url": url, "status": "failed", "file": None}
            continue

        if on_progress:
            on_progress(idx_overall, total, f"Resolving: {url}", "processing")

        try:
            meta = await asyncio.to_thread(resolve_single_file, url, token)
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
                async for chunk in iter_remote_bytes(dl):
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
    on_progress: Optional[OnProgress] = None,
    should_cancel: Optional[Callable[[], bool]] = None,
    batch_offset: int = 0,
    overall_total: int | None = None,
):
    """
    Like iter_batch_files but consumes a pre-resolved manifest (so we don't
    re-resolve each share). Each item must have 'name', 'download_url' (and
    optionally 'path' for nested arcname).
    """
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
                async for chunk in iter_remote_bytes(dl):
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
    token_getter: TokenGetter,
) -> tuple[str, AsyncIterator[bytes]]:
    """
    Resolve one share link and return (filename, async byte iterator).
    Used by /api/sharepoint/single.bin endpoint.
    """
    token = token_getter()
    if not token:
        raise RuntimeError("Not signed in to Microsoft. Sign in first.")
    meta = await asyncio.to_thread(resolve_single_file, url, token)
    name = _safe_filename(meta.get("name") or "sharepoint_file")
    dl = meta.get("download_url") or ""
    if not dl:
        raise RuntimeError("No download URL on item.")
    return name, iter_remote_bytes(dl)
