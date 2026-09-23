"""
Google Drive downloader for public "Anyone with the link" shares.

No OAuth or sign-in required. Files are fetched via drive.google.com export
URLs (with virus-scan confirm handling for large files). Public folders are
walked by parsing the shared folder page HTML.

Mirrors the public surface of sharepoint_graph.py:
    parse_drive_url, is_folder_url, looks_like_media
    parse_links_from_csv_bytes, items_to_csv_bytes
    crawl_folder(folder_url, on_progress, media_only, should_cancel)
        -> {"folder_name": str, "items": list[dict]}
    iter_remote_bytes(file_id) -> AsyncIterator[bytes]
    iter_batch_files(links, on_progress, should_cancel, batch_offset, overall_total)
    resolve_single_stream(url) -> (filename, AsyncIterator[bytes])
"""

from __future__ import annotations

import asyncio
import csv
import io
import os
import random
import re
from html import unescape as _html_unescape
from typing import AsyncIterator, Callable, Optional
from urllib.parse import urlparse, parse_qs

import httpx
import pandas as pd
import requests


VIDEO_EXT = {".mp4", ".mov", ".m4v", ".webm", ".mkv", ".avi", ".wmv", ".flv"}
MEDIA_EXT = VIDEO_EXT | {".jpg", ".jpeg", ".png", ".gif", ".heic", ".webp", ".mp3", ".wav"}

GOOGLE_NATIVE_PREFIX = "application/vnd.google-apps"
FOLDER_MIME = "application/vnd.google-apps.folder"

_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

_FILE_ID_RE = re.compile(r"/file/d/([a-zA-Z0-9_-]+)")
_FOLDER_ID_RE = re.compile(r"/folders/([a-zA-Z0-9_-]+)")
_DOC_ID_RE = re.compile(r"/document/d/([a-zA-Z0-9_-]+)")
_PRES_ID_RE = re.compile(r"/presentation/d/([a-zA-Z0-9_-]+)")
_SHEET_ID_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9_-]+)")
_OG_TITLE_RE = re.compile(r'<meta property="og:title" content="([^"]+)"')
_TITLE_RE = re.compile(r"<title>([^<]+)</title>")

# Public folders are listed via the `embeddedfolderview` endpoint. Its HTML is
# stable and easy to parse, unlike the main folder page which embeds contents in
# a frequently-changing JS blob (`_DRIVE_ivd`). Each child renders as a
# `<div class="flip-entry" id="entry-<ID>">` block with one anchor (href tells
# file/folder/native-doc) and one `flip-entry-title` holding the display name.
_EMBED_URL = "https://drive.google.com/embeddedfolderview?id={fid}#list"
_EMBED_ID_RE = re.compile(r'\bid="entry-([-\w]{10,})"')
_EMBED_HREF_RE = re.compile(r'<a\s+href="([^"]+)"')
_EMBED_NAME_RE = re.compile(r'flip-entry-title">([^<]*)<')


# ---------------------------------------------------------------------------
# URL helpers
# ---------------------------------------------------------------------------

def parse_drive_url(url: str) -> dict:
    """Return {'kind': 'file'|'folder'|'unknown', 'id': str|None}."""
    if not url:
        return {"kind": "unknown", "id": None}
    p = urlparse(url.strip())
    qs = parse_qs(p.query or "")
    if "id" in qs and qs["id"]:
        fid = qs["id"][0]
        if "/folders/" in (p.path or ""):
            return {"kind": "folder", "id": fid}
        return {"kind": "file", "id": fid}
    m = _FOLDER_ID_RE.search(p.path or "")
    if m:
        return {"kind": "folder", "id": m.group(1)}
    for pat in (_FILE_ID_RE, _DOC_ID_RE, _PRES_ID_RE, _SHEET_ID_RE):
        m = pat.search(p.path or "")
        if m:
            return {"kind": "file", "id": m.group(1)}
    return {"kind": "unknown", "id": None}


def is_folder_url(url: str) -> bool:
    return parse_drive_url(url)["kind"] == "folder"


def looks_like_media(name: str) -> bool:
    return os.path.splitext(name)[1].lower() in MEDIA_EXT


def _safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", name).strip() or "file"


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
    Accept a CSV with a `Links`/`Link` column, the manifest CSV that we
    auto-generate (`web_url` / `id`), or a single-column CSV of Drive URLs.
    """
    df = read_csv_flexible(content)
    out: list[str] = []

    for col in ("Links", "Link", "web_url", "download_url"):
        if col in df.columns:
            for u in df[col].dropna().tolist():
                s = str(u).strip()
                if s.startswith("http"):
                    out.append(s)
            break
    else:
        for u in df[df.columns[0]].dropna().tolist():
            s = str(u).strip()
            if s.startswith("http"):
                out.append(s)

    if "id" in df.columns:
        for fid in df["id"].dropna().tolist():
            s = str(fid).strip()
            if s and re.fullmatch(r"[a-zA-Z0-9_-]{10,}", s):
                url = f"https://drive.google.com/file/d/{s}/view"
                if url not in out:
                    out.append(url)

    return out


def _title_only(name: str) -> str:
    """File name without its extension, e.g. 'मेरी फ़ाइल.mp4' -> 'मेरी फ़ाइल'."""
    return os.path.splitext(name or "")[0]


def items_to_csv_bytes(items: list[dict]) -> bytes:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["name", "path", "size_mb", "mime", "id", "web_url"])
    for it in items:
        size_mb = round((int(it.get("size") or 0)) / (1024 * 1024), 2)
        w.writerow([
            _title_only(it.get("name", "")),
            it.get("path", ""),
            size_mb,
            it.get("mime", ""),
            it.get("id", ""),
            it.get("web_url", ""),
        ])
    # UTF-8 with BOM so Excel renders Hindi/regional-language names correctly.
    return buf.getvalue().encode("utf-8-sig")


# ---------------------------------------------------------------------------
# Public-link HTTP helpers (sync, run via asyncio.to_thread)
# ---------------------------------------------------------------------------

def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": _USER_AGENT})
    return s


def _page_title(html: str, fallback: str = "") -> str:
    m = _OG_TITLE_RE.search(html)
    if m:
        return m.group(1).strip()
    m = _TITLE_RE.search(html)
    if m:
        return m.group(1).replace(" - Google Drive", "").strip()
    return fallback


def _is_login_wall(html: str) -> bool:
    """True only when the page is an actual sign-in gate, not a public share."""
    if _OG_TITLE_RE.search(html):
        return False
    lowered = html.lower()
    if "sign in to continue to google drive" in lowered:
        return True
    if "accounts.google.com/servicelogin" in lowered and "drive.google.com/file" not in lowered:
        return True
    return False


def _probe_file_accessible(file_id: str, session: requests.Session) -> bool:
    """HEAD/GET the export URL; public files return bytes or a virus-scan page."""
    r = session.get(
        f"https://drive.google.com/uc?export=download&id={file_id}",
        stream=True,
        timeout=30,
    )
    ct = (r.headers.get("Content-Type") or "").lower()
    if "text/html" not in ct:
        return True
    snippet = (r.text or "")[:8000].lower()
    if "google drive can't scan this file for viruses" in snippet:
        return True
    if "download anyway" in snippet or "confirm=" in snippet:
        return True
    return not _is_login_wall(r.text or "")


def build_public_download_url(file_id: str,
                              session: requests.Session | None = None
                              ) -> tuple[str, requests.Session]:
    """
    Resolve the final uc?export=download URL for a public file, handling the
    virus-scan confirmation page for large files.
    """
    session = session or _session()
    base = f"https://drive.google.com/uc?export=download&id={file_id}"
    resp = session.get(base, stream=True, timeout=60)

    ct = (resp.headers.get("Content-Type") or "").lower()
    if "text/html" not in ct:
        return base, session

    token = None
    for key, value in resp.cookies.items():
        if key.startswith("download_warning"):
            token = value
            break
    if not token:
        m = re.search(r"confirm=([0-9A-Za-z_]+)", resp.text)
        token = m.group(1) if m else "t"
    return f"{base}&confirm={token}", session


def public_file_meta(file_id: str, session: requests.Session | None = None) -> dict:
    """Read name and mime hints from a public file's view page."""
    session = session or _session()
    r = session.get(f"https://drive.google.com/file/d/{file_id}/view", timeout=60)
    if r.status_code == 404:
        raise FileNotFoundError(
            "File not found or not shared as 'Anyone with the link'."
        )
    if r.status_code != 200:
        r.raise_for_status()

    html = r.text
    name = _safe_filename(_page_title(html, ""))
    if not name:
        if not _probe_file_accessible(file_id, session):
            raise PermissionError(
                "This file is not publicly accessible. "
                "Set sharing to 'Anyone with the link' and try again."
            )
        name = f"gdrive_{file_id[:8]}"
    mime = ""
    if GOOGLE_NATIVE_PREFIX in html:
        m = re.search(r'(application/vnd\.google-apps\.[a-z.]+)', html)
        if m:
            mime = m.group(1)

    return {
        "id": file_id,
        "name": name,
        "mime": mime,
        "size": 0,
        "web_url": f"https://drive.google.com/file/d/{file_id}/view",
    }


def _parse_folder_entries(html: str, folder_id: str) -> list[dict]:
    """Extract child file/folder entries from an embeddedfolderview page."""
    items: dict[str, dict] = {}

    # Splitting on the exact class string isolates each child block. Related
    # classes (flip-entry-info, flip-entries) don't match because the trailing
    # quote is part of the delimiter.
    for chunk in html.split('class="flip-entry"')[1:]:
        m_id = _EMBED_ID_RE.search(chunk)
        m_name = _EMBED_NAME_RE.search(chunk)
        if not m_id or not m_name:
            continue
        fid = m_id.group(1)
        if fid == folder_id or fid in items:
            continue
        name = _html_unescape(m_name.group(1)).strip()
        if not name:
            continue

        m_href = _EMBED_HREF_RE.search(chunk)
        href = m_href.group(1) if m_href else ""
        if "/drive/folders/" in href:
            mime = FOLDER_MIME
        elif "docs.google.com" in href:
            mime = GOOGLE_NATIVE_PREFIX  # Docs/Sheets/Slides: skipped downstream.
        else:
            mime = ""  # Regular binary file; media detection is by extension.

        items[fid] = {"id": fid, "name": name, "mime": mime}

    return list(items.values())


def list_public_folder_children(folder_id: str,
                                session: requests.Session | None = None) -> list[dict]:
    session = session or _session()
    r = session.get(_EMBED_URL.format(fid=folder_id), timeout=90)
    if r.status_code == 404:
        raise FileNotFoundError(
            "Folder not found or not shared as 'Anyone with the link'."
        )
    r.raise_for_status()
    return _parse_folder_entries(r.text, folder_id)


def _walk_public_folder(folder_id: str, session: requests.Session,
                        path_prefix: str = "") -> list[dict]:
    out: list[dict] = []
    try:
        kids = list_public_folder_children(folder_id, session)
    except Exception as e:
        print(f"   List error at {path_prefix or '/'}: {e}")
        return out

    for child in kids:
        name = child.get("name", "") or ""
        rel = f"{path_prefix}/{name}".lstrip("/")
        mime = child.get("mimeType") or child.get("mime") or ""
        fid = child["id"]
        if mime == FOLDER_MIME:
            out.extend(_walk_public_folder(fid, session, rel))
        else:
            if mime.startswith(GOOGLE_NATIVE_PREFIX):
                continue
            out.append({
                "id": fid,
                "name": name,
                "path": rel,
                "size": int(child.get("size", 0) or 0),
                "mime": mime,
                "web_url": f"https://drive.google.com/file/d/{fid}/view",
            })
    return out


def extract_media_from_folder(folder_id: str, media_only: bool = True) -> list[dict]:
    session = _session()
    items = _walk_public_folder(folder_id, session)
    if media_only:
        items = [i for i in items if looks_like_media(i["name"])]
    return items


def public_folder_name(folder_id: str, session: requests.Session | None = None) -> str:
    session = session or _session()
    r = session.get(_EMBED_URL.format(fid=folder_id), timeout=60)
    if r.status_code != 200:
        return ""
    m = _TITLE_RE.search(r.text)
    return m.group(1).strip() if m else ""


# ---------------------------------------------------------------------------
# Async streaming
# ---------------------------------------------------------------------------

OnProgress = Callable[..., None]


async def iter_remote_bytes(file_id: str,
                            chunk_size: int = 1024 * 1024) -> AsyncIterator[bytes]:
    """Stream bytes from a public Drive file (no auth)."""
    url, session = await asyncio.to_thread(build_public_download_url, file_id)
    cookies = dict(session.cookies)
    headers = dict(session.headers)

    async with httpx.AsyncClient(
        timeout=httpx.Timeout(300.0, connect=30.0),
        follow_redirects=True,
        cookies=cookies,
        headers=headers,
    ) as client:
        async with client.stream("GET", url) as r:
            if r.status_code != 200:
                try:
                    body = await r.aread()
                    msg = body[:200].decode("utf-8", "ignore")
                except Exception:
                    msg = ""
                raise RuntimeError(f"HTTP {r.status_code} downloading file. {msg}")
            async for chunk in r.aiter_bytes(chunk_size):
                if chunk:
                    yield chunk


# ---------------------------------------------------------------------------
# Folder crawl (manifest only)
# ---------------------------------------------------------------------------

async def crawl_folder(folder_url: str,
                       on_progress: Optional[OnProgress] = None,
                       media_only: bool = True,
                       should_cancel: Optional[Callable[[], bool]] = None) -> dict:
    info = parse_drive_url(folder_url)
    if info["kind"] != "folder" or not info["id"]:
        raise RuntimeError("URL doesn't look like a Google Drive folder share.")
    folder_id = info["id"]

    if on_progress:
        on_progress(0, 0, "Resolving folder name...", "info")
    folder_name = await asyncio.to_thread(public_folder_name, folder_id)

    if on_progress:
        on_progress(0, 0, "Crawling public Google Drive folder...", "info")
    items = await asyncio.to_thread(extract_media_from_folder, folder_id, media_only)

    if on_progress:
        on_progress(
            0, len(items),
            f"Found {len(items)} media file(s) in '{folder_name or 'folder'}'.",
            "info",
        )
    if should_cancel and should_cancel() and on_progress:
        on_progress(0, len(items), "Cancelled by user.", "cancelled")

    return {"folder_name": folder_name, "items": items}


# ---------------------------------------------------------------------------
# Streaming batch helpers
# ---------------------------------------------------------------------------

async def resolve_single_stream(url: str) -> tuple[str, AsyncIterator[bytes]]:
    info = parse_drive_url(url)
    if info["kind"] == "unknown" or not info["id"]:
        raise RuntimeError("URL doesn't look like a Google Drive file.")

    meta = await asyncio.to_thread(public_file_meta, info["id"])
    mime = meta.get("mime") or ""
    if mime.startswith(GOOGLE_NATIVE_PREFIX):
        raise RuntimeError(
            "This is a Google-native document (Docs/Sheets/Slides). "
            "Export it manually or use a media file instead."
        )
    name = _safe_filename(meta.get("name") or "gdrive_file")
    return name, iter_remote_bytes(info["id"])


async def iter_batch_files(links: list[str],
                           on_progress: Optional[OnProgress] = None,
                           should_cancel: Optional[Callable[[], bool]] = None,
                           batch_offset: int = 0,
                           overall_total: int | None = None):
    """
    Yield (filename, async-byte-iterator, result-dict) per link.
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

        info = parse_drive_url(url)
        if info["kind"] == "unknown" or not info["id"]:
            if on_progress:
                on_progress(idx_overall, total, f"Unrecognised URL: {url}", "failed")
            yield None, None, {"url": url, "status": "failed", "file": None}
            continue

        try:
            meta = await asyncio.to_thread(public_file_meta, info["id"])
        except Exception as e:
            if on_progress:
                on_progress(idx_overall, total, f"Resolve failed: {e}", "failed")
            yield None, None, {"url": url, "status": "failed", "file": None}
            continue

        name = _safe_filename(meta.get("name") or f"file_{idx_overall}")
        mime = meta.get("mime") or ""

        if mime.startswith(GOOGLE_NATIVE_PREFIX):
            if on_progress:
                on_progress(idx_overall, total, f"Skipped Google-native doc: {name}",
                            "skipped", name)
            yield None, None, {"url": url, "status": "skipped", "file": name}
            continue

        if on_progress:
            on_progress(idx_overall, total, f"Streaming: {name}", "processing")

        result = {"url": url, "status": "pending", "file": name}

        async def _stream_with_status(fid=info["id"], res=result,
                                      idx=idx_overall, nm=name):
            try:
                async for chunk in iter_remote_bytes(fid):
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
