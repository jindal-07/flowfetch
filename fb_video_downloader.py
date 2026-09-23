import asyncio
import io
import os
import random
import re
import time
from typing import AsyncIterator
from urllib.parse import parse_qs, urlparse

import httpx
import pandas as pd
import requests
from playwright.async_api import async_playwright


# ---------------------------------------------------------------------------
# URL / Post-ID normalization
# ---------------------------------------------------------------------------

_VIDEO_ID_PATTERNS = [
    # /videos/<id>/  or /videos/<slug>/<id>/
    re.compile(r"/videos/(?:[^/]+/)?(\d{6,})"),
    # /reel/<id>
    re.compile(r"/reel/(\d{6,})"),
    # /watch/?v=<id>   or  ?v=<id>
    re.compile(r"[?&]v=(\d{6,})"),
    # /share/v/<token>  (token, not numeric) -> handled separately
]


def _resolve_short_url(url: str) -> str:
    """Follow fb.watch / share short links to the canonical URL."""
    try:
        r = requests.head(url, allow_redirects=True, timeout=15, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        })
        return r.url or url
    except Exception:
        return url


def extract_post_info(url: str) -> dict:
    """
    Parse any Facebook video URL into a normalized form.

    Returns: {"id": str|None, "kind": "reel"|"video"|"unknown", "canonical": str}
    """
    url = url.strip()

    # Resolve short URLs first (fb.watch, share links).
    parsed = urlparse(url)
    host = (parsed.netloc or "").lower()
    if host in ("fb.watch", "www.fb.watch") or "/share/" in parsed.path:
        url = _resolve_short_url(url)
        parsed = urlparse(url)

    path = parsed.path or ""

    # Reel detection.
    m = re.search(r"/reel/(\d{6,})", path)
    if m:
        vid = m.group(1)
        return {
            "id": vid,
            "kind": "reel",
            "canonical": f"https://www.facebook.com/reel/{vid}",
        }

    # /videos/.../<id>
    m = re.search(r"/videos/(?:[^/]+/)?(\d{6,})", path)
    if m:
        vid = m.group(1)
        return {
            "id": vid,
            "kind": "video",
            "canonical": f"https://www.facebook.com/watch/?v={vid}",
        }

    # /watch?v=<id>
    qs = parse_qs(parsed.query or "")
    if "v" in qs and qs["v"]:
        vid = qs["v"][0]
        if vid.isdigit():
            return {
                "id": vid,
                "kind": "video",
                "canonical": f"https://www.facebook.com/watch/?v={vid}",
            }

    return {"id": None, "kind": "unknown", "canonical": url}


# ---------------------------------------------------------------------------
# Playwright helpers
# ---------------------------------------------------------------------------

DESKTOP_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)


async def _new_browser(p):
    return await p.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"],
    )


async def _new_page(browser):
    context = await browser.new_context(
        viewport={"width": 1366, "height": 768},
        user_agent=DESKTOP_UA,
        locale="en-US",
    )
    # Auto-dismiss the cookie / login popups by blocking common login dialogs.
    await context.add_init_script(
        """() => {
            try { localStorage.setItem('locale', 'en_US'); } catch(e) {}
        }"""
    )
    page = await context.new_page()
    return context, page


def _is_video_response(resp):
    url = resp.url.lower()
    if url.startswith("blob:") or url.startswith("data:"):
        return False
    if ".mp4" in url:
        return True
    try:
        ctype = (resp.headers or {}).get("content-type", "").lower()
    except Exception:
        ctype = ""
    return ctype.startswith("video/")


# Patterns FB embeds in <script> JSON. Ordered: HD first, then SD.
_HTML_VIDEO_PATTERNS = [
    r'"browser_native_hd_url":"([^"]+\.mp4[^"]*)"',
    r'"playable_url_quality_hd":"([^"]+\.mp4[^"]*)"',
    r'"hd_src_no_ratelimit":"([^"]+\.mp4[^"]*)"',
    r'"hd_src":"([^"]+\.mp4[^"]*)"',
    r'"browser_native_sd_url":"([^"]+\.mp4[^"]*)"',
    r'"playable_url":"([^"]+\.mp4[^"]*)"',
    r'"sd_src_no_ratelimit":"([^"]+\.mp4[^"]*)"',
    r'"sd_src":"([^"]+\.mp4[^"]*)"',
]


def _decode_url(s: str) -> str:
    """Decode FB's JSON-escaped URLs (\\/, \\u002F, \\u0026 etc.)."""
    import json
    try:
        return json.loads(f'"{s}"')
    except Exception:
        return s.replace("\\/", "/").replace("\\u0026", "&")


def _scan_html_for_video(html: str) -> str | None:
    for pattern in _HTML_VIDEO_PATTERNS:
        m = re.search(pattern, html)
        if m:
            return _decode_url(m.group(1))
    return None


async def extract_video_src(page, url: str) -> str | None:
    """
    Navigate to a canonical Facebook video URL and return a direct .mp4 URL.

    Strategy:
      1. Normalize the URL via post-ID extraction (/watch/?v=ID or /reel/ID).
      2. Load with a desktop UA; scrape FB's embedded JSON for hd/sd mp4 URLs
         (fastest path, no need to start playback).
      3. Fall back to sniffing the network for the actual .mp4 request that
         fires once the <video> starts buffering.
    """
    info = extract_post_info(url)
    target = info["canonical"]

    captured: list[str] = []

    def on_response(resp):
        try:
            if _is_video_response(resp):
                captured.append(resp.url)
        except Exception:
            pass

    page.on("response", on_response)

    try:
        try:
            await page.goto(target, wait_until="domcontentloaded", timeout=45000)
        except Exception as nav_err:
            print(f"   Navigation warning: {nav_err}")

        # Quick first pass: scrape the embedded JSON.
        try:
            html = await page.content()
            found = _scan_html_for_video(html)
            if found:
                return found
        except Exception:
            pass

        # Wait a bit for late scripts / lazy hydration, then re-scan.
        await page.wait_for_timeout(1500)
        try:
            html = await page.content()
            found = _scan_html_for_video(html)
            if found:
                return found
        except Exception:
            pass

        # Trigger playback so the mp4 segment request fires.
        try:
            await page.evaluate(
                """() => {
                    const v = document.querySelector('video');
                    if (v) { v.muted = true; v.play().catch(() => {}); }
                }"""
            )
        except Exception:
            pass

        # Network sniff fallback.
        for _ in range(20):  # up to 10 seconds
            if captured:
                return captured[0]
            await page.wait_for_timeout(500)

        # Last-ditch: if <video>.src is a real URL, use it.
        try:
            direct = await page.evaluate(
                "() => { const v=document.querySelector('video'); return v ? v.src : null; }"
            )
            if direct and not direct.startswith(("blob:", "data:")):
                return direct
        except Exception:
            pass

        return None
    finally:
        try:
            page.remove_listener("response", on_response)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# HTTP download
# ---------------------------------------------------------------------------

def download_file(src_url: str, filename: str) -> bool:
    if not src_url or src_url.startswith(("blob:", "data:")):
        print(f"   Cannot download non-HTTP URL: {src_url}")
        return False

    headers = {
        "User-Agent": DESKTOP_UA,
        "Referer": "https://www.facebook.com/",
        "Accept": "*/*",
        "Accept-Language": "en-US,en;q=0.9",
        "Range": "bytes=0-",
    }

    try:
        response = requests.get(src_url, stream=True, timeout=60, headers=headers)
        if response.status_code in (200, 206):
            with open(filename, "wb") as f:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        f.write(chunk)
            return True
        print(f"   HTTP {response.status_code} for video URL")
    except Exception as e:
        print(f"   Network error: {e}")
    return False


# ---------------------------------------------------------------------------
# Async streaming (memory-bounded, no disk writes)
# ---------------------------------------------------------------------------

_FB_HEADERS = {
    "User-Agent": DESKTOP_UA,
    "Referer": "https://www.facebook.com/",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Range": "bytes=0-",
}


async def iter_remote_bytes(src_url: str, chunk_size: int = 1024 * 1024) -> AsyncIterator[bytes]:
    """Yield bytes from a Facebook CDN URL without writing to disk."""
    if not src_url or src_url.startswith(("blob:", "data:")):
        raise RuntimeError(f"Cannot stream non-HTTP URL: {src_url}")
    async with httpx.AsyncClient(timeout=httpx.Timeout(120.0, connect=30.0), follow_redirects=True) as client:
        async with client.stream("GET", src_url, headers=_FB_HEADERS) as r:
            if r.status_code not in (200, 206):
                raise RuntimeError(f"HTTP {r.status_code} for {src_url[:80]}")
            async for chunk in r.aiter_bytes(chunk_size):
                if chunk:
                    yield chunk


def filename_for(url: str, index: int = 1) -> str:
    """Public helper so app.py can suggest filenames without disk side-effects."""
    return _filename_for(url, index)


async def resolve_video_src(url: str) -> str | None:
    """Public single-shot helper: launch Playwright, extract the .mp4 URL, close."""
    async with async_playwright() as p:
        browser = await _new_browser(p)
        context, page = await _new_page(browser)
        try:
            return await extract_video_src(page, url)
        finally:
            await context.close()
            await browser.close()


# Markers Facebook shows when a video is gone or restricted (vs. a public one).
_UNAVAILABLE_MARKERS = (
    "content isn't available",
    "content isn’t available",
    "isn't available right now",
    "isn’t available right now",
    "the link you followed may be broken",
    "this page isn't available",
    "this page isn’t available",
    "this content isn't available",
    "video unavailable",
)


def classify_inaccessible(html: str, final_url: str = "") -> str:
    """
    Explain why a Facebook video couldn't be resolved. Called only after
    extraction already failed, so it always returns a non-empty reason.
    """
    low = (html or "").lower()
    if any(m in low for m in _UNAVAILABLE_MARKERS):
        return ("This Facebook video isn't available — it may have been removed, "
                "set to private, or is region-restricted.")
    if 'name="email"' in low and ("log in" in low or "login" in (final_url or "").lower()):
        return ("This Facebook video isn't publicly accessible — it requires signing in. "
                "Only public video/reel links can be downloaded.")
    return ("Couldn't find a downloadable video at this link. It may be private or "
            "login-gated, or Facebook rate-limited this request — try again shortly.")


async def resolve_video_with_reason(url: str) -> tuple[str | None, str]:
    """
    Accessibility-aware resolve: returns (src, "") when the video is reachable,
    or (None, reason) with a human-readable explanation when it is not. Reuses a
    single browser so the check and the resolve are one navigation, not two.
    """
    async with async_playwright() as p:
        browser = await _new_browser(p)
        context, page = await _new_page(browser)
        try:
            src = await extract_video_src(page, url)
            if src:
                return src, ""
            try:
                html = await page.content()
            except Exception:
                html = ""
            return None, classify_inaccessible(html, page.url)
        finally:
            await context.close()
            await browser.close()


# ---------------------------------------------------------------------------
# CSV parsing
# ---------------------------------------------------------------------------

def _parse_links_from_csv(csv_path):
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"'{csv_path}' not found.")
    with open(csv_path, "rb") as f:
        df = read_csv_flexible(f.read())
    if "Links" in df.columns:
        links = df["Links"].dropna().tolist()
    else:
        links = df[df.columns[0]].dropna().tolist()
    return [str(u).strip() for u in links if str(u).strip().startswith("http")]


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
    df = read_csv_flexible(content)
    if "Links" in df.columns:
        links = df["Links"].dropna().tolist()
    else:
        links = df[df.columns[0]].dropna().tolist()
    return [str(u).strip() for u in links if str(u).strip().startswith("http")]


# ---------------------------------------------------------------------------
# File-naming
# ---------------------------------------------------------------------------

def _filename_for(url: str, index: int) -> str:
    """Prefer using the FB post ID so files are recognizable."""
    info = extract_post_info(url)
    if info["id"]:
        prefix = "reel" if info["kind"] == "reel" else "video"
        return f"fb_{prefix}_{info['id']}.mp4"
    return f"fb_video_{index}.mp4"


# ---------------------------------------------------------------------------
# Job runners
# ---------------------------------------------------------------------------

async def process_links(links, output_folder="downloads", on_progress=None, should_cancel=None):
    """
    Process a list of Facebook video URLs.

    on_progress(index, total, message, status, filename=None) is called as work happens.
    status: 'skipped' | 'success' | 'failed' | 'processing' | 'info' | 'done'
    """
    os.makedirs(output_folder, exist_ok=True)
    total = len(links)

    if on_progress:
        on_progress(0, total, f"Found {total} links to process.", "info")

    results = []

    async with async_playwright() as p:
        browser = await _new_browser(p)
        context, page = await _new_page(browser)

        for index, url in enumerate(links, start=1):
            if should_cancel and should_cancel():
                if on_progress:
                    on_progress(index - 1, total, "Cancelled by user.", "cancelled")
                break

            basename = _filename_for(url, index)
            output_filename = os.path.join(output_folder, basename)

            if os.path.exists(output_filename):
                if on_progress:
                    on_progress(index, total, f"Already downloaded: {basename}", "skipped", basename)
                results.append({"url": url, "status": "skipped", "file": basename})
                continue

            info = extract_post_info(url)
            label = f"{info['kind']} {info['id']}" if info["id"] else url
            if on_progress:
                on_progress(index, total, f"Processing: {label}", "processing")

            video_src = await extract_video_src(page, url)

            if not video_src:
                if on_progress:
                    on_progress(index, total, "Could not find video. Skipping.", "failed")
                results.append({"url": url, "status": "failed", "file": None})
                continue

            if on_progress:
                on_progress(index, total, "Downloading video...", "processing")

            success = download_file(video_src, output_filename)

            if success:
                if on_progress:
                    on_progress(index, total, f"Saved: {basename}", "success", basename)
                results.append({"url": url, "status": "success", "file": basename})
            else:
                if on_progress:
                    on_progress(index, total, "Failed to write file.", "failed")
                results.append({"url": url, "status": "failed", "file": None})

            delay = random.uniform(3.0, 7.0)
            if on_progress:
                on_progress(index, total, f"Waiting {delay:.1f}s (rate-limit protection)...", "info")
            await asyncio.sleep(delay)

            if index % 50 == 0:
                if on_progress:
                    on_progress(index, total, "Restarting browser to free memory...", "info")
                await context.close()
                await browser.close()
                browser = await _new_browser(p)
                context, page = await _new_page(browser)

        await context.close()
        await browser.close()

    if on_progress:
        on_progress(total, total, "All downloads complete!", "done")

    return results


async def process_single_url(url, output_folder="downloads", on_progress=None, should_cancel=None):
    """Download a single Facebook video URL."""
    url = url.strip()
    if not url.startswith("http"):
        raise ValueError("Invalid URL. Must start with http.")

    os.makedirs(output_folder, exist_ok=True)

    info = extract_post_info(url)
    if info["id"]:
        prefix = "reel" if info["kind"] == "reel" else "video"
        basename = f"fb_{prefix}_{info['id']}.mp4"
    else:
        basename = f"fb_video_single_{int(time.time())}.mp4"
    output_filename = os.path.join(output_folder, basename)

    label = f"{info['kind']} {info['id']}" if info["id"] else url
    if on_progress:
        on_progress(1, 1, f"Processing: {label}", "processing")

    async with async_playwright() as p:
        browser = await _new_browser(p)
        context, page = await _new_page(browser)
        try:
            video_src = await extract_video_src(page, url)
        finally:
            await context.close()
            await browser.close()

    if not video_src:
        if on_progress:
            on_progress(1, 1, "Could not find video.", "failed")
        return {"url": url, "status": "failed", "file": None}

    if on_progress:
        on_progress(1, 1, "Downloading video...", "processing")

    success = download_file(video_src, output_filename)

    if success:
        if on_progress:
            on_progress(1, 1, f"Saved: {basename}", "success", basename)
            on_progress(1, 1, "Download complete!", "done")
        return {"url": url, "status": "success", "file": basename}

    if on_progress:
        on_progress(1, 1, "Failed to write file.", "failed")
    return {"url": url, "status": "failed", "file": None}


async def process_large_csv(csv_path, output_folder="downloads", on_progress=None):
    links = _parse_links_from_csv(csv_path)
    return await process_links(links, output_folder, on_progress)


# ---------------------------------------------------------------------------
# Batch streaming: yield (filename, byte_iter, result_dict) per link
# ---------------------------------------------------------------------------

async def iter_batch_files(
    links: list[str],
    on_progress=None,
    should_cancel=None,
    batch_offset: int = 0,
    overall_total: int | None = None,
):
    """
    Async generator that, for each Facebook URL in `links`, yields a tuple of:
        (filename, byte_iterator, info_dict)

    where byte_iterator is an async iterator over the video bytes (suitable for
    feeding into stream-zip).  info_dict has the final 'status' and 'url' once
    the consumer has finished iterating bytes for that entry. The generator
    skips failed links (still emits an "info_dict" with status='failed' BEFORE
    yielding the file tuple, via on_progress).

    Re-uses a single Playwright browser across all links in the batch.
    """
    total = overall_total or len(links)
    if on_progress:
        on_progress(batch_offset, total, f"Starting batch of {len(links)} link(s)...", "info")

    async with async_playwright() as p:
        browser = await _new_browser(p)
        context, page = await _new_page(browser)

        try:
            for local_idx, url in enumerate(links, start=1):
                idx_overall = batch_offset + local_idx
                if should_cancel and should_cancel():
                    if on_progress:
                        on_progress(idx_overall - 1, total, "Cancelled by user.", "cancelled")
                    return

                basename = _filename_for(url, idx_overall)
                info = extract_post_info(url)
                label = f"{info['kind']} {info['id']}" if info["id"] else url
                if on_progress:
                    on_progress(idx_overall, total, f"Resolving: {label}", "processing")

                try:
                    video_src = await extract_video_src(page, url)
                except Exception as e:
                    if on_progress:
                        on_progress(idx_overall, total, f"Resolve error: {e}", "failed")
                    yield None, None, {"url": url, "status": "failed", "file": None}
                    continue

                if not video_src:
                    try:
                        reason = classify_inaccessible(await page.content(), page.url)
                    except Exception:
                        reason = "Could not find a downloadable video (link may be private or removed)."
                    if on_progress:
                        on_progress(idx_overall, total, reason, "failed")
                    yield None, None, {"url": url, "status": "failed", "file": None}
                    continue

                if on_progress:
                    on_progress(idx_overall, total, f"Streaming: {basename}", "processing")

                # Wrap so we report success only after iteration completes.
                result = {"url": url, "status": "pending", "file": basename}

                async def _stream_with_status(src=video_src, res=result, idx=idx_overall, name=basename):
                    try:
                        async for chunk in iter_remote_bytes(src):
                            yield chunk
                        res["status"] = "success"
                        if on_progress:
                            on_progress(idx, total, f"Saved: {name}", "success", name)
                    except Exception as e:
                        res["status"] = "failed"
                        res["file"] = None
                        if on_progress:
                            on_progress(idx, total, f"Stream failed: {e}", "failed")
                        # Re-raise so the ZIP-stream consumer knows this entry broke.
                        raise

                yield basename, _stream_with_status(), result

                # Small delay between videos to be polite to FB CDN.
                if local_idx < len(links):
                    delay = random.uniform(1.5, 3.5)
                    await asyncio.sleep(delay)
        finally:
            try:
                await context.close()
            except Exception:
                pass
            try:
                await browser.close()
            except Exception:
                pass


if __name__ == "__main__":
    CSV_FILE = "fb_exports.csv"
    asyncio.run(process_large_csv(CSV_FILE))
