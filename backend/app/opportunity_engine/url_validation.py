from __future__ import annotations

import re
import urllib.error
import urllib.request
from typing import Any


URL_RE = re.compile(r"https?://[^\s<>\"]+", flags=re.IGNORECASE)
REQUEST_TIMEOUT_SECONDS = 12
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
)


def extract_urls(value: str | None) -> list[str]:
    if not value:
        return []
    urls: list[str] = []
    for raw in URL_RE.findall(str(value)):
        url = raw.rstrip(").,;]")
        if url not in urls:
            urls.append(url)
    return urls


def classify_url(url: str) -> str:
    lowered = url.lower()
    if "amazon.com/s?" in lowered or "amazon.com/s?k=" in lowered:
        return "search page"
    if "/sb2/" in lowered or "/b/" in lowered or "/cat/" in lowered:
        return "category/filter page"
    return "product/listing page"


def validate_url(url: str) -> dict[str, Any]:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
            status = int(response.status)
            final_url = response.geturl()
    except urllib.error.HTTPError as exc:
        status = int(exc.code)
        final_url = exc.geturl() or url
    except Exception as exc:
        return {
            "url": url,
            "status": None,
            "state": "unverified",
            "note": f"unverified: {type(exc).__name__}",
            "page_type": classify_url(url),
        }

    if 200 <= status < 400:
        state = "verified"
        note = f"verified live URL ({status})"
    elif status in {403, 429}:
        state = "blocked"
        note = f"retailer/CDN blocked automated verification ({status}); manual browser check required"
    elif status in {404, 410}:
        state = "invalid"
        note = f"invalid/dead URL ({status}); replace before using as evidence"
    else:
        state = "unverified"
        note = f"unverified HTTP status ({status})"

    return {
        "url": url,
        "status": status,
        "state": state,
        "note": note,
        "final_url": final_url,
        "page_type": classify_url(final_url),
    }


def validate_urls(urls: list[str], cache: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for url in urls:
        if url not in cache:
            cache[url] = validate_url(url)
        results.append(cache[url])
    return results


def format_validation_notes(results: list[dict[str, Any]]) -> str | None:
    if not results:
        return None
    parts: list[str] = []
    for result in results:
        page_type = result.get("page_type") or "page"
        parts.append(f"{result.get('url')} -> {result.get('note')} ({page_type})")
    return "\n".join(parts)
