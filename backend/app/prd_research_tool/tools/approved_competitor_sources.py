from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus, urlparse


APPROVED_COMPETITOR_SOURCE_PATH = (
    Path(__file__).resolve().parents[3]
    / "source_data"
    / "competitor_sources"
    / "approved_competitor_sources.json"
)

GENERIC_CATEGORY_TERMS = {
    "and",
    "commercial",
    "fixture",
    "fixtures",
    "led",
    "light",
    "lighting",
    "residential",
}


def normalize_domain(value: Any) -> str:
    """Return a lower-case registrable-ish host for matching scraped URLs."""
    text = str(value or "").strip()
    if not text:
        return ""
    parsed = urlparse(text if "://" in text else f"//{text}")
    host = parsed.netloc or parsed.path
    host = host.lower().strip().strip("/")
    return re.sub(r"^www\.", "", host)


def domain_matches(candidate: Any, approved_domain: Any) -> bool:
    """Return whether a scraped host belongs to an approved source domain."""
    candidate_domain = normalize_domain(candidate)
    approved = normalize_domain(approved_domain)
    return bool(candidate_domain and approved and (candidate_domain == approved or candidate_domain.endswith(f".{approved}")))


def category_terms(category_slug: Any) -> list[str]:
    """Return useful category terms for matching approved-source scope notes."""
    text = re.sub(r"[^a-z0-9]+", " ", str(category_slug or "").lower())
    return [term for term in text.split() if len(term) >= 3 and term not in GENERIC_CATEGORY_TERMS]


@lru_cache(maxsize=1)
def load_approved_competitor_sources() -> list[dict[str, Any]]:
    """Load the approved competitor-source registry generated from the workbook."""
    path = APPROVED_COMPETITOR_SOURCE_PATH
    if not path.exists():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    sources = payload.get("sources") if isinstance(payload, dict) else payload
    if not isinstance(sources, list):
        return []
    output: list[dict[str, Any]] = []
    seen: set[str] = set()
    for source in sources:
        if not isinstance(source, dict):
            continue
        domain = normalize_domain(source.get("domain") or source.get("website"))
        if not domain or domain in seen:
            continue
        seen.add(domain)
        clean = dict(source)
        clean["domain"] = domain
        output.append(clean)
    return output


def approved_competitor_source_for_domain(domain: Any) -> dict[str, Any] | None:
    """Return the approved source matching a scraped URL/domain, if any."""
    for source in load_approved_competitor_sources():
        if domain_matches(domain, source.get("domain")):
            return source
    return None


def approved_domain_set() -> set[str]:
    """Return all approved source domains."""
    return {source["domain"] for source in load_approved_competitor_sources() if source.get("domain")}


def approved_source_scope_score(source: dict[str, Any], category_slug: Any) -> tuple[int, int, str]:
    """Score how relevant an approved source appears for the active category."""
    terms = category_terms(category_slug)
    scope_text = " ".join(
        str(source.get(key) or "")
        for key in ["focus", "scope", "product_subcategories_in_scope", "strengths", "weaknesses_notes"]
    ).lower()
    tier_text = str(source.get("tier") or "").lower()
    rank_text = str(source.get("tier_rank") or "")
    try:
        tier_rank = int(float(rank_text))
    except ValueError:
        tier_rank = 999

    term_hits = sum(1 for term in terms if term in scope_text)
    broad_lighting = 1 if any(token in scope_text for token in ["all categories", "large lighting", "commercial lighting", "lighting distributor"]) else 0
    tier_score = 3 if "tier 1" in tier_text else 2 if "tier 2" in tier_text else 1 if "tier 3" in tier_text else 0
    return (term_hits * 10 + broad_lighting + tier_score, -tier_rank, str(source.get("competitor") or source.get("domain") or ""))


def approved_sources_for_category(category_slug: Any, limit: int | None = None) -> list[dict[str, Any]]:
    """Return approved sources prioritized for the active category."""
    sources = sorted(
        load_approved_competitor_sources(),
        key=lambda source: approved_source_scope_score(source, category_slug),
        reverse=True,
    )
    return sources[:limit] if limit else sources


def source_search_link(source: dict[str, Any], category_slug: Any, ideation_name: Any = "") -> str:
    """Build a search URL for discovery without presenting it as a verified PDP."""
    domain = normalize_domain(source.get("domain") or source.get("website"))
    terms = " ".join(category_terms(category_slug))
    if ideation_name:
        terms = f"{terms} {ideation_name}".strip()
    query = f"site:{domain} {terms}".strip()
    return f"https://www.google.com/search?q={quote_plus(query)}"
