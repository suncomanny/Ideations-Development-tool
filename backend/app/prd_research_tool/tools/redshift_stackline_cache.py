"""Load Redshift-derived Stackline cache files for PRD research packets.

The double-click workflow cannot call Codex MCP tools directly, so Redshift
extracts are stored as JSON snapshots under backend/source_data. Packet
generation reads these snapshots before falling back to local Stackline CSVs.
"""

from __future__ import annotations

import json
import math
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any


TOOLS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = TOOLS_DIR.parents[3]
DEFAULT_CACHE_DIR = PROJECT_ROOT / "backend" / "source_data" / "redshift_stackline_cache"
CACHE_MAX_AGE_DAYS = 30
CHANNEL_PRIORITY = {
    "amazon": 4,
    "home_depot": 3,
    "walmart": 2,
    "lowes": 2,
    "all_retailers": 1,
}


def clean_number(value: Any, digits: int = 2) -> float | None:
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return round(number, digits)


def pct_delta(current: Any, prior: Any) -> float | None:
    current_number = clean_number(current, 8)
    prior_number = clean_number(prior, 8)
    if current_number is None or prior_number in (None, 0):
        return None
    return clean_number((current_number / prior_number - 1) * 100)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).lower()
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"[^a-z0-9\s]+", " ", text)
    parts = []
    for part in text.split():
        if len(part) > 4 and part.endswith("ies"):
            part = part[:-3] + "y"
        elif len(part) > 4 and part.endswith("s"):
            part = part[:-1]
        parts.append(part)
    return " ".join(parts)


def text_tokens(value: Any) -> set[str]:
    return {token for token in normalize_text(value).split() if token}


def match_score(subcategory: str, cache: dict[str, Any]) -> float:
    query_tokens = text_tokens(subcategory)
    if not query_tokens:
        return 0.0

    candidates = [
        cache.get("subcategory"),
        (cache.get("segment") or {}).get("segment_name"),
        cache.get("segment_name"),
    ]
    best = 0.0
    for candidate in candidates:
        normalized = normalize_text(candidate)
        if not normalized:
            continue
        if normalized == normalize_text(subcategory):
            best = max(best, 1.0)
            continue
        candidate_tokens = text_tokens(candidate)
        if not candidate_tokens:
            continue
        overlap = len(query_tokens & candidate_tokens)
        if overlap:
            coverage = overlap / len(query_tokens)
            precision = overlap / len(candidate_tokens)
            score = max(coverage, (coverage + precision) / 2)
            if query_tokens.issubset(candidate_tokens):
                score += 0.25
            best = max(best, min(score, 0.99))
    return best


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def cache_created_date(cache: dict[str, Any], path: Path) -> date:
    created_at = cache.get("created_at")
    if isinstance(created_at, str) and len(created_at) >= 10:
        try:
            return date.fromisoformat(created_at[:10])
        except ValueError:
            pass
    return datetime.fromtimestamp(path.stat().st_mtime).date()


def find_cache_file(subcategory: str, cache_dir: Path = DEFAULT_CACHE_DIR) -> tuple[Path, dict[str, Any], float]:
    if not cache_dir.exists():
        raise FileNotFoundError(f"No Redshift Stackline cache folder exists at {cache_dir}.")

    matches: list[tuple[float, date, Path, dict[str, Any]]] = []
    for path in cache_dir.glob("*_stackline_redshift_*.json"):
        try:
            cache = load_json(path)
        except Exception:
            continue
        score = match_score(subcategory, cache)
        if score <= 0:
            continue
        matches.append((score, cache_created_date(cache, path), path, cache))

    if not matches:
        raise FileNotFoundError(
            f"No Redshift Stackline cache could be matched to subcategory '{subcategory}'."
        )

    score, _created, path, cache = max(
        matches,
        key=lambda item: (item[0], item[1], item[2].stat().st_mtime),
    )
    return path, cache, score


def build_reference_family(reference_sku: str | None, products: list[dict[str, Any]]) -> dict[str, Any]:
    if not reference_sku or "TBD" in str(reference_sku).upper():
        return {
            "reference_sku": reference_sku,
            "reference_family": None,
            "found": False,
            "note": "No usable Sunco reference SKU was provided for this ideation row.",
        }

    family = str(reference_sku).upper()
    matches = [
        product
        for product in products
        if family in str(product.get("model_number") or "").upper()
        or family in str(product.get("retailer_sku") or "").upper()
    ]
    return {
        "reference_sku": reference_sku,
        "reference_family": family,
        "found": bool(matches),
        "main": {
            "variants": matches[:10],
            "variant_count": len(matches),
        }
        if matches
        else None,
    }


def build_performance_context(
    cache: dict[str, Any],
    channel_key: str,
    channel: dict[str, Any],
    reference_sku: str | None,
    brand_name: str,
    cache_file: Path,
    cache_match_score: float,
) -> dict[str, Any]:
    segment = cache.get("segment") or {}
    periods = channel.get("periods") or {}
    main = periods.get("Main") or {}
    comparison = periods.get("Comparison") or {}
    products = channel.get("top_competitor_products") or []
    top_brands = channel.get("top_brands") or []
    leader_brand = top_brands[0] if top_brands else {}
    price_percentiles = channel.get("price_percentiles") or {}

    main_snapshot = {
        "retail_sales": clean_number(main.get("retail_sales")),
        "units_sold": clean_number(main.get("units_sold")),
        "avg_retail_price": clean_number(main.get("avg_retail_price")),
        "total_traffic": clean_number(main.get("total_traffic")),
        "conversion_rate_pct": clean_number(main.get("conversion_rate_pct")),
        "catalog_product_count": main.get("catalog_product_count"),
        "brand_count": main.get("brand_count"),
        "price_percentiles": price_percentiles,
    }
    market_momentum = {
        "retail_sales": pct_delta(main.get("retail_sales"), comparison.get("retail_sales")),
        "units_sold": pct_delta(main.get("units_sold"), comparison.get("units_sold")),
        "avg_retail_price": pct_delta(main.get("avg_retail_price"), comparison.get("avg_retail_price")),
        "traffic": pct_delta(main.get("total_traffic"), comparison.get("total_traffic")),
        "conversion_rate_pct": pct_delta(
            main.get("conversion_rate_pct"),
            comparison.get("conversion_rate_pct"),
        ),
    }

    sunco_position = {
        "brand": brand_name,
        "sales_share_pct": 0.0,
        "units_share_pct": 0.0,
        "product_count": 0,
        "avg_retail_price": None,
    }
    reference_family = build_reference_family(reference_sku, products)
    share_gap_to_leader = leader_brand.get("sales_share_pct")

    opportunity_signals = []
    if market_momentum.get("retail_sales") is not None and market_momentum["retail_sales"] >= 10:
        opportunity_signals.append("segment_sales_growth_above_10pct")
    if sunco_position["sales_share_pct"] < 5:
        opportunity_signals.append("sunco_share_below_5pct")
    if share_gap_to_leader is not None and share_gap_to_leader >= 5:
        opportunity_signals.append("meaningful_share_gap_to_segment_leader")
    if (
        market_momentum.get("traffic") is not None
        and market_momentum["traffic"] >= 10
        and market_momentum.get("conversion_rate_pct") is not None
        and market_momentum["conversion_rate_pct"] <= -10
    ):
        opportunity_signals.append("traffic_up_conversion_down")
    if reference_sku and "TBD" not in str(reference_sku).upper() and not reference_family.get("found"):
        opportunity_signals.append("reference_family_absent_from_stackline_segment")

    matched_bundle = {
        "source": "redshift_cache",
        "segment_name": segment.get("segment_name"),
        "segment_id": segment.get("segment_id"),
        "retailer_id": channel.get("retailer_id"),
        "retailer_scope": channel_key,
        "latest_week_id": cache.get("latest_week_id"),
        "cache_file": str(cache_file),
        "match_score": clean_number(cache_match_score, 4),
    }

    warnings = [
        "Stackline context loaded from Redshift cache before local CSV fallback.",
        f"Stackline context is scoped to {channel_key.replace('_', ' ')} rather than an all-retailer market view.",
    ]
    warnings.extend(cache.get("warnings") or [])
    warnings.extend(channel.get("warnings") or [])

    return {
        "segment": {
            "name": segment.get("segment_name"),
            "retailer_scope": channel_key,
            "matched_bundle": matched_bundle,
            "market_snapshot": main_snapshot,
            "market_momentum_pct": market_momentum,
        },
        "sunco_position": sunco_position,
        "reference_family": reference_family,
        "estimation_inputs": {
            "price_anchor": {
                "source": "segment_average",
                "avg_retail_price": main_snapshot.get("avg_retail_price"),
                "gap_vs_segment_avg_pct": 0.0,
            },
            "segment_leader": {
                "brand": leader_brand.get("brand"),
                "sales_share_pct": leader_brand.get("sales_share_pct"),
                "share_gap_vs_sunco_pct_points": clean_number(share_gap_to_leader),
            },
            "top_competitor_products": products[:10],
            "top_brands": top_brands[:10],
            "price_percentiles": price_percentiles,
        },
        "opportunity_signals": opportunity_signals,
        "warnings": warnings,
    }


def build_channel_comparison(channel_results: dict[str, dict[str, Any]]) -> dict[str, Any]:
    channel_summary: dict[str, Any] = {}
    for channel, result in channel_results.items():
        perf = result.get("performance_estimation_context") or {}
        segment = perf.get("segment") or {}
        snapshot = segment.get("market_snapshot") or {}
        momentum = segment.get("market_momentum_pct") or {}
        sunco_position = perf.get("sunco_position") or {}
        channel_summary[channel] = {
            "segment_name": segment.get("name"),
            "retailer_scope": segment.get("retailer_scope") or channel,
            "matched_bundle": segment.get("matched_bundle"),
            "retail_sales": snapshot.get("retail_sales"),
            "units_sold": snapshot.get("units_sold"),
            "avg_retail_price": snapshot.get("avg_retail_price"),
            "total_traffic": snapshot.get("total_traffic"),
            "conversion_rate_pct": snapshot.get("conversion_rate_pct"),
            "retail_sales_growth_pct": momentum.get("retail_sales"),
            "units_sold_growth_pct": momentum.get("units_sold"),
            "avg_retail_price_growth_pct": momentum.get("avg_retail_price"),
            "traffic_growth_pct": momentum.get("traffic"),
            "sunco_sales_share_pct": sunco_position.get("sales_share_pct"),
            "sunco_units_share_pct": sunco_position.get("units_share_pct"),
            "warnings": result.get("warnings", []),
        }

    amazon = channel_summary.get("amazon")
    home_depot = channel_summary.get("home_depot")
    comparisons: dict[str, Any] = {}
    if amazon and home_depot:
        comparisons["amazon_vs_home_depot"] = {
            "avg_retail_price_gap_pct": pct_delta(
                amazon.get("avg_retail_price"),
                home_depot.get("avg_retail_price"),
            ),
            "retail_sales_gap_pct": pct_delta(
                amazon.get("retail_sales"),
                home_depot.get("retail_sales"),
            ),
            "units_sold_gap_pct": pct_delta(
                amazon.get("units_sold"),
                home_depot.get("units_sold"),
            ),
            "sunco_sales_share_gap_pct_points": clean_number(
                (amazon.get("sunco_sales_share_pct") or 0)
                - (home_depot.get("sunco_sales_share_pct") or 0)
            ),
        }

    return {
        "available_channels": list(channel_summary.keys()),
        "channels": channel_summary,
        "comparisons": comparisons,
    }


def analyze_redshift_stackline_cache_for_subcategory(
    subcategory: str,
    reference_sku: str | None = None,
    cache_dir: Path = DEFAULT_CACHE_DIR,
    brand_name: str = "Sunco Lighting",
) -> dict[str, Any]:
    """Return a Stackline batch shaped like the local CSV analyzer output."""
    cache_file, cache, cache_match_score = find_cache_file(subcategory, cache_dir)
    created = cache_created_date(cache, cache_file)
    age_days = (date.today() - created).days

    cache_warnings = []
    if age_days > CACHE_MAX_AGE_DAYS:
        cache_warnings.append(
            f"Redshift Stackline cache is {age_days} days old; refresh the cache before final decision-making."
        )

    channel_results: dict[str, dict[str, Any]] = {}
    for channel_key, channel in (cache.get("channels") or {}).items():
        performance = build_performance_context(
            cache=cache,
            channel_key=channel_key,
            channel=channel,
            reference_sku=reference_sku,
            brand_name=brand_name,
            cache_file=cache_file,
            cache_match_score=cache_match_score,
        )
        warnings = list(dict.fromkeys([*performance.get("warnings", []), *cache_warnings]))
        analysis = {
            "subcategory": subcategory,
            "segment_name": (cache.get("segment") or {}).get("segment_name"),
            "retailer_scope": channel_key,
            "matched_bundle": performance["segment"]["matched_bundle"],
            "files": {"source": "redshift_cache", "cache_file": str(cache_file)},
            "segment_metrics": {
                "main": channel.get("periods", {}).get("Main", {}),
                "comparison": channel.get("periods", {}).get("Comparison", {}),
                "deltas_pct": performance["segment"]["market_momentum_pct"],
            },
            "brand_focus": {
                "brand": brand_name,
                "main": performance["sunco_position"],
                "comparison": performance["sunco_position"],
                "deltas_pct": {},
            },
            "reference_model": performance.get("reference_family"),
            "top_brands": performance["estimation_inputs"]["top_brands"],
            "top_competitor_products": performance["estimation_inputs"]["top_competitor_products"],
            "performance_estimation_context": performance,
            "warnings": warnings,
        }
        channel_results[channel_key] = {
            "analysis": analysis,
            "performance_estimation_context": performance,
        }

    if not channel_results:
        raise FileNotFoundError(
            f"Redshift Stackline cache for '{subcategory}' does not contain channel data."
        )

    primary_channel = sorted(
        channel_results,
        key=lambda channel: CHANNEL_PRIORITY.get(channel, 0),
        reverse=True,
    )[0]
    comparison = build_channel_comparison(channel_results)
    warnings = []
    for payload in channel_results.values():
        warnings.extend((payload.get("analysis") or {}).get("warnings", []))
    warnings = list(dict.fromkeys(warnings))

    return {
        "subcategory": subcategory,
        "source_system": "redshift_cache",
        "source_priority": "redshift_first",
        "segment_name": (cache.get("segment") or {}).get("segment_name"),
        "primary_channel": primary_channel,
        "primary_analysis": channel_results[primary_channel]["analysis"],
        "channels": channel_results,
        "channel_comparison": comparison,
        "warnings": warnings,
    }
