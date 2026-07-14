from __future__ import annotations

import re
import hashlib
import json
import mimetypes
import os
import shutil
import sys
import urllib.error
import urllib.parse
import urllib.request
from html import unescape
from collections import defaultdict
from io import BytesIO
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

BACKEND_APP = Path(__file__).resolve().parents[2] / "backend" / "app"
if str(BACKEND_APP) not in sys.path:
    sys.path.insert(0, str(BACKEND_APP))

from ecommerce_evidence import ecommerce_rows_to_step1_rows, is_grow_light_candidate, load_or_refresh_ecommerce_snapshot
from luminaire_performance import is_luminaire_category, normalize_luminaire_performance, performance_note
from odbc_client import execute_odbc_sql, redshift_connection_source, redshift_connection_string, sanitize_connection_error
from opportunity_engine.stackline_segments import STACKLINE_SEGMENT_OVERRIDES, sql_values, stackline_segments_for_category
from product_demand_categories import choose_product_demand_category
from sunco_catalog_coverage import catalog_coverage_analysis, load_catalog_context_from_cache_or_snapshot


WEIGHT_PROFILES = {
    "source_rows": {
        "name": "Shopify/ecommerce launch priority",
        "stackline": 0.10,
        "sunco": 0.25,
        "inventory": 0.35,
        "performance": 0.20,
        "quality": 0.10,
        "description": "competitor ecommerce/inventory movement 35%; Sunco coverage/gap 25%; luminaire performance fit 20%; Stackline/Amazon demand 10%; data quality 10%",
    },
    "amazon_rows": {
        "name": "Amazon/Stackline priority",
        "stackline": 0.50,
        "sunco": 0.25,
        "inventory": 0.05,
        "performance": 0.10,
        "quality": 0.10,
        "description": "Stackline/Amazon demand 50%; Sunco coverage/gap 25%; luminaire performance fit 10%; competitor inventory movement 5%; data quality 10%",
    },
}

IMAGE_CONTENT_TYPES = {
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/png": ".png",
    "image/gif": ".gif",
}

REDSHIFT_ODBC_CONNECTION = "DSN=Redshift"

def _existing_step1_imports():
    from opportunity_engine.category_intelligence import format_intelligence_audit
    from opportunity_engine.gap_generator import (
        SUCCESS_PROXY_TEXT,
        _category_run_lock,
        _clear_true_gap_workbook,
        _decision_tree_text,
        _has_gap_language,
        _has_market_signal,
        _template_path,
        _write_amazon_rows,
        _write_source_audit,
        _write_source_rows,
        _write_summary,
        apply_freshness_rows,
        build_step1_freshness_rows,
        load_category_data,
        summarize_freshness_rows,
    )
    from opportunity_engine.line_review import (
        prepare_line_review_context,
        run_audit_rows,
        write_line_review_sheet,
    )
    from opportunity_engine.paths import ProjectPaths
    from opportunity_engine.sql_audit import collect_sql_text
    from opportunity_engine.utils import timestamp
    from opportunity_engine.validation import try_excel_com_open_save, validate_workbook
    from opportunity_engine.workbook_common import add_or_replace_audit_sheet, ensure_template_copy

    return {
        "format_intelligence_audit": format_intelligence_audit,
        "SUCCESS_PROXY_TEXT": SUCCESS_PROXY_TEXT,
        "_category_run_lock": _category_run_lock,
        "_clear_true_gap_workbook": _clear_true_gap_workbook,
        "_decision_tree_text": _decision_tree_text,
        "_has_gap_language": _has_gap_language,
        "_has_market_signal": _has_market_signal,
        "_template_path": _template_path,
        "_write_amazon_rows": _write_amazon_rows,
        "_write_source_audit": _write_source_audit,
        "_write_source_rows": _write_source_rows,
        "_write_summary": _write_summary,
        "apply_freshness_rows": apply_freshness_rows,
        "build_step1_freshness_rows": build_step1_freshness_rows,
        "load_category_data": load_category_data,
        "summarize_freshness_rows": summarize_freshness_rows,
        "line_review_run_audit_rows": run_audit_rows,
        "prepare_line_review_context": prepare_line_review_context,
        "write_line_review_sheet": write_line_review_sheet,
        "ProjectPaths": ProjectPaths,
        "collect_sql_text": collect_sql_text,
        "timestamp": timestamp,
        "try_excel_com_open_save": try_excel_com_open_save,
        "validate_workbook": validate_workbook,
        "add_or_replace_audit_sheet": add_or_replace_audit_sheet,
        "ensure_template_copy": ensure_template_copy,
    }


def _text(value: Any) -> str:
    text = unescape(str(value or ""))
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _date_from_text(value: Any) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def _date_from_path(path: Path | None) -> date | None:
    if not path or not path.exists():
        return None
    if path.suffix.lower() == ".json":
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            for key in ["generated_at", "generated", "created_at"]:
                generated = _date_from_text(payload.get(key))
                if generated:
                    return generated
        except Exception:
            pass
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).date()
    except OSError:
        return None


def _age_days(generated: date | None) -> int | None:
    return (date.today() - generated).days if generated else None


def _freshness_record(
    source: str,
    path: Path | None,
    source_system: str,
    role: str,
    active: bool = True,
    note: str = "",
    generated: date | None = None,
) -> dict[str, Any]:
    resolved_generated = generated or _date_from_path(path)
    return {
        "source": source,
        "source_system": source_system,
        "role": role,
        "path": str(path) if path else "",
        "generated": resolved_generated.isoformat() if resolved_generated else "Unknown",
        "age_days": _age_days(resolved_generated),
        "active": active,
        "note": note,
    }


def _attach_product_demand_freshness(
    *,
    existing: dict[str, Any],
    paths: Any,
    category: Any,
    data: dict[str, Any],
    inventory_snapshot_path: Path | None,
    inventory_source: str,
    catalog_snapshot_path: Path | None,
    catalog_source: str,
    stackline_live: bool,
) -> None:
    rows = existing["build_step1_freshness_rows"](
        paths=paths,
        category=category,
        generated=data.get("generated"),
        source_path=Path(data["source_path"]) if data.get("source_path") else None,
        amazon_rerun_path=Path(data["amazon_rerun_path"]) if data.get("amazon_rerun_path") else None,
        amazon_path=Path(data["amazon_path"]) if data.get("amazon_path") else None,
        include_legacy_active=False,
    )
    rows.append(
        _freshness_record(
            "Redshift ecommerce competitor snapshot",
            inventory_snapshot_path,
            inventory_source,
            "Main Recommendations demand and inventory movement",
            active=bool(inventory_snapshot_path),
            note="Loaded or refreshed by Product Demand Step 1B.",
        )
    )
    rows.append(
        _freshness_record(
            "Sunco catalog coverage snapshot/cache",
            catalog_snapshot_path,
            catalog_source,
            "Existing Sunco catalog coverage check",
            active=bool(catalog_snapshot_path),
            note="Loaded or refreshed by Product Demand Step 1B.",
        )
    )
    rows.append(
        _freshness_record(
            "Live Redshift Stackline Step 1B query",
            None,
            data.get("product_demand_stackline_source") or "redshift_stackline",
            "Amazon recommendation demand",
            active=stackline_live,
            note=data.get("product_demand_stackline_note") or "",
            generated=date.today() if stackline_live else None,
        )
    )
    existing["apply_freshness_rows"](data, rows)


FEATURE_DISPLAY_NAMES = {
    "emergency backup": "Emergency Backup",
    "motion sensor": "Motion Sensor",
    "daylight harvesting": "Daylight Harvesting",
    "control ready": "Control Ready",
    "smart controls": "Smart Controls",
}


def _display_feature(value: Any) -> str:
    text = _text(value)
    return FEATURE_DISPLAY_NAMES.get(text.lower(), text[:1].upper() + text[1:] if text else "")


def _display_features(values: Any) -> list[str]:
    if not values:
        return []
    return [_display_feature(value) for value in values if _display_feature(value)]


def _recommendation_body(value: Any) -> str:
    text = _text(value)
    for marker in [
        "Shopify ecommerce candidate:",
        "Shopify technology gap candidate",
        "Shopify partial-coverage technology gap",
        "Merchandising/running-change candidate:",
        "Possible feature gap:",
        "Existing Sunco coverage, but missing feature:",
        "Product Revision or merchandising review:",
        "New variant opportunity:",
        "Strategic outlier watchlist:",
        "Strategic outlier / High-output watchlist:",
    ]:
        if text.lower().startswith(marker.lower()):
            if ":" in text:
                return text.split(":", 1)[1].strip()
    return text


def _set_recommendation_label(row: dict[str, Any], label: str, features: list[str] | None = None) -> None:
    body = _recommendation_body(row.get("recommendation"))
    feature_note = f" ({', '.join(features)})" if features else ""
    row["recommendation"] = f"{label}{feature_note}: {body}" if body else f"{label}{feature_note}"


def _tokens(value: str) -> set[str]:
    raw = re.findall(r"[a-z0-9]+", value.lower())
    ignored = {"led", "light", "lights", "fixture", "fixtures", "for", "and", "with", "the", "max"}
    return {token for token in raw if len(token) >= 2 and token not in ignored}


def _row_text(row: dict[str, Any]) -> str:
    return " ".join(
        _text(row.get(key))
        for key in [
            "recommendation",
            "classification",
            "example",
            "evidence",
            "why_gap",
            "source_url",
            "review_url",
            "sunco_check",
            "pm_action",
            "action",
        ]
    )


def _inventory_text(row: dict[str, Any]) -> str:
    return " ".join(
        _text(row.get(key))
        for key in [
            "name",
            "brand",
            "sku",
            "url",
            "category",
            "product_type",
            "wattage",
            "lumens",
            "cct",
            "voltage",
        ]
    )


def _row_performance_text(row: dict[str, Any]) -> str:
    return " ".join(
        _text(row.get(key))
        for key in [
            "recommendation",
            "classification",
            "example",
            "evidence",
            "why_gap",
            "sunco_check",
            "pm_action",
            "action",
        ]
    )


def _load_ecommerce_context(root: Path, category_slug: str) -> tuple[list[dict[str, Any]], Path | None, str]:
    exports = root / "product_demand_ideation" / "experiments" / category_slug / "exports"
    snapshot, snapshot_path = load_or_refresh_ecommerce_snapshot(exports, category_slug)
    return list(snapshot.get("rows") or []), snapshot_path, str(snapshot.get("source_system") or "unknown")


def _load_catalog_context(root: Path, category_slug: str) -> tuple[list[dict[str, Any]], Path | None, str]:
    return load_catalog_context_from_cache_or_snapshot(root, category_slug)


def _best_inventory_match(row: dict[str, Any], inventory_rows: list[dict[str, Any]]) -> tuple[dict[str, Any] | None, float]:
    if not inventory_rows:
        return None, 0.0
    row_tokens = _tokens(_row_text(row))
    if not row_tokens:
        return None, 0.0
    best: dict[str, Any] | None = None
    best_score = 0.0
    for item in inventory_rows:
        item_tokens = _tokens(_inventory_text(item))
        if not item_tokens:
            continue
        overlap = row_tokens & item_tokens
        score = len(overlap) / max(6, min(len(row_tokens), len(item_tokens)))
        if score > best_score:
            best = item
            best_score = score
    return best, best_score


def _performance_alignment_score(row: dict[str, Any], match: dict[str, Any] | None, category_slug: str) -> tuple[float, str]:
    row_perf = normalize_luminaire_performance(_row_performance_text(row))
    match_perf = normalize_luminaire_performance(_inventory_text(match or {}))
    if not is_luminaire_category(category_slug) and not (row_perf.lumen_values or match_perf.lumen_values):
        return 0.0, "Luminaire performance logic not applied because no lumen/brightness signal was parsed."
    if not row_perf.has_signal and not match_perf.has_signal:
        return 0.0, "No lumen/wattage signal parsed from Step 1 row or inventory match."
    score = 0.0
    if row_perf.lumen_values and match_perf.lumen_values:
        row_lumen = max(row_perf.lumen_values)
        match_lumen = max(match_perf.lumen_values)
        if row_lumen > 0:
            ratio = min(row_lumen, match_lumen) / max(row_lumen, match_lumen)
            tier_bonus = 10.0 if row_perf.brightness_tier == match_perf.brightness_tier else 0.0
            score += min(45.0, max(0.0, (ratio * 35.0) + tier_bonus))
    elif row_perf.brightness_tier and match_perf.brightness_tier and row_perf.brightness_tier == match_perf.brightness_tier:
        score += 35.0
    if row_perf.wattage_class and match_perf.wattage_class and row_perf.wattage_class == match_perf.wattage_class:
        score += 25.0
    elif row_perf.watt_values and match_perf.watt_values:
        row_watt = max(row_perf.watt_values)
        match_watt = max(match_perf.watt_values)
        if row_watt > 0:
            ratio = min(row_watt, match_watt) / max(row_watt, match_watt)
            score += max(0.0, ratio * 25.0)
    if row_perf.efficacy_values:
        score += 15.0
    if match_perf.efficacy_values:
        score += 15.0
    note = (
        performance_note("Step 1 target", row_perf)
        + " "
        + performance_note("Inventory/competitor match", match_perf)
    )
    return min(100.0, round(score, 1)), note


def _score_quality(row: dict[str, Any]) -> float:
    priority = _text(row.get("priority")).lower()
    confidence = _text(row.get("confidence")).lower()
    score = 40.0
    if priority == "high":
        score += 25.0
    elif priority == "medium":
        score += 15.0
    if confidence == "high":
        score += 25.0
    elif confidence == "medium":
        score += 15.0
    if row.get("source_url") or row.get("review_url"):
        score += 10.0
    return min(score, 100.0)


def _has_stackline_or_amazon_support(row: dict[str, Any]) -> bool:
    haystack = " ".join(
        _text(row.get(field))
        for field in ["example", "evidence", "why_gap", "source_url", "review_url", "source_systems"]
    ).lower()
    return any(term in haystack for term in ["stackline", "amazon.com", "asin", "bsr", "retail sales"])


def _weighted_score(
    row: dict[str, Any],
    source_kind: str,
    profile_key: str,
    inventory_rows: list[dict[str, Any]],
    max_inventory_decrease: float,
    max_inventory_velocity: float,
    existing: dict[str, Any],
    category_slug: str,
) -> dict[str, Any]:
    profile = WEIGHT_PROFILES[profile_key]
    market_signal = existing["_has_market_signal"](row, source_kind)
    sunco_gap = existing["_has_gap_language"](row.get("sunco_check"))
    coverage_score = float(row.get("_sunco_catalog_coverage_score") or 0)
    has_technology_gap = bool(row.get("_technology_gap_features"))
    is_supplemental = bool(row.get("_supplemental")) or _text(row.get("recommendation")).lower().startswith("[supplemental candidate]")
    match, match_quality = (None, 0.0) if is_supplemental else _best_inventory_match(row, inventory_rows)
    decrease = float((match or {}).get("observed_stock_decrease") or 0)
    velocity = max(
        float((match or {}).get("avg_units_per_week_observed_window") or 0),
        float((match or {}).get("avg_units_per_week_decrease_window") or 0),
    )
    events = float((match or {}).get("decrease_events") or 0)
    observations = float((match or {}).get("observation_count") or 0)
    days_since_last_decrease = (match or {}).get("days_since_last_decrease")
    inventory_score = 0.0
    if match and match_quality >= 0.18 and max_inventory_decrease > 0:
        relative_decrease_score = min(100.0, (decrease / max_inventory_decrease) * 100.0)
        relative_velocity_score = min(100.0, (velocity / max_inventory_velocity) * 100.0) if max_inventory_velocity > 0 else 0.0
        absolute_velocity_score = min(100.0, (velocity / 40.0) * 100.0)
        event_score = min(100.0, (events / 6.0) * 100.0)
        observation_score = min(100.0, (observations / 10.0) * 100.0)
        recency_score = 0.0
        if days_since_last_decrease is not None:
            days = float(days_since_last_decrease or 0)
            if days <= 7:
                recency_score = 100.0
            elif days <= 21:
                recency_score = 70.0
            elif days <= 45:
                recency_score = 40.0
            else:
                recency_score = 15.0
        inventory_score = min(
            100.0,
            (max(relative_velocity_score, absolute_velocity_score) * 0.35)
            + (event_score * 0.25)
            + (recency_score * 0.20)
            + (relative_decrease_score * 0.10)
            + (observation_score * 0.10),
        )
    performance_score, performance_signal = _performance_alignment_score(row, match, category_slug)
    if profile_key == "source_rows":
        stackline_score = 100.0 if _has_stackline_or_amazon_support(row) else 35.0
    else:
        stackline_score = 100.0 if market_signal else 35.0
    sunco_score = 100.0 if sunco_gap else 40.0
    if coverage_score >= 75 and not has_technology_gap:
        sunco_score = 0.0
        inventory_score = min(inventory_score, 35.0)
    elif has_technology_gap:
        sunco_score = 100.0
    quality_score = _score_quality(row)
    total = (
        (stackline_score * profile["stackline"])
        + (sunco_score * profile["sunco"])
        + (inventory_score * profile["inventory"])
        + (performance_score * profile["performance"])
        + (quality_score * profile["quality"])
    )
    return {
        "total": round(total, 1),
        "profile_name": profile["name"],
        "profile_description": profile["description"],
        "stackline_weight": int(profile["stackline"] * 100),
        "stackline_score": round(stackline_score, 1),
        "sunco_weight": int(profile["sunco"] * 100),
        "sunco_score": round(sunco_score, 1),
        "inventory_weight": int(profile["inventory"] * 100),
        "inventory_score": round(inventory_score, 1),
        "inventory_velocity": round(velocity, 1),
        "inventory_days_since_last_decrease": days_since_last_decrease,
        "performance_weight": int(profile["performance"] * 100),
        "performance_score": round(performance_score, 1),
        "performance_signal": performance_signal,
        "quality_weight": int(profile["quality"] * 100),
        "quality_score": round(quality_score, 1),
        "inventory_match": match,
        "inventory_match_quality": round(match_quality, 3),
        "supplemental": is_supplemental,
    }


def _append_note(existing_value: Any, note: str) -> str:
    base = _text(existing_value).strip()
    if not base:
        return note
    if note in base:
        return base
    return f"{base}\n\n{note}"


def _cache_competitor_image(paths: Any, image_url: Any, category_slug: str) -> str | None:
    url = _text(image_url)
    if not url.startswith(("http://", "https://")):
        return None
    image_root = paths.cache / "images" / "product_demand_competitors" / category_slug
    image_root.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha1(url.encode("utf-8")).hexdigest()[:16]
    for existing in image_root.glob(f"{digest}.*"):
        if existing.is_file():
            return str(existing)
    parsed = urllib.parse.urlsplit(url)
    safe_url = urllib.parse.urlunsplit(
        (
            parsed.scheme,
            parsed.netloc,
            urllib.parse.quote(parsed.path, safe="/%:"),
            urllib.parse.quote(parsed.query, safe="=&?/:+%"),
            parsed.fragment,
        )
    )
    request = urllib.request.Request(safe_url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            data = response.read(4_000_000)
    except Exception:
        return None
    if not data:
        return None
    content_type = ""
    extension = None
    try:
        from PIL import Image

        with Image.open(BytesIO(data)) as image:
            image_format = (image.format or "").upper()
            if image_format in {"JPEG", "JPG"}:
                extension = ".jpg"
                image = image.convert("RGB")
                target = image_root / f"{digest}{extension}"
                image.save(target, format="JPEG", quality=90)
                return str(target)
            if image_format == "PNG":
                extension = ".png"
                target = image_root / f"{digest}{extension}"
                image.save(target, format="PNG")
                return str(target)
            if image_format == "GIF":
                extension = ".gif"
                target = image_root / f"{digest}{extension}"
                image.save(target, format="GIF")
                return str(target)
            target = image_root / f"{digest}.png"
            image.convert("RGBA").save(target, format="PNG")
            return str(target)
    except Exception:
        pass
    if data.startswith(b"\xff\xd8\xff"):
        extension = ".jpg"
    elif data.startswith(b"\x89PNG\r\n\x1a\n"):
        extension = ".png"
    elif data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
        extension = ".gif"
    else:
        extension = IMAGE_CONTENT_TYPES.get(content_type)
        if not extension:
            extension = mimetypes.guess_extension(content_type) if content_type else None
        if extension == ".jpe":
            extension = ".jpg"
        if extension not in {".jpg", ".jpeg", ".png", ".gif"}:
            return None
    target = image_root / f"{digest}{extension}"
    target.write_bytes(data)
    return str(target)


def _attach_competitor_images(paths: Any, rows: list[dict[str, Any]], category_slug: str) -> None:
    for row in rows:
        if row.get("local_image"):
            continue
        local_image = _cache_competitor_image(paths, row.get("image_url"), category_slug)
        if local_image:
            row["local_image"] = local_image


def _as_float(value: Any) -> float:
    try:
        return float(str(value or "0").replace(",", ""))
    except ValueError:
        return 0.0


def build_stackline_redshift_sql(category_slug: str, category_name: str, limit: int = 25) -> str:
    segments = stackline_segments_for_category(category_slug, category_name)
    if not segments:
        return ""
    exact_segments = STACKLINE_SEGMENT_OVERRIDES.get(category_slug)
    if exact_segments:
        segment_filter = f"cs.segment_name in ({sql_values(exact_segments)})"
    else:
        segment_filter = " or ".join(
            "lower(cs.segment_name) like '%" + term.replace("'", "''") + "%'"
            for term in segments
        )
    return f"""
with latest as (
  select max(week_id) as max_week
  from public.tb_stackline_atlas_gold_sales
),
target_segments as (
  select distinct segment_name, retailer_id, retailer_sku
  from public.tb_stackline_atlas_current_segment cs
  where cs.retailer_id = 1
    and ({segment_filter})
),
ranked as (
  select
    ts.segment_name,
    s.retailer_sku,
    p.model_number,
    p.title,
    p.brand_name,
    p.category_name,
    p.subcategory_name,
    min(s.week_id) as min_week_id,
    max(s.week_id) as max_week_id,
    sum(s.retail_sales) as retail_sales,
    sum(s.units_sold) as units_sold,
    avg(s.retail_price) as avg_retail_price,
    max(s.reviews_count) as reviews_count,
    avg(s.reviews_rating) as reviews_rating,
    avg(s.content_score) as content_score,
    avg(s.title_score) as title_score,
    avg(s.image_score) as image_score,
    count(distinct s.week_id) as weeks
  from target_segments ts
  join public.tb_stackline_atlas_gold_sales s
    on s.retailer_id = ts.retailer_id
   and s.retailer_sku = ts.retailer_sku
  left join public.tb_stackline_atlas_products p
    on p.retailer_id = s.retailer_id
   and p.retailer_sku = s.retailer_sku
  cross join latest
  where s.week_id >= greatest(202601, latest.max_week - 25)
  group by 1,2,3,4,5,6,7
)
select *
from ranked
where coalesce(retail_sales, 0) > 0
order by retail_sales desc, units_sold desc
limit {int(limit)};
""".strip()


def _industrial_design_cues(text: str) -> list[str]:
    raw = text.lower()
    cues: list[str] = []
    checks = [
        ("combo exit sign with heads", ("exit sign", "emergency light", "head")),
        ("bug-eye emergency light", ("bug-eye", "bug eye", "dual adjustable led lamp heads")),
        ("edge-lit design", ("edge-lit", "edge lit")),
        ("wet-location rated", ("wet location", "wet rated")),
        ("red lettering", ("red",)),
        ("green lettering", ("green",)),
        ("black housing", ("black",)),
        ("white housing", ("white",)),
        ("thermoplastic/ABS housing", ("thermoplastic", "abs")),
        ("90-minute backup", ("90-minute", "90 minute", "90 min")),
        ("UL listed", ("ul listed", " ul ")),
        ("multi-pack", ("2 pack", "4 pack", "6 pack", "12 pack", "2-pack", "4-pack", "6-pack", "12-pack")),
    ]
    for label, terms in checks:
        if all(term in raw for term in terms) if label == "combo exit sign with heads" else any(term in raw for term in terms):
            cues.append(label)
    return cues[:5]


def _stackline_items_to_amazon_rows(items: list[dict[str, Any]], source_label: str, limit: int = 10, category_slug: str | None = None) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for item in items:
        title = _text(item.get("title"))
        brand = _text(item.get("brand_name") or item.get("brand"))
        model = _text(item.get("model_number") or item.get("model"))
        segment = _text(item.get("segment_name") or item.get("segment"))
        asin = _text(item.get("retailer_sku") or item.get("asin"))
        if category_slug in {"commercial_grow_lights", "residential_grow_lights"} and not is_grow_light_candidate(
            {
                "name": title,
                "description": f"{brand} {model} {segment}",
                "category": item.get("category_name"),
                "product_type": item.get("subcategory_name"),
                "brand": brand,
                "sku": asin,
            }
        ):
            continue
        retail_sales = _as_float(item.get("retail_sales"))
        units_sold = _as_float(item.get("units_sold"))
        avg_price = _as_float(item.get("avg_retail_price") or item.get("price"))
        weeks = int(_as_float(item.get("weeks")))
        cues = _industrial_design_cues(f"{title} {model}")
        is_sunco = "sunco" in brand.lower() or "sunco" in title.lower()
        priority = "High" if retail_sales >= 100_000 or units_sold >= 1_000 else "Medium"
        recommendation_prefix = "Defend/optimize Sunco Amazon winner" if is_sunco else "Amazon Stackline opportunity"
        recommendation = f"{recommendation_prefix}: {segment or 'Stackline segment'} - {', '.join(cues) if cues else brand or 'top product'}"
        evidence = (
            f"{source_label} found ${retail_sales:,.0f} retail sales and {units_sold:,.0f} units "
            f"across {weeks} week(s). Average observed retail price ${avg_price:,.2f}. "
            f"Segment: {segment or 'Unknown'}."
        )
        if item.get("min_week_id") or item.get("max_week_id"):
            evidence += f" Week range: {item.get('min_week_id')} to {item.get('max_week_id')}."
        if cues:
            evidence += f" Industrial-design cues: {', '.join(cues)}."
        if item.get("reviews_count"):
            evidence += f" Reviews: {float(item.get('reviews_count')):,.0f}; rating: {_as_float(item.get('reviews_rating')):.1f}."
        rows.append(
            {
                "subcategory": segment or "Amazon",
                "recommendation": recommendation,
                "priority": priority,
                "confidence": "High" if weeks >= 8 else "Medium",
                "classification": "Stackline/Amazon leader" if not is_sunco else "Sunco Amazon incumbent",
                "evidence": evidence,
                "sunco_check": (
                    "Sunco already has this Amazon design/pack position; use as defend-and-optimize benchmark."
                    if is_sunco
                    else "Needs Sunco Amazon exact-design and pack-size coverage check before treating as a launch gap."
                ),
                "example": f"{brand} {model}: {title}".strip(),
                "review_url": f"https://www.amazon.com/dp/{asin}" if asin.startswith("B") else "",
                "action": (
                    "Audit PDP imagery, pack strategy, title/spec copy, price position, and ad/rank support for this proven Amazon design."
                    if is_sunco
                    else "Compare Sunco's Amazon catalog against this design cue and pack/price architecture; scope listing or SKU gap if not covered."
                ),
                "source_systems": ["redshift:stackline_atlas_gold_sales", "redshift:stackline_atlas_current_segment", "redshift:stackline_atlas_products"],
                "_stackline_local_sales": retail_sales,
                "_stackline_local_units": units_sold,
            }
        )
        if len(rows) >= limit:
            break
    return rows


def _stackline_amazon_rows_from_redshift(category_slug: str, category_name: str, limit: int = 10) -> tuple[list[dict[str, Any]], str]:
    sql = build_stackline_redshift_sql(category_slug, category_name, limit=max(limit * 3, 25))
    if not sql:
        return [], "No Redshift Stackline segment mapping for this category."
    try:
        items = execute_odbc_sql(redshift_connection_string(REDSHIFT_ODBC_CONNECTION), sql, timeout_seconds=240)
    except Exception as exc:
        return [], (
            f"Redshift Stackline query failed through {redshift_connection_source(REDSHIFT_ODBC_CONNECTION)}: "
            f"{sanitize_connection_error(exc)}"
        )
    return _stackline_items_to_amazon_rows(items, "Live Redshift Stackline Atlas", limit=limit, category_slug=category_slug), sql


def _apply_catalog_coverage_to_amazon_rows(rows: list[dict[str, Any]], catalog_rows: list[dict[str, Any]], category_slug: str) -> None:
    generic_checks = (
        "Needs Sunco Amazon exact-design and pack-size coverage check before treating as a launch gap.",
        "Sunco already has this Amazon design/pack position; use as defend-and-optimize benchmark.",
        "Pending Sunco Amazon coverage join.",
    )
    for row in rows:
        coverage = catalog_coverage_analysis(row, catalog_rows, category_slug)
        coverage_score = float(coverage.get("score") or 0)
        missing_features = coverage.get("missing_features") or []
        note = coverage.get("note") or ""
        current_note = _text(row.get("sunco_check"))
        if row.get("classification") == "Sunco Amazon incumbent":
            note = _append_note("Sunco-owned Amazon row: use this as a defend-and-optimize benchmark.", note)
        elif coverage_score >= 75:
            note = _append_note(
                note,
                "Amazon action: likely optimize listing, pack visibility, image story, pricing, or search placement before creating a new SKU.",
            )
        elif coverage_score >= 50:
            note = _append_note(
                note,
                "Amazon action: treat as listing-depth or variant/spec-gap candidate based on the Sunco coverage note.",
            )
        else:
            note = _append_note(
                note,
                "Amazon action: likely launch/assortment candidate where Sunco active-catalog coverage shows no equivalent offer.",
            )
        if missing_features:
            note = _append_note(
                note,
                "Technology gap signal: competitor/Amazon evidence includes "
                f"{', '.join(missing_features)} that was not found on the closest Sunco active-catalog match.",
            )
        if current_note and current_note not in generic_checks:
            row["sunco_check"] = _append_note(current_note, note)
        else:
            row["sunco_check"] = note
        row["_sunco_catalog_coverage_score"] = coverage_score
        row["_technology_gap_features"] = missing_features


def _apply_product_demand_overlay(
    paths: Any,
    data: dict[str, Any],
    inventory_rows: list[dict[str, Any]],
    catalog_rows: list[dict[str, Any]],
    existing: dict[str, Any],
    category_slug: str,
    category_name: str,
) -> dict[str, Any]:
    output = deepcopy(data)
    ecommerce_source_rows = ecommerce_rows_to_step1_rows(category_name, inventory_rows)
    if ecommerce_source_rows:
        for row in ecommerce_source_rows:
            coverage = catalog_coverage_analysis(row, catalog_rows, category_slug)
            coverage_score = float(coverage.get("score") or 0)
            missing_features = coverage.get("missing_features") or []
            display_missing_features = _display_features(missing_features)
            note = coverage.get("note") or ""
            if missing_features:
                note = _append_note(
                    note,
                    "Feature gap signal: competitor evidence includes searched/customer term(s) "
                    f"{', '.join(display_missing_features)} that were not found on the closest Sunco active-catalog match. "
                    "Treat this as a feature hypothesis before treating it as a base SKU duplicate.",
                )
            row["sunco_check"] = note
            row["_sunco_catalog_coverage_score"] = coverage_score
            row["_technology_gap_features"] = missing_features
            if row.get("_strategic_outlier"):
                row["classification"] = "Strategic outlier / High-output watchlist"
                row["priority"] = "Medium"
                row["confidence"] = "Directional"
                _set_recommendation_label(row, "Strategic outlier / High-output watchlist")
                row["sunco_check"] = _append_note(
                    note,
                    (
                        "Strategic outlier lane: source-backed high-output concept kept visible despite falling below "
                        "the normal Step 1B rank cutoff. Treat as a watchlist/RFQ-screening hypothesis, not a standard panel replacement."
                    ),
                )
                row["pm_action"] = (
                    "Review as a strategic high-output outlier. Confirm application fit, installation constraints, heat/load profile, "
                    "certifications, vendor feasibility, and whether demand is broader than one competitor PDP before PRD/RFQ."
                )
            elif coverage_score >= 75 and missing_features:
                row["classification"] = "Existing Sunco coverage, but missing feature"
                _set_recommendation_label(row, "Existing Sunco coverage, but missing feature", display_missing_features)
                row["pm_action"] = (
                    f"Position {', '.join(display_missing_features)} as a feature opportunity on the matched active Sunco family. "
                    "Use as an add-variant, running-change, or PDP merchandising callout path rather than a base SKU duplicate."
                )
            elif coverage_score >= 75:
                row["classification"] = "Product Revision or merchandising review"
                row["priority"] = "Medium"
                row["confidence"] = "Coverage exists"
                _set_recommendation_label(row, "Product Revision or merchandising review")
                row["pm_action"] = (
                    "Prioritize PDP/category merchandising or product revision on the matched active SKU. "
                    "Use the competitor demand signal to improve title/spec coverage, pack-size visibility, price position, search/category placement, or running-change planning."
                )
            elif coverage_score >= 50:
                if missing_features:
                    row["classification"] = "Possible feature gap"
                    _set_recommendation_label(row, "Possible feature gap", display_missing_features)
                else:
                    row["classification"] = "New variant opportunity"
                    _set_recommendation_label(row, "New variant opportunity")
                row["pm_action"] = (
                    "Scope the missing variant or feature depth shown by this row. "
                    "Route PDP/pack visibility findings to merchandising instead of new SKU setup."
                )
            else:
                row["classification"] = "New variant opportunity"
                _set_recommendation_label(row, "New variant opportunity")
                row["pm_action"] = (
                    "Scope as a new variant around this normalized spec pattern. "
                    "Use the coverage result and demand evidence in this row as the launch rationale."
                )
        output["source_rows"] = ecommerce_source_rows
        output["source_exact_count"] = len(ecommerce_source_rows)
        output["source_from_amazon_count"] = 0
        output["source_supplemental_count"] = 0
    else:
        output["source_rows"] = []
        output["source_exact_count"] = 0
        output["source_from_amazon_count"] = 0
        output["source_supplemental_count"] = 0
        output["source_supplemental_warnings"] = [
            "No Redshift ecommerce competitor rows were available for this category. Product Demand Step 1B did not use legacy local recommendation seeds."
        ]
    redshift_stackline_rows, redshift_stackline_audit = _stackline_amazon_rows_from_redshift(category_slug, category_name)
    if redshift_stackline_rows:
        output["amazon_rows"] = redshift_stackline_rows
        output["amazon_exact_count"] = len(redshift_stackline_rows)
        output["amazon_supplemental_count"] = 0
        output["product_demand_stackline_source"] = "redshift_odbc_stackline_atlas"
        output["product_demand_stackline_audit"] = redshift_stackline_audit
        output["product_demand_stackline_note"] = (
            f"Live Redshift Stackline Atlas returned {len(redshift_stackline_rows)} Amazon rows. "
            "Legacy local Step 1 Amazon rows were excluded from Product Demand Step 1B."
        )
    else:
        output["amazon_rows"] = []
        output["amazon_exact_count"] = 0
        output["amazon_supplemental_count"] = 0
        output["product_demand_stackline_audit"] = redshift_stackline_audit
        output["product_demand_stackline_note"] = redshift_stackline_audit
        output["product_demand_stackline_note"] = _append_note(
            redshift_stackline_audit,
            "Step 1B requires Redshift/ODBC Stackline data and did not use legacy local Amazon recommendation seeds.",
        )
    _apply_catalog_coverage_to_amazon_rows(output.get("amazon_rows", []), catalog_rows, category_slug)
    max_inventory_decrease = max((float(row.get("observed_stock_decrease") or 0) for row in inventory_rows), default=0.0)
    max_inventory_velocity = max(
        (
            max(
                float(row.get("avg_units_per_week_observed_window") or 0),
                float(row.get("avg_units_per_week_decrease_window") or 0),
            )
            for row in inventory_rows
        ),
        default=0.0,
    )
    overlay_rows: list[dict[str, Any]] = []
    for key, source_kind in [("source_rows", "Sunco.com/ecommerce"), ("amazon_rows", "Amazon")]:
        scored_rows = []
        for row in output.get(key, []):
            score = _weighted_score(row, source_kind, key, inventory_rows, max_inventory_decrease, max_inventory_velocity, existing, category_slug)
            match = score.get("inventory_match") or {}
            inventory_note = "no strong competitor inventory match"
            if match:
                inventory_note = (
                    f"{match.get('brand') or 'Unknown'} {match.get('sku') or ''} | "
                    f"observed stock decrease {match.get('observed_stock_decrease')} | "
                    f"decrease events {match.get('decrease_events')} | "
                    f"velocity {score.get('inventory_velocity')} units/week | "
                    f"days since last decrease {score.get('inventory_days_since_last_decrease')} | "
                    f"match quality {score['inventory_match_quality']}"
                )
            elif score.get("supplemental"):
                inventory_note = "not applied because this is a supplemental adjacent-category row"
            note = (
                f"Product Demand overlay score: {score['total']}/100. "
                f"Score profile: {score['profile_name']}. "
                f"Weights: Stackline/Amazon demand {score['stackline_weight']}%={score['stackline_score']}; "
                f"Sunco coverage/gap {score['sunco_weight']}%={score['sunco_score']}; "
                f"competitor ecommerce inventory movement {score['inventory_weight']}%={score['inventory_score']}; "
                f"luminaire performance fit {score['performance_weight']}%={score['performance_score']}; "
                f"data quality {score['quality_weight']}%={score['quality_score']}. "
                f"Inventory support: {inventory_note}. "
                f"{score['performance_signal']}"
            )
            row["_product_demand_score"] = score["total"]
            row["_product_demand_overlay"] = note
            if key == "amazon_rows":
                row["evidence"] = _append_note(row.get("evidence"), note)
            else:
                row["why_gap"] = _append_note(row.get("why_gap"), note)
            overlay_rows.append(
                {
                    "tab": key,
                    "recommendation": row.get("recommendation"),
                    "score": score["total"],
                    "profile": score["profile_name"],
                    "inventory_support": inventory_note,
                    "performance_signal": score["performance_signal"],
                }
            )
            scored_rows.append(row)
        if key == "amazon_rows" and any(row.get("_stackline_local_sales") for row in scored_rows):
            scored_rows.sort(
                key=lambda item: (
                    float(item.get("_stackline_local_sales") or 0),
                    float(item.get("_stackline_local_units") or 0),
                    float(item.get("_product_demand_score") or 0),
                ),
                reverse=True,
            )
        else:
            if key == "source_rows":
                scored_rows.sort(
                    key=lambda item: (
                        1 if item.get("_strategic_outlier") else 0,
                        -float(item.get("_product_demand_score") or 0),
                    )
                )
            else:
                scored_rows.sort(key=lambda item: float(item.get("_product_demand_score") or 0), reverse=True)
        output[key] = scored_rows
    _attach_competitor_images(paths, output.get("source_rows", []), category_slug)
    output["product_demand_overlay"] = overlay_rows
    return output


def generate_product_demand_step1b(root: Path | str) -> tuple[Path, list[str], dict[str, Any]]:
    root = Path(root)
    existing = _existing_step1_imports()
    paths = existing["ProjectPaths"].from_root(root)
    paths.ensure()
    category = choose_product_demand_category(paths)
    inventory_rows, inventory_snapshot_path, inventory_source = _load_ecommerce_context(root, category.slug)
    catalog_rows, catalog_snapshot_path, catalog_source = _load_catalog_context(root, category.slug)
    with existing["_category_run_lock"](paths, category):
        data = existing["load_category_data"](paths, category)
        data = _apply_product_demand_overlay(paths, data, inventory_rows, catalog_rows, existing, category.slug, category.run_name)
        if not data.get("amazon_rows"):
            raise RuntimeError(
                "Product Demand Step 1B stopped before writing a workbook because Amazon/Stackline rows are missing. "
                "This would create an incomplete combined-model report. Fix the local Redshift ODBC/env config and rerun. "
                f"Stackline note: {data.get('product_demand_stackline_note') or 'No Stackline note was returned.'}"
            )
        line_review_context = existing["prepare_line_review_context"](paths, category)
        data["line_review_context"] = line_review_context
        _attach_product_demand_freshness(
            existing=existing,
            paths=paths,
            category=category,
            data=data,
            inventory_snapshot_path=inventory_snapshot_path,
            inventory_source=inventory_source,
            catalog_snapshot_path=catalog_snapshot_path,
            catalog_source=catalog_source,
            stackline_live=data.get("product_demand_stackline_source") == "redshift_odbc_stackline_atlas",
        )
        template = existing["_template_path"](paths)
        run_stamp = existing["timestamp"]()
        output_folder = root / "product_demand_ideation" / "experiments" / category.slug / "outputs"
        output_folder.mkdir(parents=True, exist_ok=True)
        output = output_folder / f"{category.slug}_product_demand_step1b_{run_stamp}.xlsx"
        step2_handoff_folder = paths.gap_category_outputs(category.slug)
        step2_handoff = step2_handoff_folder / f"{category.slug}_true_gaps_{run_stamp}_product_demand_step1b.xlsx"
        existing["ensure_template_copy"](template, output)

        workbook = load_workbook(output)
        existing["_clear_true_gap_workbook"](workbook)
        image_status = []
        existing["_write_summary"](workbook, category, data)
        image_status.extend(existing["_write_source_rows"](paths, workbook, data["source_rows"]))
        image_status.extend(existing["_write_amazon_rows"](paths, workbook, data["amazon_rows"]))
        existing["_write_source_audit"](workbook, category, data)
        existing["write_line_review_sheet"](workbook, line_review_context)

        overlay_summary = "\n".join(
            f"{row['tab']}: {row['score']} | {row['profile']} | {row['recommendation']} | {row['inventory_support']}"
            for row in data.get("product_demand_overlay", [])[:30]
        )
        audit_rows = [
            ("Project", "sunco-product-opportunity-engine / Product Demand Step 1B"),
            ("Category", category.run_name),
            ("Owner", category.owner),
            ("Generated", run_stamp),
            ("Original Step 1 reused", "Yes - category data, Stackline/Amazon evidence, Sunco coverage checks, workbook writer, line review, and Step 2 contract are reused."),
            ("Step 1B change", "Uses Redshift ecommerce competitor PDP evidence for the main Recommendations tab, keeps Amazon Recommendations Stackline-led, saves to the isolated product_demand_ideation output folder, and publishes a clearly labeled Step 2 handoff copy."),
            ("Step 2 handoff workbook", str(step2_handoff)),
            ("Weighting rule", f"Recommendations tab: {WEIGHT_PROFILES['source_rows']['description']}. Amazon Recommendations tab: {WEIGHT_PROFILES['amazon_rows']['description']}."),
            ("Luminaire performance rule", "For light-producing categories, compare lumens as the user-facing brightness target and wattage as the installer/load target. Prefer recommendations that meet a brightness tier at equal or lower wattage."),
            ("Ecommerce competitor source", inventory_source),
            ("Ecommerce competitor snapshot", str(inventory_snapshot_path) if inventory_snapshot_path else "No ecommerce competitor overlay snapshot for this category."),
            ("Amazon / Stackline source", data.get("product_demand_stackline_source") or "Inherited from original Step 1 category data."),
            ("Amazon / Stackline note", data.get("product_demand_stackline_note") or "No live Redshift Stackline replacement was applied."),
            ("Sunco catalog source", catalog_source),
            ("Sunco catalog snapshot", str(catalog_snapshot_path) if catalog_snapshot_path else "No Sunco catalog coverage snapshot for this category."),
            ("Active source age days", "Unknown" if data.get("age_days") is None else str(data.get("age_days"))),
            ("Source freshness detail", data.get("freshness_summary") or "No source freshness records were available."),
            ("Main Recommendations source rule", "If Redshift ecommerce PDP rows exist for the selected category, they lead the main Recommendations tab. Amazon-derived display rows are kept out of the Shopify/front-end tab."),
            ("Inventory role", "Inventory movement is a Shopify/ecommerce demand proxy for the main tab and a supporting signal only for the Amazon tab."),
            ("Product Demand overlay rows", overlay_summary or "No overlay rows."),
            ("Image status", "\n".join(image_status) if image_status else "No images embedded for this category run."),
            ("Category intelligence audit", existing["format_intelligence_audit"](data["category_intelligence"])),
        ]
        audit_rows.extend(existing["line_review_run_audit_rows"](line_review_context))
        existing["add_or_replace_audit_sheet"](
            workbook,
            audit_rows,
            existing["collect_sql_text"](paths, category.slug),
        )
        workbook.save(output)
        workbook.close()
        step2_handoff_folder.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(output, step2_handoff)

    issues = existing["validate_workbook"](output, ["Summary", "Recommendations", "Sources and Audit", "Amazon Recommendations", "Amazon Source Audit", "Run Audit"])
    handoff_issues = existing["validate_workbook"](step2_handoff, ["Summary", "Recommendations", "Sources and Audit", "Amazon Recommendations", "Amazon Source Audit", "Run Audit"])
    issues.extend(f"Step 2 handoff: {issue}" for issue in handoff_issues)
    ok, message = existing["try_excel_com_open_save"](output)
    if message:
        issues.append(message if not ok else message)
    metadata = {
        "category": category.run_name,
        "category_slug": category.slug,
        "inventory_snapshot": str(inventory_snapshot_path) if inventory_snapshot_path else None,
        "inventory_source": inventory_source,
        "catalog_snapshot": str(catalog_snapshot_path) if catalog_snapshot_path else None,
        "catalog_source": catalog_source,
        "source_rows": len(data.get("source_rows", [])),
        "amazon_rows": len(data.get("amazon_rows", [])),
        "step2_handoff": str(step2_handoff),
    }
    return output, issues, metadata
