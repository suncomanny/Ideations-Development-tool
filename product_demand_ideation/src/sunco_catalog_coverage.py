from __future__ import annotations

import json
import os
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from ecommerce_evidence import profile_for
from luminaire_performance import normalize_luminaire_performance
from sku_classification_cache import enrich_catalog_rows_with_classification

try:
    from opportunity_engine.db_mcp import McpRemoteClient, sanitize_mcp_error
except ImportError:  # pragma: no cover - user-facing scripts add backend/app before importing this module.
    McpRemoteClient = None  # type: ignore[assignment]
    sanitize_mcp_error = None  # type: ignore[assignment]


CATALOG_COLUMNS = [
    "master_sku",
    "name",
    "product_status",
    "pack_size",
    "category_name",
    "category_path",
    "shopify_title",
    "shopify_status",
    "shopify_product_type",
    "shopify_sku",
    "shopify_price",
    "multi_cct",
    "multi_wattage",
    "advertised_lumens",
    "advertised_wattage",
    "cct",
    "voltage",
]

CATALOG_SNAPSHOT_MAX_AGE_HOURS_ENV = "PRODUCT_DEMAND_CATALOG_SNAPSHOT_MAX_AGE_HOURS"
ALLOW_STALE_CATALOG_ENV = "PRODUCT_DEMAND_ALLOW_STALE_CATALOG_CACHE"


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _env_flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "y"}


def _generated_age_hours(generated_at: Any) -> float | None:
    if not generated_at:
        return None
    try:
        timestamp = datetime.fromisoformat(str(generated_at).replace("Z", "+00:00"))
    except ValueError:
        return None
    if timestamp.tzinfo is None:
        timestamp = timestamp.replace(tzinfo=timezone.utc)
    return (datetime.now(timezone.utc) - timestamp).total_seconds() / 3600


def _max_catalog_age_hours() -> float:
    return float(os.environ.get(CATALOG_SNAPSHOT_MAX_AGE_HOURS_ENV) or 24)


def _catalog_payload_is_stale(payload: dict[str, Any]) -> bool:
    if _env_flag("PRODUCT_DEMAND_FORCE_CATALOG_REFRESH"):
        return True
    source = str(payload.get("source_system") or "").lower()
    if not source.startswith("postgres_mcp"):
        return True
    age_hours = _generated_age_hours(payload.get("generated_at"))
    return age_hours is None or age_hours > _max_catalog_age_hours()


def _text(value: Any) -> str:
    return str(value or "").strip()


def _sql_like_term(value: str) -> str:
    return value.replace("%", "").replace("'", "''").lower()


def _like_clause(field: str, terms: tuple[str, ...]) -> str:
    return " or ".join(f"lower(coalesce({field}, '')) like '%{_sql_like_term(term)}%'" for term in terms)


def build_sunco_catalog_sql(category_slug: str, limit: int = 1200) -> str:
    profile = profile_for(category_slug)
    terms = profile.include_terms or tuple(term for term in category_slug.replace("_", " ").split() if len(term) >= 3)
    haystack = "coalesce(c.name,'') || ' ' || coalesce(c.path,'') || ' ' || coalesce(p.name,'') || ' ' || coalesce(sp.title,'')"
    include = _like_clause(haystack, terms) if terms else "1 = 1"
    excludes = profile.exclude_terms
    exclude_clause = "\n".join(
        f"  and lower({haystack}) not like '%{_sql_like_term(term)}%'"
        for term in excludes
    )
    return f"""
select
  p.master_sku,
  p.name,
  p.status as product_status,
  p.pack_size,
  c.name as category_name,
  c.path as category_path,
  sp.title as shopify_title,
  sp.status as shopify_status,
  sp.product_type as shopify_product_type,
  sv.sku as shopify_sku,
  sv.price as shopify_price,
  lp.multi_cct,
  lp.multi_wattage,
  lum.name as advertised_lumens,
  watt.name as advertised_wattage,
  cct.name as cct,
  volt.name as voltage
from products_product p
left join products_category c on c.id = p.category_id
left join shopify_productvariantatshopify sv on sv.product_id = p.id
left join shopify_shopifyproduct sp on sp.id = sv.shopify_product_id
left join products_lightingproduct lp on lp.product_id = p.id
left join product_misc_lumen lum on lum.id = lp.advertised_lumens_id
left join product_misc_wattage watt on watt.id = lp.advertised_wattage_id
left join product_misc_correlatedcolortemperature cct on cct.id = lp.cct_id
left join product_misc_voltage volt on volt.id = lp.voltage_id
where ({include})
{exclude_clause}
order by p.master_sku
limit {int(limit)};
""".strip()


def build_full_sunco_catalog_sql(limit: int = 50000) -> str:
    return f"""
select
  p.master_sku,
  p.name,
  p.status as product_status,
  p.pack_size,
  c.name as category_name,
  c.path as category_path,
  sp.title as shopify_title,
  sp.status as shopify_status,
  sp.product_type as shopify_product_type,
  sv.sku as shopify_sku,
  sv.price as shopify_price,
  lp.multi_cct,
  lp.multi_wattage,
  lum.name as advertised_lumens,
  watt.name as advertised_wattage,
  cct.name as cct,
  volt.name as voltage
from products_product p
left join products_category c on c.id = p.category_id
left join shopify_productvariantatshopify sv on sv.product_id = p.id
left join shopify_shopifyproduct sp on sp.id = sv.shopify_product_id
left join products_lightingproduct lp on lp.product_id = p.id
left join product_misc_lumen lum on lum.id = lp.advertised_lumens_id
left join product_misc_wattage watt on watt.id = lp.advertised_wattage_id
left join product_misc_correlatedcolortemperature cct on cct.id = lp.cct_id
left join product_misc_voltage volt on volt.id = lp.voltage_id
order by p.master_sku
limit {int(limit)};
""".strip()


def latest_catalog_snapshot(exports_dir: Path, category_slug: str) -> Path | None:
    snapshots = sorted(exports_dir.glob(f"{category_slug}_sunco_catalog_coverage_*.json"), key=lambda path: path.stat().st_mtime)
    return snapshots[-1] if snapshots else None


def write_catalog_snapshot(exports_dir: Path, category_slug: str, source_system: str, sql: str, rows: list[dict[str, Any]]) -> Path:
    exports_dir.mkdir(parents=True, exist_ok=True)
    generated_at = utc_now()
    stamp = generated_at.replace("-", "").replace(":", "").split("+", 1)[0].replace("T", "_")
    target = exports_dir / f"{category_slug}_sunco_catalog_coverage_{stamp}.json"
    payload = {
        "source_system": source_system,
        "category_slug": category_slug,
        "generated_at": generated_at,
        "row_count": len(rows),
        "sql": sql,
        "rows": rows,
    }
    target.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return target


def _postgres_mcp_client(timeout_seconds: int) -> McpRemoteClient:
    if McpRemoteClient is None:
        raise RuntimeError("Postgres MCP refresh requires backend/app on sys.path.")
    return McpRemoteClient(timeout_seconds=timeout_seconds, client_name="sunco-product-demand-catalog")


def refresh_catalog_snapshot_via_mcp(exports_dir: Path, category_slug: str, timeout_seconds: int = 240) -> Path:
    sql = build_sunco_catalog_sql(category_slug)
    try:
        with _postgres_mcp_client(timeout_seconds) as client:
            rows = client.execute_sql(sql, timeout_seconds=timeout_seconds)
    except Exception as exc:
        detail = sanitize_mcp_error(exc) if sanitize_mcp_error else str(exc)
        raise RuntimeError(
            "Sunco catalog snapshot refresh requires Postgres MCP access. "
            f"{detail}"
        ) from exc
    return write_catalog_snapshot(exports_dir, category_slug, "postgres_mcp_sunco_catalog_snapshot", sql, rows)


def load_or_refresh_catalog_snapshot(exports_dir: Path, category_slug: str) -> tuple[dict[str, Any], Path]:
    path = latest_catalog_snapshot(exports_dir, category_slug)
    payload = json.loads(path.read_text(encoding="utf-8")) if path else {}
    if path is None or _catalog_payload_is_stale(payload):
        path = refresh_catalog_snapshot_via_mcp(exports_dir, category_slug)
        payload = json.loads(path.read_text(encoding="utf-8"))
    return payload, path


def catalog_cache_path(root: Path) -> Path:
    return root / "product_demand_ideation" / "cache" / "sunco_catalog.sqlite"


def _sqlite_value(value: Any) -> str | int | float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float, str)):
        return value
    return str(value)


def write_catalog_cache(cache_path: Path, rows: list[dict[str, Any]], source_system: str, sql: str) -> Path:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(cache_path)
    try:
        connection.execute("drop table if exists catalog_products")
        connection.execute("drop table if exists cache_metadata")
        connection.execute(
            "create table catalog_products ("
            + ", ".join(f"{column} text" for column in CATALOG_COLUMNS)
            + ")"
        )
        placeholders = ", ".join("?" for _ in CATALOG_COLUMNS)
        connection.executemany(
            "insert into catalog_products values (" + placeholders + ")",
            [
                tuple(_sqlite_value(row.get(column)) for column in CATALOG_COLUMNS)
                for row in rows
            ],
        )
        connection.execute("create table cache_metadata (key text primary key, value text)")
        metadata = {
            "source_system": source_system,
            "generated_at": utc_now(),
            "row_count": str(len(rows)),
            "sql": sql,
        }
        connection.executemany("insert into cache_metadata values (?, ?)", metadata.items())
        connection.execute("create index idx_catalog_master_sku on catalog_products(master_sku)")
        connection.execute("create index idx_catalog_category_name on catalog_products(category_name)")
        connection.execute("create index idx_catalog_product_status on catalog_products(product_status)")
        connection.commit()
    finally:
        connection.close()
    return cache_path


def refresh_catalog_cache_via_mcp(cache_path: Path, timeout_seconds: int = 300) -> Path:
    sql = build_full_sunco_catalog_sql()
    try:
        with _postgres_mcp_client(timeout_seconds) as client:
            rows = client.execute_sql(sql, timeout_seconds=timeout_seconds)
    except Exception as exc:
        detail = sanitize_mcp_error(exc) if sanitize_mcp_error else str(exc)
        raise RuntimeError(
            "Sunco catalog cache refresh requires Postgres MCP access. "
            f"{detail}"
        ) from exc
    return write_catalog_cache(cache_path, rows, "postgres_mcp_local_sqlite_sunco_catalog_cache", sql)


def _cache_metadata(cache_path: Path) -> dict[str, str]:
    connection = sqlite3.connect(cache_path)
    try:
        return dict(connection.execute("select key, value from cache_metadata").fetchall())
    finally:
        connection.close()


def _catalog_cache_is_stale(cache_path: Path) -> bool:
    try:
        metadata = _cache_metadata(cache_path)
    except Exception:
        return True
    return _catalog_payload_is_stale(metadata)


def _profile_text(row: dict[str, Any]) -> str:
    return " ".join(
        _text(row.get(key))
        for key in [
            "name",
            "category_name",
            "category_path",
            "shopify_title",
            "shopify_product_type",
            "master_sku",
            "classification_category",
            "classification_pm_responsible",
            "classification_series",
        ]
    ).lower()


def _matches_category_profile(row: dict[str, Any], category_slug: str) -> bool:
    profile = profile_for(category_slug)
    text = _profile_text(row)
    product_type = _text(row.get("shopify_product_type")).lower()
    if profile.exclude_terms and any(term.lower() in text for term in profile.exclude_terms):
        return False
    product_type_match = bool(profile.product_types) and product_type in {item.lower() for item in profile.product_types}
    term_match = bool(profile.include_terms) and any(term.lower() in text for term in profile.include_terms)
    fallback_terms = tuple(term for term in category_slug.replace("_", " ").split() if len(term) >= 3)
    fallback_match = not profile.include_terms and any(term.lower() in text for term in fallback_terms)
    return product_type_match or term_match or fallback_match


def load_catalog_rows_from_cache(cache_path: Path, category_slug: str) -> list[dict[str, Any]]:
    connection = sqlite3.connect(cache_path)
    connection.row_factory = sqlite3.Row
    try:
        rows = [dict(row) for row in connection.execute("select * from catalog_products").fetchall()]
    finally:
        connection.close()
    return [row for row in rows if _matches_category_profile(row, category_slug)]


def load_all_catalog_rows_from_cache(cache_path: Path) -> list[dict[str, Any]]:
    connection = sqlite3.connect(cache_path)
    connection.row_factory = sqlite3.Row
    try:
        return [dict(row) for row in connection.execute("select * from catalog_products").fetchall()]
    finally:
        connection.close()


def load_catalog_context_from_cache_or_snapshot(root: Path, category_slug: str) -> tuple[list[dict[str, Any]], Path | None, str]:
    cache_path = catalog_cache_path(root)
    cache_missing = not cache_path.exists()
    cache_stale = False if cache_missing else _catalog_cache_is_stale(cache_path)
    if cache_missing or cache_stale:
        try:
            refresh_catalog_cache_via_mcp(cache_path)
        except Exception:
            if cache_path.exists() and _env_flag(ALLOW_STALE_CATALOG_ENV):
                pass
            elif cache_path.exists() and cache_stale:
                raise
            else:
                exports = root / "product_demand_ideation" / "experiments" / category_slug / "exports"
                snapshot, snapshot_path = load_or_refresh_catalog_snapshot(exports, category_slug)
                rows, classification_path, classification_source = enrich_catalog_rows_with_classification(root, list(snapshot.get("rows") or []))
                source = str(snapshot.get("source_system") or "unknown")
                if classification_source:
                    source = f"{source}; SKU classification source: {classification_source}"
                    return rows, classification_path or snapshot_path, source
                return rows, snapshot_path, source
        if not cache_path.exists():
            exports = root / "product_demand_ideation" / "experiments" / category_slug / "exports"
            snapshot, snapshot_path = load_or_refresh_catalog_snapshot(exports, category_slug)
            rows, classification_path, classification_source = enrich_catalog_rows_with_classification(root, list(snapshot.get("rows") or []))
            source = str(snapshot.get("source_system") or "unknown")
            if classification_source:
                source = f"{source}; SKU classification source: {classification_source}"
                return rows, classification_path or snapshot_path, source
            return rows, snapshot_path, source
    if category_slug == "smart_lighting":
        rows = load_all_catalog_rows_from_cache(cache_path)
        rows, classification_path, classification_source = enrich_catalog_rows_with_classification(root, rows)
        rows = [row for row in rows if _matches_category_profile(row, category_slug)]
    else:
        rows = load_catalog_rows_from_cache(cache_path, category_slug)
        rows, classification_path, classification_source = enrich_catalog_rows_with_classification(root, rows)
    metadata = _cache_metadata(cache_path)
    source = metadata.get("source_system") or "local_sqlite_sunco_catalog_cache"
    generated_at = metadata.get("generated_at")
    if generated_at:
        source = f"{source} generated_at={generated_at}"
    if classification_source:
        source = f"{source}; SKU classification source: {classification_source}"
    return rows, cache_path, source


def _tokens(value: str) -> set[str]:
    raw = re.findall(r"[a-z0-9]+", value.lower())
    ignored = {"led", "light", "lights", "fixture", "fixtures", "for", "and", "with", "the", "max", "selectable"}
    return {token for token in raw if len(token) >= 2 and token not in ignored}


def _size_signal(text: str) -> str | None:
    normalized = text.lower().replace("×", "x")
    match = re.search(r"\b([124])\s*x\s*([124])\b", normalized)
    return f"{match.group(1)}x{match.group(2)}" if match else None


def _pack_size_signal(text: str) -> int | None:
    normalized = text.lower()
    match = re.search(r"\b(\d{1,3})\s*(?:pack|pk|pc|piece|pieces)\b", normalized)
    if match:
        return int(match.group(1))
    match = re.search(r"\b(\d{1,3})\s*-\s*(?:pack|pk|pc)\b", normalized)
    return int(match.group(1)) if match else None


def _model_tokens(text: str) -> set[str]:
    tokens = re.findall(r"\b[A-Z0-9]+(?:[_-][A-Z0-9]+){1,}\b", text.upper())
    ignored = {"120-277V", "0-10V"}
    return {token for token in tokens if token not in ignored and len(token) >= 4}


def _cct_set(text: str) -> set[str]:
    values = re.findall(r"\b(?:2700|3000|3500|4000|4500|5000|5700|6000|6500)\s*k\b|\b(?:27|30|35|40|45|50|57|60|65)00\s*k\b", text, flags=re.IGNORECASE)
    return {value.lower().replace(" ", "").upper() for value in values}


def _feature_set(text: str) -> set[str]:
    raw = text.lower()
    features = set()
    if any(term in raw for term in ["emergency", "battery backup", "em backup", "90 minute"]):
        features.add("emergency backup")
    if any(term in raw for term in ["motion sensor", "occupancy sensor", "pir sensor", "microwave sensor"]):
        features.add("motion sensor")
    if any(term in raw for term in ["daylight harvesting", "daylight sensor", "ambient light sensor", "photocell"]):
        features.add("daylight harvesting")
    if any(term in raw for term in ["sensor ready", "control ready", "controls ready", "c-max", "control-ready"]):
        features.add("control ready")
    if any(term in raw for term in ["bluetooth mesh", "mesh", "wireless control", "smart control"]):
        features.add("smart controls")
    if "0-10v" in raw or "0-10 volt" in raw:
        features.add("0-10v")
    if "triac" in raw:
        features.add("triac")
    if "dimmable" in raw or "dimming" in raw:
        features.add("dimmable")
    return features


def _assortment_cues(text: str) -> set[str]:
    raw = text.lower()
    cues = set()
    checks = [
        ("combo exit sign with heads", ("exit sign", "head")),
        ("exit sign", ("exit sign",)),
        ("bug-eye emergency light", ("bug-eye", "bug eye", "adjustable led lamp heads")),
        ("edge-lit", ("edge-lit", "edge lit")),
        ("red lettering", ("red",)),
        ("green lettering", ("green",)),
        ("black housing", ("black",)),
        ("white housing", ("white",)),
        ("wet/damp rated", ("wet location", "wet rated", "damp rated", "damp location")),
        ("thermoplastic/ABS housing", ("thermoplastic", "abs")),
        ("battery backup", ("battery", "backup", "90-minute", "90 minute", "90 min")),
        ("UL listed", ("ul listed", " ul ")),
    ]
    for label, terms in checks:
        if label == "combo exit sign with heads":
            if all(term in raw for term in terms):
                cues.add(label)
        elif any(term in raw for term in terms):
            cues.add(label)
    return cues


TECHNOLOGY_GAP_FEATURES = {
    "emergency backup",
    "motion sensor",
    "daylight harvesting",
    "control ready",
    "smart controls",
}


def _catalog_text(row: dict[str, Any]) -> str:
    return " ".join(
        _text(row.get(key))
        for key in [
            "master_sku",
            "name",
            "category_name",
            "shopify_title",
            "shopify_product_type",
            "advertised_lumens",
            "advertised_wattage",
            "cct",
            "voltage",
            "classification_category",
            "classification_pm_responsible",
            "classification_series",
        ]
    )


def _candidate_text(row: dict[str, Any]) -> str:
    return " ".join(
        _text(row.get(key))
        for key in [
            "subcategory",
            "recommendation",
            "example",
            "evidence",
            "why_gap",
            "pm_action",
            "action",
        ]
    )


def _is_active_catalog_row(row: dict[str, Any]) -> bool:
    product_status = _text(row.get("product_status")).lower()
    shopify_status = _text(row.get("shopify_status")).lower()
    return product_status == "active" or shopify_status == "active"


def _match_score(candidate: dict[str, Any], catalog_row: dict[str, Any]) -> float:
    candidate_text = _candidate_text(candidate)
    catalog_text = _catalog_text(catalog_row)
    candidate_perf = normalize_luminaire_performance(candidate_text)
    catalog_perf = normalize_luminaire_performance(catalog_text)
    score = 0.0

    master_sku = _text(catalog_row.get("master_sku")).upper()
    if master_sku and master_sku in candidate_text.upper():
        score += 55.0
    else:
        model_overlap = _model_tokens(candidate_text) & _model_tokens(catalog_text)
        if model_overlap:
            score += 30.0

    candidate_size = _size_signal(candidate_text)
    catalog_size = _size_signal(catalog_text)
    if candidate_size and catalog_size and candidate_size == catalog_size:
        score += 25.0

    candidate_pack = _pack_size_signal(candidate_text)
    catalog_pack = _pack_size_signal(catalog_text) or int(_text(catalog_row.get("pack_size")) or 0)
    if candidate_pack and catalog_pack:
        if candidate_pack == catalog_pack:
            score += 15.0
        elif min(candidate_pack, catalog_pack) / max(candidate_pack, catalog_pack) >= 0.5:
            score += 6.0

    if candidate_perf.lumen_values and catalog_perf.lumen_values:
        candidate_lm = max(candidate_perf.lumen_values)
        catalog_lm = max(catalog_perf.lumen_values)
        ratio = min(candidate_lm, catalog_lm) / max(candidate_lm, catalog_lm)
        score += min(25.0, ratio * 25.0)
        if candidate_perf.brightness_tier == catalog_perf.brightness_tier:
            score += 10.0

    if candidate_perf.watt_values and catalog_perf.watt_values:
        candidate_w = max(candidate_perf.watt_values)
        catalog_w = max(catalog_perf.watt_values)
        ratio = min(candidate_w, catalog_w) / max(candidate_w, catalog_w)
        score += min(15.0, ratio * 15.0)
        if candidate_perf.wattage_class == catalog_perf.wattage_class:
            score += 5.0

    candidate_cct = _cct_set(candidate_text)
    catalog_cct = _cct_set(catalog_text)
    if candidate_cct and catalog_cct:
        overlap = len(candidate_cct & catalog_cct) / max(1, len(candidate_cct))
        score += min(10.0, overlap * 10.0)

    candidate_features = _feature_set(candidate_text)
    catalog_features = _feature_set(catalog_text)
    if candidate_features:
        score += min(10.0, (len(candidate_features & catalog_features) / len(candidate_features)) * 10.0)

    candidate_cues = _assortment_cues(candidate_text)
    catalog_cues = _assortment_cues(catalog_text)
    if candidate_cues and catalog_cues:
        score += min(35.0, (len(candidate_cues & catalog_cues) / len(candidate_cues)) * 35.0)

    candidate_tokens = _tokens(candidate_text)
    catalog_tokens = _tokens(catalog_text)
    if candidate_tokens and catalog_tokens:
        score += min(10.0, (len(candidate_tokens & catalog_tokens) / max(5, len(candidate_tokens))) * 10.0)

    final_score = min(100.0, score)
    if _text(catalog_row.get("classification_series")).lower() == "nsl":
        # NSL rows are limited-line coverage and should not fully clear a new-SKU gap by themselves.
        final_score = min(final_score, 62.0)
    return round(final_score, 1)


def best_catalog_match(candidate: dict[str, Any], catalog_rows: list[dict[str, Any]]) -> tuple[dict[str, Any] | None, float]:
    active_rows = [row for row in catalog_rows if _is_active_catalog_row(row)]
    best_row: dict[str, Any] | None = None
    best_score = 0.0
    for row in active_rows:
        score = _match_score(candidate, row)
        if score > best_score:
            best_row = row
            best_score = score
    return best_row, best_score


def catalog_coverage_analysis(candidate: dict[str, Any], catalog_rows: list[dict[str, Any]], category_slug: str | None = None) -> dict[str, Any]:
    match, score = best_catalog_match(candidate, catalog_rows)
    candidate_features = _feature_set(_candidate_text(candidate))
    catalog_features = _feature_set(_catalog_text(match or {}))
    gap_features = set(TECHNOLOGY_GAP_FEATURES)
    if category_slug == "emergency":
        gap_features.discard("emergency backup")
    missing_features = sorted((candidate_features - catalog_features) & gap_features)
    note, _ = catalog_coverage_note(candidate, catalog_rows)
    return {
        "note": note,
        "score": score,
        "match": match,
        "candidate_features": sorted(candidate_features),
        "catalog_features": sorted(catalog_features),
        "missing_features": missing_features,
        "has_technology_gap": bool(missing_features),
    }


def catalog_coverage_note(candidate: dict[str, Any], catalog_rows: list[dict[str, Any]]) -> tuple[str, float]:
    match, score = best_catalog_match(candidate, catalog_rows)
    if not match or score < 50:
        return (
            "Active title match: 0 strong Sunco active-catalog exact-spec matches found from Postgres product/catalog snapshot. "
            "Treat as likely Shopify assortment gap based on current active-catalog coverage.",
            score,
        )

    match_label = f"{_text(match.get('master_sku'))}: {_text(match.get('name'))}"
    status = _text(match.get("product_status")) or _text(match.get("shopify_status")) or "active"
    is_limited_series = _text(match.get("classification_series")).lower() == "nsl"
    classification_context = ""
    if match.get("classification_category") or match.get("classification_pm_responsible") or match.get("classification_series"):
        parts = []
        if match.get("classification_category"):
            parts.append(f"PowerBI category: {_text(match.get('classification_category'))}")
        if match.get("classification_pm_responsible"):
            parts.append(f"PM: {_text(match.get('classification_pm_responsible'))}")
        if match.get("classification_series"):
            parts.append(f"Series: {_text(match.get('classification_series'))}")
        classification_context = " " + "; ".join(parts) + "."
    if score >= 75:
        return (
            f"Sunco likely already has comparable active coverage ({score}/100 match): {match_label}. "
            f"Status: {status}.{classification_context} Treat as merchandising/listing-depth opportunity, not a confirmed new-product launch need.",
            score,
        )
    if is_limited_series:
        return (
            f"Limited-line Sunco coverage found ({score}/100 match): {match_label}. "
            f"Status: {status}.{classification_context} NSL/limited-line coverage does not fully clear the assortment gap; consider a standard Sunco line when demand evidence supports this spec.",
            score,
        )
    return (
        f"Partial Sunco active coverage found ({score}/100 match): {match_label}. "
        f"Status: {status}.{classification_context} Treat as partial coverage; the missing spec appears to be size/output/CCT/feature depth rather than a fully covered active product.",
        score,
    )
