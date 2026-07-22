from __future__ import annotations

import json
import os
import re
from html import unescape
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from luminaire_performance import normalize_luminaire_performance
from redshift_query import execute_redshift_sql, is_redshift_mcp_source


@dataclass(frozen=True)
class CategoryProfile:
    product_types: tuple[str, ...]
    include_terms: tuple[str, ...]
    exclude_terms: tuple[str, ...] = ()
    include_description: bool = False


STRATEGIC_OUTLIER_LIMIT = 2
HIGH_OUTPUT_PANEL_LUMEN_THRESHOLD = 15_000
HIGH_OUTPUT_PANEL_WATT_THRESHOLD = 100


DEFAULT_EXCLUDE_TERMS = (
    "accessory",
    "accessor",
    "adapter",
    "adaptor",
    "ballast",
    "bracket",
    "clip",
    "face plate",
    "driver",
    "frame kit",
    "junction box",
    "mounting kit",
    "replacement lens",
)

SMART_EXCLUDE_TERMS = DEFAULT_EXCLUDE_TERMS + (
    "advance smartmate",
    "smartmate",
    "ballast",
    "control module",
    "controller",
    "dimming module",
    "fluorescent",
    "incandescent",
    "replacement",
    "remote",
)

FIXTURE_EXCLUDE_TERMS = DEFAULT_EXCLUDE_TERMS + (
    "a19",
    "appliance bulb",
    "br30",
    "bulb",
    "chandelier bulb",
    "filament bulb",
    "hid retrofit",
    "lamp",
    "lampholder",
    "mr16",
    "pigtail",
    "par30",
    "par38",
    "pl lamp",
    "socket",
    "tube",
)

ELECTRICAL_EXCLUDE_TERMS = DEFAULT_EXCLUDE_TERMS + (
    "bulb",
    "fixture",
    "lamp",
    "light bulb",
    "lumens",
    "watt",
)

GROW_LIGHT_EXCLUDE_TERMS = DEFAULT_EXCLUDE_TERMS + (
    "capacitor",
    "controller",
    "control",
    "controls",
    "digital programmable timer",
    "hid lighting",
    "high pressure sodium",
    "hps",
    "cfl",
    "fluorescent",
    "incandescent",
    "metal halide",
    "timer",
    "outlet",
    "nutrient",
    "fertilizer",
    "hydroponic",
    "tent",
    "tray",
    "rope ratchet",
    "hanger",
    "reflector",
    "accessory",
)

COMMERCIAL_GROW_LIGHT_EXCLUDE_TERMS = GROW_LIGHT_EXCLUDE_TERMS + (
    "a19",
    "br30",
    "bulb",
    "e26",
    "lamp - br",
    "lamp - par",
    "lamp - type",
    "par38",
)

GROW_LIGHT_STRONG_TERMS = (
    "grow light",
    "grow lights",
    "grow lamp",
    "grow bulb",
    "growing lamp",
    "plant light",
    "plant lights",
    "plant grow",
    "horticulture led",
    "horticulture luminaire",
    "horticulture lamp",
    "horticulture vapor tight",
    "horticulture top light",
    "horticulture starlight",
    "led horticulture",
)

GROW_LIGHT_SUPPORT_TERMS = (
    "full spectrum",
    "seed starting",
    "veg",
    "vegetable",
    "bloom",
    "greenhouse",
    "ppfd",
    " ppf ",
    "umol",
    "par efficacy",
    "indoor plants",
    "indoor growing",
)

GROW_LIGHT_ALWAYS_REJECT_TERMS = (
    "surge protector",
    "power strip",
    "oil filled capacitor",
    "capacitor",
    "cfl",
    "fluorescent",
    "high pressure sodium",
    "reflector",
)

GROW_LIGHT_IDENTITY_REJECT_TERMS = (
    "foundry lighting",
    "high temperature led",
    "high temperature light",
    "high temperature lights",
    "power plant lighting",
    "steel plant lighting",
)

GROW_LIGHT_CONTEXT_REJECT_TERMS = (
    "outlet timer",
    "mechanical timer",
    "programmable timer",
    "nutrient",
    "fertilizer",
    "rope ratchet",
    "hanger",
    "replacement cord",
)


CATEGORY_PROFILES: dict[str, CategoryProfile] = {
    "area_lights": CategoryProfile(("area light",), (), FIXTURE_EXCLUDE_TERMS + ("slipfitter mount", "tenon")),
    "bathroom_fans": CategoryProfile((), ("bathroom fan", "exhaust fan", "ventilation fan", "cfm", "sone"), FIXTURE_EXCLUDE_TERMS + ("rv", "trailer", "vehicle", "porch", "utility light")),
    "bulbs_plus_tubes": CategoryProfile(("bulb", "tube"), ("bulb", "lamp", "tube", "t8", "a19", "br30", "par")),
    "canopy": CategoryProfile(("canopy",), ("canopy", "gas station"), FIXTURE_EXCLUDE_TERMS),
    "ceiling_fixtures": CategoryProfile(("surface mount",), ("flush mount", "ceiling fixture", "ceiling light"), FIXTURE_EXCLUDE_TERMS + ("emergency vehicle", "hideaway strobe")),
    "chandeliers": CategoryProfile(("chandelier",), ("chandelier fixture", "chandelier light", "chandeliers"), FIXTURE_EXCLUDE_TERMS),
    "clean_room_lighting": CategoryProfile((), ("clean room", "cleanroom"), FIXTURE_EXCLUDE_TERMS),
    "commercial_fans": CategoryProfile((), ("commercial fan", "ceiling fan", "industrial fan", "destratification fan"), FIXTURE_EXCLUDE_TERMS),
    "commercial_grow_lights": CategoryProfile(
        ("grow light",),
        (
            "grow light",
            "grow fixture",
            "grow lamp",
            "horticulture led",
            "horticulture luminaire",
            "horticulture light",
            "horticulture",
            "plant light",
            "full spectrum plant",
        ),
        COMMERCIAL_GROW_LIGHT_EXCLUDE_TERMS,
        True,
    ),
    "commercial_landscape": CategoryProfile(("pathway/landscape", "bollard"), ("landscape fixture", "path light", "pathway light", "bollard"), FIXTURE_EXCLUDE_TERMS + ("deck light", "step light")),
    "commercial_security": CategoryProfile(("security", "flood light"), ("security", "wall light"), FIXTURE_EXCLUDE_TERMS + ("lampholder",)),
    "dimmers": CategoryProfile(("dimmer/switch",), ("dimmer", "switch", "occupancy sensor", "motion sensor"), ELECTRICAL_EXCLUDE_TERMS),
    "emergency": CategoryProfile(("emergency/exit",), ("emergency", "exit sign", "battery backup", "90 min"), FIXTURE_EXCLUDE_TERMS + ("vehicle", "strobe")),
    "explosion_proof": CategoryProfile(("explosion proof",), ("explosion proof", "hazardous location", "class 1 div"), FIXTURE_EXCLUDE_TERMS + ("flame bulb",)),
    "flood_lights": CategoryProfile(("flood light",), ("flood light", "security flood"), FIXTURE_EXCLUDE_TERMS),
    "lamps": CategoryProfile((), ("table lamp", "floor lamp", "desk lamp"), DEFAULT_EXCLUDE_TERMS + ("dimming module", "appliance module", "plug-in module")),
    "led_ready_fixtures": CategoryProfile((), ("led ready", "ready fixture", "lamp ready"), FIXTURE_EXCLUDE_TERMS),
    "linears": CategoryProfile(
        (),
        (
            "linear high bay",
            "recessed linear",
            "architectural led linear",
            "architectural linear",
            "linear fixture",
        ),
        FIXTURE_EXCLUDE_TERMS
        + (
            "channel",
            "extrusion",
            "led strip",
            "power strip",
            "shop light",
            "strip fixture",
            "strip light",
            "tape light",
            "rope light",
            "connector",
        ),
        True,
    ),
    "low_voltage_transformers": CategoryProfile((), ("low voltage transformer", "12v transformer", "24v transformer", "landscape transformer"), ELECTRICAL_EXCLUDE_TERMS),
    "outdoor_ceiling": CategoryProfile(("surface mount",), ("outdoor ceiling", "porch ceiling", "ceiling mount outdoor"), FIXTURE_EXCLUDE_TERMS + ("emergency vehicle", "hideaway strobe")),
    "outdoor_security": CategoryProfile(("security", "flood light"), ("security", "motion flood", "flood light"), FIXTURE_EXCLUDE_TERMS + ("lampholder",)),
    "panels": CategoryProfile(
        ("panel", "troffer"),
        ("panel", "troffer", "flat panel"),
        DEFAULT_EXCLUDE_TERMS
        + (
            "automotive",
            "charging port",
            "downlight",
            "fluorescent",
            "flush mount",
            "lampholder",
            "magnetic strip",
            "photocell",
            "rocker switch",
            "sensor",
            "socket",
            "trailer",
            "tube",
            "usb",
            "vehicle",
        ),
        False,
    ),
    "pendants": CategoryProfile(("pendant",), ("pendant fixture", "pendant light", "pendants"), FIXTURE_EXCLUDE_TERMS),
    "residential_fans": CategoryProfile((), ("ceiling fan", "residential fan", "fan light kit"), FIXTURE_EXCLUDE_TERMS),
    "residential_grow_lights": CategoryProfile(
        ("grow light",),
        (
            "grow light",
            "grow bulb",
            "plant light",
            "plant lamp",
            "grow lamp",
            "plant grow",
            "full spectrum plant",
            "seed starting light",
        ),
        GROW_LIGHT_EXCLUDE_TERMS,
        True,
    ),
    "residential_landscape": CategoryProfile(("pathway/landscape", "bollard"), ("landscape fixture", "path light", "pathway light", "garden light"), FIXTURE_EXCLUDE_TERMS + ("deck light", "step light")),
    "residential_security": CategoryProfile(("security", "flood light"), ("security", "motion flood", "flood light"), FIXTURE_EXCLUDE_TERMS + ("lampholder",)),
    "retros": CategoryProfile((), ("retrofit downlight", "downlight retrofit", "recessed retrofit", "gimbal", "eyeball"), FIXTURE_EXCLUDE_TERMS + ("track light",)),
    "rough_ins": CategoryProfile((), ("rough in", "rough-in", "new construction frame", "mounting frame"), FIXTURE_EXCLUDE_TERMS),
    "sensors": CategoryProfile(("sensor",), ("occupancy sensor", "vacancy sensor", "motion sensor", "photocell"), ELECTRICAL_EXCLUDE_TERMS + ("controller", "remote", "dimming module")),
    "slims": CategoryProfile((), ("slim", "wafer", "canless"), FIXTURE_EXCLUDE_TERMS + ("panel", "troffer", "gimbal")),
    "solar": CategoryProfile((), ("solar flood", "solar area", "solar wall", "solar roadway", "solar powered led"), FIXTURE_EXCLUDE_TERMS),
    "smart_lighting": CategoryProfile(
        (),
        (
            "smart",
            "smart lighting",
            "smart led",
            "smart wifi",
            "smart wi-fi",
            "wifi led",
            "wi-fi led",
            "bluetooth led",
            "rgbw",
            "rgb+w",
            "alexa",
            "google assistant",
            "app controlled",
            "app-control",
            "matter compatible",
        ),
        SMART_EXCLUDE_TERMS,
    ),
    "shop_light": CategoryProfile(
        (),
        ("shop light", "shop lights", "garage light", "garage lights", "utility shop light", "linkable shop light"),
        FIXTURE_EXCLUDE_TERMS
        + (
            "architectural",
            "channel",
            "downlight",
            "emergency",
            "flood",
            "grow",
            "outdoor",
            "power strip",
            "recessed",
            "sconce",
            "security",
            "solar",
            "stairwell",
            "step light",
            "strip fixture",
            "under cabinet",
            "wall light",
        ),
        True,
    ),
    "sport_lights": CategoryProfile(("sport light",), ("sport light", "stadium light", "field light"), FIXTURE_EXCLUDE_TERMS + ("slipfitter mount",)),
    "striplights": CategoryProfile(
        ("strip/linear",),
        ("strip fixture", "strip light fixture", "linear strip", "general purpose strip"),
        FIXTURE_EXCLUDE_TERMS
        + (
            "architectural",
            "channel",
            "cob led strip",
            "cove",
            "extrusion",
            "light bar",
            "power strip",
            "rope light",
            "spool",
            "stairwell",
            "tape light",
            "under cabinet",
            "vapor tight",
            "vaportight",
            "wall wash",
        ),
    ),
    "string_lights": CategoryProfile(("string light",), ("string light",)),
    "tape_rope_light": CategoryProfile((), ("tape light", "rope light", "led strip"), DEFAULT_EXCLUDE_TERMS + ("linear fixture", "strip fixture", "track", "connector", "pigtail", "light bar", "power cable", "power cord", "channel", "raceway")),
    "ufo": CategoryProfile(
        (),
        ("ufo", "ufo high bay", "ufo led high bay", "ufo lighting"),
        FIXTURE_EXCLUDE_TERMS
        + (
            "linear high bay",
            "linear led high bay",
            "led ready",
            "lamp ready",
            "t8",
        ),
        True,
    ),
    "under_cabinet": CategoryProfile(("under cabinet",), ("under cabinet", "task lighting")),
    "vanity": CategoryProfile(("vanity & wall - linear vanity", "sconce"), ("vanity",)),
    "vaportights": CategoryProfile(("vapor tight",), ("vapor tight", "vaportight")),
    "wall_packs": CategoryProfile(("wall pack",), ("wall pack",)),
    "wall_sconces": CategoryProfile(("sconce",), ("sconce", "wall light")),
    "well_lights": CategoryProfile((), ("well light", "in-ground", "in ground"), FIXTURE_EXCLUDE_TERMS + ("automotive", "wedge bulb", "deck light", "step light")),
    "wire": CategoryProfile((), ("electrical wire", "building wire", "fixture wire", "low voltage wire", "cable spool"), ELECTRICAL_EXCLUDE_TERMS + ("track", "connector", "sensor", "strip", "license plate", "light bar", "hardwired", "dimmer", "wireless")),
    "wire_connectors": CategoryProfile((), ("wire connector", "pigtail connector", "cable connector"), ELECTRICAL_EXCLUDE_TERMS + ("track", "trailer", "light bar")),
    "wraparounds": CategoryProfile((), ("wraparound", "wrap around"), FIXTURE_EXCLUDE_TERMS),
    "cans": CategoryProfile((), ("recessed can", "can light", "ceiling can"), FIXTURE_EXCLUDE_TERMS),
}


def _is_grow_light_category(category_name: str) -> bool:
    normalized = category_name.lower().replace("_", " ")
    return "grow light" in normalized or "grow lights" in normalized


def _grow_light_text(row: dict[str, Any]) -> str:
    return " ".join(
        _text(row.get(key))
        for key in ["name", "description", "category", "product_type", "brand", "sku", "url"]
    ).lower()


def is_grow_light_candidate(row: dict[str, Any]) -> bool:
    text = f" {_grow_light_text(row)} "
    identity_text = " ".join(
        _text(row.get(key))
        for key in ["name", "category", "product_type", "brand", "sku", "url"]
    ).lower()
    if any(term in identity_text for term in GROW_LIGHT_IDENTITY_REJECT_TERMS):
        return False
    if any(term in text for term in GROW_LIGHT_ALWAYS_REJECT_TERMS):
        return False
    if any(term in text for term in ["grow bulb", "grow bulbs", "plant grow lamps"]) and " led " not in text:
        return False
    has_strong = any(term in text for term in GROW_LIGHT_STRONG_TERMS)
    if any(term in text for term in GROW_LIGHT_CONTEXT_REJECT_TERMS) and not has_strong:
        return False
    if has_strong:
        return True
    has_support = any(term in text for term in GROW_LIGHT_SUPPORT_TERMS)
    has_plant_context = any(term in text for term in [" plant", " plants", " grow", " growing", " seed", " bloom", " ppfd", " ppf ", " par "])
    return has_support and has_plant_context and any(
        light_term in text for light_term in [" led ", " lamp", " bulb", "fixture", "light", "luminaire"]
    )


def is_residential_grow_light_candidate(row: dict[str, Any]) -> bool:
    text = f" {_grow_light_text(row)} "
    perf = normalize_luminaire_performance(
        " ".join(_text(row.get(key)) for key in ["name", "description", "wattage"])
    )
    max_watt = max(perf.watt_values, default=0)
    small_linear_form = any(
        term in text
        for term in [
            "seed starting",
            "shelf",
            "clip",
            "desk",
            "under cabinet",
            "shop light",
            "strip",
            "t5",
            "t8",
            "2ft",
            "2 ft",
            "4ft",
            "4 ft",
        ]
    )
    bulb_or_lamp_form = any(term in text for term in [" bulb", " lamp", "br30", "par38"])
    if max_watt > 150 and not small_linear_form:
        return False
    if small_linear_form or bulb_or_lamp_form:
        return True
    return max_watt == 0 or max_watt <= 150


def is_shop_light_candidate(row: dict[str, Any]) -> bool:
    name = _text(row.get("name")).lower()
    product_type = _text(row.get("product_type")).lower()
    category = _text(row.get("category")).lower()
    full_text = f" {name} {product_type} {category} "
    reject_terms = (
        "canopy",
        "cob led strip",
        "downlight",
        "emergency",
        "flood",
        "picture light",
        "recessed",
        "sconce",
        "security",
        "solar",
        "stairwell",
        "step light",
        "strip light spool",
        "ufo high bay",
        "under cabinet",
        "wall light",
    )
    if any(term in full_text for term in reject_terms):
        return False
    if "linear high bay" in full_text and "shop light" not in name:
        return False
    return (
        "shop light" in name
        or "shop strip light" in name
        or "shop light" in product_type
        or ("wraparound" in name and "shop light" in category)
    )


def is_emergency_candidate(row: dict[str, Any]) -> bool:
    name = _text(row.get("name")).lower()
    product_type = _text(row.get("product_type")).lower()
    category = _text(row.get("category")).lower()
    full_text = f" {name} {product_type} {category} "
    reject_terms = (
        "canopy",
        "flat panel",
        "high bay",
        "panel fixture",
        "panel light",
        "shop light",
        "stairwell",
        "strip fixture",
        "troffer",
        "vapor tight",
        "wall pack",
    )
    if any(term in full_text for term in reject_terms):
        return False
    allow_terms = (
        "bug-eye",
        "bug eye",
        "combo exit",
        "emergency battery backup micro inverter",
        "emergency inverter",
        "emergency light",
        "emergency lights",
        "exit sign",
        "remote head",
    )
    return product_type == "emergency/exit" or any(term in full_text for term in allow_terms)


def is_striplights_candidate(row: dict[str, Any]) -> bool:
    name = _text(row.get("name")).lower()
    product_type = _text(row.get("product_type")).lower()
    category = _text(row.get("category")).lower()
    full_text = f" {name} {product_type} {category} "
    reject_terms = (
        "architectural",
        "channel",
        "cob led strip",
        "cove",
        "extrusion",
        "light bar",
        "power strip",
        "rope light",
        "spool",
        "stairwell",
        "tape light",
        "under cabinet",
        "vapor tight",
        "vaportight",
        "wall wash",
    )
    if any(term in full_text for term in reject_terms):
        return False
    return (
        "strip fixture" in name
        or "strip light fixture" in name
        or ("strip light" in name and "fixture" in full_text)
        or ("general purpose strip" in full_text and "fixture" in full_text)
        or ("linear strip" in full_text and "fixture" in full_text)
        or (product_type == "strip/linear" and "fixture" in full_text and "strip" in full_text)
    )


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _slug_terms(category_slug: str) -> tuple[str, ...]:
    ignored = {"led", "light", "lights", "lighting", "fixture", "fixtures", "commercial", "residential"}
    return tuple(term for term in category_slug.replace("_", " ").split() if len(term) >= 3 and term not in ignored)


def profile_for(category_slug: str) -> CategoryProfile:
    return CATEGORY_PROFILES.get(category_slug) or CategoryProfile((), _slug_terms(category_slug), DEFAULT_EXCLUDE_TERMS)


def _sql_literal(value: str) -> str:
    return "'" + value.replace("'", "''").lower() + "'"


def _sql_like_term(value: str) -> str:
    return value.replace("%", "").replace("'", "''").lower()


def _like_clause(field: str, terms: tuple[str, ...]) -> str:
    return " or ".join(f"lower(coalesce({field}, '')) like '%{_sql_like_term(term)}%'" for term in terms)


def build_ecommerce_sql(category_slug: str, limit: int = 600) -> str:
    profile = profile_for(category_slug)
    filters: list[str] = []
    if profile.product_types:
        values = ", ".join(_sql_literal(value) for value in profile.product_types)
        filters.append(f"lower(coalesce(latest.product_type, '')) in ({values})")
    if profile.include_terms:
        include_fields = ["latest.name", "latest.category"]
        if profile.include_description:
            include_fields.append("latest.description")
        filters.append("(" + " or ".join(_like_clause(field, profile.include_terms) for field in include_fields) + ")")
    where_clause = " or ".join(filters) or "1 = 1"
    excludes = profile.exclude_terms or DEFAULT_EXCLUDE_TERMS
    exclude_clause = "\n".join(
        f"    and lower(coalesce(latest.name, '') || ' ' || coalesce(latest.category, '') || ' ' || coalesce(latest.product_type, '')) not like '%{_sql_like_term(term)}%'"
        for term in excludes
    )
    return f"""
with movement as (
  select
    url,
    max(scrape_date) as latest_inventory_scrape_date,
    max(stock_qty) as latest_stock_qty,
    max(price) as latest_inventory_price,
    sum(case when stock_qty_delta < 0 then abs(stock_qty_delta) else 0 end) as observed_stock_decrease,
    sum(case when stock_qty_delta > 0 then stock_qty_delta else 0 end) as observed_restock,
    count(case when stock_qty_delta < 0 then 1 end) as decrease_events,
    count(case when stock_qty_delta > 0 then 1 end) as restock_events,
    count(*) as observation_count,
    min(scrape_date) as first_observed_date,
    max(scrape_date) as last_observed_date,
    datediff(day, min(scrape_date), max(scrape_date)) + 1 as observation_window_days,
    min(case when stock_qty_delta < 0 then scrape_date end) as first_decrease_date,
    max(case when stock_qty_delta < 0 then scrape_date end) as last_decrease_date,
    datediff(day, min(case when stock_qty_delta < 0 then scrape_date end), max(case when stock_qty_delta < 0 then scrape_date end)) + 1 as decrease_window_days,
    round(sum(case when stock_qty_delta < 0 then abs(stock_qty_delta) else 0 end) / nullif(datediff(day, min(scrape_date), max(scrape_date)) + 1, 0) * 7, 1) as avg_units_per_week_observed_window,
    round(sum(case when stock_qty_delta < 0 then abs(stock_qty_delta) else 0 end) / nullif(datediff(day, min(case when stock_qty_delta < 0 then scrape_date end), max(case when stock_qty_delta < 0 then scrape_date end)) + 1, 0) * 7, 1) as avg_units_per_week_decrease_window,
    datediff(day, max(case when stock_qty_delta < 0 then scrape_date end), max(scrape_date)) as days_since_last_decrease
  from public.v_competitors_inventory_daily
  group by url
)
select
  latest.name,
  latest.brand,
  latest.sku,
  latest.category,
  latest.product_type,
  latest.description,
  latest.wattage,
  latest.lumens,
  latest.cct,
  latest.cri,
  latest.ip_rating,
  latest.voltage,
  latest.life_hours,
  latest.dimmable,
  latest.price,
  latest.price_high,
  latest.currency,
  latest.availability,
  latest.stock_qty,
  latest.stock_status,
  latest.image,
  latest.url,
  latest.scraped_at,
  split_part(replace(replace(latest.url,'https://',''),'http://',''), '/', 1) as domain,
  coalesce(movement.observed_stock_decrease, 0) as observed_stock_decrease,
  coalesce(movement.observed_restock, 0) as observed_restock,
  coalesce(movement.decrease_events, 0) as decrease_events,
  coalesce(movement.restock_events, 0) as restock_events,
  coalesce(movement.observation_count, 0) as observation_count,
  movement.first_observed_date,
  movement.last_observed_date,
  coalesce(movement.observation_window_days, 0) as observation_window_days,
  movement.first_decrease_date,
  movement.last_decrease_date,
  coalesce(movement.decrease_window_days, 0) as decrease_window_days,
  coalesce(movement.avg_units_per_week_observed_window, 0) as avg_units_per_week_observed_window,
  coalesce(movement.avg_units_per_week_decrease_window, 0) as avg_units_per_week_decrease_window,
  movement.days_since_last_decrease,
  movement.latest_stock_qty,
  movement.latest_inventory_scrape_date
from public.v_competitors_scrapping_latest latest
left join movement on movement.url = latest.url
where ({where_clause})
{exclude_clause}
order by coalesce(movement.avg_units_per_week_observed_window, 0) desc,
         coalesce(movement.observed_stock_decrease, 0) desc,
         coalesce(movement.decrease_events, 0) desc,
         latest.scraped_at desc
limit {int(limit)};
""".strip()


def latest_ecommerce_snapshot(exports_dir: Path, category_slug: str) -> Path | None:
    snapshots = sorted(exports_dir.glob(f"{category_slug}_ecommerce_competitor_evidence_*.json"), key=lambda path: path.stat().st_mtime)
    return snapshots[-1] if snapshots else None


def _snapshot_age_hours(payload: dict[str, Any]) -> float | None:
    generated_at = payload.get("generated_at")
    if not generated_at:
        return None
    try:
        timestamp = datetime.fromisoformat(str(generated_at).replace("Z", "+00:00"))
    except ValueError:
        return None
    if timestamp.tzinfo is None:
        timestamp = timestamp.replace(tzinfo=timezone.utc)
    return (datetime.now(timezone.utc) - timestamp).total_seconds() / 3600


def _is_redshift_snapshot(payload: dict[str, Any]) -> bool:
    return str(payload.get("source_system") or "").lower().startswith(("redshift_mcp", "redshift_odbc"))


def _should_refresh_ecommerce_snapshot(payload: dict[str, Any]) -> bool:
    if os.environ.get("PRODUCT_DEMAND_FORCE_ECOMMERCE_REFRESH", "").strip() == "1":
        return True
    if not _is_redshift_snapshot(payload):
        return True
    max_age_hours = float(os.environ.get("PRODUCT_DEMAND_ECOMMERCE_SNAPSHOT_MAX_AGE_HOURS") or 24)
    age_hours = _snapshot_age_hours(payload)
    return age_hours is None or age_hours > max_age_hours


def write_ecommerce_snapshot(
    exports_dir: Path,
    category_slug: str,
    source_system: str,
    sql: str,
    rows: list[dict[str, Any]],
    source_connector: str | None = None,
) -> Path:
    exports_dir.mkdir(parents=True, exist_ok=True)
    generated_at = utc_now()
    stamp = generated_at.replace("-", "").replace(":", "").split("+", 1)[0].replace("T", "_")
    target = exports_dir / f"{category_slug}_ecommerce_competitor_evidence_{stamp}.json"
    payload = {
        "source_system": source_system,
        "category_slug": category_slug,
        "generated_at": generated_at,
        "row_count": len(rows),
        "sql": sql,
        "rows": rows,
    }
    if source_connector:
        payload["source_connector"] = source_connector
    target.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return target


def refresh_ecommerce_snapshot_via_redshift(exports_dir: Path, category_slug: str, timeout_seconds: int = 240) -> Path:
    sql = build_ecommerce_sql(category_slug)
    rows, source_connector = execute_redshift_sql(sql, timeout_seconds=timeout_seconds, client_name="sunco-ecommerce-competitor-redshift")
    source_system = (
        "redshift_mcp_ecommerce_competitor_snapshot"
        if is_redshift_mcp_source(source_connector)
        else "redshift_odbc_dsn_ecommerce_competitor_snapshot"
    )
    return write_ecommerce_snapshot(exports_dir, category_slug, source_system, sql, rows, source_connector=source_connector)


def load_or_refresh_ecommerce_snapshot(exports_dir: Path, category_slug: str) -> tuple[dict[str, Any], Path]:
    path = latest_ecommerce_snapshot(exports_dir, category_slug)
    payload = json.loads(path.read_text(encoding="utf-8")) if path else {}
    if path is None or _should_refresh_ecommerce_snapshot(payload):
        path = refresh_ecommerce_snapshot_via_redshift(exports_dir, category_slug)
        payload = json.loads(path.read_text(encoding="utf-8"))
    return payload, path


def _text(value: Any) -> str:
    text = unescape(str(value or ""))
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _display_price(value: Any) -> str:
    text = _text(value)
    if not text or text.lower() in {"n/a", "na", "none", "null", "nan"}:
        return "n/a"
    return f"${text}"


def _domain(row: dict[str, Any]) -> str:
    return _text(row.get("domain")) or re.sub(r"^https?://", "", _text(row.get("url"))).split("/", 1)[0]


def _size_signal(text: str) -> str | None:
    normalized = text.lower().replace("×", "x")
    match = re.search(r"\b([124])\s*x\s*([124])\b", normalized)
    if match:
        return f"{match.group(1)}x{match.group(2)}"
    match = re.search(r"\b(\d{1,2})\s*(?:inch|in\.)\b", normalized)
    if match:
        return f"{match.group(1)} inch"
    return None


def _pack_signal(text: str) -> str | None:
    match = re.search(r"\b(\d{1,3})\s*(?:pack|pk)\b", text.lower())
    return f"{match.group(1)} pack" if match else None


def _cct_signal(row: dict[str, Any]) -> str | None:
    text = " ".join(_text(row.get(key)) for key in ["cct", "name", "description"])
    values = re.findall(r"\b(?:27|30|35|40|45|50|57|60|65)00\s*k\b|\b(?:2700|3000|3500|4000|4500|5000|5700|6000|6500)k\b", text, flags=re.IGNORECASE)
    cleaned = []
    for value in values:
        item = value.lower().replace(" ", "").upper()
        if item not in cleaned:
            cleaned.append(item)
    if len(cleaned) >= 5:
        return "5CCT"
    if len(cleaned) >= 3:
        return "3CCT"
    if cleaned:
        return "/".join(cleaned[:3])
    if "selectable" in text.lower():
        return "selectable CCT"
    return None


def _feature_signals(text: str) -> tuple[str, ...]:
    raw = text.lower()
    features = []
    for label, terms in [
        ("Emergency Backup", ("emergency", "battery backup", "90 minute")),
        ("Motion Sensor", ("motion sensor", "occupancy sensor", "pir sensor", "microwave sensor")),
        ("Daylight Harvesting", ("daylight harvesting", "daylight sensor", "ambient light sensor", "photocell")),
        ("Control Ready", ("sensor ready", "control ready", "controls ready", "c-max", "control-ready")),
        ("Smart Controls", ("bluetooth mesh", "wireless control", "smart control")),
        ("0-10V Dimming", ("0-10v",)),
        ("TRIAC Dimming", ("triac",)),
        ("dimmable", ("dimmable", "dimming")),
        ("wet/damp rated", ("wet rated", "damp rated", "ip65", "ip66")),
    ]:
        if any(term in raw for term in terms):
            features.append(label)
    return tuple(dict.fromkeys(features))


def _product_type_signal(row: dict[str, Any]) -> str:
    text = " ".join(_text(row.get(key)) for key in ["name", "category"]).lower()
    if "vapor tight" in text or "vaportight" in text or "vapor-proof" in text or "vapor proof" in text:
        if "strip" in text or "linear" in text:
            return "Linear Vapor Tight"
        if "led ready" in text or "lamp ready" in text:
            return "LED Ready Vapor Tight"
        return "Vapor Tight"
    if "grow" in text or "horticulture" in text:
        if "vapor tight" in text:
            return "Vapor Tight Grow Fixture"
        if "top light" in text:
            return "Horticulture Top Light"
        if "strip" in text or "shop light" in text:
            return "Linear/Shop Grow Light"
        if "panel" in text:
            return "Panel Grow Light"
        if "bulb" in text or "lamp" in text:
            return "Grow Bulb/Lamp"
        if "fixture" in text or "luminaire" in text or "light" in text:
            return "Grow Light Fixture"
    product_type = _text(row.get("product_type"))
    if product_type:
        return product_type
    if "wraparound" in text or "wrap around" in text:
        return "Wraparound"
    if "troffer" in text:
        return "Troffer"
    if "panel" in text:
        return "Panel"
    return "Ecommerce product"


def _spec_signature(row: dict[str, Any], product_type_override: str = "") -> tuple[str, ...]:
    product_spec_text = " ".join(
        _text(row.get(key)) for key in ["name", "wattage", "lumens", "cct", "voltage"]
    )
    feature_text = " ".join(
        _text(row.get(key)) for key in ["name", "description", "category", "product_type", "wattage", "lumens", "cct", "voltage"]
    )
    perf = normalize_luminaire_performance(product_spec_text)
    if not perf.has_signal:
        product_spec_text = " ".join(
            _text(row.get(key)) for key in ["name", "description", "wattage", "lumens", "cct", "voltage"]
        )
        perf = normalize_luminaire_performance(product_spec_text)
    parts = [
        product_type_override or _product_type_signal(row),
        _size_signal(product_spec_text),
        perf.user_brightness_target,
        perf.installer_load_target,
        _cct_signal(row),
    ]
    parts.extend(_feature_signals(feature_text)[:2])
    return tuple(part for part in parts if part)


def _row_completeness(row: dict[str, Any]) -> int:
    fields = ["name", "brand", "product_type", "wattage", "lumens", "cct", "voltage", "price", "availability", "url"]
    return sum(1 for field in fields if _text(row.get(field)))


def _as_float(value: Any) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _as_int(value: Any) -> int:
    return int(_as_float(value))


def _movement_metrics(rows: list[dict[str, Any]]) -> dict[str, Any]:
    decrease = sum(_as_float(row.get("observed_stock_decrease")) for row in rows)
    events = sum(_as_int(row.get("decrease_events")) for row in rows)
    observations = sum(_as_int(row.get("observation_count")) for row in rows)
    windows = [_as_int(row.get("observation_window_days")) for row in rows if _as_int(row.get("observation_window_days")) > 0]
    decrease_windows = [_as_int(row.get("decrease_window_days")) for row in rows if _as_int(row.get("decrease_window_days")) > 0]
    observed_velocity = sum(_as_float(row.get("avg_units_per_week_observed_window")) for row in rows)
    decrease_velocity = sum(_as_float(row.get("avg_units_per_week_decrease_window")) for row in rows)
    recencies = [_as_int(row.get("days_since_last_decrease")) for row in rows if row.get("days_since_last_decrease") is not None]
    dates = []
    for key in ["first_observed_date", "last_observed_date", "first_decrease_date", "last_decrease_date"]:
        values = sorted(_text(row.get(key)) for row in rows if _text(row.get(key)))
        dates.append((key, values[0] if values and key.startswith("first") else values[-1] if values else ""))
    return {
        "decrease": decrease,
        "events": events,
        "observations": observations,
        "observation_window_days": max(windows, default=0),
        "decrease_window_days": max(decrease_windows, default=0),
        "avg_units_per_week_observed_window": observed_velocity,
        "avg_units_per_week_decrease_window": decrease_velocity,
        "days_since_last_decrease": min(recencies) if recencies else None,
        **dict(dates),
    }


def _feature_search_terms(rows: list[dict[str, Any]]) -> tuple[str, ...]:
    text = " ".join(
        " ".join(_text(row.get(key)) for key in ["name", "description", "product_type", "category"])
        for row in rows
    )
    return tuple(
        feature
        for feature in _feature_signals(text)
        if feature not in {"dimmable", "wet/damp rated"}
    )


def _vendor_bias_risk(domains: list[str], brands: list[str]) -> str:
    if len(domains) >= 2 and len(brands) >= 2:
        return "Low"
    if len(domains) >= 2 or len(brands) >= 2:
        return "Medium"
    return "High"


def _spec_demand_confidence(
    feature_terms: tuple[str, ...],
    group_count: int,
    domain_count: int,
    brand_count: int,
    decrease: float,
    events: int,
) -> str:
    if decrease <= 0:
        return "Low"
    if feature_terms and domain_count >= 2 and brand_count >= 2 and events >= 3:
        return "High"
    if feature_terms and events >= 2 and (group_count >= 2 or domain_count >= 2 or brand_count >= 2):
        return "Medium"
    if domain_count >= 2 and events >= 3:
        return "Medium"
    return "Directional"


def _spec_demand_note(
    feature_terms: tuple[str, ...],
    domains: list[str],
    brands: list[str],
    group_count: int,
    decrease: float,
    events: int,
) -> str:
    confidence = _spec_demand_confidence(feature_terms, group_count, len(domains), len(brands), decrease, events)
    bias = _vendor_bias_risk(domains, brands)
    search_terms = ", ".join(feature_terms) if feature_terms else "no distinct feature term isolated beyond the base spec pattern"
    return (
        f"Spec demand confidence: {confidence}. "
        f"Searched/customer terms: {search_terms}. "
        f"Vendor bias risk: {bias} ({len(domains)} competitor domain(s), {len(brands)} brand(s)). "
        "This is a demand hypothesis: inventory movement shows the product/spec cluster is moving, "
        "but it does not prove the feature alone caused the sales lift."
    )


def _movement_score(metrics: dict[str, Any]) -> float:
    decrease = _as_float(metrics.get("decrease"))
    events = _as_float(metrics.get("events"))
    velocity = max(
        _as_float(metrics.get("avg_units_per_week_observed_window")),
        _as_float(metrics.get("avg_units_per_week_decrease_window")),
    )
    observations = _as_float(metrics.get("observations"))
    days_since = metrics.get("days_since_last_decrease")
    recency_score = 0.0
    if days_since is not None:
        days = _as_float(days_since)
        if days <= 7:
            recency_score = 100.0
        elif days <= 21:
            recency_score = 70.0
        elif days <= 45:
            recency_score = 40.0
        else:
            recency_score = 15.0
    return min(
        100.0,
        (min(100.0, velocity / 40.0 * 100.0) * 0.35)
        + (min(100.0, events / 6.0 * 100.0) * 0.25)
        + (recency_score * 0.20)
        + (min(100.0, decrease / 120.0 * 100.0) * 0.10)
        + (min(100.0, observations / 10.0 * 100.0) * 0.10),
    )


def _candidate_score(rows: list[dict[str, Any]]) -> float:
    metrics = _movement_metrics(rows)
    movement_score = _movement_score(metrics)
    domains = len({_domain(row) for row in rows if _domain(row)})
    completeness = max((_row_completeness(row) for row in rows), default=0)
    feature_text = " ".join(
        " ".join(_feature_signals(" ".join(_text(row.get(key)) for key in ["name", "description", "product_type", "category"])))
        for row in rows
    ).lower()
    technology_bonus = 12.0 if any(
        term in feature_text
        for term in ["emergency backup", "motion sensor", "daylight harvesting", "control ready", "smart controls"]
    ) else 0.0
    return round(min(100.0, (movement_score * 0.55) + (domains * 8.0) + (completeness * 3.2) + technology_bonus), 1)


def _best_example(rows: list[dict[str, Any]]) -> dict[str, Any]:
    return sorted(
        rows,
        key=lambda row: (
            _as_float(row.get("avg_units_per_week_observed_window")),
            _as_float(row.get("observed_stock_decrease")),
            _as_float(row.get("decrease_events")),
            _row_completeness(row),
        ),
        reverse=True,
    )[0]


def _priority_confidence(score: float, rows: list[dict[str, Any]]) -> tuple[str, str]:
    domains = len({_domain(row) for row in rows if _domain(row)})
    metrics = _movement_metrics(rows)
    decrease = _as_float(metrics.get("decrease"))
    movement_score = _movement_score(metrics)
    if decrease <= 0:
        return "Medium", "Needs demand evidence"
    if score >= 70 or movement_score >= 70 or (domains >= 2 and decrease > 0):
        return "High", "High"
    if score >= 45 or decrease > 0:
        return "High", "Medium"
    return "Medium", "Needs more evidence"


def _candidate_identity(row: dict[str, Any]) -> tuple[str, str]:
    return (_text(row.get("source_url") or row.get("review_url")), _text(row.get("recommendation")).lower())


def _recommendation_body(value: Any) -> str:
    text = _text(value)
    for marker in [
        "Shopify ecommerce candidate:",
        "Strategic outlier watchlist:",
    ]:
        if text.lower().startswith(marker.lower()) and ":" in text:
            return text.split(":", 1)[1].strip()
    return text


def _strategic_high_output_reason(category_name: str, row: dict[str, Any]) -> str:
    row_text = " ".join(
        _text(row.get(key))
        for key in [
            "subcategory",
            "recommendation",
            "classification",
            "example",
            "why_gap",
            "sunco_check",
        ]
    )
    if "panel" not in f"{category_name} {row_text}".lower():
        return ""
    if not _text(row.get("source_url") or row.get("review_url")):
        return ""
    perf = normalize_luminaire_performance(row_text)
    max_lumens = max(perf.lumen_values, default=0.0)
    max_watts = max(perf.watt_values, default=0.0)
    if max_lumens < HIGH_OUTPUT_PANEL_LUMEN_THRESHOLD or max_watts < HIGH_OUTPUT_PANEL_WATT_THRESHOLD:
        return ""
    metrics = row.get("_movement_metrics") or {}
    decrease = _as_float(metrics.get("decrease"))
    events = _as_int(metrics.get("events"))
    days_since_value = metrics.get("days_since_last_decrease")
    days_since = _as_float(days_since_value) if days_since_value is not None else None
    if decrease <= 0 or events < 3:
        return ""
    if days_since is not None and days_since > 30:
        return ""
    return (
        f"High-output panel outlier with {max_lumens:g} lumen / {max_watts:g} watt class evidence, "
        f"{decrease:g} observed stock decrease units, {events} decrease event(s)"
        f"{f', and last decrease {int(days_since)} day(s) from latest scrape' if days_since is not None else ''}."
    )


def _mark_strategic_outlier(row: dict[str, Any], reason: str) -> dict[str, Any]:
    marked = dict(row)
    body = _recommendation_body(marked.get("recommendation"))
    marked["_strategic_outlier"] = True
    marked["_strategic_outlier_reason"] = reason
    marked["recommendation"] = f"Strategic outlier watchlist: {body}" if body else "Strategic outlier watchlist"
    marked["classification"] = "Strategic outlier / High-output watchlist"
    marked["priority"] = "Medium"
    marked["confidence"] = "Directional"
    marked["why_gap"] = (
        f"{_text(marked.get('why_gap'))}\n\n"
        f"Strategic outlier note: {reason} It fell below the normal Step 1B rank cutoff but remains visible for PM review."
    ).strip()
    marked["pm_action"] = (
        "Review as a strategic high-output outlier. Confirm application fit, installation constraints, heat/load profile, "
        "certifications, vendor feasibility, and whether demand is broader than one competitor PDP before PRD/RFQ."
    )
    return marked


def _select_candidates_with_outliers(category_name: str, candidates: list[dict[str, Any]], limit: int) -> list[dict[str, Any]]:
    selected = candidates[:limit]
    selected_keys = {_candidate_identity(row) for row in selected}
    outliers: list[tuple[float, dict[str, Any]]] = []
    for row in candidates[limit:]:
        if _candidate_identity(row) in selected_keys:
            continue
        reason = _strategic_high_output_reason(category_name, row)
        if not reason:
            continue
        row_text = " ".join(_text(row.get(key)) for key in ["recommendation", "example", "why_gap"])
        perf = normalize_luminaire_performance(row_text)
        outlier_rank = max(perf.lumen_values, default=0.0) + _as_float(row.get("_ecommerce_score"))
        outliers.append((outlier_rank, _mark_strategic_outlier(row, reason)))
    outliers.sort(key=lambda item: item[0], reverse=True)
    return selected + [row for _, row in outliers[:STRATEGIC_OUTLIER_LIMIT]]


def ecommerce_rows_to_step1_rows(category_name: str, rows: list[dict[str, Any]], limit: int = 16) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, ...], list[dict[str, Any]]] = defaultdict(list)
    grow_light_category = _is_grow_light_category(category_name)
    residential_grow_category = grow_light_category and "residential" in category_name.lower()
    shop_light_category = category_name.lower().replace("_", " ") == "shop light"
    emergency_category = category_name.lower().replace("_", " ") == "emergency"
    striplights_category = category_name.lower().replace("_", " ") == "striplights"
    for row in rows:
        if not _text(row.get("url")) or not _text(row.get("name")):
            continue
        if grow_light_category and not is_grow_light_candidate(row):
            continue
        if residential_grow_category and not is_residential_grow_light_candidate(row):
            continue
        if shop_light_category and not is_shop_light_candidate(row):
            continue
        if emergency_category and not is_emergency_candidate(row):
            continue
        if striplights_category and not is_striplights_candidate(row):
            continue
        signature = _spec_signature(row, product_type_override="Strip/Linear" if striplights_category else "")
        perf = normalize_luminaire_performance(
            " ".join(_text(row.get(key)) for key in ["name", "description", "wattage", "lumens", "cct"])
        )
        if len(signature) < 3 or not (perf.lumen_values or perf.watt_values):
            continue
        grouped[signature].append(row)

    candidates: list[dict[str, Any]] = []
    for signature, group_rows in grouped.items():
        score = _candidate_score(group_rows)
        example = _best_example(group_rows)
        priority, confidence = _priority_confidence(score, group_rows)
        domains = sorted({_domain(row) for row in group_rows if _domain(row)})
        brands = sorted({_text(row.get("brand")) for row in group_rows if _text(row.get("brand"))})
        metrics = _movement_metrics(group_rows)
        decrease = _as_float(metrics.get("decrease"))
        events = _as_int(metrics.get("events"))
        feature_terms = _feature_search_terms(group_rows)
        spec_demand_note = _spec_demand_note(feature_terms, domains, brands, len(group_rows), decrease, events)
        observed_window = _as_int(metrics.get("observation_window_days"))
        decrease_window = _as_int(metrics.get("decrease_window_days"))
        weekly_velocity = max(
            _as_float(metrics.get("avg_units_per_week_observed_window")),
            _as_float(metrics.get("avg_units_per_week_decrease_window")),
        )
        days_since = metrics.get("days_since_last_decrease")
        velocity_note = ""
        if observed_window:
            velocity_note = (
                f" Movement was observed over {observed_window} day(s)"
                f"{f', with decrease activity over {decrease_window} day(s)' if decrease_window else ''}"
                f", averaging about {weekly_velocity:.1f} units/week."
            )
            if days_since is not None:
                velocity_note += f" Last decrease was {int(_as_float(days_since))} day(s) from the latest scrape."
        spec_text = "; ".join(signature)
        example_text = (
            f"{_text(example.get('brand'))} {_text(example.get('sku'))}: {_text(example.get('name'))} | "
            f"{_text(example.get('lumens'))} | {_text(example.get('wattage'))} | {_text(example.get('cct'))} | "
            f"{_text(example.get('availability'))} | {_display_price(example.get('price'))}"
        ).strip()
        candidates.append(
            {
                "subcategory": category_name,
                "recommendation": f"Shopify ecommerce candidate: {spec_text}",
                "classification": "Ecommerce competitor demand candidate",
                "priority": priority,
                "confidence": confidence,
                "example": example_text,
                "source_url": example.get("url"),
                "review_url": example.get("url"),
                "sunco_check": "Sunco active-catalog exact-spec coverage will be applied by the Product Demand overlay before this row is published.",
                "why_gap": (
                    f"Redshift ecommerce competitor evidence found {len(group_rows)} PDP listing(s) across {len(domains)} domain(s) "
                    f"and {len(brands)} brand(s) for this spec pattern. Observed stock decrease totals {decrease:g} units across "
                    f"{events} decrease event(s).{velocity_note} {spec_demand_note} "
                    "This should lead the Shopify/front-end launch review before Amazon-only evidence."
                ),
                "pm_action": (
                    "Scope as a new variant around this normalized spec pattern. Use the coverage result and demand evidence in this row as the launch rationale; do not copy the competitor SKU one-for-one."
                ),
                "source_systems": ["redshift:v_competitors_scrapping_latest", "redshift:v_competitors_inventory_daily"],
                "image_url": example.get("image"),
                "_ecommerce_score": score,
                "_ecommerce_group_size": len(group_rows),
                "_ecommerce_domains": domains,
                "_ecommerce_brands": brands,
                "_feature_search_terms": feature_terms,
                "_spec_demand_confidence": _spec_demand_confidence(feature_terms, len(group_rows), len(domains), len(brands), decrease, events),
                "_vendor_bias_risk": _vendor_bias_risk(domains, brands),
                "_movement_metrics": metrics,
            }
        )
    candidates.sort(key=lambda row: (float(row.get("_ecommerce_score") or 0), int(row.get("_ecommerce_group_size") or 0)), reverse=True)
    return _select_candidates_with_outliers(category_name, candidates, limit)
