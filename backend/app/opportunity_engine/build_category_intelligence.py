from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sqlite3
from collections import Counter, defaultdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from .paths import ProjectPaths
from .utils import slugify


ALIASES = {
    "ufo": ["high bay", "ufo high bay", "warehouse light"],
    "linears": ["linear", "linear fixture", "shop light", "strip light"],
    "slims": ["slim", "slim light", "canless slim"],
    "retros": ["retrofit", "retrofit kit", "downlight retrofit"],
    "cans": ["recessed can", "can light", "downlight can"],
    "rough_ins": ["rough in", "rough-in", "junction box"],
    "under_cabinet": ["Under cabinet / tape", "under-cabinet", "under counter", "Under Cabinet Light Fixture"],
    "tape_rope_light": ["Under cabinet / tape", "tape light", "rope light", "LED Neon Rope"],
    "wall_sconces": ["Wall sconces", "sconce", "wall sconce"],
    "ceiling_fixtures": ["Ceiling fixtures", "Residential decor fixtures", "fixture", "flush mount", "ceiling light"],
    "panels": ["Ceiling panels", "Ceiling panels / fixtures", "Ceiling Panel Lights", "panel lights", "panel light", "flat panels"],
    "emergency": ["Emergency Lights", "Emergency Signs", "exit signs", "exit sign"],
    "striplights": ["striplight", "strip light", "strip lights"],
    "vaportights": ["vapor tight", "vaportight", "vapor tights"],
    "wraparounds": ["Wraparound", "Wraparound / utility ceiling", "Wraparound / residential utility"],
    "wall_packs": ["wall pack", "wall packs"],
    "flood_lights": ["flood light", "flood lights"],
    "commercial_landscape": ["landscape light", "commercial landscape"],
    "outdoor_security": ["security light", "outdoor security"],
    "area_lights": ["area light", "area lights", "shoebox"],
    "canopy": ["canopy light", "canopy"],
    "string_lights": ["string light", "string lights"],
    "outdoor_ceiling": ["outdoor ceiling"],
    "wire": ["wire", "cable"],
    "low_voltage_transformers": ["transformer", "low voltage transformer"],
    "wire_connectors": ["wire connector", "cable connector", "connector"],
    "sensors": ["sensor", "occupancy sensor", "pir sensor"],
    "dimmers": ["dimmer", "light switch", "switches", "light switches"],
    "bathroom_fans": ["bathroom fan", "exhaust fan"],
    "bulbs_plus_tubes": ["bulb", "bulbs", "led bulbs", "tube", "tubes", "t8", "lamp", "lamps"],
    "pendants": ["pendant", "pendant light"],
    "vanity": ["vanity", "vanity light"],
    "commercial_grow_lights": ["commercial grow light", "grow lights", "grow light"],
    "residential_grow_lights": ["residential grow light"],
    "commercial_security": ["commercial security"],
    "residential_security": ["residential security"],
    "residential_landscape": ["path lights", "path light", "spotlight", "well light", "landscape"],
}

ATTRIBUTE_HINTS = {
    "wattage": r"\b\d+(?:\.\d+)?\s*W\b",
    "lumens": r"\b\d{3,6}\s*(?:lm|lumens)\b",
    "cct": r"\b(?:27|30|35|40|50|65)00K\b|\b\d{4}\s*K\b",
    "cri": r"\bCRI\s*\d{2,3}\+?|\b\d{2,3}\+\s*CRI\b",
    "finish": r"\b(?:black|white|bronze|brass|gold|nickel|rattan|wood|matte black)\b",
    "mounting": r"\b(?:hardwired|plug[- ]?in|magnetic|adhesive|surface mount|recessed|chain|rod)\b",
    "dimming": r"\b(?:dimmable|0-10V|triac|remote|motion|sensor)\b",
    "certifications": r"\b(?:ETL|UL|DLC|Energy Star|FCC|RoHS|cETLus|cULus)\b",
    "pack_qty": r"\b\d+\s*(?:pack|pk|count)\b",
    "socket_base": r"\bE(?:12|17|26|39)\b",
    "length": r"\b\d+(?:\.\d+)?\s*(?:ft|feet|in|inch)\b",
}

PROFILE_SIGNAL_FIELDS = {
    "style_keywords",
    "finish_keywords",
    "material_keywords",
    "mounting_keywords",
}

INCLUDE_LEGACY_GAP_EVIDENCE_ENV = "OPPORTUNITY_ENGINE_INCLUDE_LEGACY_GAP_EVIDENCE"
GENERATED_FEATURE_PROFILE_FILENAME = "generated_category_feature_signal_profiles.json"
GENERATED_ATTRIBUTE_DECISION_PROFILE_FILENAME = "generated_category_attribute_decision_profiles.json"

NOISE_ATTRIBUTE_NAMES = {
    "brand",
    "components",
    "dimensions",
    "efficacy",
    "lifetime",
    "model_number",
    "product_type",
    "warranty",
}

PLACEHOLDER_VALUES = {
    "",
    "0",
    "na",
    "n/a",
    "none",
    "not applicable",
    "null",
}

ATTRIBUTE_GROUPS = {
    "base_type": "Physical / Compatibility",
    "beam_angle": "Performance Configuration",
    "bulb_base": "Physical / Compatibility",
    "certifications": "Ratings / Certifications",
    "color_temp": "CCT Strategy",
    "cct": "CCT Strategy",
    "cri": "Performance Configuration",
    "daylight_sensor_auto_dimming": "Controls",
    "dimmable": "Controls",
    "dimming_details": "Controls",
    "finish": "Physical / Compatibility",
    "ic_rated": "Ratings / Certifications",
    "ip_rating": "Ratings / Certifications",
    "linkable": "Installation / Linking",
    "lumens": "Performance Configuration",
    "material": "Physical / Compatibility",
    "moisture_rating": "Ratings / Certifications",
    "motion_sensor": "Controls",
    "mount_type": "Installation / Linking",
    "power_type": "Installation / Linking",
    "sensor_details": "Controls",
    "usage": "Ratings / Certifications",
    "voltage": "Performance Configuration",
    "wattage": "Performance Configuration",
    "wiring_type": "Installation / Linking",
}

GROUP_TYPES = {
    "CCT Strategy": "Feature",
    "Controls": "Feature",
    "Installation / Linking": "Feature",
    "Performance Configuration": "Feature",
    "Physical / Compatibility": "Feature",
    "Ratings / Certifications": "Mixed",
}

BROAD_GENERATED_MATCH_TERMS = {
    "beam angle",
    "certifications",
    "color temp",
    "cri",
    "dimming details",
    "ip rating",
    "linkable",
    "lumens",
    "moisture rating",
    "mount type",
    "power type",
    "usage",
    "voltage",
    "wattage",
    "yes",
}


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def json_text(value: Any) -> str:
    return json.dumps(value, ensure_ascii=True, sort_keys=True)


def parse_generated_date(payload: dict[str, Any], path: Path) -> str | None:
    for key in ("generated", "generated_at", "run_date", "created_at"):
        value = payload.get(key)
        if value:
            return str(value)[:10]
    match = re.search(r"(20\d{2}-\d{2}-\d{2})", path.name)
    return match.group(1) if match else None


def age_days(generated: str | None) -> int | None:
    if not generated:
        return None
    try:
        return (date.today() - date.fromisoformat(generated[:10])).days
    except ValueError:
        return None


def load_category_rows(paths: ProjectPaths) -> list[dict[str, Any]]:
    with (paths.templates / "category_reference.csv").open("r", encoding="utf-8-sig", newline="") as handle:
        rows = []
        for row in csv.DictReader(handle):
            if (row.get("active") or "").strip().lower() not in {"yes", "true", "1"}:
                continue
            run_name = (row.get("run_name") or row.get("category") or "").strip()
            slug = slugify(run_name)
            aliases = sorted(set([run_name, row.get("category") or "", *ALIASES.get(slug, [])]))
            rows.append(
                {
                    "owner": (row.get("owner") or "").strip(),
                    "category": (row.get("category") or "").strip(),
                    "run_name": run_name,
                    "slug": slug,
                    "active": 1,
                    "aliases": [alias for alias in aliases if alias],
                    "notes": (row.get("notes") or "").strip(),
                }
            )
    return rows


def category_lookup(category_rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    lookup: dict[str, dict[str, Any]] = {}
    for row in category_rows:
        keys = {row["slug"], slugify(row["run_name"]), slugify(row["category"])}
        keys.update(slugify(alias) for alias in row["aliases"])
        for key in keys:
            if key:
                lookup[key] = row
    return lookup


def resolve_category(lookup: dict[str, dict[str, Any]], value: str | None) -> dict[str, Any] | None:
    if not value:
        return None
    raw = str(value).strip()
    keys = [slugify(raw)]
    keys.extend(slugify(part.strip()) for part in re.split(r"/|\|", raw) if part.strip())
    for key in keys:
        if key in lookup:
            return lookup[key]
    for key, row in lookup.items():
        if key and key in slugify(raw):
            return row
    return None


def resolve_category_from_values(lookup: dict[str, dict[str, Any]], *values: Any) -> dict[str, Any] | None:
    for value in values:
        category = resolve_category(lookup, str(value) if value is not None else None)
        if category:
            return category
    return None


def extract_attribute_values(text: str) -> dict[str, list[str]]:
    output: dict[str, list[str]] = {}
    for name, pattern in ATTRIBUTE_HINTS.items():
        values = []
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            value = re.sub(r"\s+", " ", match.group(0)).strip()
            if value and value.lower() not in {item.lower() for item in values}:
                values.append(value)
        if values:
            output[name] = values
    return output


def reset_database(connection: sqlite3.Connection, schema_path: Path) -> None:
    connection.execute("PRAGMA foreign_keys = OFF")
    connection.execute("DROP VIEW IF EXISTS category_intelligence_summary")
    for table in [
        "refresh_audit",
        "gap_evidence",
        "category_feature_signal_profile",
        "stackline_top_products",
        "stackline_segments",
        "sku_decoder_codes",
        "category_attribute_distribution",
        "shopify_spec_attributes",
        "shopify_category_products",
        "categories",
    ]:
        connection.execute(f"DROP TABLE IF EXISTS {table}")
    connection.execute("PRAGMA foreign_keys = ON")
    connection.executescript(schema_path.read_text(encoding="utf-8"))
    connection.commit()


def insert_categories(connection: sqlite3.Connection, category_rows: list[dict[str, Any]]) -> dict[str, int]:
    now = utc_now()
    ids = {}
    for row in category_rows:
        cursor = connection.execute(
            """
            INSERT INTO categories(owner, category, run_name, slug, active, aliases_json, notes, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                row["owner"],
                row["category"],
                row["run_name"],
                row["slug"],
                row["active"],
                json_text(row["aliases"]),
                row["notes"],
                now,
                now,
            ),
        )
        ids[row["slug"]] = int(cursor.lastrowid)
    return ids


def insert_gap_manifests(
    connection: sqlite3.Connection,
    paths: ProjectPaths,
    lookup: dict[str, dict[str, Any]],
    category_ids: dict[str, int],
) -> tuple[int, dict[str, int], list[str]]:
    if os.environ.get(INCLUDE_LEGACY_GAP_EVIDENCE_ENV, "").strip().lower() not in {"1", "true", "yes", "y"}:
        warning = (
            "Legacy schema-reference gap manifests were skipped. "
            f"Set {INCLUDE_LEGACY_GAP_EVIDENCE_ENV}=1 only for archival comparison rebuilds."
        )
        now = utc_now()
        connection.execute(
            """
            INSERT INTO refresh_audit(source, run_date, data_age_days, sql_used, row_counts_json, warnings_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "legacy_schema_reference_gap_manifests",
                date.today().isoformat(),
                None,
                "Skipped by default so stale local schema-reference manifests cannot become active recommendation evidence.",
                json_text({}),
                json_text([warning]),
                now,
            ),
        )
        return 0, {}, [warning]

    manifests = [
        ("Sunco.com/ecommerce", paths.source_data / "schema_references" / "source_manifest_indoor_residential_2026-05-13.json"),
        ("Amazon", paths.source_data / "schema_references" / "amazon_rerun_evidence_indoor_residential_2026-05-13.json"),
        ("Amazon", paths.source_data / "schema_references" / "amazon_recommendations_manifest_indoor_residential_2026-05-13.json"),
    ]
    now = utc_now()
    count = 0
    row_counts: dict[str, int] = {}
    warnings: list[str] = []
    attribute_counter: dict[int, Counter[tuple[str, str]]] = defaultdict(Counter)
    product_counts: Counter[int] = Counter()

    for channel, path in manifests:
        if not path.exists():
            warnings.append(f"Missing local manifest: {path.name}")
            continue
        payload = read_json(path)
        generated = parse_generated_date(payload, path)
        rows = list(payload.get("recommendations") or [])
        row_counts[path.name] = len(rows)
        for item in rows:
            category = resolve_category(lookup, item.get("subcategory"))
            if not category:
                warnings.append(f"Could not map manifest row category '{item.get('subcategory')}' in {path.name}")
                continue
            category_id = category_ids[category["slug"]]
            source_systems = item.get("source_systems") or [path.name]
            connection.execute(
                """
                INSERT INTO gap_evidence(
                    category_id, source_channel, recommendation, classification, priority, confidence,
                    competitor_example, review_url, sunco_coverage_check, gap_rationale, pm_action,
                    source_systems_json, source_reference, local_image, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    category_id,
                    channel,
                    item.get("recommendation") or item.get("Amazon-channel recommendation") or "Unnamed recommendation",
                    item.get("classification"),
                    item.get("priority"),
                    item.get("confidence"),
                    item.get("example"),
                    item.get("review_url") or item.get("source_url"),
                    item.get("sunco_check"),
                    item.get("why_gap") or item.get("evidence"),
                    item.get("pm_action") or item.get("action"),
                    json_text(source_systems),
                    str(path),
                    item.get("local_image"),
                    now,
                ),
            )
            count += 1
            product_counts[category_id] += 1
            text = " ".join(str(item.get(key) or "") for key in ["recommendation", "example", "evidence", "why_gap", "pm_action", "action"])
            for attribute, values in extract_attribute_values(text).items():
                for value in values:
                    attribute_counter[category_id][(attribute, value)] += 1
        connection.execute(
            """
            INSERT INTO refresh_audit(source, run_date, data_age_days, sql_used, row_counts_json, warnings_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                path.name,
                generated or date.today().isoformat(),
                age_days(generated),
                "Local JSON manifest import; no SQL used.",
                json_text({path.name: len(rows)}),
                json_text([]),
                now,
            ),
        )

    for category_id, counter in attribute_counter.items():
        total = product_counts[category_id] or 1
        for (attribute, value), occurrence_count in counter.items():
            coverage = round((occurrence_count / total) * 100, 2)
            signal = "table_stakes" if coverage >= 60 else "common" if coverage >= 25 else "rare"
            connection.execute(
                """
                INSERT INTO category_attribute_distribution(
                    category_id, attribute_name, attribute_value, occurrence_count, product_count,
                    coverage_pct, signal_class, source, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (category_id, attribute, value, occurrence_count, total, coverage, signal, "gap_evidence_text", now),
            )

    return count, row_counts, warnings


def insert_attribute_profiles(
    connection: sqlite3.Connection,
    paths: ProjectPaths,
    lookup: dict[str, dict[str, Any]],
    category_ids: dict[str, int],
) -> int:
    path = paths.app / "opportunity_engine" / "category_attribute_profiles.json"
    if not path.exists():
        return 0
    payload = read_json(path)
    now = utc_now()
    count = 0
    for key, profile in payload.items():
        category = lookup.get(slugify(key))
        if not category:
            continue
        category_id = category_ids[category["slug"]]
        defaults = profile.get("defaults") or {}
        for attribute, value in defaults.items():
            if value in (None, "", [], {}):
                continue
            connection.execute(
                """
                INSERT INTO category_attribute_distribution(
                    category_id, attribute_name, attribute_value, occurrence_count, product_count,
                    coverage_pct, signal_class, source, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (category_id, attribute, str(value), 1, 1, 100.0, "table_stakes", str(path), now),
            )
            count += 1
    return count


def flatten_profile_signals(profile: dict[str, Any]) -> list[str]:
    signals = list(profile.get("static_features") or [])
    for field in PROFILE_SIGNAL_FIELDS:
        for label in (profile.get(field) or {}).values():
            if label and label not in signals:
                signals.append(label)
    return signals


def insert_feature_profiles(
    connection: sqlite3.Connection,
    paths: ProjectPaths,
    lookup: dict[str, dict[str, Any]],
    category_ids: dict[str, int],
) -> int:
    path = paths.prd_tool / "config" / "category_feature_signal_profiles.json"
    if not path.exists():
        return 0
    payload = read_json(path)
    now = utc_now()
    count = 0
    for profile in payload.get("profiles") or []:
        category_matches = profile.get("category_matches") or profile.get("subcategory_matches") or []
        matched_categories = [resolve_category(lookup, value) for value in category_matches]
        matched_categories = [item for item in matched_categories if item]
        if not matched_categories and profile.get("id") == "generic_lighting":
            matched_categories = [None]
        for category in matched_categories:
            category_id = category_ids[category["slug"]] if category else None
            connection.execute(
                """
                INSERT INTO category_feature_signal_profile(
                    category_id, profile_key, label, applies_to_json, feature_signals_json,
                    source, source_reference, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    category_id,
                    profile.get("id") or "unnamed_profile",
                    profile.get("label") or profile.get("id") or "Unnamed profile",
                    json_text(category_matches),
                    json_text(flatten_profile_signals(profile)),
                    "category_feature_signal_profiles.json",
                    str(path),
                    now,
                ),
            )
            count += 1
    return count


def clean_generated_profile_value(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text


def generated_value_is_usable(attribute: str, value: Any) -> bool:
    attribute_key = slugify(attribute).replace("-", "_")
    if attribute_key in NOISE_ATTRIBUTE_NAMES:
        return False
    text = clean_generated_profile_value(value)
    if text.lower() in PLACEHOLDER_VALUES:
        return False
    if text.lower().startswith(("see specification", "refer to specification")):
        return False
    return True


def split_generated_profile_values(attribute: str, value: Any) -> list[str]:
    text = clean_generated_profile_value(value)
    if not generated_value_is_usable(attribute, text):
        return []
    attribute_key = slugify(attribute).replace("-", "_")
    if attribute_key in {"certifications", "color_temp", "cct"}:
        parts = [part.strip() for part in re.split(r",|;", text) if part.strip()]
        return parts or [text]
    return [text]


def normalize_generated_label(attribute: str, value: Any) -> str | None:
    attribute_key = slugify(attribute).replace("-", "_")
    text = clean_generated_profile_value(value)
    lowered = text.lower()
    if not generated_value_is_usable(attribute_key, text):
        return None

    if attribute_key == "certifications":
        cert_patterns = [
            ("cETLus", r"\bc\s*etlus\b|\bcetlus\b"),
            ("cULus", r"\bc\s*ulus\b|\bculus\b"),
            ("Energy Star", r"\benergy\s+star\b"),
            ("ETL", r"\betl\b"),
            ("UL", r"\bul\b"),
            ("DLC Premium", r"\bdlc\s+premium\b"),
            ("DLC", r"\bdlc\b"),
            ("FCC", r"\bfcc\b"),
            ("RoHS", r"\brohs\b"),
        ]
        for label, pattern in cert_patterns:
            if re.search(pattern, lowered, flags=re.IGNORECASE):
                return label
        return text
    if attribute_key in {"color_temp", "cct"}:
        ccts = re.findall(r"\b\d{4}\s*K\b", text, flags=re.IGNORECASE)
        if len(ccts) >= 5:
            return "5CCT selectable"
        if len(ccts) >= 3:
            return "3CCT selectable"
        if "select" in lowered:
            return "Selectable CCT"
        if ccts:
            return ccts[0].upper().replace(" ", "")
        return text
    if attribute_key in {"dimming_details", "dimming"}:
        if "0-10" in lowered or "0 10" in lowered:
            return "0-10V dimming"
        if "triac" in lowered:
            return "TRIAC dimming"
        if "dimm" in lowered:
            return "Dimmable"
        return text
    if attribute_key == "dimmable":
        if "non" in lowered or lowered in {"no", "false"}:
            return "Non-dimmable"
        if "dimm" in lowered or lowered in {"yes", "true"}:
            return "Dimmable"
        return text
    if attribute_key == "moisture_rating":
        if "wet" in lowered:
            return "Wet rated"
        if "damp" in lowered:
            return "Damp rated"
        if "dry" in lowered:
            return "Dry rated"
        return text
    if attribute_key == "ip_rating":
        return text.upper().replace(" ", "")
    if attribute_key in {"mount_type", "mounting"}:
        if "wall" in lowered:
            return "Wall mount"
        if "ceiling" in lowered:
            return "Ceiling mount"
        if "hang" in lowered or "suspend" in lowered:
            return "Hanging mount"
        if "recess" in lowered:
            return "Recessed mount"
        if "surface" in lowered:
            return "Surface mount"
        return text
    if attribute_key in {"power_type", "wiring_type"}:
        if "hardwire" in lowered:
            return "Hardwired"
        if "plug" in lowered:
            return "Plug-in"
        if "battery" in lowered:
            return "Battery powered"
        return text
    if attribute_key == "linkable":
        if lowered in PLACEHOLDER_VALUES or lowered in {"no", "false"}:
            return None
        return "Linkable / daisy-chain"
    if attribute_key == "motion_sensor":
        if lowered in {"yes", "true", "included"} or "motion" in lowered:
            return "Motion sensor"
        return None
    if attribute_key == "daylight_sensor_auto_dimming":
        if lowered in {"yes", "true", "included"} or "daylight" in lowered:
            return "Daylight sensor / auto-dimming"
        return None
    if attribute_key == "beam_angle":
        number = re.search(r"\d+(?:\.\d+)?", text)
        return f"{number.group(0)} degree beam angle" if number else text
    if attribute_key == "voltage":
        voltage_match = re.search(r"\b\d{2,3}\s*-\s*\d{2,3}\s*V?\b|\b\d{2,3}\s*V\b", text, flags=re.IGNORECASE)
        voltage = (voltage_match.group(0) if voltage_match else text).upper().replace(" ", "")
        if re.search(r"\d", voltage) and "V" not in voltage:
            voltage = f"{voltage}V"
        return voltage
    if attribute_key == "wattage":
        wattage = text.upper().replace(" ", "")
        if re.search(r"\d", wattage) and "W" not in wattage:
            wattage = f"{wattage}W"
        return wattage
    if attribute_key == "lumens":
        lumens = text.upper().replace(" ", "")
        if re.search(r"\d", lumens) and "LM" not in lumens and "LUMEN" not in lumens:
            lumens = f"{lumens}LM"
        return lumens
    if attribute_key == "cri":
        number = re.search(r"\d{2,3}", text)
        return f"CRI {number.group(0)}+" if number else text
    if attribute_key in {"bulb_base", "base_type", "socket_base"}:
        base = re.search(r"\bE(?:12|17|26|39)\b", text, flags=re.IGNORECASE)
        return f"{base.group(0).upper()} socket" if base else text
    if attribute_key == "finish":
        return text if "finish" in lowered else f"{text} finish"
    return text


def generated_match_terms(attribute: str, raw_value: Any, label: str) -> list[str]:
    terms = [label]
    attribute_key = slugify(attribute).replace("-", "_")
    raw_text = clean_generated_profile_value(raw_value)
    raw_key = raw_text.lower()
    numeric_attributes = {"beam_angle", "cri", "lumens", "voltage", "wattage"}
    if (
        raw_text
        and raw_key not in BROAD_GENERATED_MATCH_TERMS
        and not (attribute_key in numeric_attributes and re.fullmatch(r"[\d,.\s/-]+", raw_text))
    ):
        terms.append(raw_text)
    if label.endswith(" selectable"):
        terms.append(label.replace(" selectable", ""))
    return sorted({term for term in terms if term})


def generated_attribute_recommendation(group: str, label: str) -> str:
    if group == "CCT Strategy":
        return f"Keep {label} in scope when direct competitors show it as a repeated buying filter."
    if group == "Controls":
        return f"Include {label} when the comparable set or channel use case supports the added controls complexity."
    if group == "Installation / Linking":
        return f"Treat {label} as an RFQ compatibility requirement when it matches the intended install path."
    if group == "Ratings / Certifications":
        return f"Request {label} support when the target channel or installation environment requires it."
    if group == "Performance Configuration":
        return f"Use {label} when it supports the target performance tier without weakening price or margin."
    return f"Use {label} when it is part of the target SKU positioning and competitor evidence supports it."


def generated_profile_candidates(connection: sqlite3.Connection, category_id: int) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str]] = set()
    rows = connection.execute(
        """
        SELECT attribute_name, attribute_value, occurrence_count, product_count, coverage_pct, signal_class, source
        FROM category_attribute_distribution
        WHERE category_id = ?
        ORDER BY occurrence_count DESC, coverage_pct DESC, attribute_name, attribute_value
        """,
        (category_id,),
    ).fetchall()
    for attribute, raw_value, occurrence_count, product_count, coverage_pct, signal_class, source in rows:
        if "category_attribute_profiles.json" in str(source or "").replace("\\", "/"):
            continue
        attribute_key = slugify(attribute).replace("-", "_")
        group = ATTRIBUTE_GROUPS.get(attribute_key)
        if not group:
            continue
        for value in split_generated_profile_values(attribute_key, raw_value):
            label = normalize_generated_label(attribute_key, value)
            if not label:
                continue
            if label.lower() in {"non-dimmable"}:
                continue
            marker = (group, label.lower(), attribute_key)
            if marker in seen:
                continue
            seen.add(marker)
            candidates.append(
                {
                    "attribute": attribute_key,
                    "raw_value": value,
                    "label": label,
                    "group": group,
                    "occurrence_count": int(occurrence_count or 0),
                    "product_count": int(product_count or 0),
                    "coverage_pct": float(coverage_pct or 0),
                    "signal_class": signal_class or "",
                }
            )
    return candidates


def generated_keyword_maps(candidates: list[dict[str, Any]]) -> dict[str, dict[str, str]]:
    maps = {
        "style_keywords": {},
        "finish_keywords": {},
        "material_keywords": {},
        "mounting_keywords": {},
        "power_keywords": {},
        "control_keywords": {},
        "emergency_keywords": {},
        "color_mode_keywords": {},
        "certification_keywords": {},
    }
    for item in candidates:
        label = str(item["label"])
        attribute = str(item["attribute"])
        raw_value = str(item["raw_value"])
        target_map = None
        if item["group"] == "CCT Strategy":
            target_map = maps["color_mode_keywords"]
        elif item["group"] == "Controls":
            target_map = maps["control_keywords"]
        elif item["group"] == "Installation / Linking":
            target_map = maps["mounting_keywords"]
        elif item["group"] == "Ratings / Certifications":
            target_map = maps["certification_keywords"]
        elif item["group"] == "Performance Configuration":
            target_map = maps["power_keywords"]
        elif attribute == "finish":
            target_map = maps["finish_keywords"]
        elif attribute == "material":
            target_map = maps["material_keywords"]
        if target_map is None:
            continue
        for term in generated_match_terms(attribute, raw_value, label):
            target_map.setdefault(term.lower(), label)
    return maps


def category_is_accessory(slug: str) -> bool:
    return slug in {
        "dimmers",
        "goof_ring",
        "low_voltage_transformers",
        "rough_ins",
        "sensors",
        "trim",
        "wire",
        "wire_connectors",
    }


def generated_static_features(slug: str, candidates: list[dict[str, Any]]) -> list[str]:
    labels = [
        str(item["label"])
        for item in candidates
        if item["coverage_pct"] >= 60
        and item["group"] in {"Controls", "Ratings / Certifications", "Installation / Linking"}
        and item["attribute"] != "usage"
        and str(item["label"]).lower() not in {"non-dimmable", "dry rated"}
    ]
    if not labels and not category_is_accessory(slug):
        labels.append("dimmable")
    return list(dict.fromkeys(labels))[:8]


def category_filtered_candidates(slug: str, candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not category_is_accessory(slug):
        return candidates
    allowed_groups = {"Physical / Compatibility", "Ratings / Certifications", "Performance Configuration"}
    blocked_attributes = {
        "daylight_sensor_auto_dimming",
        "dimmable",
        "dimming_details",
        "linkable",
        "motion_sensor",
        "mount_type",
        "power_type",
        "sensor_details",
        "wiring_type",
    }
    return [
        item
        for item in candidates
        if item["group"] in allowed_groups and item["attribute"] not in blocked_attributes
    ]


def generated_feature_fields() -> list[dict[str, Any]]:
    return [
        {
            "path": "identity.ideation_name",
            "extractors": [
                "size_range",
                "style_keywords",
                "mounting_keywords",
                "color_mode_keywords",
                "power_keywords",
                "control_keywords",
                "certification_keywords",
            ],
        },
        {
            "path": "physical.size_form_factor",
            "extractors": ["size_range", "style_keywords", "mounting_keywords"],
        },
        {"path": "physical.mounting_type", "extractors": ["mounting_keywords"]},
        {"path": "physical.finish_color", "extractors": ["finish_keywords"]},
        {"path": "physical.material", "extractors": ["material_keywords"]},
        {"path": "electrical.cct_primary", "extractors": ["color_mode_keywords"]},
        {"path": "electrical.wattage_primary", "extractors": ["power_keywords"]},
        {"path": "controls.dimming_type", "extractors": ["control_keywords"]},
        {"path": "features.certifications", "extractors": ["certification_keywords"]},
    ]


def generated_baseline_attributes(candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows = []
    for item in candidates:
        if item["coverage_pct"] < 60:
            continue
        label = str(item["label"])
        if label.lower() in {"dry rated", "non-dimmable"}:
            continue
        group = str(item["group"])
        rows.append(
            {
                "attribute": group,
                "label": label,
                "type": "Baseline",
                "match_terms": generated_match_terms(str(item["attribute"]), item["raw_value"], label),
                "recommendation": f"Treat {label} as a baseline RFQ expectation for this category unless the concept intentionally differs.",
                "notes": "Generated from current category catalog/spec distribution.",
            }
        )
    rows.sort(key=lambda row: (row["attribute"], row["label"]))
    return rows[:10]


def generic_generated_attribute_groups() -> list[dict[str, Any]]:
    return [
        {
            "attribute": "Core Feature",
            "type": "Feature",
            "decision_question": "Which repeated feature is important enough to keep in scope?",
            "values": [
                {
                    "label": "Dimmable",
                    "match_terms": ["dimmable", "dimming"],
                    "recommendation": "Keep dimmability when direct competitor examples or the use case make it important.",
                },
                {
                    "label": "Selectable CCT",
                    "match_terms": ["selectable cct", "cct selectable", "color selectable", "selectable color temperature"],
                    "recommendation": "Keep selectable CCT when direct competitor examples show it as a repeated buying filter.",
                },
            ],
        }
    ]


def generated_attribute_groups(candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    seen: set[tuple[str, str]] = set()
    for item in candidates:
        if item["coverage_pct"] < 10:
            continue
        group = str(item["group"])
        label = str(item["label"])
        marker = (group, label.lower())
        if marker in seen:
            continue
        seen.add(marker)
        grouped[group].append(item)

    output = []
    for group in [
        "CCT Strategy",
        "Controls",
        "Installation / Linking",
        "Performance Configuration",
        "Physical / Compatibility",
        "Ratings / Certifications",
    ]:
        values = []
        for item in sorted(grouped.get(group, []), key=lambda row: (-row["coverage_pct"], row["label"]))[:8]:
            label = str(item["label"])
            entry = {
                "label": label,
                "match_terms": generated_match_terms(str(item["attribute"]), item["raw_value"], label),
                "recommendation": generated_attribute_recommendation(group, label),
            }
            if group == "Ratings / Certifications" and label in {"ETL", "UL", "DLC", "FCC", "RoHS", "Energy Star"}:
                entry["matcher"] = "certification"
            values.append(entry)
        if values:
            output.append(
                {
                    "attribute": group,
                    "type": GROUP_TYPES.get(group, "Feature"),
                    "decision_question": f"Which {group.lower()} attributes should this SKU carry?",
                    "values": values,
                }
            )
    return output or generic_generated_attribute_groups()


def build_generated_profiles(
    connection: sqlite3.Connection,
    category_rows: list[dict[str, Any]],
    category_ids: dict[str, int],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    feature_profiles: list[dict[str, Any]] = []
    decision_profiles: list[dict[str, Any]] = []
    for category in category_rows:
        slug = category["slug"]
        category_id = category_ids.get(slug)
        if not category_id:
            continue
        candidates = category_filtered_candidates(slug, generated_profile_candidates(connection, category_id))
        aliases = list(dict.fromkeys([category["run_name"], category["category"], *category.get("aliases", [])]))
        match_keywords = [
            keyword
            for alias in aliases
            for keyword in [alias, alias.replace(" lights", ""), alias.replace(" lighting", "")]
            if keyword
        ]
        maps = generated_keyword_maps(candidates)
        feature_profile = {
            "id": f"generated_{slug}",
            "label": f"{category['run_name']} generated SKU signals",
            "category_matches": [category["run_name"]],
            "subcategory_matches": aliases,
            "match_keywords": sorted(set(match_keywords)),
            "static_features": generated_static_features(slug, candidates),
            "feature_fields": generated_feature_fields(),
            **maps,
            "max_features": 14,
            "_category_id": category_id,
        }
        decision_profile = {
            "id": f"generated_{slug}_attributes",
            "label": f"{category['run_name']} Generated Attribute Decisions",
            "profile_ids": [f"generated_{slug}"],
            "category_matches": [category["run_name"]],
            "match_keywords": sorted(set(match_keywords)),
            "baseline_attributes": generated_baseline_attributes(candidates),
            "attribute_groups": generated_attribute_groups(candidates),
            "_category_id": category_id,
        }
        feature_profiles.append(feature_profile)
        decision_profiles.append(decision_profile)
    return feature_profiles, decision_profiles


def write_generated_profile_configs(
    paths: ProjectPaths,
    feature_profiles: list[dict[str, Any]],
    decision_profiles: list[dict[str, Any]],
) -> tuple[Path, Path]:
    target_dir = paths.source_data / "category_intelligence"
    target_dir.mkdir(parents=True, exist_ok=True)

    def public_profile(profile: dict[str, Any]) -> dict[str, Any]:
        return {key: value for key, value in profile.items() if not key.startswith("_")}

    feature_path = target_dir / GENERATED_FEATURE_PROFILE_FILENAME
    feature_path.write_text(
        json.dumps(
            {
                "version": date.today().isoformat(),
                "description": "Generated baseline feature-signal profiles from current category intelligence data.",
                "profiles": [public_profile(profile) for profile in feature_profiles],
            },
            indent=2,
            ensure_ascii=True,
        ),
        encoding="utf-8",
    )
    decision_path = target_dir / GENERATED_ATTRIBUTE_DECISION_PROFILE_FILENAME
    decision_path.write_text(
        json.dumps(
            {
                "version": date.today().isoformat(),
                "description": "Generated baseline Section F attribute decision profiles from current category intelligence data.",
                "profiles": [public_profile(profile) for profile in decision_profiles],
            },
            indent=2,
            ensure_ascii=True,
        ),
        encoding="utf-8",
    )
    return feature_path, decision_path


def insert_generated_feature_profiles(
    connection: sqlite3.Connection,
    paths: ProjectPaths,
    category_rows: list[dict[str, Any]],
    category_ids: dict[str, int],
) -> int:
    feature_profiles, decision_profiles = build_generated_profiles(connection, category_rows, category_ids)
    feature_path, _ = write_generated_profile_configs(paths, feature_profiles, decision_profiles)
    now = utc_now()
    count = 0
    for profile in feature_profiles:
        category_id = int(profile["_category_id"])
        connection.execute(
            """
            INSERT INTO category_feature_signal_profile(
                category_id, profile_key, label, applies_to_json, feature_signals_json,
                source, source_reference, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                category_id,
                profile["id"],
                profile["label"],
                json_text(profile.get("category_matches") or []),
                json_text(flatten_profile_signals(profile)),
                "generated_category_intelligence_profile",
                str(feature_path),
                now,
            ),
        )
        count += 1
    return count


def insert_sku_decoder_codes(
    connection: sqlite3.Connection,
    paths: ProjectPaths,
    category_ids: dict[str, int],
) -> int:
    path = paths.source_data / "sku_decoder" / "sku_decoder_clean.csv"
    if not path.exists():
        return 0
    now = utc_now()
    count = 0
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            mapped_slug = (row.get("mapped_category_slug") or "").strip()
            category_id = category_ids.get(mapped_slug) if mapped_slug else None
            connection.execute(
                """
                INSERT INTO sku_decoder_codes(
                    category_id, code_category, normalized_code_category, code, match_prefix,
                    code_meaning, mapped_category_slug, mapped_attribute, line_review_match,
                    source, source_reference, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    category_id,
                    row.get("code_category") or "",
                    row.get("normalized_code_category") or "",
                    row.get("code") or "",
                    row.get("match_prefix") or "",
                    row.get("code_meaning") or "",
                    mapped_slug,
                    row.get("mapped_attribute") or "",
                    1 if str(row.get("line_review_match") or "").strip().lower() in {"1", "true", "yes"} else 0,
                    "sku_decoder_clean.csv",
                    str(path),
                    now,
                ),
            )
            count += 1
    return count


def insert_private_catalog_exports(
    connection: sqlite3.Connection,
    paths: ProjectPaths,
    lookup: dict[str, dict[str, Any]],
    category_ids: dict[str, int],
) -> tuple[int, int, list[str]]:
    folder = paths.source_data / "postgres_exports"
    if not folder.exists():
        return 0, 0, []
    now = utc_now()
    product_count = 0
    spec_count = 0
    warnings: list[str] = []
    for path in folder.glob("*catalog*.json"):
        payload = read_json(path)
        rows = payload.get("products") if isinstance(payload, dict) else payload
        if not isinstance(rows, list):
            warnings.append(f"Skipped unsupported catalog export shape: {path.name}")
            continue
        for item in rows:
            if not isinstance(item, dict):
                continue
            category = resolve_category_from_values(
                lookup,
                item.get("category"),
                item.get("category_mapping"),
                item.get("product_type"),
                item.get("title"),
                item.get("product_title"),
                item.get("name"),
            )
            if not category:
                warnings.append(f"Could not map catalog export row SKU '{item.get('sku') or item.get('master_sku')}' from {path.name}")
                continue
            category_id = category_ids[category["slug"]]
            sku = item.get("sku") or item.get("master_sku")
            title = item.get("title") or item.get("product_title") or item.get("name") or sku or "Unnamed catalog product"
            connection.execute(
                """
                INSERT INTO shopify_category_products(
                    category_id, sku, product_title, product_type, tags_json, shopify_url,
                    image_url, active_status, category_mapping, source, source_reference, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    category_id,
                    sku,
                    title,
                    item.get("product_type"),
                    json_text(item.get("tags") or []),
                    item.get("shopify_url") or item.get("url"),
                    item.get("image_url"),
                    item.get("status") or item.get("active_status"),
                    item.get("category_mapping") or item.get("category"),
                    "private_postgres_export",
                    str(path),
                    now,
                ),
            )
            product_count += 1
            specs = item.get("specs") or {}
            specs.update(extract_attribute_values(" ".join(str(value or "") for value in [title, item.get("tags")])))
            for attribute, value in specs.items():
                values = value if isinstance(value, list) else [value]
                for one_value in values:
                    if one_value in (None, "", [], {}):
                        continue
                    connection.execute(
                        """
                        INSERT INTO shopify_spec_attributes(
                            category_id, sku, attribute_name, attribute_value, source, source_reference, created_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (category_id, sku, str(attribute), str(one_value), "private_postgres_export", str(path), now),
                    )
                    spec_count += 1
    return product_count, spec_count, warnings


def iter_clean_catalog_product_rows(paths: ProjectPaths) -> list[dict[str, str]]:
    folder = paths.source_data / "catalog_specs"
    path = folder / "sunco_catalog_products_clean.csv"
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def iter_clean_catalog_attribute_rows(paths: ProjectPaths) -> list[dict[str, str]]:
    folder = paths.source_data / "catalog_specs"
    path = folder / "sunco_catalog_spec_attributes_long.csv"
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def insert_clean_catalog_specs(
    connection: sqlite3.Connection,
    paths: ProjectPaths,
    lookup: dict[str, dict[str, Any]],
    category_ids: dict[str, int],
) -> tuple[int, int, list[str]]:
    """Import the normalized backend catalog/spec reference CSVs."""
    product_rows = iter_clean_catalog_product_rows(paths)
    attribute_rows = iter_clean_catalog_attribute_rows(paths)
    if not product_rows and not attribute_rows:
        return 0, 0, []

    now = utc_now()
    warnings: list[str] = []
    product_count = 0
    spec_count = 0
    category_by_key: dict[str, dict[str, Any]] = {}
    category_by_sku: dict[str, dict[str, Any]] = {}

    for item in product_rows:
        catalog_key = item.get("catalog_key") or item.get("sku") or item.get("source_handle") or item.get("title")
        category = resolve_category_from_values(
            lookup,
            item.get("category_mapping"),
            item.get("product_type"),
            item.get("title"),
            item.get("raw_info"),
            item.get("source_handle"),
        )
        if not category:
            warnings.append(f"Could not map clean catalog row '{catalog_key}' to a system category.")
            continue
        category_id = category_ids[category["slug"]]
        sku = item.get("sku") or None
        connection.execute(
            """
            INSERT INTO shopify_category_products(
                category_id, sku, product_title, product_type, tags_json, shopify_url,
                image_url, active_status, category_mapping, source, source_reference, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                category_id,
                sku,
                item.get("title") or item.get("raw_info") or sku or "Unnamed catalog product",
                item.get("product_type"),
                json_text([]),
                item.get("shopify_url"),
                item.get("image_url"),
                item.get("active_status") or "catalog_reference",
                item.get("category_mapping") or item.get("product_type"),
                "clean_catalog_specs_reference",
                "backend/source_data/catalog_specs/sunco_catalog_products_clean.csv",
                now,
            ),
        )
        product_count += 1
        if catalog_key:
            category_by_key[str(catalog_key)] = category
        if sku:
            category_by_sku[str(sku)] = category

    for item in attribute_rows:
        sku = item.get("sku") or None
        category = None
        if sku:
            category = category_by_sku.get(str(sku))
        if not category and item.get("catalog_key"):
            category = category_by_key.get(str(item.get("catalog_key")))
        if not category:
            category = resolve_category_from_values(
                lookup,
                item.get("category_mapping"),
                item.get("product_type"),
                item.get("title"),
            )
        if not category:
            warnings.append(f"Could not map clean catalog spec row '{item.get('catalog_key') or sku}' to a system category.")
            continue
        attribute = (item.get("attribute_name") or "").strip()
        value = (item.get("attribute_value") or "").strip()
        if not attribute or not value:
            continue
        connection.execute(
            """
            INSERT INTO shopify_spec_attributes(
                category_id, sku, attribute_name, attribute_value, source, source_reference, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                category_ids[category["slug"]],
                sku,
                attribute,
                value,
                "clean_catalog_specs_reference",
                "backend/source_data/catalog_specs/sunco_catalog_spec_attributes_long.csv",
                now,
            ),
        )
        spec_count += 1

    return product_count, spec_count, warnings


def insert_catalog_attribute_distribution(connection: sqlite3.Connection) -> int:
    """Roll SKU/spec rows into category-level attribute distributions."""
    now = utc_now()
    product_counts = {
        int(row[0]): int(row[1])
        for row in connection.execute(
            """
            SELECT category_id, COUNT(*)
            FROM shopify_category_products
            GROUP BY category_id
            """
        )
    }
    inserted = 0
    rows = connection.execute(
        """
        SELECT category_id, attribute_name, attribute_value, COUNT(*) AS occurrence_count
        FROM shopify_spec_attributes
        GROUP BY category_id, attribute_name, attribute_value
        """
    ).fetchall()
    for category_id, attribute, value, occurrence_count in rows:
        total = product_counts.get(int(category_id), 0)
        if total <= 0:
            continue
        coverage = round((int(occurrence_count) / total) * 100, 2)
        if coverage >= 60:
            signal = "table_stakes"
        elif coverage >= 25:
            signal = "common"
        elif coverage >= 5:
            signal = "emerging"
        else:
            signal = "rare"
        connection.execute(
            """
            INSERT INTO category_attribute_distribution(
                category_id, attribute_name, attribute_value, occurrence_count, product_count,
                coverage_pct, signal_class, source, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(category_id),
                str(attribute),
                str(value),
                int(occurrence_count),
                total,
                coverage,
                signal,
                "clean_catalog_specs_reference",
                now,
            ),
        )
        inserted += 1
    return inserted


def insert_redshift_cache_inventory(
    connection: sqlite3.Connection,
    paths: ProjectPaths,
    lookup: dict[str, dict[str, Any]],
    category_ids: dict[str, int],
) -> int:
    folder = paths.source_data / "redshift_stackline_cache"
    if not folder.exists():
        return 0
    now = utc_now()
    count = 0
    for path in folder.glob("*_stackline_redshift_*.json"):
        category = resolve_category(lookup, path.name.split("_stackline_redshift_")[0].replace("_", " "))
        if not category:
            continue
        payload = read_json(path)
        channels = payload.get("channels") if isinstance(payload, dict) else None
        channel_names = channels.keys() if isinstance(channels, dict) else ["redshift_cache"]
        for channel in channel_names:
            cursor = connection.execute(
                """
                INSERT INTO stackline_segments(
                    category_id, channel, segment_name, source, source_reference, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (category_ids[category["slug"]], str(channel), path.stem, "redshift_stackline_cache", str(path), now),
            )
            count += 1
            segment_id = int(cursor.lastrowid)
            for item in ((payload.get("top_competitor_products") if isinstance(payload, dict) else None) or [])[:25]:
                connection.execute(
                    """
                    INSERT INTO stackline_top_products(
                        category_id, segment_id, channel, brand, title, asin_sku, price,
                        reviews, rating, retail_sales, units_sold, sales_share_pct, source, source_reference, created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        category_ids[category["slug"]],
                        segment_id,
                        str(channel),
                        item.get("brand"),
                        item.get("product_title") or item.get("title") or "Unnamed Stackline product",
                        item.get("sku") or item.get("asin") or item.get("model_number"),
                        item.get("avg_retail_price") or item.get("price"),
                        item.get("review_count"),
                        item.get("rating"),
                        item.get("retail_sales"),
                        item.get("units_sold"),
                        item.get("sales_share_pct"),
                        "redshift_stackline_cache",
                        str(path),
                        now,
                    ),
                )
    return count


def build_category_intelligence_database(paths: ProjectPaths, output_path: Path | None = None) -> dict[str, Any]:
    paths.ensure()
    target = output_path or (paths.source_data / "category_intelligence" / "sunco_category_intelligence.sqlite")
    target.parent.mkdir(parents=True, exist_ok=True)
    schema_path = paths.app / "opportunity_engine" / "category_intelligence_schema.sql"
    category_rows = load_category_rows(paths)
    lookup = category_lookup(category_rows)
    warnings: list[str] = []
    counts: dict[str, int] = {}

    with sqlite3.connect(target) as connection:
        reset_database(connection, schema_path)
        category_ids = insert_categories(connection, category_rows)
        counts["categories"] = len(category_ids)
        gap_count, manifest_counts, manifest_warnings = insert_gap_manifests(connection, paths, lookup, category_ids)
        counts["gap_evidence"] = gap_count
        warnings.extend(manifest_warnings)
        product_count, spec_count, catalog_warnings = insert_private_catalog_exports(connection, paths, lookup, category_ids)
        counts["shopify_catalog_products"] = product_count
        counts["shopify_catalog_spec_attributes"] = spec_count
        warnings.extend(catalog_warnings)
        clean_product_count, clean_spec_count, clean_catalog_warnings = insert_clean_catalog_specs(connection, paths, lookup, category_ids)
        counts["clean_catalog_products"] = clean_product_count
        counts["clean_catalog_spec_attributes"] = clean_spec_count
        warnings.extend(clean_catalog_warnings[:100])
        counts["catalog_attribute_distribution"] = insert_catalog_attribute_distribution(connection)
        counts["attribute_profile_defaults"] = insert_attribute_profiles(connection, paths, lookup, category_ids)
        counts["feature_profiles"] = insert_feature_profiles(connection, paths, lookup, category_ids)
        counts["sku_decoder_codes"] = insert_sku_decoder_codes(connection, paths, category_ids)
        counts["stackline_segments_from_redshift_cache_inventory"] = insert_redshift_cache_inventory(connection, paths, lookup, category_ids)
        counts["generated_feature_profiles"] = insert_generated_feature_profiles(connection, paths, category_rows, category_ids)

        if not (product_count or clean_product_count) and not any((paths.source_data / folder).exists() for folder in ["shopify", "catalog", "postgres_exports", "catalog_specs"]):
            warnings.append("No local Shopify/catalog product export folder was found; shopify_category_products remains empty until a private export is supplied.")
        connection.execute(
            """
            INSERT INTO refresh_audit(source, run_date, data_age_days, sql_used, row_counts_json, warnings_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "category_intelligence_builder",
                date.today().isoformat(),
                0,
                "SQLite builder from local CSV/JSON source data.",
                json_text({**counts, **manifest_counts}),
                json_text(warnings),
                utc_now(),
            ),
        )
        connection.commit()

    return {"database": str(target), "row_counts": counts, "warnings": warnings}


def main() -> None:
    parser = argparse.ArgumentParser(description="Build the local Sunco category intelligence SQLite database.")
    parser.add_argument("--root", default=str(Path.cwd()), help="Ideation Development project root.")
    parser.add_argument("--output", default=None, help="Optional SQLite output path.")
    args = parser.parse_args()
    paths = ProjectPaths.from_root(args.root)
    result = build_category_intelligence_database(paths, Path(args.output) if args.output else None)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
