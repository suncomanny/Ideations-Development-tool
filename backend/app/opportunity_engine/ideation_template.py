from __future__ import annotations

import json
import re
from copy import deepcopy
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.views import Pane, Selection

from .categories import Category
from .category_intelligence import format_intelligence_audit, load_category_intelligence
from .paths import ProjectPaths
from .sql_audit import collect_sql_text
from .url_validation import extract_urls, format_validation_notes, validate_urls
from .utils import newest_file, timestamp
from .validation import try_excel_com_open_save, validate_workbook
from .workbook_common import add_or_replace_audit_sheet, clear_row_values, ensure_template_copy


IDEATION_MAX_COL = 54


COLUMN_MAP = {
    "category": 1,
    "subcategory": 2,
    "ideation_name": 3,
    "sunco_reference_sku": 4,
    "reference_sku_source": 5,
    "strategy": 6,
    "voltage": 7,
    "wattage_primary": 8,
    "wattage_max": 9,
    "selectable_wattage": 10,
    "cct_primary": 11,
    "cct_max": 12,
    "selectable_cct": 13,
    "cri": 14,
    "lumens_target": 15,
    "efficiency": 16,
    "power_factor": 17,
    "dimmable": 18,
    "dimming_type": 19,
    "frequency": 20,
    "driver_type": 21,
    "size_form_factor": 22,
    "mounting_type": 23,
    "material": 24,
    "finish_color": 25,
    "ip_rating": 26,
    "moisture_rating": 27,
    "indoor_outdoor_use": 28,
    "operating_temperature": 29,
    "wiring_type": 30,
    "emergency_battery": 31,
    "run_time": 32,
    "charge_time": 33,
    "switching_time": 34,
    "motion_sensor": 35,
    "motion_duration": 36,
    "daylight_sensor_auto_dimming": 37,
    "smart_connected": 38,
    "linkable": 39,
    "bulb_base_type": 40,
    "bulb_shape": 41,
    "beam_angle": 42,
    "target_msrp": 43,
    "target_margin_shopify": 44,
    "target_margin_amazon": 45,
    "cost_type": 46,
    "target_vendor_cost": 47,
    "certifications": 48,
    "lifetime_hours": 49,
    "warranty": 50,
    "known_competitors": 51,
    "priority_channels": 52,
    "stackline_data": 53,
    "research_notes": 54,
}


def _template_path(paths: ProjectPaths) -> Path:
    template = paths.templates / "PRD_Research_Ideation_Template.xlsx"
    if template.exists():
        return template
    fallback = paths.source_data / "schema_references" / "PRD_Research_Indoor_Residential_High_Confidence_Ideations_2026-05-13.xlsx"
    if fallback.exists():
        return fallback
    raise FileNotFoundError("Missing PRD ideation template.")


def _attribute_profile_path() -> Path:
    return Path(__file__).with_name("category_attribute_profiles.json")


def _load_attribute_profiles() -> dict[str, Any]:
    profile_path = _attribute_profile_path()
    if not profile_path.exists():
        return {}
    return json.loads(profile_path.read_text(encoding="utf-8"))


def _merged_attribute_profile(category: Category, intelligence_defaults: dict[str, Any] | None = None) -> dict[str, Any]:
    profiles = _load_attribute_profiles()
    selected = profiles.get(category.slug) or {}
    inherited_key = selected.get("inherits")
    merged: dict[str, Any] = {}
    if inherited_key and inherited_key in profiles:
        merged.update(profiles[inherited_key])
        merged["defaults"] = dict(profiles[inherited_key].get("defaults") or {})
    merged.update({key: value for key, value in selected.items() if key != "defaults"})
    defaults = dict(merged.get("defaults") or {})
    defaults.update(selected.get("defaults") or {})
    defaults.update({key: value for key, value in (intelligence_defaults or {}).items() if value not in (None, "", [], {})})
    merged["defaults"] = defaults
    return merged


def _load_market_price_samples(paths: ProjectPaths, category: Category) -> tuple[list[dict[str, Any]], Path | None]:
    folder = paths.source_data / "market_price_samples"
    sample_file = newest_file(folder, [f"{category.slug}_market_price_samples_*.json"])
    if sample_file is None:
        return [], None
    payload = json.loads(sample_file.read_text(encoding="utf-8"))
    return list(payload.get("samples") or []), sample_file


def latest_gap_workbook(paths: ProjectPaths, category: Category) -> Path | None:
    return newest_file(
        paths.gap_category_outputs(category.slug),
        [f"{category.slug}_true_gaps_*.xlsx"],
    ) or newest_file(paths.gap_outputs, [f"{category.slug}_true_gaps_*.xlsx"])


def _clear_ideation_template(workbook) -> None:
    ws = workbook["Ideations"]
    ws["B1"] = None
    ws["E1"] = None
    for row in range(4, 104):
        clear_row_values(ws, row, IDEATION_MAX_COL)
    if "Source Mapping" in workbook.sheetnames:
        mapping = workbook["Source Mapping"]
        if mapping.max_row > 1:
            mapping.delete_rows(2, mapping.max_row - 1)


def _reset_sheet_views(workbook) -> None:
    """Keep Excel pane metadata valid after editing a copied template."""
    ws = workbook["Ideations"]
    ws.freeze_panes = None
    ws.sheet_view.pane = Pane(
        xSplit=3,
        ySplit=2,
        topLeftCell="D6",
        activePane="bottomRight",
        state="frozen",
    )
    ws.sheet_view.selection = [
        Selection(pane="topRight", activeCell="D1", sqref="D1"),
        Selection(pane="bottomLeft", activeCell="A6", sqref="A6"),
        Selection(pane="bottomRight", activeCell="D6", sqref="D6"),
    ]

    if "Source Mapping" in workbook.sheetnames:
        mapping = workbook["Source Mapping"]
        mapping.freeze_panes = None
        mapping.sheet_view.pane = Pane(
            ySplit=1,
            topLeftCell="A2",
            activePane="bottomLeft",
            state="frozen",
        )
        mapping.sheet_view.selection = [Selection(pane="bottomLeft", activeCell="A2", sqref="A2")]


def _read_gap_rows(gap_workbook: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(gap_workbook, data_only=False)
    rows: list[dict[str, Any]] = []
    exact_source_names: set[str] = set()

    if "Recommendations" in workbook.sheetnames:
        ws = workbook["Recommendations"]
        headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
        for row_index in range(2, ws.max_row + 1):
            record = {str(headers[col - 1]): ws.cell(row_index, col).value for col in range(1, len(headers) + 1)}
            recommendation = record.get("Recommendation")
            is_supplemental = str(recommendation or "").strip().lower().startswith("[supplemental candidate]")
            if recommendation and ((record.get("Priority") == "High" and record.get("Confidence") == "High") or is_supplemental):
                why_text = str(record.get("Why This Is A True Gap") or "").lower()
                source_label = "Amazon" if "amazon/stackline-derived display row" in why_text else "Sunco.com/ecommerce"
                exact_source_names.add(str(record.get("Recommendation")).strip().lower())
                rows.append({
                    "source": source_label,
                    "classification": "True Gap",
                    "subcategory": record.get("Subcategory"),
                    "name": record.get("Recommendation"),
                    "example": record.get("Competitor / Ecommerce Example"),
                    "url": record.get("Review Link"),
                    "sunco_check": record.get("Sunco Active-Catalog Check"),
                    "why": record.get("Why This Is A True Gap"),
                    "action": record.get("PM Action"),
                })

    if "Amazon Recommendations" in workbook.sheetnames:
        ws = workbook["Amazon Recommendations"]
        headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
        for row_index in range(2, ws.max_row + 1):
            record = {str(headers[col - 1]): ws.cell(row_index, col).value for col in range(1, len(headers) + 1)}
            recommendation = record.get("Amazon-channel recommendation")
            confidence = str(record.get("Confidence") or "").strip().lower()
            is_supporting_evidence = str(recommendation or "").strip().lower() in exact_source_names
            is_supplemental = str(recommendation or "").strip().lower().startswith("[supplemental candidate]")
            if recommendation and ((record.get("Priority") == "High" and (confidence == "high" or is_supporting_evidence)) or is_supplemental):
                classification = record.get("Amazon classification") or "Amazon recommendation"
                rows.append({
                    "source": "Amazon",
                    "classification": classification,
                    "subcategory": record.get("Subcategory"),
                    "name": record.get("Amazon-channel recommendation"),
                    "example": record.get("Example listing"),
                    "url": record.get("Review link"),
                    "sunco_check": record.get("Sunco Amazon coverage check"),
                    "why": record.get("Stackline / Amazon evidence"),
                    "action": record.get("PM action"),
                })

    workbook.close()
    return _dedupe_gap_rows(rows)


def _append_unique(existing: str | None, new: str | None) -> str | None:
    if not new:
        return existing
    if not existing:
        return new
    existing_parts = [part.strip() for part in str(existing).split("\n") if part.strip()]
    if str(new).strip() in existing_parts:
        return existing
    return f"{existing}\n{new}"


def _append_unique_source(existing: str | None, new: str | None) -> str | None:
    if not new:
        return existing
    if not existing:
        return new
    parts = [part.strip() for part in str(existing).split(" + ") if part.strip()]
    if str(new).strip() not in parts:
        parts.append(str(new).strip())
    return " + ".join(parts)


def _dedupe_gap_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    deduped: dict[str, dict[str, Any]] = {}
    for item in rows:
        key = str(item.get("name") or "").strip().lower()
        if not key:
            continue
        if key not in deduped:
            deduped[key] = item.copy()
            continue
        existing = deduped[key]
        existing["source"] = _append_unique_source(existing.get("source"), item.get("source"))
        for field in ["classification", "example", "url", "sunco_check", "why", "action"]:
            existing[field] = _append_unique(existing.get(field), item.get(field))
    return list(deduped.values())


def _strategy_from_classification(value: str | None) -> str:
    text = (value or "").lower()
    if "style" in text:
        return "Style Extension"
    if "feature" in text or "variant" in text:
        return "Feature Expansion"
    return "New Product"


def _evidence_text(item: dict[str, Any]) -> str:
    return "\n".join(
        str(item.get(key) or "")
        for key in ["name", "classification", "example", "why", "action", "sunco_check", "url"]
    )


def _unique_csv(parts: list[str]) -> str:
    output: list[str] = []
    for part in parts:
        text = part.strip()
        if text and text not in output:
            output.append(text)
    return ", ".join(output)


def _extract_bulb_bases(text: str) -> str | None:
    bases = re.findall(r"\bE(?:12|17|26|39)\b", text, flags=re.IGNORECASE)
    return _unique_csv([base.upper() for base in bases]) or None


def _extract_light_count(text: str) -> str | None:
    match = re.search(r"\b(\d+)\s*(?:-|to)\s*(\d+)\s*light\b", text, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}-{match.group(2)} light"
    matches = re.findall(r"\b(\d+)\s*-?\s*light\b", text, flags=re.IGNORECASE)
    if matches:
        first = f"{matches[0]}-light"
        if len(matches) > 1:
            return f"{first}; optional {matches[1]}-light"
        return first
    return None


def _extract_light_count_values(text: str) -> list[str]:
    values: list[str] = []
    for match in re.finditer(r"\b(\d+)\s*(?:-|to)\s*(\d+)\s*light\b", text, flags=re.IGNORECASE):
        for value in [match.group(1), match.group(2)]:
            if value not in values:
                values.append(value)
    for value in re.findall(r"\b(\d+)\s*-?\s*light\b", text, flags=re.IGNORECASE):
        if value not in values:
            values.append(value)
    return values


def _extract_pack_count_values(text: str) -> list[str]:
    values: list[str] = []
    for value in re.findall(r"\b(\d+)\s*[-/]?\s*(?:pack|pk|count)\b", text, flags=re.IGNORECASE):
        if value not in values:
            values.append(value)
    return values


def _extract_sizes(text: str) -> list[str]:
    return [
        re.sub(r"\s+", " ", match.group(0)).replace("inch", "in").strip()
        for match in re.finditer(r"\b\d+(?:\.\d+)?\s*(?:-|to)\s*\d+(?:\.\d+)?\s*(?:in|inch)\b", text, flags=re.IGNORECASE)
    ]


def _extract_panel_size_tokens(text: str) -> list[str]:
    values: list[str] = []
    for match in re.finditer(r"\b(?:1|2|4)\s*x\s*(?:2|4)\b", text, flags=re.IGNORECASE):
        value = re.sub(r"\s+", "", match.group(0)).lower()
        if value not in values:
            values.append(value)
    return values


def _extract_wattage_text(text: str) -> str | None:
    values: list[str] = []
    for match in re.finditer(r"\b\d+(?:\.\d+)?\s*W\b", text, flags=re.IGNORECASE):
        value = match.group(0).replace(" ", "").upper()
        if value not in values:
            values.append(value)
    if not values:
        return None
    return " / ".join(values[:5]) + " target from Step 1 evidence"


def _extract_lumens_text(text: str) -> str | None:
    values: list[str] = []
    for match in re.finditer(r"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\s*(?:lm|lumens)\b", text, flags=re.IGNORECASE):
        value = re.sub(r"\s+", " ", match.group(0)).strip()
        value = re.sub(r"\blm\b", "lumens", value, flags=re.IGNORECASE)
        if value not in values:
            values.append(value)
    if not values:
        return None
    return " / ".join(values[:4]) + " target from Step 1 evidence"


def _extract_cct_text(text: str) -> tuple[str | None, str | None]:
    lowered = text.lower()
    values: list[str] = []
    for match in re.finditer(r"\b(?:27|30|35|40|50|65)00K\b|\b\d{4}\s*K\b", text, flags=re.IGNORECASE):
        value = match.group(0).replace(" ", "").upper()
        if value not in values:
            values.append(value)
    if "5cct" in lowered:
        values = values or ["3000K", "3500K", "4000K", "5000K", "6500K"]
        return "5CCT selectable", "/".join(values[:5]) if values else "5CCT selectable"
    if "3cct" in lowered:
        values = values or ["3000K", "4000K", "5000K"]
        return "3CCT selectable", "/".join(values[:3]) if values else "3CCT selectable"
    if values:
        return "/".join(values[:4]), values[-1]
    if "selectable cct" in lowered or "switchable cct" in lowered:
        return "Selectable CCT target from Step 1 evidence", None
    return None, None


def _infer_integrated_mounting(text: str, default: str | None = None) -> str:
    lowered = text.lower()
    if "surface mount" in lowered or "surface-mount" in lowered:
        return "Surface mount"
    if "direct mount" in lowered or "direct-mount" in lowered:
        return "Direct ceiling mount"
    if "drop ceiling" in lowered or "lay in" in lowered or "lay-in" in lowered:
        return "Drop ceiling / lay-in grid"
    if "troffer" in lowered or "center basket" in lowered:
        return "Recessed grid / troffer"
    if "flat panel" in lowered or "panel" in lowered:
        return "Commercial ceiling panel mount"
    return default or "Mounting type TBD from Step 1 evidence and supplier file"


def _infer_integrated_form_factor(item: dict[str, Any], text: str) -> str:
    lowered = text.lower()
    pieces: list[str] = []
    sizes = _extract_panel_size_tokens(text)
    if sizes:
        pieces.append("/".join(sizes))
    if "center basket" in lowered:
        pieces.append("center basket troffer")
    elif "troffer" in lowered:
        pieces.append("troffer")
    elif "surface mount" in lowered or "surface-mount" in lowered:
        pieces.append("surface-mount panel")
    elif "grid frame" in lowered:
        pieces.append("grid frame panel")
    elif "flat panel" in lowered:
        pieces.append("flat panel")
    elif "panel" in lowered:
        pieces.append("panel")
    if not pieces:
        pieces.append(str(item.get("name") or "").strip() or "Integrated LED fixture")
    return ", ".join(piece for piece in pieces if piece)


def _is_decorative_profile(category: Category, profile: dict[str, Any]) -> bool:
    return profile.get("inherits") == "decorative_fixture" or category.slug in {"chandeliers"}


def _apply_integrated_led_enrichment(enriched: dict[str, Any], category: Category, item: dict[str, Any], text: str) -> None:
    cct_primary, cct_max = _extract_cct_text(text)
    lowered = text.lower()
    enriched["size_form_factor"] = _infer_integrated_form_factor(item, text)
    enriched["mounting_type"] = _infer_integrated_mounting(text, enriched.get("mounting_type"))
    enriched["material"] = enriched.get("material")
    enriched["finish_color"] = _infer_finish(text) or enriched.get("finish_color")
    enriched["bulb_base_type"] = None
    enriched["bulb_shape"] = None
    enriched["wattage_primary"] = _extract_wattage_text(text) or enriched.get("wattage_primary") or "Integrated LED wattage TBD from supplier file"
    enriched["wattage_max"] = enriched.get("wattage_max") or enriched["wattage_primary"]
    enriched["cct_primary"] = cct_primary or enriched.get("cct_primary") or "CCT TBD from Step 1 evidence and supplier file"
    enriched["cct_max"] = cct_max or enriched.get("cct_max")
    enriched["lumens_target"] = _extract_lumens_text(text) or enriched.get("lumens_target") or "Lumen output TBD from Step 1 evidence and supplier file"
    enriched["efficiency"] = enriched.get("efficiency") or "Validate lm/W from final wattage and lumen target"
    enriched["power_factor"] = enriched.get("power_factor") or ">=0.9 target for commercial integrated LED driver"
    enriched["operating_temperature"] = enriched.get("operating_temperature") or "Commercial indoor ambient; supplier rating TBD"
    if "triac" in lowered:
        enriched["dimming_type"] = "TRIAC dimming"
    elif "0-10v" in lowered or "0-10 v" in lowered:
        enriched["dimming_type"] = "0-10V dimming"
    if "selectable wattage" in lowered or "wattage selectable" in lowered or "multi-wattage" in lowered or "5-wattage" in lowered:
        enriched["selectable_wattage"] = "Yes"
    if "selectable cct" in lowered or "switchable cct" in lowered or "3cct" in lowered or "5cct" in lowered:
        enriched["selectable_cct"] = "Yes"
    if "emergency backup" in lowered or "emergency battery" in lowered:
        enriched["emergency_battery"] = "Yes - emergency backup named in Step 1 evidence"
    if "motion" in lowered or "sensor" in lowered:
        enriched["motion_sensor"] = "Yes - sensor/control feature named in Step 1 evidence"


def _infer_finish(text: str) -> str | None:
    lowered = text.lower()
    if "black/gold" in lowered or ("black" in lowered and "gold" in lowered):
        return "Matte black / gold or brass accent target"
    if "matte black" in lowered and "brass" in lowered:
        return "Matte black / brass"
    if "black/wood" in lowered or ("black" in lowered and "wood" in lowered):
        return "Black with wood-look accent target"
    if "rattan" in lowered or "woven" in lowered:
        return "Natural rattan / woven shade target"
    if "black" in lowered:
        return "Black finish target"
    return None


def _infer_material(text: str) -> str | None:
    lowered = text.lower()
    if "rattan" in lowered or "woven" in lowered:
        return "Metal frame with natural rattan/woven shade target"
    if "wood" in lowered:
        return "Metal frame with wood-look accents target"
    if any(term in lowered for term in ["chandelier", "linear", "wagon wheel", "black", "brass", "gold"]):
        return "Metal fixture body target"
    return None


def _infer_mounting(text: str) -> str:
    lowered = text.lower()
    if re.search(r"\brods?\b|\brod[- ]hung\b", lowered) or "linear" in lowered or "island" in lowered:
        return "Ceiling / rod-hung"
    if "wagon wheel" in lowered:
        return "Ceiling / chain-hanging"
    return "Ceiling / chain-hanging or rod-hung"


def _infer_form_factor_label(text: str) -> str:
    lowered = text.lower()
    if "wagon wheel" in lowered:
        return "wagon wheel chandelier"
    if "linear" in lowered or "island" in lowered:
        return "linear island chandelier"
    if "rattan" in lowered or "woven" in lowered:
        return "rattan/woven pendant"
    if "chandelier" in lowered:
        return "chandelier"
    return "fixture"


def _infer_size_form_factor(item: dict[str, Any], text: str) -> str:
    name = str(item.get("name") or "").strip()
    lowered = text.lower()
    pieces: list[str] = []
    light_count = _extract_light_count(text)
    sizes = _extract_sizes(text)
    if light_count:
        pieces.append(light_count)
    if "wagon wheel" in lowered:
        pieces.append("wagon wheel chandelier")
    elif "linear" in lowered or "island" in lowered:
        pieces.append("linear island chandelier")
    else:
        pieces.append(name)
    if sizes:
        pieces.append("; ".join(sizes) + " target")
    return ", ".join(piece for piece in pieces if piece)


def _sample_mentions_light_count(sample: dict[str, Any], light_count: str | None) -> bool:
    if not light_count:
        return False
    number = re.sub(r"\D+", "", str(light_count))
    if not number:
        return False
    haystack = " ".join(str(sample.get(key) or "") for key in ["product", "note", "url"]).lower()
    return bool(re.search(rf"\b{re.escape(number)}\s*[- ]?light\b", haystack, flags=re.IGNORECASE))


def _build_known_competitors(item: dict[str, Any]) -> str:
    parts = []
    for key in ["example", "url"]:
        value = item.get(key)
        if value:
            parts.append(str(value))
    return "\n".join(parts) or "Amazon and competitor pool from templates/Competitors.md"


def _normalize_match_text(value: str | None) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def _market_samples_for_item(item: dict[str, Any], samples: list[dict[str, Any]]) -> list[dict[str, Any]]:
    name = _normalize_match_text(item.get("name"))
    matched: list[dict[str, Any]] = []
    for sample in samples:
        match = _normalize_match_text(sample.get("ideation_match"))
        if not match:
            continue
        if match in name or name in match:
            matched.append(sample)
    return matched


def _percentile(values: list[float], percentile: float) -> float:
    if not values:
        raise ValueError("No values provided.")
    ordered = sorted(values)
    if len(ordered) == 1:
        return ordered[0]
    position = (len(ordered) - 1) * percentile
    lower = int(position)
    upper = min(lower + 1, len(ordered) - 1)
    weight = position - lower
    return ordered[lower] + (ordered[upper] - ordered[lower]) * weight


def _retail_price_point(value: float) -> float:
    rounded = int((value + 9.9999) // 10 * 10) - 0.01
    return max(0.99, round(rounded, 2))


def _format_money(value: float) -> str:
    return f"${value:,.2f}"


def _market_msrp_target(item: dict[str, Any], market_samples: list[dict[str, Any]]) -> dict[str, Any] | None:
    matched_samples = [
        sample for sample in _market_samples_for_item(item, market_samples)
        if isinstance(sample.get("price"), (int, float))
    ]
    target_light_count = str(item.get("_target_light_count") or "").strip()
    pricing_basis = "comparable listings"
    if target_light_count:
        light_count_samples = [sample for sample in matched_samples if _sample_mentions_light_count(sample, target_light_count)]
        pricing_basis = (
            f"exact {target_light_count}-light comparable listings"
            if light_count_samples
            else f"nearest parent-family comparable listings; no exact {target_light_count}-light sample was found in the local cache"
        )
        samples = light_count_samples or matched_samples
    else:
        samples = [sample for sample in matched_samples if sample.get("use_for_base_msrp", True) is not False]
    prices = [float(sample["price"]) for sample in samples]
    if not prices:
        return None

    p50 = _percentile(prices, 0.50)
    p55 = _percentile(prices, 0.55)
    target = _retail_price_point((p50 + p55) / 2)
    sample_labels = [
        f"{sample.get('retailer')} {sample.get('brand')}: {_format_money(float(sample['price']))}"
        for sample in samples
    ]
    sample_links = [
        f"{sample.get('retailer')} {sample.get('brand')}: {sample.get('url')}"
        for sample in samples
        if sample.get("url")
    ]
    return {
        "target": target,
        "text": (
            f"{_format_money(target)} market MSRP target; based on p50-p55 of {pricing_basis} "
            f"({_format_money(p50)}-{_format_money(p55)}). Validate against live pricing before launch."
        ),
        "vendor_cost_text": (
            f"Backsolve from AQ MSRP target {_format_money(target)} after supplier quotes; "
            "do not use margin targets as the initial MSRP anchor."
        ),
        "notes": f"Market MSRP pricing basis: {pricing_basis}. Samples: " + "; ".join(sample_labels),
        "links": "\n".join(sample_links),
    }


def _variant_size_form_factor(text: str, light_count: str | None, size: str | None) -> str:
    form_factor = _infer_form_factor_label(text)
    lowered = text.lower()
    if light_count == "1" and ("rattan" in lowered or "woven" in lowered):
        form_factor = "dome rattan/woven pendant"
    elif light_count == "3" and ("rattan" in lowered or "woven" in lowered):
        form_factor = "linear island rattan/woven pendant"
    pieces = []
    if light_count:
        pieces.append(f"{light_count}-light")
    pieces.append(form_factor)
    if size:
        pieces.append(size)
    return ", ".join(pieces)


def _variant_label(text: str, light_count: str | None, size: str | None) -> str:
    form_factor = _infer_form_factor_label(text)
    lowered = text.lower()
    if light_count == "1" and ("rattan" in lowered or "woven" in lowered):
        form_factor = "dome rattan/woven pendant"
    elif light_count == "3" and ("rattan" in lowered or "woven" in lowered):
        form_factor = "linear rattan/woven island pendant"
    finish = _infer_finish(text)
    parts: list[str] = []
    if light_count:
        parts.append(f"{light_count}-light")
    if size:
        parts.append(size)
    parts.append(form_factor)
    if finish:
        parts.append(finish.replace(" target", ""))
    return " ".join(parts)


def _candidate_name(base_name: str, label: str) -> str:
    supplemental_prefix = "[Supplemental candidate] "
    prefix = supplemental_prefix if base_name.startswith(supplemental_prefix) else ""
    clean_base = base_name.removeprefix(supplemental_prefix).strip()
    return f"{prefix}{clean_base} - {label}"


def _sku_candidate_items(category: Category, item: dict[str, Any]) -> list[dict[str, Any]]:
    text = _evidence_text(item)
    base_name = str(item.get("name") or "").strip()
    light_counts = _extract_light_count_values(text)
    pack_counts = _extract_pack_count_values(text)
    sizes = _extract_sizes(text)
    if category.slug == "under_cabinet" and pack_counts:
        candidates: list[dict[str, Any]] = []
        for pack_count in pack_counts[:4]:
            candidate = deepcopy(item)
            label = f"{pack_count}-pack"
            candidate["name"] = _candidate_name(base_name, label)
            candidate["_parent_opportunity"] = base_name
            candidate["_target_pack_count"] = pack_count
            candidate["_sku_candidate_rationale"] = (
                "Split from the parent opportunity because Step 1 evidence or PM action named this as a "
                "distinct kit pack-count option. Step 3 should research this row as one candidate SKU."
            )
            candidate["_field_overrides"] = {
                "size_form_factor": f"{pack_count}-pack under-cabinet kit",
                "research_notes": f"Pack-count target: {pack_count}-pack; validate battery capacity, mounting accessories, and Amazon pack economics.",
            }
            candidates.append(candidate)
        return candidates
    if category.slug != "chandeliers":
        return [item]
    if not light_counts and len(sizes) <= 1:
        return [item]

    candidates: list[dict[str, Any]] = []
    max_candidates = 4
    for index, light_count in enumerate(light_counts[:max_candidates]):
        size = sizes[index] if index < len(sizes) else None
        label = _variant_label(text, light_count, size)
        candidate = deepcopy(item)
        candidate["name"] = _candidate_name(base_name, label)
        candidate["_parent_opportunity"] = base_name
        candidate["_target_light_count"] = light_count
        candidate["_target_size"] = size or "TBD from supplier and live market refresh"
        candidate["_sku_candidate_rationale"] = (
            "Split from the parent opportunity because Step 1 evidence or PM action named this as a "
            "distinct SKU-level option. Step 3 should research this row as one candidate SKU, not as a broad family."
        )
        candidate["_field_overrides"] = {
            "size_form_factor": _variant_size_form_factor(text, light_count, size),
            "wattage_primary": f"Bulb-dependent; target {light_count}-light replaceable LED bulb design",
        }
        candidates.append(candidate)

    return candidates or [item]


def _url_validation_summary(cache: dict[str, dict[str, Any]]) -> str:
    if not cache:
        return "No review URLs found to validate."
    counts: dict[str, int] = {}
    flagged: list[str] = []
    for result in cache.values():
        state = str(result.get("state") or "unverified")
        counts[state] = counts.get(state, 0) + 1
        if state != "verified":
            flagged.append(f"{result.get('url')} -> {result.get('note')}")
    count_text = ", ".join(f"{state}: {count}" for state, count in sorted(counts.items()))
    if flagged:
        return count_text + "\nFlagged links:\n" + "\n".join(flagged)
    return count_text


def _build_research_notes(
    item: dict[str, Any],
    enriched: dict[str, Any],
    profile: dict[str, Any],
    pricing: dict[str, Any] | None,
    link_validation: str | None,
) -> str:
    notes = [
        f"Source: {item.get('source')}",
        f"Classification: {item.get('classification')}",
        f"Parent opportunity: {item.get('_parent_opportunity')}" if item.get("_parent_opportunity") else None,
        f"SKU candidate target: {item.get('_target_light_count')}-light; size target {item.get('_target_size')}" if item.get("_target_light_count") else None,
        f"SKU candidate target: {item.get('_target_pack_count')}-pack kit" if item.get("_target_pack_count") else None,
        f"SKU candidate rationale: {item.get('_sku_candidate_rationale')}" if item.get("_sku_candidate_rationale") else None,
        f"Evidence: {item.get('why')}",
        f"Sunco check: {item.get('sunco_check')}",
        f"Recommended action: {item.get('action')}",
        f"Review link: {item.get('url')}",
        f"Link validation: {link_validation}" if link_validation else None,
        "Attribute enrichment: fields were filled from Step 1 evidence first, then category-profile defaults when safe.",
        "Confidence note: validate final electrical specs, certifications, packaging, and claims against supplier files before PRD lock.",
    ]
    if profile.get("notes"):
        notes.append(f"Profile note: {profile.get('notes')}")
    if pricing:
        notes.append(pricing["notes"])
        if pricing.get("links"):
            notes.append("Market pricing links:\n" + pricing["links"])
    assumed = [
        key for key in [
            "voltage",
            "frequency",
            "dimmable",
            "dimming_type",
            "mounting_type",
            "wiring_type",
            "certifications",
        ]
        if enriched.get(key)
    ]
    if assumed:
        notes.append("Filled fields to validate: " + ", ".join(assumed))
    return "\n".join(part for part in notes if part and not str(part).endswith("None"))


def _enrich_item(
    category: Category,
    item: dict[str, Any],
    profile: dict[str, Any],
    market_samples: list[dict[str, Any]],
    url_validation_cache: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    defaults = dict(profile.get("defaults") or {})
    text = _evidence_text(item)
    bulb_base = _extract_bulb_bases(text)
    light_count = _extract_light_count(text)
    is_decorative = _is_decorative_profile(category, profile)
    pricing = _market_msrp_target(item, market_samples)
    link_validation = format_validation_notes(validate_urls(extract_urls(item.get("url")), url_validation_cache))
    enriched = {
        "category": category.run_name,
        "subcategory": defaults.get("subcategory") or category.run_name,
        "ideation_name": str(item.get("name") or "").strip(),
        "sunco_reference_sku": "TBD - use best-selling adjacent Sunco family by category revenue",
        "reference_sku_source": f"No exact Sunco {category.run_name} SKU identified from Step 1 evidence; use revenue proxy during refresh",
        "strategy": _strategy_from_classification(item.get("classification")),
        "known_competitors": _build_known_competitors(item),
    }
    enriched.update(defaults)

    if is_decorative:
        enriched["size_form_factor"] = _infer_size_form_factor(item, text)
        enriched["mounting_type"] = _infer_mounting(text)
        enriched["material"] = _infer_material(text) or enriched.get("material")
        enriched["finish_color"] = _infer_finish(text) or enriched.get("finish_color")
        enriched["bulb_base_type"] = bulb_base or "E12/E26 target; validate by design size and supplier options"
        if light_count and bulb_base:
            normalized_count = light_count.replace("; optional", " / optional")
            enriched["wattage_primary"] = f"Bulb-dependent; target {normalized_count} using {bulb_base} LED bulbs"
        elif light_count:
            enriched["wattage_primary"] = f"Bulb-dependent; target {light_count} replaceable LED bulbs"
        else:
            enriched["wattage_primary"] = "Bulb-dependent; confirm bulb count during supplier research"
        enriched["wattage_max"] = "Socket max TBD by supplier and certification file"
        enriched["cct_primary"] = "2700K-3000K warm white target if bulbs are included"
        enriched["cct_max"] = "3000K target; avoid selectable CCT unless integrated LED"
        enriched["lumens_target"] = "Bulb-dependent; validate total fixture output by included bulb bundle"
        enriched["efficiency"] = "Bulb-dependent"
        enriched["power_factor"] = "N/A for replaceable bulbs; validate if integrated LED"
        enriched["operating_temperature"] = "Residential indoor ambient; supplier rating TBD"
    elif category.slug == "under_cabinet":
        enriched["size_form_factor"] = _infer_integrated_form_factor(item, text)
        enriched["mounting_type"] = _infer_integrated_mounting(text, enriched.get("mounting_type"))
        enriched["material"] = enriched.get("material")
        enriched["finish_color"] = _infer_finish(text) or enriched.get("finish_color")
        enriched["bulb_base_type"] = None
        enriched["bulb_shape"] = None
    else:
        _apply_integrated_led_enrichment(enriched, category, item, text)

    if category.slug == "under_cabinet":
        lowered = text.lower()
        if "magnetic" in lowered or "rechargeable" in lowered or "puck" in lowered:
            enriched["mounting_type"] = "Magnetic and adhesive under-cabinet mount target"
            enriched["wiring_type"] = "Rechargeable battery / no-wire installation"
            enriched["motion_sensor"] = "Yes for rechargeable puck/bar variants"
        elif "tape" in lowered or "strip" in lowered:
            enriched["mounting_type"] = "Adhesive tape/channel under-cabinet mount target"
            enriched["wiring_type"] = "Plug-in or hardwired low-voltage driver options"
            enriched["motion_sensor"] = "No unless selected as a motion-kit variant"
            enriched["linkable"] = "Yes; extension/corner accessories expected"
        enriched["size_form_factor"] = str(enriched.get("size_form_factor") or "").replace("Ceiling / chain-hanging or rod-hung", "under-cabinet kit")
        if item.get("_target_pack_count"):
            enriched["size_form_factor"] = f"{item.get('_target_pack_count')}-pack under-cabinet kit"
        enriched["wattage_primary"] = "Integrated LED; wattage by puck/bar/tape length and pack count"
        enriched["lumens_target"] = "Task-light output by kit type; validate per-light and total kit lumens"
        enriched["bulb_base_type"] = None
        enriched["bulb_shape"] = None
    if pricing:
        enriched["target_msrp"] = pricing["text"]
        enriched["target_vendor_cost"] = pricing["vendor_cost_text"]
    for key, value in dict(item.get("_field_overrides") or {}).items():
        if key in COLUMN_MAP and value:
            enriched[key] = value
    enriched["research_notes"] = _build_research_notes(item, enriched, profile, pricing, link_validation)
    enriched["_link_validation"] = link_validation
    return enriched


def _fill_ideation_row(
    ws,
    row_index: int,
    category: Category,
    item: dict[str, Any],
    market_samples: list[dict[str, Any]],
    url_validation_cache: dict[str, dict[str, Any]],
    intelligence_defaults: dict[str, Any] | None = None,
) -> dict[str, Any]:
    profile = _merged_attribute_profile(category, intelligence_defaults)
    enriched = _enrich_item(category, item, profile, market_samples, url_validation_cache)
    values = {
        COLUMN_MAP[key]: value
        for key, value in enriched.items()
        if key in COLUMN_MAP
    }
    for col, value in values.items():
        cell = ws.cell(row_index, col)
        cell.value = value
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    return enriched


def generate_prd_ideation_workbook(paths: ProjectPaths, category: Category, gap_workbook: Path | None = None) -> tuple[Path, list[str]]:
    paths.ensure()
    intelligence = load_category_intelligence(paths, category)
    selected_gap = gap_workbook or latest_gap_workbook(paths, category)
    if selected_gap is None:
        raise FileNotFoundError(f"No Step 1 gap workbook found for {category.run_name}. Run Step 1 first.")

    output_folder = paths.prd_ideation_category_outputs(category.slug)
    output_folder.mkdir(parents=True, exist_ok=True)
    output = output_folder / f"{category.slug}_prd_ideations_{timestamp()}.xlsx"
    ensure_template_copy(_template_path(paths), output)
    workbook = load_workbook(output)
    _clear_ideation_template(workbook)

    ws = workbook["Ideations"]
    ws["B1"] = category.owner
    ws["E1"] = f"Prepared from {selected_gap.name} on {timestamp()}"
    rows = _read_gap_rows(selected_gap)
    candidate_rows: list[dict[str, Any]] = []
    for item in rows:
        candidate_rows.extend(_sku_candidate_items(category, item))
    market_samples, market_sample_path = _load_market_price_samples(paths, category)
    url_validation_cache: dict[str, dict[str, Any]] = {}
    enriched_rows: list[dict[str, Any]] = []
    for offset, item in enumerate(candidate_rows[:100], start=4):
        enriched_rows.append(_fill_ideation_row(ws, offset, category, item, market_samples, url_validation_cache, intelligence.attribute_defaults))

    if "Source Mapping" in workbook.sheetnames:
        mapping = workbook["Source Mapping"]
        if mapping.max_row < 1:
            mapping.append(["Ideation Name", "Why selected", "Key evidence used", "Reference SKU rationale"])
        for item, enriched in zip(candidate_rows[:100], enriched_rows):
            evidence_parts = [str(item.get("why") or ""), str(item.get("url") or "")]
            if enriched.get("_link_validation"):
                evidence_parts.append("Link validation: " + str(enriched.get("_link_validation")))
            if item.get("_sku_candidate_rationale"):
                evidence_parts.append("SKU candidate: " + str(item.get("_sku_candidate_rationale")))
            mapping.append([
                enriched.get("ideation_name") or item.get("name"),
                (
                    f"{item.get('classification')} with High priority and High confidence"
                    + (
                        f"; parent opportunity: {item.get('_parent_opportunity')}"
                        if item.get("_parent_opportunity")
                        else ""
                    )
                ),
                " | ".join(part for part in evidence_parts if part),
                "Use Sunco.com/Shopify SKU first when available; otherwise refresh should use best-selling item by category revenue.",
            ])
            for cell in mapping[mapping.max_row]:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        mapping.column_dimensions["A"].width = 42
        mapping.column_dimensions["B"].width = 48
        mapping.column_dimensions["C"].width = 90
        mapping.column_dimensions["D"].width = 72

    add_or_replace_audit_sheet(
        workbook,
        [
            ("Project", "sunco-product-opportunity-engine"),
            ("Category", category.run_name),
            ("Owner", category.owner),
            ("Generated", timestamp()),
            ("Source Step 1 workbook", str(selected_gap)),
            ("Step 1 opportunities selected", str(len(rows))),
            ("Step 2 SKU candidate rows", str(len(candidate_rows[:100]))),
            ("SKU expansion rule", "Step 2 creates one row per candidate SKU when Step 1 evidence or PM action names distinct SKU-level options such as light count, form factor, or size. It does not force a minimum row count."),
            ("Selection rule", "Priority = High and Confidence = High; true, feature, and style gaps are all allowed. Matching ideation names are merged across source tabs."),
            ("Reference SKU rule", "Sunco.com/Shopify SKU first; if absent, use best-selling item by category revenue."),
            ("MSRP target rule", "When market price samples exist, Target MSRP is based on the 50th-55th percentile of comparable listings and is independent from margin targets."),
            ("URL validation rule", "Step 2 checks Step 1 review URLs live when the workbook is generated. 2xx/3xx means verified; 403/429 means blocked by retailer/CDN and needs manual browser review; 404/410 means invalid and should be replaced."),
            ("URL validation summary", _url_validation_summary(url_validation_cache)),
            ("Market price sample file", str(market_sample_path) if market_sample_path else "No category market price sample file found."),
            ("Category intelligence audit", format_intelligence_audit(intelligence)),
        ],
        collect_sql_text(paths, category.slug),
    )

    _reset_sheet_views(workbook)
    for col in range(1, IDEATION_MAX_COL + 1):
        ws.cell(3, col).font = Font(bold=True)

    workbook.save(output)
    workbook.close()

    issues = validate_workbook(output, ["Instructions", "Ideations", "Source Mapping", "Run Audit"])
    ok, message = try_excel_com_open_save(output)
    if message:
        issues.append(message if not ok else message)
    return output, issues
