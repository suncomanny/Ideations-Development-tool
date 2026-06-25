from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any


LUMINAIRE_CATEGORY_SLUGS = {
    "area_lights",
    "bulbs_plus_tubes",
    "canopy",
    "ceiling_fixtures",
    "chandeliers",
    "clean_room_lighting",
    "commercial_grow_lights",
    "commercial_landscape",
    "commercial_security",
    "emergency",
    "explosion_proof",
    "flood_lights",
    "lamps",
    "linears",
    "outdoor_ceiling",
    "outdoor_security",
    "panels",
    "pendants",
    "residential_grow_lights",
    "residential_landscape",
    "residential_security",
    "slims",
    "smart_lighting",
    "sport_lights",
    "striplights",
    "tape_rope_light",
    "under_cabinet",
    "ufo",
    "vanity",
    "vaportights",
    "wall_packs",
    "wall_sconces",
    "well_lights",
    "wraparounds",
}


@dataclass(frozen=True)
class LuminairePerformance:
    lumen_values: tuple[int, ...]
    watt_values: tuple[int, ...]
    efficacy_values: tuple[float, ...]
    brightness_tier: str | None
    wattage_class: str | None
    user_brightness_target: str | None
    installer_load_target: str | None
    efficiency_note: str | None

    @property
    def has_signal(self) -> bool:
        return bool(self.lumen_values or self.watt_values)


def is_luminaire_category(category_slug: str) -> bool:
    return category_slug in LUMINAIRE_CATEGORY_SLUGS


def _clean_number(value: str) -> int | None:
    try:
        return int(float(value.replace(",", "")))
    except ValueError:
        return None


def _has_valid_comma_groups(value: str) -> bool:
    if "," not in value:
        return True
    head, *groups = value.split(".", 1)[0].split(",")
    return bool(head) and all(len(group) == 3 and group.isdigit() for group in groups)


def _clean_float(value: str) -> float | None:
    try:
        return float(value.replace(",", ""))
    except ValueError:
        return None


def _unique_ordered(values: list[int]) -> tuple[int, ...]:
    output: list[int] = []
    for value in values:
        if value > 0 and value not in output:
            output.append(value)
    return tuple(output)


def extract_lumens(text: str) -> tuple[int, ...]:
    values: list[int] = []
    for match in re.finditer(r"\b([\d,]+(?:\.\d+)?)\s*(?:lm|lumen|lumens|klm)\b(?!\s*/\s*w)", text, flags=re.IGNORECASE):
        raw = match.group(1)
        if not _has_valid_comma_groups(raw):
            continue
        unit = match.group(0).lower()
        if "klm" in unit:
            k_value = _clean_float(raw)
            if k_value is None:
                continue
            values.append(int(round(k_value * 1000)))
            continue
        value = _clean_number(raw)
        if value is None:
            continue
        values.append(value)
    for match in re.finditer(r"\b(\d+(?:\.\d+)?)\s*k\s*lm\b", text, flags=re.IGNORECASE):
        value = _clean_float(match.group(1))
        if value is not None:
            values.append(int(round(value * 1000)))
    return _unique_ordered(values)


def extract_watts(text: str) -> tuple[int, ...]:
    values: list[int] = []
    for match in re.finditer(r"\b(\d{1,4}(?:\.\d+)?)\s*w(?:att|atts)?\b", text, flags=re.IGNORECASE):
        value = _clean_number(match.group(1))
        if value is not None:
            values.append(value)
    return _unique_ordered(values)


def brightness_tier(lumens: int | None) -> str | None:
    if lumens is None:
        return None
    if lumens < 4000:
        return "low brightness"
    if lumens < 6500:
        return "standard brightness"
    if lumens < 8500:
        return "high output"
    if lumens <= 10000:
        return "ultra high output"
    return "extreme output"


def wattage_class(watts: int | None) -> str | None:
    if watts is None:
        return None
    if watts <= 20:
        return "20W class"
    if watts <= 30:
        return "30W class"
    if watts <= 40:
        return "40W class"
    if watts <= 50:
        return "50W class"
    if watts <= 60:
        return "60W class"
    if watts <= 75:
        return "72-75W class"
    return f"{watts}W+ class"


def normalize_luminaire_performance(*texts: Any) -> LuminairePerformance:
    text = " ".join(str(value or "") for value in texts)
    lumens = extract_lumens(text)
    watts = extract_watts(text)
    max_lumens = max(lumens) if lumens else None
    max_watts = max(watts) if watts else None
    efficacy_values: tuple[float, ...] = ()
    if max_lumens and max_watts:
        efficacy_values = (round(max_lumens / max_watts, 1),)
    tier = brightness_tier(max_lumens)
    w_class = wattage_class(max_watts)
    user_target = f"{max_lumens:,}lm {tier}" if max_lumens and tier else None
    installer_target = f"{w_class}; max parsed wattage {max_watts}W" if max_watts and w_class else None
    efficiency_note = None
    if max_lumens and max_watts:
        efficiency_note = (
            f"Treat {max_lumens:,}lm as the user-facing brightness target and {max_watts}W as the installer/load target. "
            f"Prefer matching the brightness tier at equal or lower wattage"
        )
    elif max_lumens:
        efficiency_note = f"Treat {max_lumens:,}lm as the user-facing brightness target; wattage/load target is not present in the parsed evidence"
    elif max_watts:
        efficiency_note = f"Treat {max_watts}W as the installer/load target; lumen brightness target is not present in the parsed evidence"
    return LuminairePerformance(
        lumen_values=lumens,
        watt_values=watts,
        efficacy_values=efficacy_values,
        brightness_tier=tier,
        wattage_class=w_class,
        user_brightness_target=user_target,
        installer_load_target=installer_target,
        efficiency_note=efficiency_note,
    )


def performance_note(label: str, performance: LuminairePerformance) -> str:
    if not performance.has_signal:
        return f"{label}: no lumen/wattage signal parsed."
    parts: list[str] = []
    if performance.user_brightness_target:
        parts.append(f"brightness target {performance.user_brightness_target}")
    if performance.installer_load_target:
        parts.append(f"load target {performance.installer_load_target}")
    if performance.efficacy_values:
        best = max(performance.efficacy_values)
        parts.append(f"parsed efficacy at max output/load {best} lm/W")
    if performance.efficiency_note:
        parts.append(performance.efficiency_note)
    return f"{label}: " + "; ".join(parts) + "."
