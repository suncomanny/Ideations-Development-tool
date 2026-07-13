from __future__ import annotations

import re


STACKLINE_SEGMENT_OVERRIDES: dict[str, tuple[str, ...]] = {
    "emergency": ("Emergency Signs", "Emergency Lights - Core"),
    "panels": ("Ceiling Panel Lights", "Flat Panel Ceiling Lights", "Flat Panels", "2x4 PANEL", "1x4 Flat Panel"),
    "wraparounds": ("Wraparound LED - Broad V2", "Wraparound LED"),
    "ufo": ("UFO High Bays", "UFO High Bay Overview", "UFO Lighting", "UFO"),
    "linears": (
        "Linear High Bays",
        "Linear High Bay Overview",
        "4ft Linear High Bay",
        "Linear High Bay Selectable",
        "Top 25 4ft Linear High Bay",
        "4ft Linear High Bay Dual Selectable (200W/260W/320W)",
        "Plug and Play Linear High Bay",
    ),
    "shop_light": ("SHOP LIGHT",),
    "striplights": ("Striplights",),
    "wall_packs": ("Wall Packs", "wall Packs", "White Wall Packs"),
    "commercial_grow_lights": (
        "2FT Backless Grow Light Fixture",
        "2ft Backless Grow Fixtures",
        "2x2 Bar Grow Light Fixture",
        "2x2 Flat Panel Grow Light Fixture",
        "2x4 Bar Grow Light Fixture",
        "2x4 Flat Panel Grow Light Fixture",
        "3x3 Bar Grow Light Fixture",
        "3x3 Flat Panel Grow Light Fixture",
        "4ft 45W V-Shape Reflector T8 Grow Fixture",
        "4ft Grow Light with Timer & Selectable Spectrum",
        "4x4 Bar Grow Light Fixture",
        "4x4 Flat Panel Grow Light Fixture",
        "5x5 Bar Grow Light Fixture",
        "Bar Grow Light Fixture",
        "Flat Panel Grow Light Fixture",
        "Full Spectrum Strip T8 Grow Light",
        "Full Spectrum V-Shaped Strip Grow Light",
        "KA - Grow Lights",
        "SHOP LIGHT - GROW LIGHTS",
    ),
    "residential_grow_lights": (
        "4ft Grow Light with Timer & Selectable Spectrum",
        "Full Spectrum Strip T8 Grow Light",
        "Full Spectrum V-Shaped Strip Grow Light",
        "Grow Light [Archived]",
        "KA - Grow Lights",
        "SHOP LIGHT - GROW LIGHTS",
    ),
    "slims": ("RECESSED - Slims",),
    "recessed": ("CC Broad - Recessed Lighting", "RECESSED - ALL", "RECESSED - Retrofits", "RECESSED - Slims"),
    "under_cabinet": ("Under Cabinet Light Fixture", "Under Cabinet Light Fixture - 24 inch"),
    "ceiling_fixtures": ("Ceiling Fixtures",),
    "chandeliers": ("Chandeliers",),
    "vanity": ("Vanity",),
    "bulbs_plus_tubes": ("A19 - Non-Smart / RGB", "DATA-A19-9W", "T8", "DATA-T8", "T8 Bulbs"),
}


def slugish(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


def stackline_segments_for_category(
    category_slug: str,
    category_name: str,
    aliases: tuple[str, ...] | list[str] = (),
) -> tuple[str, ...]:
    override = STACKLINE_SEGMENT_OVERRIDES.get(category_slug)
    if override:
        return override

    raw = " ".join([category_slug, category_name, *aliases])
    terms = [term for term in slugish(raw).split() if len(term) >= 4]
    return tuple(dict.fromkeys(terms))


def sql_values(values: tuple[str, ...] | list[str]) -> str:
    return ", ".join("'" + str(value).replace("'", "''") + "'" for value in values)
