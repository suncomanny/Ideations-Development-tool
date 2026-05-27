from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .paths import ProjectPaths


DEFAULT_SOURCE = (
    Path.home()
    / "OneDrive - Sunco Lighting"
    / "Documents"
    / "Claude Workbook"
    / "Manny Sunco"
    / "Resources"
    / "SKU Decoder- Manny Version.csv"
)

OUTPUT_FOLDER = Path("source_data") / "sku_decoder"
CLEAN_FILENAME = "sku_decoder_clean.csv"
MANIFEST_FILENAME = "sku_decoder_manifest.json"

CATEGORY_NORMALIZATION = {
    "customer": "customer",
    "product type": "product_type",
    "_decorative": "decorative",
    "_communication": "communication",
    "_special feature": "special_feature",
    "color": "color",
    "cct range": "cct_range",
    "pack size": "pack_size",
    "non-sunco skus": "non_sunco_skus",
}

ATTRIBUTE_BY_CATEGORY = {
    "customer": "customer_code",
    "product_type": "product_type",
    "decorative": "style_or_form_factor",
    "communication": "control_or_sensor",
    "special_feature": "special_feature",
    "color": "finish_color",
    "cct_range": "cct_range",
    "pack_size": "pack_size",
    "non_sunco_skus": "non_sunco_prefix",
}

# Conservative category mappings. These help classify known SKU prefixes without
# treating ambiguous codes such as DL/downlight as a single category.
PRODUCT_TYPE_CATEGORY_MAP = {
    "CL": "ceiling_fixtures",
    "CP": "canopy",
    "DM": "dimmers",
    "EX": "emergency",
    "FL": "flood_lights",
    "GD": "residential_landscape",
    "HB_UFO": "ufo",
    "LRF": "led_ready_fixtures",
    "MWP": "wall_packs",
    "OCL": "outdoor_ceiling",
    "ODS": "outdoor_security",
    "OSPL": "residential_landscape",
    "PD": "pendants",
    "PN12": "panels",
    "PN14": "panels",
    "PN22": "panels",
    "PN24": "panels",
    "PW": "residential_landscape",
    "PWL": "residential_landscape",
    "RP": "rough_ins",
    "SH": "linears",
    "SHW": "wraparounds",
    "SL": "residential_landscape",
    "SNS": "sensors",
    "SPL": "residential_landscape",
    "ST": "striplights",
    "STL##": "string_lights",
    "TL": "lamps",
    "UFO": "ufo",
    "VL": "vanity",
    "VT": "vaportights",
    "VTHB_": "vaportights",
    "WL": "well_lights",
    "WP": "wall_packs",
    "WR": "wraparounds",
    "WR2": "wraparounds",
    "WS": "wall_sconces",
}

DECORATIVE_CATEGORY_MAP = {
    "#ST": "striplights",
    "#VT": "vaportights",
    "#WR": "wraparounds",
    "BF4": "retros",
    "BF6": "retros",
    "DF": "emergency",
    "EL": "emergency",
    "FBF3": "slims",
    "FBF4": "slims",
    "FBF6": "slims",
    "FH": "ceiling_fixtures",
    "G4": "retros",
    "G56": "retros",
    "MSL": "slims",
    "PD": "pendants",
    "RE4": "retros",
    "RE6": "retros",
    "RT1": "ceiling_fixtures",
    "RT2": "ceiling_fixtures",
    "RTR22": "led_ready_fixtures",
    "RTR24": "led_ready_fixtures",
    "SF": "emergency",
    "SL4": "slims",
    "SL6": "slims",
    "SL7": "slims",
    "SSL": "slims",
    "T2": "area_lights",
    "T3": "area_lights",
    "TR22": "led_ready_fixtures",
    "TR24": "led_ready_fixtures",
}

OUTPUT_HEADERS = [
    "code_category",
    "normalized_code_category",
    "code",
    "match_prefix",
    "code_meaning",
    "mapped_category_slug",
    "mapped_attribute",
    "line_review_match",
    "source_file",
]


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def normalize_code_category(value: str) -> str:
    clean = re.sub(r"\s+", " ", value.strip().lower())
    return CATEGORY_NORMALIZATION.get(clean, clean.replace(" ", "_").replace("-", "_"))


def match_prefix(code: str) -> str:
    clean = str(code or "").strip().upper()
    clean = re.sub(r"#.*$", "", clean)
    clean = clean.strip("_- ")
    return clean


def mapped_category(normalized_category: str, code: str) -> str:
    if normalized_category == "product_type":
        return PRODUCT_TYPE_CATEGORY_MAP.get(code.strip().upper(), "")
    if normalized_category == "decorative":
        return DECORATIVE_CATEGORY_MAP.get(code.strip().upper(), "")
    return ""


def iter_decoder_rows(source: Path) -> list[dict[str, Any]]:
    with source.open("r", encoding="utf-8-sig", newline="") as handle:
        raw_rows = list(csv.reader(handle))

    rows: list[dict[str, Any]] = []
    for raw in raw_rows:
        if len(raw) < 3:
            continue
        code_category = raw[0].strip()
        code = raw[1].strip()
        code_meaning = raw[2].strip()
        if not code_category or code_category == "# nimbalyst: {\"hasHeaders\":false":
            continue
        if code_category == "Code category" and code == "Code":
            continue
        if not code and not code_meaning:
            continue

        normalized_category = normalize_code_category(code_category)
        prefix = match_prefix(code)
        category_slug = mapped_category(normalized_category, code)
        line_review_match = (
            normalized_category == "product_type"
            and bool(category_slug)
            and len(prefix) >= 2
            and prefix not in {"DL", "TM", "TR", "KT", "PL", "PLG"}
        )
        rows.append(
            {
                "code_category": code_category,
                "normalized_code_category": normalized_category,
                "code": code,
                "match_prefix": prefix,
                "code_meaning": code_meaning,
                "mapped_category_slug": category_slug,
                "mapped_attribute": ATTRIBUTE_BY_CATEGORY.get(normalized_category, normalized_category),
                "line_review_match": "true" if line_review_match else "false",
                "source_file": str(source),
            }
        )
    return rows


def clean_sku_decoder(paths: ProjectPaths, source: Path = DEFAULT_SOURCE) -> dict[str, Any]:
    if not source.exists():
        raise FileNotFoundError(f"SKU decoder source file was not found: {source}")

    rows = iter_decoder_rows(source)
    output_folder = paths.backend / OUTPUT_FOLDER
    output_folder.mkdir(parents=True, exist_ok=True)
    clean_path = output_folder / CLEAN_FILENAME
    manifest_path = output_folder / MANIFEST_FILENAME

    with clean_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=OUTPUT_HEADERS)
        writer.writeheader()
        writer.writerows(rows)

    mapped_category_counts: dict[str, int] = {}
    for row in rows:
        slug = row["mapped_category_slug"]
        if slug:
            mapped_category_counts[slug] = mapped_category_counts.get(slug, 0) + 1

    manifest = {
        "generated_at": utc_now(),
        "source_file": str(source),
        "clean_csv": str(clean_path),
        "row_count": len(rows),
        "line_review_match_count": sum(1 for row in rows if row["line_review_match"] == "true"),
        "mapped_category_counts": dict(sorted(mapped_category_counts.items())),
        "notes": [
            "Cleaned from SKU Decoder source export.",
            "Only conservative Product Type mappings are used for line-review SKU-prefix matching.",
            "Original source remains outside git; cleaned backend files live under ignored backend/source_data.",
        ],
    }
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser(description="Clean the SKU Decoder CSV into backend source data.")
    parser.add_argument("--root", default=str(Path.cwd()), help="Ideation Development project root.")
    parser.add_argument("--source", default=str(DEFAULT_SOURCE), help="Path to SKU Decoder CSV.")
    args = parser.parse_args()
    manifest = clean_sku_decoder(ProjectPaths.from_root(args.root), Path(args.source).expanduser())
    print(json.dumps(manifest, indent=2))


if __name__ == "__main__":
    main()
