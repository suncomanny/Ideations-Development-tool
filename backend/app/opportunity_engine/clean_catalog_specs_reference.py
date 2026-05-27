from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .paths import ProjectPaths


DEFAULT_SOURCE = Path(
    r"C:\Users\Sunco\OneDrive - Sunco Lighting\Documents\Claude Workbook\Manny Sunco\Resources\SUNCO ALL SPECS REFERENCE.csv"
)

PRODUCT_HEADERS = [
    "catalog_key",
    "sku",
    "source_handle",
    "title",
    "product_type",
    "category_mapping",
    "brand",
    "model_number",
    "warranty",
    "short_description",
    "global_title_tag",
    "global_description_tag",
    "active_status",
    "raw_info",
]

ATTRIBUTE_HEADERS = [
    "catalog_key",
    "sku",
    "title",
    "product_type",
    "category_mapping",
    "attribute_name",
    "attribute_value",
    "source_column",
]


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_sku(value: Any) -> str:
    text = clean_text(value)
    if not text or text.lower() == "not found":
        return ""
    return text


def parse_title_and_type(raw_info: str) -> tuple[str, str]:
    text = clean_text(raw_info)
    match = re.search(r"\(([^()]*)\)\s*$", text)
    if not match:
        return text, ""
    product_type = clean_text(match.group(1))
    title = re.sub(r"\s*\([^()]*\)\s*$", "", text).strip()
    return title or text, product_type


def normalized_attribute_name(header: str) -> str:
    name = header.removeprefix("specs.")
    name = name.replace(".", "_").replace("-", "_")
    name = re.sub(r"[^a-zA-Z0-9_]+", "_", name)
    return name.strip("_").lower()


def is_useful_attribute(header: str, value: str) -> bool:
    if not header.startswith("specs."):
        return False
    if not value:
        return False
    # The source SKU/model fields are useful metadata, but not SKU-defining spec signals.
    if header in {"specs.sku"}:
        return False
    return True


def clean_catalog_specs_reference(
    root: Path,
    source: Path,
    output_folder: Path | None = None,
) -> dict[str, Any]:
    if not source.exists():
        raise FileNotFoundError(f"Missing source catalog spec CSV: {source}")

    paths = ProjectPaths.from_root(root)
    destination = output_folder or (paths.source_data / "catalog_specs")
    destination.mkdir(parents=True, exist_ok=True)
    product_path = destination / "sunco_catalog_products_clean.csv"
    attributes_path = destination / "sunco_catalog_spec_attributes_long.csv"
    manifest_path = destination / "sunco_catalog_specs_manifest.json"

    product_count = 0
    attribute_count = 0
    skipped_attribute_cells = 0
    source_headers: list[str] = []

    with source.open("r", encoding="utf-8-sig", newline="") as source_handle:
        reader = csv.DictReader(source_handle)
        source_headers = list(reader.fieldnames or [])
        with product_path.open("w", encoding="utf-8", newline="") as product_handle, attributes_path.open(
            "w", encoding="utf-8", newline=""
        ) as attribute_handle:
            product_writer = csv.DictWriter(product_handle, fieldnames=PRODUCT_HEADERS)
            attribute_writer = csv.DictWriter(attribute_handle, fieldnames=ATTRIBUTE_HEADERS)
            product_writer.writeheader()
            attribute_writer.writeheader()

            for row in reader:
                sku = normalize_sku(row.get("SKU"))
                source_handle_value = clean_text(row.get("_id"))
                raw_info = clean_text(row.get("_info"))
                title, product_type = parse_title_and_type(raw_info)
                catalog_key = sku or source_handle_value or title
                if not catalog_key or not title:
                    continue

                product_writer.writerow(
                    {
                        "catalog_key": catalog_key,
                        "sku": sku,
                        "source_handle": source_handle_value,
                        "title": title,
                        "product_type": product_type,
                        "category_mapping": product_type,
                        "brand": clean_text(row.get("specs.brand")),
                        "model_number": clean_text(row.get("specs.model_number")),
                        "warranty": clean_text(row.get("specs.warranty")),
                        "short_description": clean_text(row.get("info.short_description")),
                        "global_title_tag": clean_text(row.get("global.title_tag")),
                        "global_description_tag": clean_text(row.get("global.description_tag")),
                        "active_status": "catalog_reference",
                        "raw_info": raw_info,
                    }
                )
                product_count += 1

                for header in source_headers:
                    value = clean_text(row.get(header))
                    if not is_useful_attribute(header, value):
                        if header.startswith("specs.") and not value:
                            skipped_attribute_cells += 1
                        continue
                    attribute_writer.writerow(
                        {
                            "catalog_key": catalog_key,
                            "sku": sku,
                            "title": title,
                            "product_type": product_type,
                            "category_mapping": product_type,
                            "attribute_name": normalized_attribute_name(header),
                            "attribute_value": value,
                            "source_column": header,
                        }
                    )
                    attribute_count += 1

    manifest = {
        "generated_at": utc_now(),
        "source_csv": str(source),
        "product_output": str(product_path),
        "attribute_output": str(attributes_path),
        "source_header_count": len(source_headers),
        "product_rows": product_count,
        "attribute_rows": attribute_count,
        "skipped_empty_spec_cells": skipped_attribute_cells,
        "format": {
            "products": PRODUCT_HEADERS,
            "attributes": ATTRIBUTE_HEADERS,
        },
    }
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser(description="Create clean backend catalog/spec reference CSVs.")
    parser.add_argument("--root", default=str(Path.cwd()), help="Ideation Development project root.")
    parser.add_argument("--source", default=str(DEFAULT_SOURCE), help="Source SUNCO ALL SPECS REFERENCE.csv path.")
    parser.add_argument("--output-folder", default=None, help="Optional output folder for clean backend CSVs.")
    args = parser.parse_args()
    result = clean_catalog_specs_reference(
        Path(args.root),
        Path(args.source),
        Path(args.output_folder) if args.output_folder else None,
    )
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
