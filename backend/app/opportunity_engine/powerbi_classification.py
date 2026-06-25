from __future__ import annotations

import csv
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .categories import Category
from .paths import ProjectPaths
from .utils import slugify


MAPPING_FILENAME = "powerbi_category_designation_map.csv"


@dataclass(frozen=True)
class PowerBiClassificationContext:
    cache_path: Path | None
    mapping_path: Path | None
    families: list[str]
    skus: list[str]
    source_categories: list[str]
    rule_count: int
    row_count: int
    notes: list[str]

    @property
    def has_matches(self) -> bool:
        return bool(self.families or self.skus)

    def summary(self) -> str:
        if self.has_matches:
            categories = ", ".join(self.source_categories) if self.source_categories else "mapped categories"
            return (
                f"{len(self.families)} family match(es), {len(self.skus)} SKU match(es) "
                f"from PowerBI Families category designation ({categories})."
            )
        if self.notes:
            return " ".join(self.notes)
        return "No PowerBI Families category-designation matches available."


def classification_cache_path(paths: ProjectPaths) -> Path:
    return paths.source_data / "sharepoint_exports" / "sku_classification" / "sku_classification.sqlite"


def classification_workbook_path(paths: ProjectPaths) -> Path:
    return paths.source_data / "sharepoint_exports" / "sku_classification" / "SkuClassification_PowerBIFamilies_latest.xlsx"


def designation_mapping_path(paths: ProjectPaths) -> Path:
    return paths.templates / MAPPING_FILENAME


def _norm(value: Any) -> str:
    return str(value or "").strip()


def _norm_sku(value: Any) -> str:
    return re.sub(r"\s+", "", _norm(value).upper())


def _split_list(value: Any) -> list[str]:
    output: list[str] = []
    for item in re.split(r"[;,\n]+", _norm(value)):
        clean = item.strip()
        if clean and clean not in output:
            output.append(clean)
    return output


def _prefix_matches(value: str, prefixes: list[str]) -> bool:
    if not prefixes:
        return True
    normalized = _norm_sku(value)
    for prefix in prefixes:
        candidate = _norm_sku(prefix)
        if candidate and normalized.startswith(candidate):
            return True
    return False


def _load_rules(paths: ProjectPaths, category: Category) -> tuple[Path | None, list[dict[str, Any]], list[str]]:
    path = designation_mapping_path(paths)
    if not path.exists():
        return None, [], [f"Missing PowerBI category designation map: {path}."]

    rows: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            active = (row.get("active") or "yes").strip().lower()
            if active not in {"yes", "true", "1"}:
                continue
            rule_slug = slugify(row.get("tool_category_slug") or row.get("tool_category") or "")
            if rule_slug != category.slug:
                continue
            source_category = _norm(row.get("powerbi_high_level_category"))
            if not source_category:
                continue
            rows.append(
                {
                    "source_category": source_category,
                    "source_category_slug": slugify(source_category),
                    "family_prefixes": _split_list(row.get("family_prefixes")),
                    "sku_prefixes": _split_list(row.get("sku_prefixes")),
                    "notes": _norm(row.get("notes")),
                }
            )
    if not rows:
        return path, [], [f"No PowerBI category designation rules configured for {category.slug}."]
    return path, rows, []


def _load_classification_rows(cache_path: Path) -> list[dict[str, Any]]:
    connection = sqlite3.connect(cache_path)
    connection.row_factory = sqlite3.Row
    try:
        return [dict(row) for row in connection.execute("select * from sku_classification").fetchall()]
    finally:
        connection.close()


def load_powerbi_classification_matches(paths: ProjectPaths, category: Category) -> PowerBiClassificationContext:
    cache_path = classification_cache_path(paths)
    mapping_path, rules, notes = _load_rules(paths, category)
    if not cache_path.exists():
        notes.append(f"Missing PowerBI Families classification cache: {cache_path}.")
        return PowerBiClassificationContext(cache_path, mapping_path, [], [], [], len(rules), 0, notes)
    if not rules:
        return PowerBiClassificationContext(cache_path, mapping_path, [], [], [], 0, 0, notes)

    rows = _load_classification_rows(cache_path)
    families: list[str] = []
    skus: list[str] = []
    source_categories: list[str] = []
    matched_rows = 0

    for row in rows:
        row_category = _norm(row.get("category"))
        row_category_slug = slugify(row_category)
        sku = _norm_sku(row.get("sku"))
        parent_sku = _norm_sku(row.get("parent_sku"))
        family = _norm_sku(row.get("family") or row.get("parent_sku") or sku)
        for rule in rules:
            if row_category_slug != rule["source_category_slug"]:
                continue
            family_ok = _prefix_matches(family, rule["family_prefixes"])
            sku_ok = _prefix_matches(sku, rule["sku_prefixes"])
            if not (family_ok and sku_ok):
                continue
            matched_rows += 1
            if family and family not in families:
                families.append(family)
            if sku and sku not in skus:
                skus.append(sku)
            if parent_sku and parent_sku not in skus:
                skus.append(parent_sku)
            if row_category and row_category not in source_categories:
                source_categories.append(row_category)
            break

    return PowerBiClassificationContext(
        cache_path=cache_path,
        mapping_path=mapping_path,
        families=sorted(families),
        skus=sorted(skus),
        source_categories=sorted(source_categories),
        rule_count=len(rules),
        row_count=matched_rows,
        notes=notes,
    )
