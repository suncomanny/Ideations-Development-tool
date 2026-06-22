from __future__ import annotations

import argparse
import csv
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from .paths import ProjectPaths
from .powerbi_classification import classification_cache_path, designation_mapping_path
from .utils import slugify


FIELDNAMES = ["owner", "category", "run_name", "active", "powerbi_categories", "notes"]
OWNER_SHORT_TO_FULL = {
    "Manny": "Manny Hernandez",
    "Jesse": "Jesse Harper",
    "Stephanie": "Stephanie",
}
RUN_NAME_OVERRIDES = {
    "ufo": "UFO",
    "led_ready_fixtures": "LED Ready Fixtures",
}


def _text(value: Any) -> str:
    return str(value or "").strip()


def _norm_sku(value: Any) -> str:
    return re.sub(r"\s+", "", _text(value).upper())


def _split_list(value: Any) -> list[str]:
    output: list[str] = []
    for item in re.split(r"[;,\n]+", _text(value)):
        clean = item.strip()
        if clean and clean not in output:
            output.append(clean)
    return output


def _prefix_matches(value: Any, prefixes: list[str]) -> bool:
    if not prefixes:
        return True
    normalized = _norm_sku(value)
    return any(normalized.startswith(_norm_sku(prefix)) for prefix in prefixes if _norm_sku(prefix))


def _run_name_from_slug(slug: str) -> str:
    return RUN_NAME_OVERRIDES.get(slug, slug.replace("_", " ").title())


def _read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return [dict(row) for row in csv.DictReader(handle)]


def _load_rules(paths: ProjectPaths) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for row in _read_csv(designation_mapping_path(paths)):
        if _text(row.get("active")).lower() not in {"yes", "true", "1"}:
            continue
        item = dict(row)
        item["tool_slug"] = slugify(row.get("tool_category_slug") or row.get("tool_category") or "")
        item["family_prefixes_list"] = _split_list(row.get("family_prefixes"))
        item["sku_prefixes_list"] = _split_list(row.get("sku_prefixes"))
        rows.append(item)
    return rows


def _load_classification_rows(cache_path: Path) -> list[dict[str, Any]]:
    if not cache_path.exists():
        raise FileNotFoundError(
            f"Missing PowerBI Families classification cache: {cache_path}. "
            "Run Refresh Product Demand SKU Classification Cache.py first."
        )
    connection = sqlite3.connect(cache_path)
    connection.row_factory = sqlite3.Row
    try:
        return [dict(row) for row in connection.execute("select * from sku_classification").fetchall()]
    finally:
        connection.close()


def _category_rows_by_slug(existing_rows: list[dict[str, str]], rules: list[dict[str, Any]]) -> dict[str, dict[str, str]]:
    by_slug = {
        slugify(row.get("run_name") or row.get("category")): dict(row)
        for row in existing_rows
        if _text(row.get("run_name") or row.get("category"))
    }
    rules_by_slug: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for rule in rules:
        if rule["tool_slug"]:
            rules_by_slug[rule["tool_slug"]].append(rule)
    for slug, group in rules_by_slug.items():
        if slug in by_slug:
            continue
        run_name = _text(group[0].get("powerbi_high_level_category")) or _run_name_from_slug(slug)
        by_slug[slug] = {
            "owner": "",
            "category": run_name,
            "run_name": run_name,
            "active": "yes",
            "powerbi_categories": "",
            "notes": "Added from current PowerBI Families category designation map.",
        }
    return by_slug


def _owner_and_note(
    slug: str,
    row: dict[str, str],
    rules: list[dict[str, Any]],
    classification_rows: list[dict[str, Any]],
) -> tuple[str, str, str]:
    relevant = [rule for rule in rules if rule["tool_slug"] == slug]
    source_categories: list[str] = []
    owner_counts: Counter[str] = Counter()
    matched_count = 0

    for rule in relevant:
        source_category = _text(rule.get("powerbi_high_level_category"))
        if source_category and source_category not in source_categories:
            source_categories.append(source_category)
        source_slug = slugify(source_category)
        for classification in classification_rows:
            if slugify(classification.get("category")) != source_slug:
                continue
            if not _prefix_matches(classification.get("family"), rule["family_prefixes_list"]):
                continue
            if not _prefix_matches(classification.get("sku"), rule["sku_prefixes_list"]):
                continue
            matched_count += 1
            owner = _text(classification.get("pm_responsible"))
            if owner:
                owner_counts[owner] += 1

    old_owner = _text(row.get("owner"))
    owner = owner_counts.most_common(1)[0][0] if owner_counts else OWNER_SHORT_TO_FULL.get(old_owner, old_owner)
    categories_text = "; ".join(source_categories)
    if source_categories and owner_counts:
        owner_text = ", ".join(f"{owner_name}={count}" for owner_name, count in owner_counts.most_common())
        note = (
            f"PowerBI Families mapped categories: {', '.join(source_categories)}. "
            f"Owner from nonblank PM Responsible rows: {owner_text}. Matched rows={matched_count}."
        )
    elif source_categories:
        note = (
            f"PowerBI Families mapped categories: {', '.join(source_categories)}. "
            "No nonblank PM Responsible rows in current PowerBI data; retained prior owner pending Data Team assignment. "
            f"Matched rows={matched_count}."
        )
    else:
        note = "No current PowerBI Families mapping configured; retained active category pending Data Team category assignment."
    return owner, categories_text, note


def refresh_category_reference(paths: ProjectPaths) -> Path:
    category_path = paths.templates / "category_reference.csv"
    existing_rows = _read_csv(category_path)
    rules = _load_rules(paths)
    classification_rows = _load_classification_rows(classification_cache_path(paths))
    by_slug = _category_rows_by_slug(existing_rows, rules)

    ordered_slugs = [
        slugify(row.get("run_name") or row.get("category"))
        for row in existing_rows
        if _text(row.get("run_name") or row.get("category"))
    ]
    ordered_slugs.extend(slug for slug in sorted(by_slug) if slug not in ordered_slugs)

    output: list[dict[str, str]] = []
    for slug in ordered_slugs:
        row = by_slug[slug]
        owner, powerbi_categories, note = _owner_and_note(slug, row, rules, classification_rows)
        run_name = _text(row.get("run_name") or row.get("category"))
        output.append(
            {
                "owner": owner,
                "category": _text(row.get("category")) or run_name,
                "run_name": run_name,
                "active": _text(row.get("active")) or "yes",
                "powerbi_categories": powerbi_categories,
                "notes": note,
            }
        )

    category_path.parent.mkdir(parents=True, exist_ok=True)
    with category_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(output)
    return category_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Refresh templates/category_reference.csv from PowerBI Families cache.")
    parser.add_argument("--root", default=str(Path.cwd()), help="Ideation Development project root.")
    args = parser.parse_args()
    output = refresh_category_reference(ProjectPaths.from_root(args.root))
    print(f"PowerBI-backed category reference refreshed: {output}")


if __name__ == "__main__":
    main()
