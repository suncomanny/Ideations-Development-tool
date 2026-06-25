from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from .paths import ProjectPaths
from .utils import prompt_choice, slugify


POWERBI_CATEGORY_MAP_FILENAME = "powerbi_category_designation_map.csv"


@dataclass(frozen=True)
class Category:
    owner: str
    name: str
    run_name: str
    notes: str = ""
    powerbi_aliases: tuple[str, ...] = ()

    @property
    def slug(self) -> str:
        return slugify(self.run_name or self.name)

    @property
    def display(self) -> str:
        label = f"{self.run_name} ({self.owner})"
        if not self.powerbi_aliases:
            return label
        aliases = list(self.powerbi_aliases)
        shown = ", ".join(aliases[:4])
        if len(aliases) > 4:
            shown += f", +{len(aliases) - 4} more"
        return f"{label} | PowerBI: {shown}"


def _powerbi_alias_label(row: dict[str, str]) -> str:
    category = (row.get("powerbi_high_level_category") or "").strip()
    prefix = (row.get("family_prefixes") or row.get("sku_prefixes") or "").strip()
    return f"{category} [{prefix}]" if prefix else category


def _load_powerbi_aliases(paths: ProjectPaths) -> dict[str, tuple[str, ...]]:
    source = paths.templates / POWERBI_CATEGORY_MAP_FILENAME
    if not source.exists():
        return {}
    aliases: dict[str, list[str]] = {}
    with source.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            if (row.get("active") or "").strip().lower() not in {"yes", "true", "1"}:
                continue
            slug = slugify(row.get("tool_category_slug") or row.get("tool_category") or "")
            label = _powerbi_alias_label(row)
            if slug and label and label not in aliases.setdefault(slug, []):
                aliases[slug].append(label)
    return {slug: tuple(values) for slug, values in aliases.items()}


def load_categories(paths: ProjectPaths) -> list[Category]:
    source = paths.templates / "category_reference.csv"
    if not source.exists():
        raise FileNotFoundError(f"Missing category reference: {source}")
    powerbi_aliases = _load_powerbi_aliases(paths)
    with source.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = csv.DictReader(handle)
        categories = [
            Category(
                owner=(row.get("owner") or "").strip(),
                name=(row.get("category") or "").strip(),
                run_name=(row.get("run_name") or row.get("category") or "").strip(),
                notes=(row.get("notes") or "").strip(),
                powerbi_aliases=powerbi_aliases.get(slugify(row.get("run_name") or row.get("category") or ""), ()),
            )
            for row in rows
            if (row.get("active") or "").strip().lower() in {"yes", "true", "1"}
        ]
    return sorted(categories, key=lambda item: (item.owner.lower(), item.run_name.lower()))


def choose_category(paths: ProjectPaths) -> Category:
    categories = load_categories(paths)
    return prompt_choice(categories, lambda item: item.display, "Which category should run?")
