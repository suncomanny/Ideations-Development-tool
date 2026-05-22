from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from .paths import ProjectPaths
from .utils import prompt_choice, slugify


@dataclass(frozen=True)
class Category:
    owner: str
    name: str
    run_name: str
    notes: str = ""

    @property
    def slug(self) -> str:
        return slugify(self.run_name or self.name)

    @property
    def display(self) -> str:
        return f"{self.run_name} ({self.owner})"


def load_categories(paths: ProjectPaths) -> list[Category]:
    source = paths.templates / "category_reference.csv"
    if not source.exists():
        raise FileNotFoundError(f"Missing category reference: {source}")
    with source.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = csv.DictReader(handle)
        categories = [
            Category(
                owner=(row.get("owner") or "").strip(),
                name=(row.get("category") or "").strip(),
                run_name=(row.get("run_name") or row.get("category") or "").strip(),
                notes=(row.get("notes") or "").strip(),
            )
            for row in rows
            if (row.get("active") or "").strip().lower() in {"yes", "true", "1"}
        ]
    return sorted(categories, key=lambda item: (item.owner.lower(), item.run_name.lower()))


def choose_category(paths: ProjectPaths) -> Category:
    categories = load_categories(paths)
    return prompt_choice(categories, lambda item: item.display, "Which category should run?")
