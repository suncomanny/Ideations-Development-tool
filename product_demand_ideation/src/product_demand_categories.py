from __future__ import annotations

from typing import Any


PRODUCT_DEMAND_EXTRA_CATEGORIES = (
    {
        "owner": "Manny",
        "name": "Smart",
        "run_name": "Smart Lighting",
        "notes": "Category sourced from PowerBI Families high-level category.",
    },
)


def load_product_demand_categories(paths: Any) -> list[Any]:
    from opportunity_engine.categories import Category, load_categories

    categories = list(load_categories(paths))
    existing_slugs = {category.slug for category in categories}
    for item in PRODUCT_DEMAND_EXTRA_CATEGORIES:
        category = Category(**item)
        if category.slug not in existing_slugs:
            categories.append(category)
            existing_slugs.add(category.slug)
    return sorted(categories, key=lambda item: (item.owner.lower(), item.run_name.lower()))


def choose_product_demand_category(paths: Any) -> Any:
    from opportunity_engine.utils import prompt_choice

    return prompt_choice(load_product_demand_categories(paths), lambda item: item.display, "Which category should run?")
