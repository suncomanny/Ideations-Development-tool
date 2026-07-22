from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from .categories import Category, load_categories
from .build_category_intelligence import build_category_intelligence_database
from .paths import ProjectPaths
from .refresh_line_review_snapshots import refresh_line_review_snapshots
from .refresh_redshift_stackline_cache import refresh_redshift_stackline_caches


def normalize_category_name(value: str) -> str:
    return value.strip().lower().replace(" ", "_").replace("/", "_")


def select_categories(paths: ProjectPaths, names: list[str] | None) -> list[Category]:
    categories = load_categories(paths)
    if not names:
        return categories
    wanted = {normalize_category_name(name) for name in names if name.strip()}
    selected = [
        category
        for category in categories
        if category.slug in wanted
        or normalize_category_name(category.run_name) in wanted
        or normalize_category_name(category.name) in wanted
    ]
    missing = sorted(wanted - {category.slug for category in selected})
    if missing:
        print(f"Warning: no active category matched: {', '.join(missing)}")
    return selected


def write_manifest(paths: ProjectPaths, payload: dict[str, Any]) -> Path:
    target = paths.source_data / "backend_refresh_manifest.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return target


def main() -> None:
    parser = argparse.ArgumentParser(description="Refresh backend source snapshots for the Ideation Development tool.")
    parser.add_argument("--root", default=str(Path.cwd()), help="Ideation Development project root.")
    parser.add_argument("--category", action="append", help="Optional category slug/name. Repeat to refresh multiple categories.")
    parser.add_argument("--timeout-seconds", type=int, default=180, help="Per-category Postgres MCP timeout for line review.")
    parser.add_argument("--stackline-timeout-seconds", type=int, default=240, help="Per-query Redshift timeout for Stackline.")
    parser.add_argument("--dry-run", action="store_true", help="Write SQL files only; do not query Postgres/Redshift.")
    parser.add_argument("--skip-line-review", action="store_true", help="Skip Postgres line-review snapshot refresh.")
    parser.add_argument("--skip-stackline", action="store_true", help="Skip Redshift Stackline cache refresh.")
    parser.add_argument("--force-stackline", action="store_true", help="Refresh Stackline even when a fresh cache exists.")
    args = parser.parse_args()

    paths = ProjectPaths.from_root(args.root)
    categories = select_categories(paths, args.category)
    paths.ensure()

    manifest: dict[str, Any] = {
        "categories": [category.slug for category in categories],
        "line_review": None,
        "stackline": None,
        "category_intelligence": None,
    }

    if not args.skip_line_review:
        print("\nRefreshing Postgres line-review snapshots...")
        manifest["line_review"] = refresh_line_review_snapshots(
            paths=paths,
            categories=categories,
            timeout_seconds=args.timeout_seconds,
            dry_run=args.dry_run,
        )

    if not args.skip_stackline:
        print("\nRefreshing Redshift Stackline caches...")
        manifest["stackline"] = refresh_redshift_stackline_caches(
            paths=paths,
            categories=categories,
            timeout_seconds=args.stackline_timeout_seconds,
            force=args.force_stackline,
            dry_run=args.dry_run,
        )

    print("\nRebuilding category intelligence database and generated category profiles...")
    manifest["category_intelligence"] = build_category_intelligence_database(paths)

    target = write_manifest(paths, manifest)
    print(f"\nBackend refresh manifest:\n{target}")


if __name__ == "__main__":
    main()
