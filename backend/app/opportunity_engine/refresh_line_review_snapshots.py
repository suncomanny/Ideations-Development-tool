from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .categories import Category, load_categories
from .db_odbc import execute_odbc_sql, postgres_connection_source, postgres_connection_string, sanitize_connection_error
from .line_review import build_line_review_sql
from .paths import ProjectPaths


SNAPSHOT_FOLDER = Path("postgres_exports") / "line_reviews"


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def snapshot_path(paths: ProjectPaths, category: Category, generated_at: str) -> Path:
    stamp = generated_at.replace("-", "").replace(":", "").split("+", 1)[0].replace("T", "_")
    folder = paths.source_data / SNAPSHOT_FOLDER
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{category.slug}_line_review_{stamp}_postgres_odbc.json"


def write_snapshot(paths: ProjectPaths, category: Category, sql: str, rows: list[dict[str, Any]]) -> Path:
    generated_at = utc_now()
    payload = {
        "source_system": "postgres_odbc",
        "category_slug": category.slug,
        "category_run_name": category.run_name,
        "generated_at": generated_at,
        "row_count": len(rows),
        "source_reference": f"Postgres ODBC line-review refresh via {postgres_connection_source()}",
        "sql": sql,
        "rows": rows,
    }
    target = snapshot_path(paths, category, generated_at)
    target.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return target


def select_categories(paths: ProjectPaths, names: list[str] | None) -> list[Category]:
    categories = load_categories(paths)
    if not names:
        return categories
    wanted = {name.strip().lower().replace(" ", "_").replace("/", "_") for name in names if name.strip()}
    selected = [
        category
        for category in categories
        if category.slug in wanted or category.run_name.lower() in wanted or category.name.lower() in wanted
    ]
    missing = sorted(wanted - {category.slug for category in selected})
    if missing:
        print(f"Warning: no active category matched: {', '.join(missing)}")
    return selected


def _run_postgres_query(sql: str, timeout_seconds: int) -> list[dict[str, Any]]:
    connection = postgres_connection_string()
    return execute_odbc_sql(connection, sql, timeout_seconds=timeout_seconds)


def refresh_line_review_snapshots(
    paths: ProjectPaths,
    categories: list[Category],
    timeout_seconds: int = 180,
    dry_run: bool = False,
) -> dict[str, Any]:
    paths.ensure()
    results: list[dict[str, Any]] = []
    connection_source = postgres_connection_source()
    for index, category in enumerate(categories, start=1):
        sql = build_line_review_sql(category, paths)
        print(f"[{index}/{len(categories)}] Refreshing {category.run_name} through {connection_source}...")
        if dry_run:
            path = paths.cache / "ideation_data" / category.slug / "sql" / "line_review_postgres.sql"
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(sql, encoding="utf-8")
            results.append({"category": category.slug, "row_count": None, "snapshot": None, "status": "dry_run"})
            continue
        try:
            rows = _run_postgres_query(sql, timeout_seconds)
            target = write_snapshot(paths, category, sql, rows)
            print(f"  wrote {len(rows)} row(s): {target}")
            results.append({"category": category.slug, "row_count": len(rows), "snapshot": str(target), "status": "ok"})
        except Exception as exc:
            print(f"  full query failed: {sanitize_connection_error(exc)}")
            fallback_sql = build_line_review_sql(category, paths, include_purchase_order_facts=False)
            try:
                print("  retrying with fallback query without PO facts...")
                rows = _run_postgres_query(fallback_sql, timeout_seconds)
                target = write_snapshot(paths, category, fallback_sql, rows)
                print(f"  wrote {len(rows)} fallback row(s): {target}")
                results.append(
                    {
                        "category": category.slug,
                        "row_count": len(rows),
                        "snapshot": str(target),
                        "status": "ok_fallback_without_po_facts",
                        "fallback_reason": sanitize_connection_error(exc),
                    }
                )
            except Exception as fallback_exc:
                print(f"  failed: {sanitize_connection_error(fallback_exc)}")
                results.append(
                    {
                        "category": category.slug,
                        "row_count": 0,
                        "snapshot": None,
                        "status": "failed",
                        "error": sanitize_connection_error(fallback_exc),
                        "first_error": sanitize_connection_error(exc),
                    }
                )
    return {
        "generated_at": utc_now(),
        "connection_source": connection_source,
        "categories": len(categories),
        "results": results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Refresh approved Postgres line-review snapshots for Step 1.")
    parser.add_argument("--root", default=str(Path.cwd()), help="Ideation Development project root.")
    parser.add_argument("--category", action="append", help="Optional category slug/name. Repeat to refresh multiple categories.")
    parser.add_argument("--timeout-seconds", type=int, default=180, help="Per-category ODBC query timeout.")
    parser.add_argument("--dry-run", action="store_true", help="Write SQL files only; do not query Postgres.")
    args = parser.parse_args()

    paths = ProjectPaths.from_root(args.root)
    categories = select_categories(paths, args.category)
    result = refresh_line_review_snapshots(paths, categories, args.timeout_seconds, args.dry_run)
    manifest = paths.source_data / SNAPSHOT_FOLDER / "line_review_refresh_manifest.json"
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(f"\nRefresh manifest:\n{manifest}")


if __name__ == "__main__":
    main()
