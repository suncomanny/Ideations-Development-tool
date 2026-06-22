from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .categories import Category, load_categories
from .db_mcp import McpRemoteClient, sanitize_mcp_error
from .line_review import build_line_review_sql
from .paths import ProjectPaths


SNAPSHOT_FOLDER = Path("postgres_exports") / "line_reviews"


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def snapshot_path(paths: ProjectPaths, category: Category, generated_at: str) -> Path:
    stamp = generated_at.replace("-", "").replace(":", "").split("+", 1)[0].replace("T", "_")
    folder = paths.source_data / SNAPSHOT_FOLDER
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{category.slug}_line_review_{stamp}_postgres_mcp.json"


def write_snapshot(paths: ProjectPaths, category: Category, sql: str, rows: list[dict[str, Any]]) -> Path:
    generated_at = utc_now()
    payload = {
        "source_system": "postgres_mcp",
        "category_slug": category.slug,
        "category_run_name": category.run_name,
        "generated_at": generated_at,
        "row_count": len(rows),
        "source_reference": "Postgres MCP line-review refresh",
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


def refresh_line_review_snapshots(
    paths: ProjectPaths,
    categories: list[Category],
    timeout_seconds: int = 180,
    dry_run: bool = False,
) -> dict[str, Any]:
    paths.ensure()
    results: list[dict[str, Any]] = []
    connection_source = "Postgres MCP"
    client: McpRemoteClient | None = None
    try:
        if not dry_run:
            client = McpRemoteClient(timeout_seconds=timeout_seconds, client_name="sunco-line-review-refresh")
            client.__enter__()
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
                assert client is not None
                rows = client.execute_sql(sql, timeout_seconds=timeout_seconds)
                target = write_snapshot(paths, category, sql, rows)
                print(f"  wrote {len(rows)} row(s): {target}")
                results.append({"category": category.slug, "row_count": len(rows), "snapshot": str(target), "status": "ok"})
            except Exception as exc:
                print(f"  full query failed: {sanitize_mcp_error(exc)}")
                fallback_sql = build_line_review_sql(category, paths, include_purchase_order_facts=False)
                try:
                    assert client is not None
                    print("  retrying with fallback query without PO facts...")
                    rows = client.execute_sql(fallback_sql, timeout_seconds=timeout_seconds)
                    target = write_snapshot(paths, category, fallback_sql, rows)
                    print(f"  wrote {len(rows)} fallback row(s): {target}")
                    results.append(
                        {
                            "category": category.slug,
                            "row_count": len(rows),
                            "snapshot": str(target),
                            "status": "ok_fallback_without_po_facts",
                            "fallback_reason": sanitize_mcp_error(exc),
                        }
                    )
                except Exception as fallback_exc:
                    print(f"  failed: {sanitize_mcp_error(fallback_exc)}")
                    results.append(
                        {
                            "category": category.slug,
                            "row_count": 0,
                            "snapshot": None,
                            "status": "failed",
                            "error": sanitize_mcp_error(fallback_exc),
                            "first_error": sanitize_mcp_error(exc),
                        }
                    )
    finally:
        if client is not None:
            client.__exit__(None, None, None)
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
    parser.add_argument("--timeout-seconds", type=int, default=180, help="Per-category MCP query timeout.")
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
