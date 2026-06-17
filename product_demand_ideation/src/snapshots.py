from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from mcp_remote_client import McpRemoteClient


INVENTORY_SQL = """
with movement as (
  select
    url,
    brand,
    sku,
    name,
    max(scrape_date) as latest_scrape_date,
    max(stock_qty) as latest_stock_qty,
    max(price) as latest_price,
    sum(case when stock_qty_delta < 0 then abs(stock_qty_delta) else 0 end) as observed_stock_decrease,
    sum(case when stock_qty_delta > 0 then stock_qty_delta else 0 end) as observed_restock,
    count(case when stock_qty_delta < 0 then 1 end) as decrease_events,
    count(case when stock_qty_delta > 0 then 1 end) as restock_events,
    count(*) as observation_count
  from public.v_competitors_inventory_daily
  where
    (
      lower(coalesce(name, '')) like '%panel%'
      or lower(coalesce(name, '')) like '%troffer%'
    )
    and lower(coalesce(name, '')) not like '%sensor%'
    and lower(coalesce(name, '')) not like '%photocell%'
    and lower(coalesce(name, '')) not like '%mounting kit%'
    and lower(coalesce(name, '')) not like '%driver%'
    and lower(coalesce(name, '')) not like '%frame kit%'
  group by url, brand, sku, name
)
select *
from movement
where observed_stock_decrease > 0
order by observed_stock_decrease desc, decrease_events desc
limit 75;
""".strip()


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def latest_snapshot(exports_dir: Path) -> Path | None:
    snapshots = sorted(exports_dir.glob("ceiling_panels_inventory_movement_*.json"), key=lambda path: path.stat().st_mtime)
    return snapshots[-1] if snapshots else None


def write_snapshot(exports_dir: Path, source_system: str, sql: str, rows: list[dict[str, Any]]) -> Path:
    exports_dir.mkdir(parents=True, exist_ok=True)
    generated_at = utc_now()
    stamp = generated_at.replace("-", "").replace(":", "").split("+", 1)[0].replace("T", "_")
    target = exports_dir / f"ceiling_panels_inventory_movement_{stamp}.json"
    payload = {
        "source_system": source_system,
        "category_slug": "ceiling_panels",
        "category_name": "Ceiling Panels",
        "generated_at": generated_at,
        "row_count": len(rows),
        "sql": sql,
        "rows": rows,
    }
    target.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return target


def refresh_inventory_snapshot_via_mcp(exports_dir: Path, timeout_seconds: int = 180) -> Path:
    with McpRemoteClient(timeout_seconds=timeout_seconds) as client:
        rows = client.execute_sql(INVENTORY_SQL, timeout_seconds=timeout_seconds)
    return write_snapshot(exports_dir, "redshift_mcp_development_snapshot", INVENTORY_SQL, rows)


def load_or_refresh_inventory_snapshot(exports_dir: Path) -> tuple[dict[str, Any], Path]:
    path = latest_snapshot(exports_dir)
    if path is None:
        path = refresh_inventory_snapshot_via_mcp(exports_dir)
    payload = json.loads(path.read_text(encoding="utf-8"))
    return payload, path
