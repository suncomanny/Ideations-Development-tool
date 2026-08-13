from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from .categories import Category, load_categories
from .paths import ProjectPaths
from .stackline_segments import STACKLINE_SEGMENT_OVERRIDES, sql_values, stackline_segments_for_category


CACHE_FOLDER = Path("redshift_stackline_cache")
CACHE_MAX_AGE_DAYS = 30
RETAILER_CHANNELS = {
    1: "amazon",
    6: "home_depot",
}


class StacklineRefreshError(RuntimeError):
    """Raised when a required Redshift Stackline cache refresh cannot complete."""


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def today_stamp() -> str:
    return date.today().isoformat()


def product_demand_src(paths: ProjectPaths) -> Path:
    return paths.root / "product_demand_ideation" / "src"


def redshift_query_helpers(paths: ProjectPaths):
    src = product_demand_src(paths)
    if str(src) not in sys.path:
        sys.path.insert(0, str(src))
    from redshift_query import RedshiftQueryClient, sanitize_redshift_error

    return RedshiftQueryClient, sanitize_redshift_error


def cache_dir(paths: ProjectPaths) -> Path:
    target = paths.source_data / CACHE_FOLDER
    target.mkdir(parents=True, exist_ok=True)
    return target


def cache_path(paths: ProjectPaths, category: Category) -> Path:
    return cache_dir(paths) / f"{category.slug}_stackline_redshift_{today_stamp()}.json"


def sql_path(paths: ProjectPaths, category: Category) -> Path:
    return cache_dir(paths) / f"{category.slug}_stackline_redshift_{today_stamp()}.sql"


def latest_category_cache(paths: ProjectPaths, category: Category) -> Path | None:
    candidates = sorted(
        cache_dir(paths).glob(f"{category.slug}_stackline_redshift_*.json"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def cache_created_date(path: Path) -> date:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        created_at = str(payload.get("created_at") or "")[:10]
        if created_at:
            return date.fromisoformat(created_at)
    except Exception:
        pass
    return datetime.fromtimestamp(path.stat().st_mtime).date()


def cache_is_fresh(path: Path, max_age_days: int = CACHE_MAX_AGE_DAYS) -> bool:
    return (date.today() - cache_created_date(path)).days <= max_age_days


def sql_like(value: str) -> str:
    return "'%" + value.replace("'", "''").lower() + "%'"


def segment_filter_sql(category: Category) -> str:
    exact_segments = STACKLINE_SEGMENT_OVERRIDES.get(category.slug)
    if exact_segments:
        return f"cs.segment_name in ({sql_values(exact_segments)})"

    terms = stackline_segments_for_category(category.slug, category.run_name, category.powerbi_aliases)
    if not terms:
        return "1 = 0"
    return " or ".join(f"lower(cs.segment_name) like {sql_like(term)}" for term in terms)


def build_segment_discovery_sql(category: Category) -> str:
    return f"""
select
    cs.segment_name,
    cs.segment_id,
    cs.retailer_id,
    count(distinct cs.retailer_sku) as sku_count
from public.tb_stackline_atlas_current_segment cs
where cs.retailer_id in (1, 6)
  and ({segment_filter_sql(category)})
group by 1, 2, 3
order by sku_count desc, cs.segment_name, cs.retailer_id;
""".strip()


def segment_ctes(segment_id: str) -> str:
    escaped = segment_id.replace("'", "''")
    return f"""
with segment_skus as (
    select segment_name, segment_id, retailer_id, retailer_sku
    from public.tb_stackline_atlas_current_segment
    where segment_id = '{escaped}'
      and retailer_id in (1, 6)
), ranked_weeks as (
    select week_id, row_number() over (order by week_id desc) as rn
    from (
        select distinct gs.week_id
        from public.tb_stackline_atlas_gold_sales gs
        join segment_skus s
          on s.retailer_id = gs.retailer_id
         and s.retailer_sku = gs.retailer_sku
    ) w
), week_periods as (
    select week_id,
           case
               when rn between 1 and 13 then 'Main'
               when rn between 14 and 26 then 'Comparison'
           end as time_period
    from ranked_weeks
    where rn <= 26
)
""".strip()


def build_period_metrics_sql(segment_id: str) -> str:
    return f"""
{segment_ctes(segment_id)}, sales_periods as (
    select
        s.retailer_id,
        wp.time_period,
        gs.retailer_sku,
        p.brand_name,
        sum(coalesce(gs.retail_sales, 0)) as retail_sales,
        sum(coalesce(gs.units_sold, 0)) as units_sold,
        avg(nullif(gs.retail_price, 0)) as avg_retail_price
    from segment_skus s
    join public.tb_stackline_atlas_gold_sales gs
      on s.retailer_id = gs.retailer_id
     and s.retailer_sku = gs.retailer_sku
    join week_periods wp on gs.week_id = wp.week_id
    left join public.tb_stackline_atlas_products p
      on gs.retailer_id = p.retailer_id
     and gs.retailer_sku = p.retailer_sku
    where wp.time_period is not null
    group by 1, 2, 3, 4
), traffic_periods as (
    select
        s.retailer_id,
        wp.time_period,
        gt.retailer_sku,
        sum(coalesce(gt.total_traffic, 0)) as total_traffic
    from segment_skus s
    join public.tb_stackline_atlas_gold_traffic gt
      on s.retailer_id = gt.retailer_id
     and s.retailer_sku = gt.retailer_sku
    join week_periods wp on gt.week_id = wp.week_id
    where wp.time_period is not null
    group by 1, 2, 3
)
select
    (select max(week_id) from ranked_weeks) as latest_week_id,
    (select min(week_id) from week_periods where time_period = 'Main') as main_start_week_id,
    (select max(week_id) from week_periods where time_period = 'Main') as main_end_week_id,
    (select count(distinct week_id) from week_periods where time_period = 'Main') as main_week_count,
    (select min(week_id) from week_periods where time_period = 'Comparison') as comparison_start_week_id,
    (select max(week_id) from week_periods where time_period = 'Comparison') as comparison_end_week_id,
    (select count(distinct week_id) from week_periods where time_period = 'Comparison') as comparison_week_count,
    sp.retailer_id,
    sp.time_period,
    count(distinct sp.retailer_sku) as catalog_product_count,
    count(distinct nullif(sp.brand_name, '')) as brand_count,
    sum(sp.retail_sales) as retail_sales,
    sum(sp.units_sold) as units_sold,
    case when sum(sp.units_sold) = 0 then null else sum(sp.retail_sales) / sum(sp.units_sold) end as avg_retail_price,
    sum(coalesce(tp.total_traffic, 0)) as total_traffic,
    case when sum(coalesce(tp.total_traffic, 0)) = 0 then null else sum(sp.units_sold) / sum(coalesce(tp.total_traffic, 0)) * 100 end as conversion_rate_pct
from sales_periods sp
left join traffic_periods tp
  on sp.retailer_id = tp.retailer_id
 and sp.retailer_sku = tp.retailer_sku
 and sp.time_period = tp.time_period
where sp.retail_sales > 0
group by 1,2,3,4,5,6,7,8,9
order by sp.retailer_id, sp.time_period;
""".strip()


def build_price_percentiles_sql(segment_id: str) -> str:
    return f"""
{segment_ctes(segment_id)}, sku_prices as (
    select
        s.retailer_id,
        gs.retailer_sku,
        avg(nullif(gs.retail_price, 0)) as avg_price
    from segment_skus s
    join public.tb_stackline_atlas_gold_sales gs
      on s.retailer_id = gs.retailer_id
     and s.retailer_sku = gs.retailer_sku
    join week_periods wp on gs.week_id = wp.week_id
    where wp.time_period = 'Main'
      and nullif(gs.retail_price, 0) is not null
    group by 1, 2
)
select
    retailer_id,
    count(*) as priced_products,
    min(avg_price) as min_price,
    percentile_cont(0.25) within group (order by avg_price) as p25_price,
    percentile_cont(0.50) within group (order by avg_price) as p50_price,
    percentile_cont(0.55) within group (order by avg_price) as p55_price,
    percentile_cont(0.75) within group (order by avg_price) as p75_price,
    max(avg_price) as max_price
from sku_prices
group by retailer_id
order by retailer_id;
""".strip()


def build_top_brands_sql(segment_id: str, limit: int = 10) -> str:
    return f"""
{segment_ctes(segment_id)}, brand_sales as (
    select
        s.retailer_id,
        coalesce(nullif(p.brand_name, ''), 'Unknown') as brand,
        count(distinct gs.retailer_sku) as product_count,
        sum(coalesce(gs.retail_sales, 0)) as retail_sales,
        sum(coalesce(gs.units_sold, 0)) as units_sold,
        case when sum(coalesce(gs.units_sold, 0)) = 0 then null else sum(coalesce(gs.retail_sales, 0)) / sum(coalesce(gs.units_sold, 0)) end as avg_retail_price
    from segment_skus s
    join public.tb_stackline_atlas_gold_sales gs
      on s.retailer_id = gs.retailer_id
     and s.retailer_sku = gs.retailer_sku
    join week_periods wp on gs.week_id = wp.week_id
    left join public.tb_stackline_atlas_products p
      on gs.retailer_id = p.retailer_id
     and gs.retailer_sku = p.retailer_sku
    where wp.time_period = 'Main'
    group by 1, 2
), ranked as (
    select
        *,
        case when sum(retail_sales) over (partition by retailer_id) = 0 then null else retail_sales / sum(retail_sales) over (partition by retailer_id) * 100 end as sales_share_pct,
        case when sum(units_sold) over (partition by retailer_id) = 0 then null else units_sold / sum(units_sold) over (partition by retailer_id) * 100 end as units_share_pct,
        row_number() over (partition by retailer_id order by retail_sales desc, units_sold desc) as rn
    from brand_sales
    where retail_sales > 0
)
select retailer_id, brand, retail_sales, units_sold, avg_retail_price, product_count, sales_share_pct, units_share_pct
from ranked
where rn <= {int(limit)}
order by retailer_id, rn;
""".strip()


def build_top_products_sql(segment_id: str, limit: int = 15) -> str:
    return f"""
{segment_ctes(segment_id)}, product_sales as (
    select
        s.retailer_id,
        gs.retailer_sku,
        p.model_number,
        p.title as product_title,
        coalesce(nullif(p.brand_name, ''), 'Unknown') as brand,
        sum(coalesce(gs.retail_sales, 0)) as retail_sales,
        sum(coalesce(gs.units_sold, 0)) as units_sold,
        case when sum(coalesce(gs.units_sold, 0)) = 0 then null else sum(coalesce(gs.retail_sales, 0)) / sum(coalesce(gs.units_sold, 0)) end as avg_retail_price,
        max(gs.reviews_count) as reviews_count,
        avg(gs.reviews_rating) as reviews_rating,
        avg(gs.content_score) as content_score,
        avg(gs.title_score) as title_score,
        avg(gs.image_score) as image_score
    from segment_skus s
    join public.tb_stackline_atlas_gold_sales gs
      on s.retailer_id = gs.retailer_id
     and s.retailer_sku = gs.retailer_sku
    join week_periods wp on gs.week_id = wp.week_id
    left join public.tb_stackline_atlas_products p
      on gs.retailer_id = p.retailer_id
     and gs.retailer_sku = p.retailer_sku
    where wp.time_period = 'Main'
    group by 1, 2, 3, 4, 5
), ranked as (
    select
        *,
        case when sum(retail_sales) over (partition by retailer_id) = 0 then null else retail_sales / sum(retail_sales) over (partition by retailer_id) * 100 end as sales_share_pct,
        case when sum(units_sold) over (partition by retailer_id) = 0 then null else units_sold / sum(units_sold) over (partition by retailer_id) * 100 end as units_share_pct,
        row_number() over (partition by retailer_id order by retail_sales desc, units_sold desc) as rn
    from product_sales
    where retail_sales > 0
)
select retailer_id, retailer_sku, model_number, product_title, brand, retail_sales, units_sold, avg_retail_price, reviews_count, reviews_rating, content_score, title_score, image_score, sales_share_pct, units_share_pct
from ranked
where rn <= {int(limit)}
order by retailer_id, rn;
""".strip()


def combined_sql(category: Category, segment_id: str | None = None) -> str:
    parts = [
        f"-- Redshift Stackline cache refresh for {category.run_name}.",
        f"-- Generated {utc_now()}.",
        "",
        "-- Segment discovery.",
        build_segment_discovery_sql(category),
    ]
    if segment_id:
        parts.extend(
            [
                "",
                "-- Main/comparison 13-week segment metrics.",
                build_period_metrics_sql(segment_id),
                "",
                "-- Main-period price percentiles.",
                build_price_percentiles_sql(segment_id),
                "",
                "-- Main-period top brands.",
                build_top_brands_sql(segment_id),
                "",
                "-- Main-period top competitor products.",
                build_top_products_sql(segment_id),
            ]
        )
    return "\n\n".join(parts)


def choose_segment(category: Category, rows: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not rows:
        return None

    exact_segments = [value.lower() for value in STACKLINE_SEGMENT_OVERRIDES.get(category.slug, ())]
    grouped: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        key = (str(row.get("segment_id") or ""), str(row.get("segment_name") or ""))
        if not key[0] and not key[1]:
            continue
        item = grouped.setdefault(
            key,
            {
                "segment_id": key[0],
                "segment_name": key[1],
                "retailer_sku_counts": {},
                "total_sku_count": 0,
            },
        )
        retailer_id = int(row.get("retailer_id") or 0)
        count = int(float(row.get("sku_count") or 0))
        item["retailer_sku_counts"][RETAILER_CHANNELS.get(retailer_id, str(retailer_id))] = count
        item["total_sku_count"] += count

    if not grouped:
        return None

    def sort_key(item: dict[str, Any]) -> tuple[int, int]:
        name = str(item.get("segment_name") or "").lower()
        priority = exact_segments.index(name) if name in exact_segments else len(exact_segments) + 1
        return priority, -int(item.get("total_sku_count") or 0)

    return sorted(grouped.values(), key=sort_key)[0]


def rows_by_retailer(rows: list[dict[str, Any]]) -> dict[int, list[dict[str, Any]]]:
    output: dict[int, list[dict[str, Any]]] = {}
    for row in rows:
        retailer_id = int(row.get("retailer_id") or 0)
        if retailer_id:
            output.setdefault(retailer_id, []).append(row)
    return output


def strip_retailer(row: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in row.items() if key != "retailer_id"}


def first_period_row(metrics_rows: list[dict[str, Any]]) -> dict[str, Any]:
    return metrics_rows[0] if metrics_rows else {}


def build_cache_payload(
    category: Category,
    segment: dict[str, Any],
    metrics_rows: list[dict[str, Any]],
    price_rows: list[dict[str, Any]],
    brand_rows: list[dict[str, Any]],
    product_rows: list[dict[str, Any]],
    source_sql_file: str,
    connection_source: str,
) -> dict[str, Any]:
    metrics_by_retailer = rows_by_retailer(metrics_rows)
    price_by_retailer = {int(row.get("retailer_id") or 0): row for row in price_rows}
    brands_by_retailer = rows_by_retailer(brand_rows)
    products_by_retailer = rows_by_retailer(product_rows)
    period_anchor = first_period_row(metrics_rows)

    channels: dict[str, Any] = {}
    for retailer_id, channel_key in RETAILER_CHANNELS.items():
        retailer_metrics = metrics_by_retailer.get(retailer_id, [])
        if not retailer_metrics:
            continue
        periods: dict[str, dict[str, Any]] = {}
        for row in retailer_metrics:
            period = str(row.get("time_period") or "")
            if not period:
                continue
            periods[period] = strip_retailer({key: value for key, value in row.items() if not key.endswith("_week_id") and key not in {"main_week_count", "comparison_week_count"}})
            periods[period].pop("time_period", None)

        channels[channel_key] = {
            "retailer_id": retailer_id,
            "periods": periods,
            "price_percentiles": strip_retailer(price_by_retailer.get(retailer_id, {})),
            "top_brands": [strip_retailer(row) for row in brands_by_retailer.get(retailer_id, [])],
            "top_competitor_products": [strip_retailer(row) for row in products_by_retailer.get(retailer_id, [])],
        }

    return {
        "schema_version": 1,
        "source_system": "redshift",
        "source_connector": connection_source,
        "created_at": today_stamp(),
        "generated_at": utc_now(),
        "category_slug": category.slug,
        "subcategory": category.run_name,
        "segment": {
            "segment_name": segment.get("segment_name"),
            "segment_id": segment.get("segment_id"),
            "retailer_sku_counts": segment.get("retailer_sku_counts") or {},
        },
        "latest_week_id": period_anchor.get("latest_week_id"),
        "periods": {
            "Main": {
                "start_week_id": period_anchor.get("main_start_week_id"),
                "end_week_id": period_anchor.get("main_end_week_id"),
                "week_count": period_anchor.get("main_week_count"),
            },
            "Comparison": {
                "start_week_id": period_anchor.get("comparison_start_week_id"),
                "end_week_id": period_anchor.get("comparison_end_week_id"),
                "week_count": period_anchor.get("comparison_week_count"),
            },
        },
        "source_sql_file": source_sql_file,
        "warnings": [],
        "channels": channels,
    }


def refresh_category_stackline_cache(
    paths: ProjectPaths,
    category: Category,
    timeout_seconds: int = 240,
    force: bool = False,
    dry_run: bool = False,
    max_age_days: int = CACHE_MAX_AGE_DAYS,
) -> dict[str, Any]:
    paths.ensure()
    existing = latest_category_cache(paths, category)
    if existing and cache_is_fresh(existing, max_age_days=max_age_days) and not force:
        return {
            "category": category.slug,
            "status": "fresh",
            "cache": str(existing),
            "age_days": (date.today() - cache_created_date(existing)).days,
        }

    target_sql = sql_path(paths, category)
    if dry_run:
        target_sql.write_text(combined_sql(category), encoding="utf-8")
        return {"category": category.slug, "status": "dry_run", "sql": str(target_sql)}

    RedshiftQueryClient, sanitize_redshift_error = redshift_query_helpers(paths)
    connection_source = ""
    try:
        with RedshiftQueryClient(
            timeout_seconds=timeout_seconds,
            client_name="sunco-stackline-cache-redshift",
        ) as redshift:
            connection_source = redshift.connection_source
            discovery_sql = build_segment_discovery_sql(category)
            discovery_rows = redshift.execute_sql(discovery_sql, timeout_seconds=timeout_seconds)
            segment = choose_segment(category, discovery_rows)
            if not segment:
                target_sql.write_text(combined_sql(category), encoding="utf-8")
                raise StacklineRefreshError(f"No Redshift Stackline segment matched {category.run_name}.")

            segment_id = str(segment.get("segment_id") or "")
            metrics_rows = redshift.execute_sql(build_period_metrics_sql(segment_id), timeout_seconds=timeout_seconds)
            price_rows = redshift.execute_sql(build_price_percentiles_sql(segment_id), timeout_seconds=timeout_seconds)
            brand_rows = redshift.execute_sql(build_top_brands_sql(segment_id), timeout_seconds=timeout_seconds)
            product_rows = redshift.execute_sql(build_top_products_sql(segment_id), timeout_seconds=timeout_seconds)

        target_sql.write_text(combined_sql(category, segment_id), encoding="utf-8")
        payload = build_cache_payload(
            category=category,
            segment=segment,
            metrics_rows=metrics_rows,
            price_rows=price_rows,
            brand_rows=brand_rows,
            product_rows=product_rows,
            source_sql_file=target_sql.name,
            connection_source=connection_source,
        )
        target_json = cache_path(paths, category)
        target_json.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
        return {
            "category": category.slug,
            "status": "ok",
            "cache": str(target_json),
            "sql": str(target_sql),
            "segment_name": segment.get("segment_name"),
            "channels": list((payload.get("channels") or {}).keys()),
            "latest_week_id": payload.get("latest_week_id"),
        }
    except StacklineRefreshError:
        raise
    except Exception as exc:
        detail = sanitize_redshift_error(exc)
        raise StacklineRefreshError(
            f"Redshift Stackline refresh failed for {category.run_name} through {connection_source}: {detail}"
        ) from exc


def ensure_stackline_cache_for_category(
    paths: ProjectPaths,
    category: Category,
    timeout_seconds: int = 240,
    max_age_days: int = CACHE_MAX_AGE_DAYS,
) -> dict[str, Any]:
    try:
        return refresh_category_stackline_cache(
            paths=paths,
            category=category,
            timeout_seconds=timeout_seconds,
            force=False,
            dry_run=False,
            max_age_days=max_age_days,
        )
    except StacklineRefreshError as exc:
        raise StacklineRefreshError(
            "Step 3 requires a fresh Redshift-derived Stackline cache before report generation. "
            f"{exc}"
        ) from exc


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


def refresh_redshift_stackline_caches(
    paths: ProjectPaths,
    categories: list[Category],
    timeout_seconds: int = 240,
    force: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    results = []
    for index, category in enumerate(categories, start=1):
        print(f"[{index}/{len(categories)}] Refreshing Stackline cache for {category.run_name}...")
        try:
            result = refresh_category_stackline_cache(
                paths=paths,
                category=category,
                timeout_seconds=timeout_seconds,
                force=force,
                dry_run=dry_run,
            )
            print(f"  {result['status']}: {result.get('cache') or result.get('sql')}")
            results.append(result)
        except StacklineRefreshError as exc:
            print(f"  failed: {exc}")
            results.append({"category": category.slug, "status": "failed", "error": str(exc)})

    return {
        "generated_at": utc_now(),
        "connection_source": "Redshift MCP",
        "categories": len(categories),
        "results": results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Refresh Redshift-derived Stackline cache snapshots for Step 3.")
    parser.add_argument("--root", default=str(Path.cwd()), help="Ideation Development project root.")
    parser.add_argument("--category", action="append", help="Optional category slug/name. Repeat to refresh multiple categories.")
    parser.add_argument("--timeout-seconds", type=int, default=240, help="Per-query Redshift timeout.")
    parser.add_argument("--force", action="store_true", help="Refresh even when a fresh cache already exists.")
    parser.add_argument("--dry-run", action="store_true", help="Write discovery SQL files only; do not query Redshift.")
    args = parser.parse_args()

    paths = ProjectPaths.from_root(args.root)
    categories = select_categories(paths, args.category)
    result = refresh_redshift_stackline_caches(paths, categories, args.timeout_seconds, args.force, args.dry_run)
    manifest = cache_dir(paths) / "redshift_stackline_refresh_manifest.json"
    manifest.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")
    print(f"\nStackline refresh manifest:\n{manifest}")


if __name__ == "__main__":
    main()
