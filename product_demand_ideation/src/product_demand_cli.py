from __future__ import annotations

from pathlib import Path


def refresh_local_sunco_catalog_cache(root: Path | str) -> None:
    root = Path(root)
    import sys

    backend_app = root / "backend" / "app"
    if str(backend_app) not in sys.path:
        sys.path.insert(0, str(backend_app))

    from sunco_catalog_coverage import catalog_cache_path, refresh_catalog_cache_via_odbc

    print("Product Demand local Sunco catalog cache refresh")
    print()
    target = catalog_cache_path(root)
    refreshed = refresh_catalog_cache_via_odbc(target)
    print("Refresh complete:")
    print(f"  {refreshed}")
    print()
    print("Normal Step 1B runs will read this local cache for Sunco catalog coverage.")


def run_product_demand_step1b(root: Path | str) -> None:
    root = Path(root)
    import sys

    backend_app = root / "backend" / "app"
    if str(backend_app) not in sys.path:
        sys.path.insert(0, str(backend_app))

    from step1b_generator import generate_product_demand_step1b

    print("Product Demand Ideation Generator")
    print()
    print("This is the isolated Step 1B shadow tool built from the existing Step 1 pipeline.")
    print("It reuses Step 1 category data, Stackline/Amazon evidence, Sunco coverage checks, workbook format, and Step 2 handoff.")
    print()
    try:
        output, issues, metadata = generate_product_demand_step1b(root)
    except RuntimeError as exc:
        print()
        print("Run stopped:")
        print(f"  {exc}")
        raise SystemExit(1) from exc
    print()
    print("Step 1B output:")
    print(f"  {output}")
    if metadata.get("step2_handoff"):
        print("Step 2 handoff copy:")
        print(f"  {metadata['step2_handoff']}")
    print()
    print("Run details:")
    print(f"  Category: {metadata['category']}")
    print(f"  Recommendations rows: {metadata['source_rows']}")
    print(f"  Amazon rows: {metadata['amazon_rows']}")
    print(f"  Ecommerce competitor source: {metadata['inventory_source']}")
    print(f"  Ecommerce competitor snapshot: {metadata['inventory_snapshot'] or 'none'}")
    print(f"  Sunco catalog source: {metadata['catalog_source']}")
    print(f"  Sunco catalog snapshot: {metadata['catalog_snapshot'] or 'none'}")
    print()
    if issues:
        print("Validation notes:")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("Validation passed.")
