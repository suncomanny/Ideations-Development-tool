from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

from ecommerce_evidence import build_ecommerce_sql
from odbc_client import execute_odbc_sql


def _text(value: Any) -> str:
    return str(value or "").replace("\r", " ").replace("\n", " ").strip()


def _load_categories(root: Path) -> list[Any]:
    import sys

    backend_app = root / "backend" / "app"
    if str(backend_app) not in sys.path:
        sys.path.insert(0, str(backend_app))

    from opportunity_engine.paths import ProjectPaths
    from product_demand_categories import load_product_demand_categories

    return load_product_demand_categories(ProjectPaths.from_root(root))


def audit_category_profiles(root: Path | str, limit: int = 12) -> Path:
    root = Path(root)
    output_dir = root / "product_demand_ideation" / "profile_audits"
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / "category_profile_audit.csv"
    rows: list[dict[str, Any]] = []

    for category in _load_categories(root):
        result: dict[str, Any] = {
            "category": category.run_name,
            "category_slug": category.slug,
            "row_count": 0,
            "observed_stock_decrease": 0,
            "domain_count": 0,
            "sample_1": "",
            "sample_2": "",
            "sample_3": "",
            "status": "ok",
            "error": "",
        }
        try:
            evidence_rows = execute_odbc_sql("DSN=Redshift", build_ecommerce_sql(category.slug, limit=limit), timeout_seconds=90)
            domains = sorted({_text(row.get("domain")).lower() for row in evidence_rows if _text(row.get("domain"))})
            result["row_count"] = len(evidence_rows)
            result["observed_stock_decrease"] = sum(float(row.get("observed_stock_decrease") or 0) for row in evidence_rows)
            result["domain_count"] = len(domains)
            for index, row in enumerate(evidence_rows[:3], start=1):
                result[f"sample_{index}"] = _text(row.get("name"))[:160]
            if not evidence_rows:
                result["status"] = "empty"
        except Exception as exc:
            result["status"] = "error"
            result["error"] = f"{type(exc).__name__}: {exc}"
        rows.append(result)

    with output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()) if rows else [])
        writer.writeheader()
        writer.writerows(rows)
    return output
