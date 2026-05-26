from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .categories import Category
from .paths import ProjectPaths


DATABASE_RELATIVE_PATH = Path("source_data") / "category_intelligence" / "sunco_category_intelligence.sqlite"


@dataclass(frozen=True)
class CategoryIntelligence:
    category: Category
    database_path: Path
    summary: dict[str, Any]
    attribute_defaults: dict[str, Any]
    feature_signals: list[str]
    gap_evidence: list[dict[str, Any]]
    audit: list[dict[str, Any]]


def database_path(paths: ProjectPaths) -> Path:
    return paths.backend / DATABASE_RELATIVE_PATH


def connect(paths: ProjectPaths) -> sqlite3.Connection:
    connection = sqlite3.connect(database_path(paths))
    connection.row_factory = sqlite3.Row
    return connection


def ensure_category_intelligence_database(paths: ProjectPaths, force: bool = False) -> Path:
    path = database_path(paths)
    if force or not path.exists():
        from .build_category_intelligence import build_category_intelligence_database

        build_category_intelligence_database(paths, path)
    return path


def _json_loads(value: Any, default: Any) -> Any:
    if value in (None, ""):
        return default
    try:
        return json.loads(value)
    except (TypeError, json.JSONDecodeError):
        return default


def _rows(cursor: sqlite3.Cursor) -> list[dict[str, Any]]:
    return [dict(row) for row in cursor.fetchall()]


def category_summary(paths: ProjectPaths, category: Category) -> dict[str, Any]:
    ensure_category_intelligence_database(paths)
    with connect(paths) as connection:
        row = connection.execute(
            "SELECT * FROM category_intelligence_summary WHERE slug = ?",
            (category.slug,),
        ).fetchone()
    return dict(row) if row else {}


def load_category_intelligence(paths: ProjectPaths, category: Category) -> CategoryIntelligence:
    ensure_category_intelligence_database(paths)
    with connect(paths) as connection:
        summary_row = connection.execute(
            "SELECT * FROM category_intelligence_summary WHERE slug = ?",
            (category.slug,),
        ).fetchone()
        category_row = connection.execute(
            "SELECT category_id FROM categories WHERE slug = ?",
            (category.slug,),
        ).fetchone()
        if not category_row:
            return CategoryIntelligence(category, database_path(paths), {}, {}, [], [], [])

        category_id = int(category_row["category_id"])
        distributions = _rows(
            connection.execute(
                """
                SELECT attribute_name, attribute_value, signal_class, coverage_pct, occurrence_count
                FROM category_attribute_distribution
                WHERE category_id = ?
                ORDER BY signal_class, occurrence_count DESC, attribute_name
                """,
                (category_id,),
            )
        )
        profile_rows = _rows(
            connection.execute(
                """
                SELECT feature_signals_json
                FROM category_feature_signal_profile
                WHERE category_id = ? OR category_id IS NULL
                ORDER BY CASE WHEN category_id = ? THEN 0 ELSE 1 END, profile_id
                """,
                (category_id, category_id),
            )
        )
        evidence = _rows(
            connection.execute(
                """
                SELECT source_channel, recommendation, classification, priority, confidence,
                       competitor_example, review_url, sunco_coverage_check, gap_rationale,
                       pm_action, source_systems_json, source_reference, local_image
                FROM gap_evidence
                WHERE category_id = ?
                ORDER BY
                    CASE lower(priority) WHEN 'high' THEN 0 WHEN 'medium' THEN 1 ELSE 2 END,
                    CASE lower(confidence) WHEN 'high' THEN 0 WHEN 'medium' THEN 1 ELSE 2 END,
                    recommendation
                """,
                (category_id,),
            )
        )
        audit = _rows(
            connection.execute(
                """
                SELECT source, run_date, data_age_days, row_counts_json, warnings_json
                FROM refresh_audit
                ORDER BY audit_id DESC
                LIMIT 12
                """
            )
        )

    defaults: dict[str, Any] = {}
    for row in distributions:
        if row["signal_class"] in {"table_stakes", "common"} and row["attribute_value"]:
            defaults.setdefault(row["attribute_name"], row["attribute_value"])

    feature_signals: list[str] = []
    for row in profile_rows:
        for value in _json_loads(row.get("feature_signals_json"), []):
            if value and value not in feature_signals:
                feature_signals.append(value)

    for row in evidence:
        row["source_systems"] = _json_loads(row.pop("source_systems_json"), [])
    for row in audit:
        row["row_counts"] = _json_loads(row.pop("row_counts_json"), {})
        row["warnings"] = _json_loads(row.pop("warnings_json"), [])

    return CategoryIntelligence(
        category=category,
        database_path=database_path(paths),
        summary=dict(summary_row) if summary_row else {},
        attribute_defaults=defaults,
        feature_signals=feature_signals,
        gap_evidence=evidence,
        audit=audit,
    )


def format_intelligence_audit(intelligence: CategoryIntelligence) -> str:
    summary = intelligence.summary or {}
    parts = [
        f"Category intelligence DB: {intelligence.database_path}",
        f"Shopify/catalog products: {summary.get('shopify_product_count', 0)}",
        f"Stackline segments: {summary.get('stackline_segment_count', 0)}",
        f"Stackline top products: {summary.get('stackline_top_product_count', 0)}",
        f"Gap evidence rows: {summary.get('gap_evidence_count', 0)}",
    ]
    if intelligence.feature_signals:
        parts.append("Feature signals: " + ", ".join(intelligence.feature_signals[:12]))
    warnings = []
    for audit in intelligence.audit:
        warnings.extend(audit.get("warnings") or [])
    if warnings:
        parts.append("Refresh warnings: " + " | ".join(dict.fromkeys(warnings)))
    return "\n".join(parts)
