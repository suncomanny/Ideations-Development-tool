from __future__ import annotations

from pathlib import Path
from typing import Any


APPROVED_DB_SOURCE_LABELS = {
    "postgres",
    "postgres_odbc",
    "postgres_export",
    "private_postgres_export",
    "redshift",
    "redshift_cache",
    "redshift_stackline_cache",
}

FORBIDDEN_SOURCE_MARKERS = {
    "claude workbook",
    "sunco all metadata",
    "legacy_fallback",
    "legacy fallback",
    "full sunco",
    "full nsl",
    "2025 all sales",
    "sunco all specs reference",
}


def normalized_source_text(*values: Any) -> str:
    return " ".join(str(value or "") for value in values).strip().lower()


def is_approved_db_source(value: Any) -> bool:
    text = normalized_source_text(value)
    if not text:
        return False
    return any(label in text for label in APPROVED_DB_SOURCE_LABELS)


def contains_forbidden_reference(*values: Any) -> bool:
    text = normalized_source_text(*values)
    if not text:
        return False
    return any(marker in text for marker in FORBIDDEN_SOURCE_MARKERS)


def classify_source(source: Any, reference: Any = None) -> tuple[bool, str]:
    if contains_forbidden_reference(source, reference):
        return False, "Blocked by source policy: legacy local workbook/CSV reference."
    if not is_approved_db_source(source):
        return False, "Blocked by source policy: source is not marked as Postgres or Redshift."
    return True, "Approved database source."


def source_policy_text() -> str:
    labels = ", ".join(sorted(APPROVED_DB_SOURCE_LABELS))
    blocked = ", ".join(sorted(FORBIDDEN_SOURCE_MARKERS))
    return (
        "Approved line-review evidence must come from Postgres or Redshift exports/cache only. "
        f"Accepted source labels: {labels}. Blocked legacy markers: {blocked}."
    )


def path_has_forbidden_reference(path: Path) -> bool:
    return contains_forbidden_reference(str(path))
