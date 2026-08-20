"""Shared Reference SKU state helpers for Step 3 tools."""

from __future__ import annotations

from typing import Any


NO_CURRENT_SUNCO_SKU = "NO_CURRENT_SUNCO_SKU"
NO_CURRENT_SUNCO_STATUS = "no_current_sunco_sku"
NO_CURRENT_SUNCO_DISPLAY = "No current Sunco SKU"


def is_no_current_sunco_sku(value: Any) -> bool:
    """Return True when a workbook value is the no-current-Sunco sentinel."""
    return str(value or "").strip().upper() == NO_CURRENT_SUNCO_SKU


def display_reference_sku(value: Any) -> Any:
    """Return a PM-facing label for the no-current-Sunco sentinel."""
    if is_no_current_sunco_sku(value):
        return NO_CURRENT_SUNCO_DISPLAY
    return value
