"""
Step 6A: Build Excel research reports from completed analysis artifacts.

Usage:
  python tools/research_report_builder.py "C:\\path\\to\\research_session"
  python tools/research_report_builder.py "C:\\path\\to\\research_session" --rows 3,4,5
  python tools/research_report_builder.py "C:\\path\\to\\research_session" --combined
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
from decimal import Decimal, InvalidOperation
from functools import lru_cache
from pathlib import Path
from typing import Any
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from research_session_manager import artifact_path_for, packet_path_for, read_json, update_session


HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
SUBHEADER_FILL = PatternFill(fill_type="solid", fgColor="D9E2F3")
ACCENT_FILL = PatternFill(fill_type="solid", fgColor="EAF2F8")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=12)
SUBHEADER_FONT = Font(bold=True)
TITLE_FONT = Font(bold=True, size=15)
HYPERLINK_FONT = Font(color="0563C1", underline="single")
WRAP_ALIGNMENT = Alignment(vertical="top", wrap_text=True)
DEFAULT_COLUMN_WIDTHS = {
    "A": 24,
    "B": 34,
    "C": 18,
    "D": 18,
    "E": 18,
    "F": 18,
    "G": 18,
    "H": 24,
    "I": 18,
    "J": 42,
    "K": 14,
    "L": 24,
    "M": 24,
}
ALTERNATE_ROW_FILL = PatternFill(fill_type="solid", fgColor="F7FBFF")
SLIDE_TEAL = "009AA6"
SLIDE_NAVY = "26384F"
SLIDE_DARK_NAVY = "0B1F36"
SLIDE_LIGHT_BLUE = "EAF2F8"
SLIDE_LIGHT_GREY = "F4F7FA"
SLIDE_BORDER = "D7DEE8"
SLIDE_TEXT = "102033"
SLIDE_MUTED = "6B7A90"
SLIDE_TEAL_FILL = PatternFill(fill_type="solid", fgColor=SLIDE_TEAL)
SLIDE_NAVY_FILL = PatternFill(fill_type="solid", fgColor=SLIDE_NAVY)
SLIDE_LIGHT_FILL = PatternFill(fill_type="solid", fgColor=SLIDE_LIGHT_BLUE)
SLIDE_WHITE_FILL = PatternFill(fill_type="solid", fgColor="FFFFFF")
SLIDE_SECTION_FONT = Font(color="FFFFFF", bold=True, size=9)
SLIDE_TITLE_FONT = Font(color=SLIDE_DARK_NAVY, bold=True, size=18)
SLIDE_BODY_FONT = Font(color=SLIDE_TEXT, size=9)
SLIDE_SMALL_FONT = Font(color=SLIDE_TEXT, size=8)
SLIDE_MUTED_FONT = Font(color=SLIDE_MUTED, italic=True, size=7)
SLIDE_LINK_FONT = Font(color="0563C1", underline="single", size=8)
SLIDE_THIN_SIDE = Side(style="thin", color=SLIDE_BORDER)
SLIDE_MEDIUM_SIDE = Side(style="medium", color=SLIDE_TEAL)
SLIDE_TABLE_BORDER = Border(left=SLIDE_THIN_SIDE, right=SLIDE_THIN_SIDE, top=SLIDE_THIN_SIDE, bottom=SLIDE_THIN_SIDE)
AMAZON_CHANNELS = {"amazon"}
BM_DIRECT_CHANNELS = {"home_depot", "walmart", "lowes", "brand_site"}
SKU_DECODER_CACHE = Path(__file__).resolve().parents[3] / "source_data" / "sku_decoder" / "sku_decoder_clean.csv"
CHANNEL_DISPLAY_NAMES = {
    "amazon": "Amazon",
    "home_depot": "Home Depot",
    "walmart": "Walmart",
    "lowes": "Lowe's",
    "brand_site": "Brand Site",
    "stackline_seed": "Stackline",
}

SUMMARY_METRIC_GUIDE_ROWS = [
    [
        "Outlook",
        "Directional recommendation for the ideation in its current configuration, such as favorable, mixed, or cautious.",
        "Given the category, competitor, pricing, and feature context, should this concept look attractive to pursue as configured?",
    ],
    [
        "Confidence",
        "Overall confidence in the launch outlook based on the quality, consistency, and completeness of the supporting research.",
        "How much trust should I place in the outlook recommendation?",
    ],
    [
        "Amazon G2",
        "Weighted Amazon Gate 2 readiness score on a 0-10 scale using the current rubric and available evidence.",
        "How ready is this concept to move forward for Amazon Gate 2 review?",
    ],
    [
        "Amazon Evidence",
        "Confidence label for the Amazon G2 score based on how direct, complete, and well-supported the underlying data is.",
        "How strong is the evidence behind the Amazon G2 score?",
    ],
]

OPPORTUNITY_TYPES = (
    "Possible feature gap",
    "Existing Sunco coverage, but missing feature",
    "Product Revision or merchandising review",
    "New variant opportunity",
)
EVIDENCE_STRENGTH_STRONG = "Strong support"
EVIDENCE_STRENGTH_DIRECTIONAL = "Directional support"
EVIDENCE_STRENGTH_REVIEW = "Needs PM review"
MATCH_QUALITY_APPLES = "Apples-to-apples"
MATCH_QUALITY_SIMILAR = "Similar but not exact"
MATCH_QUALITY_NOT_COMPARABLE = "Not comparable"
LINK_STATUS_VERIFIED = "Verified link"
LINK_STATUS_NEEDS_CHECK = "Needs link check"
LINK_STATUS_MISSING = "No direct link"

SOURCE_LABELS = {
    "legacy_fy2025_sales_export": "Local historical sales fallback",
    "legacy_metadata_variant_price": "Local catalog price fallback",
    "legacy_fy2025_shopify_sales_csv": "Local historical Shopify sales fallback",
    "legacy_fy2025_amazon_sales_csv": "Local historical Amazon sales fallback",
    "legacy_local_metadata_and_fy2025_sales_exports": "Local catalog and historical sales fallback",
    "legacy_fy2025_amazon_export_fallback": "Local historical Amazon sales fallback",
    "mixed_legacy_fy2025_amazon_export_fallback": "Mixed live/cache plus local historical Amazon fallback",
    "target_msrp": "Target MSRP",
}

DISPLAY_TEXT_REPLACEMENTS = (
    ("max parsed wattage", "observed max wattage"),
    (" target from Step 1 evidence", " target"),
    ("No normalized competitors currently surface ", "Current competitor set does not clearly show "),
    (
        "; validate that it represents a real customer need before locking in added complexity.",
        "; treat as a hypothesis unless a PM wants this as a deliberate differentiator.",
    ),
    (
        "; validate whether that certification is a real channel requirement before adding cost.",
        "; confirm this is required for the intended channel before adding cost.",
    ),
    (
        "Autofilled locally by Codex from Stackline-backed packet seeds to reduce Claude collection cost.",
        "Autofilled from Stackline-backed market seeds to reduce manual collection work.",
    ),
    (
        "These are market-intelligence seeds, not live page fetches; use optional web enrichment for missing detailed specs/certifications.",
        "These are market-intelligence seeds, not live page fetches; use optional web enrichment only for missing specs or certifications.",
    ),
    (
        "Brand-site collection was explicitly skipped for a no-Claude/local-only clean run.",
        "Official brand-site evidence was not included in this run.",
    ),
    (
        "No brand product pages were fetched or fabricated in this artifact.",
        "No official brand product pages were used for this row.",
    ),
    (
        "Use Claude or a future web collector later if official brand-site specs are needed.",
        "Use optional web enrichment if official brand-site specs are needed.",
    ),
    ("Prioritize Claude collection next.", "Prioritize optional web collection next."),
    ("Legacy FY2025 local export fallback", "Local historical sales fallback period"),
    ("validate premium feature justification or lower price", "confirm premium feature justification or lower price"),
)


def resolved_source_channel(item: dict[str, Any]) -> str:
    """Return the effective source channel, preferring explicit retailer domains."""
    raw_channel = (optional_text(item.get("source_channel")) or "").lower()
    url = optional_text(item.get("url")) or ""
    detected_domain = urlparse(url).netloc.lower() if "://" in url else ""
    domain = detected_domain or (optional_text(item.get("source_domain")) or "")
    if domain:
        host = urlparse(domain).netloc.lower() if "://" in domain else domain.lower()
        host = host.replace("www.", "")
        if host == "amazon.com" or host.endswith(".amazon.com"):
            return "amazon"
        if host == "homedepot.com" or host.endswith(".homedepot.com"):
            return "home_depot"
        if host == "walmart.com" or host.endswith(".walmart.com"):
            return "walmart"
        if host == "lowes.com" or host.endswith(".lowes.com"):
            return "lowes"
    return raw_channel


def normalize_text(value: Any) -> str:
    """Render values as readable worksheet strings."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, list):
        return ", ".join([normalize_text(item) for item in value if normalize_text(item)])
    text = str(value)
    text = SOURCE_LABELS.get(text, text)
    for old, new in DISPLAY_TEXT_REPLACEMENTS:
        text = text.replace(old, new)
    return fix_ordinal_suffixes(text)


def ordinal_suffix(number: int) -> str:
    """Return the English ordinal suffix for small display strings."""
    if 10 <= number % 100 <= 20:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(number % 10, "th")


def fix_ordinal_suffixes(text: str) -> str:
    """Fix phrases such as 73th percentile without changing source data."""
    return re.sub(
        r"\b(\d+)(?:st|nd|rd|th)\b",
        lambda match: f"{int(match.group(1))}{ordinal_suffix(int(match.group(1)))}",
        text,
    )


def decimal_from_value(value: Any) -> Decimal | None:
    """Parse simple numeric values without keeping float representation noise."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    if isinstance(value, str):
        text = value.strip().replace("$", "").replace(",", "").replace("%", "")
        if not re.fullmatch(r"-?\d+(?:\.\d+)?", text):
            return None
        try:
            return Decimal(text)
        except InvalidOperation:
            return None
    return None


def format_decimal(value: Decimal, places: int = 2) -> str:
    """Return a readable decimal string with thousands separators."""
    quantized = value.quantize(Decimal("1") if places == 0 else Decimal("1." + ("0" * places)))
    return f"{quantized:,.{places}f}"


def display_value_for_label(label: str, value: Any) -> str:
    """Format report values using the row label/header for context."""
    text = normalize_text(value)
    label_key = normalize_text(label).lower()
    number = decimal_from_value(value if value is not None else text)
    if number is None:
        return text

    if any(term in label_key for term in ["%", "percent", "percentile", "growth", "share", "vs target", "vs market"]):
        return f"{format_decimal(number, 1)}%"
    if any(term in label_key for term in ["score", "confidence"]):
        return format_decimal(number, 2).rstrip("0").rstrip(".")
    if any(term in label_key for term in ["units", "count", "samples"]):
        return format_decimal(number, 0)
    if any(term in label_key for term in ["price", "pricing", "revenue", "sales", "msrp", "cost", "retail", "floor", "ceiling"]):
        return f"${format_decimal(number, 2)}"
    return format_decimal(number, 2).rstrip("0").rstrip(".")


def as_dict(value: Any) -> dict[str, Any]:
    """Coerce optional mappings into dicts."""
    if isinstance(value, dict):
        return value
    return {}


def as_list(value: Any) -> list[Any]:
    """Coerce optional sequences into lists."""
    if isinstance(value, list):
        return value
    return []


def slugify(value: Any) -> str:
    """Normalize category-like values into stable slugs."""
    text = normalize_text(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


@lru_cache(maxsize=1)
def sku_decoder_entries() -> list[dict[str, str]]:
    """Load the local SKU decoder cache used for concept tracking names."""
    if not SKU_DECODER_CACHE.exists():
        return []
    with SKU_DECODER_CACHE.open("r", encoding="utf-8-sig", newline="") as handle:
        return [dict(row) for row in csv.DictReader(handle)]


def decoder_rows_for(category_slug: str, normalized_code_category: str) -> list[dict[str, str]]:
    """Return SKU decoder rows that apply to a category and code family."""
    return [
        row
        for row in sku_decoder_entries()
        if row.get("mapped_category_slug") == category_slug
        and row.get("normalized_code_category") == normalized_code_category
    ]


def context_text_for_packet(packet: dict[str, Any]) -> str:
    """Collect readable packet context for token inference."""
    identity = as_dict(packet.get("identity"))
    target_profile = as_dict(packet.get("target_profile"))
    electrical = as_dict(target_profile.get("electrical"))
    physical = as_dict(target_profile.get("physical"))
    business = as_dict(target_profile.get("business_case"))
    pieces: list[str] = [
        identity.get("ideation_name"),
        identity.get("category"),
        identity.get("subcategory"),
        identity.get("sunco_reference_sku"),
        electrical.get("wattage_primary"),
        electrical.get("wattage_max"),
        electrical.get("cct_primary"),
        electrical.get("cct_max"),
        electrical.get("lumens_target"),
        physical.get("size_form_factor"),
        physical.get("finish_color"),
        physical.get("mounting_type"),
        business.get("certifications"),
    ]
    pieces.extend(as_list(target_profile.get("feature_watchlist")))
    return " ".join(normalize_text(piece) for piece in pieces if normalize_text(piece))


def preferred_product_type_code(packet: dict[str, Any]) -> str:
    """Infer the most useful product-type code for a Step 3 concept name."""
    identity = as_dict(packet.get("identity"))
    category_slug = slugify(identity.get("subcategory") or identity.get("category"))
    rows = decoder_rows_for(category_slug, "product_type")
    context = context_text_for_packet(packet).lower()
    reference_sku = normalize_text(identity.get("sunco_reference_sku")).upper()

    if category_slug == "wraparounds":
        if re.search(r"\b2\s*(?:ft|foot|feet)\b|\b2ft\b", context):
            return "WR2"
        if re.search(r"\b4\s*(?:ft|foot|feet)\b|\b4ft\b", context):
            return "WR"
    if category_slug == "panels":
        panel_map = {
            "1x2": "PN12",
            "1 x 2": "PN12",
            "1x4": "PN14",
            "1 x 4": "PN14",
            "2x2": "PN22",
            "2 x 2": "PN22",
            "2x4": "PN24",
            "2 x 4": "PN24",
        }
        for token, code in panel_map.items():
            if token in context:
                return code

    non_phasing_rows = [
        row
        for row in rows
        if "phasing" not in row.get("code_meaning", "").lower()
        and "legacy" not in row.get("code_meaning", "").lower()
    ]
    for row in non_phasing_rows:
        match_prefix = normalize_text(row.get("match_prefix")).upper()
        if match_prefix and reference_sku.startswith(match_prefix):
            return normalize_text(row.get("code")).replace("_", "")
    for row in non_phasing_rows:
        meaning = row.get("code_meaning", "").lower()
        if meaning and len(meaning) > 2 and meaning in context:
            return normalize_text(row.get("code")).replace("_", "")
    if non_phasing_rows:
        return normalize_text(non_phasing_rows[0].get("code")).replace("_", "")
    if rows:
        return normalize_text(rows[0].get("code")).replace("_", "")
    fallback = slugify(identity.get("subcategory") or identity.get("category") or "IDEA")
    return fallback[:8].upper() or "IDEA"


def first_wattage_token(text: str) -> str:
    """Return the first wattage token found in text."""
    match = re.search(r"\b(\d+(?:\.\d+)?(?:-\d+(?:\.\d+)?)?)\s*W\b", text, flags=re.IGNORECASE)
    if not match:
        return ""
    number = match.group(1).replace(".0", "")
    return f"{number}W"


def first_lumen_token(text: str) -> str:
    """Return a compact lumen token when wattage is unavailable."""
    match = re.search(r"\b([0-9][0-9,]*)\s*(?:lm|lumens?)\b", text, flags=re.IGNORECASE)
    if not match:
        return ""
    return f"{match.group(1).replace(',', '')}LM"


def cct_token_from_text(text: str) -> str:
    """Infer a SKU-decoder-style CCT token from available target text."""
    lowered = text.lower()
    cct_values = [int(value) for value in re.findall(r"\b((?:27|30|35|40|50|60|65|70)00)\s*k\b", lowered)]
    expanded = sorted(set(cct_values))
    if 3000 in expanded and 5000 in expanded:
        return "3050"
    if 3000 in expanded and 6500 in expanded:
        return "3065"
    if 4000 in expanded and 6000 in expanded:
        return "4060"
    if 2700 in expanded and 5000 in expanded:
        return "2750"
    if 2700 in expanded and 4000 in expanded:
        return "2740"
    if 2700 in expanded and 6000 in expanded:
        return "2760"
    if "3cct" in lowered and not expanded:
        return "3CCT"
    single = re.search(r"\b(2700|3000|3500|4000|5000|6000|6500|7000)\s*K?\b", text, flags=re.IGNORECASE)
    return single.group(1) if single else ""


def finish_token_from_text(text: str) -> str:
    """Infer a finish token using common SKU decoder color codes."""
    lowered = text.lower()
    finish_map = [
        ("black", "BK"),
        ("bronze", "BR"),
        ("chrome", "CR"),
        ("gold", "GD"),
        ("nickel", "NK"),
        ("pure white", "PW"),
        ("silver", "SV"),
        ("white", "WH"),
    ]
    for label, code in finish_map:
        if label in lowered:
            return code
    return ""


def feature_tokens_from_text(text: str) -> list[str]:
    """Infer compact feature/control tokens for concept tracking names."""
    lowered = text.lower()
    tokens: list[str] = []
    if "ultra high output" in lowered:
        tokens.append("UO")
    elif "high output" in lowered:
        tokens.append("HO")
    if "selectable wattage" in lowered:
        tokens.append("SW")
    if "0-10" in lowered or "0 to 10" in lowered:
        tokens.append("010V")
    elif "dimmable" in lowered:
        tokens.append("DIM")
    if "dusk to dawn" in lowered or "d2d" in lowered:
        tokens.append("DD")
    if "smart" in lowered:
        tokens.append("S")
    if "sensor" in lowered:
        tokens.append("SNS")
    return list(dict.fromkeys(tokens))


def pack_token_from_text(text: str) -> str:
    """Return a compact pack suffix such as /4 when a multi-pack is explicit."""
    match = re.search(r"\b(\d+)\s*(?:pack|pk)\b", text, flags=re.IGNORECASE)
    if not match:
        return ""
    count = int(match.group(1))
    return f"/{count}" if count > 1 else ""


def concept_tracking_name(packet: dict[str, Any]) -> str:
    """Build a compact, SKU-decoder-informed concept tracking name."""
    context = context_text_for_packet(packet)
    product_code = preferred_product_type_code(packet)
    wattage = first_wattage_token(context)
    lumen = first_lumen_token(context)
    cct = cct_token_from_text(context)
    finish = finish_token_from_text(context)
    tokens = [product_code]
    if wattage:
        tokens.append(wattage)
    elif lumen:
        tokens.append(lumen)
    if cct:
        tokens.append(cct)
    tokens.extend(feature_tokens_from_text(context))
    if finish:
        tokens.append(finish)
    name = "-".join(token for token in tokens if token)
    name = re.sub(r"-+", "-", name).strip("-")
    return f"{name}{pack_token_from_text(context)}" if name else "IDEA"


def set_default_layout(ws) -> None:
    """Apply shared column widths and wrapping."""
    ws.freeze_panes = None
    ws.sheet_view.showGridLines = False
    for column, width in DEFAULT_COLUMN_WIDTHS.items():
        ws.column_dimensions[column].width = width


def sheet_title_for_row(row_number: int, packet: dict[str, Any], used: set[str] | None = None) -> str:
    """Create a compact row-first sheet title for combined reports."""
    tracking_name = re.sub(r"[^A-Za-z0-9/_-]+", " ", concept_tracking_name(packet)).strip()
    base = f"R{row_number:02d} {tracking_name}" if tracking_name else f"R{row_number:02d} Ideation"
    return safe_sheet_title(base, used or set())


def safe_sheet_title(base: str, used: set[str]) -> str:
    """Create a workbook-safe, unique worksheet title."""
    text = re.sub(r"[\[\]\*?/\\:]", " ", base or "Report").strip()
    text = re.sub(r"\s+", " ", text)
    text = text[:31] or "Report"
    candidate = text
    suffix = 2
    while candidate in used:
        trimmed = text[: max(0, 31 - len(f" ({suffix})"))]
        candidate = f"{trimmed} ({suffix})"
        suffix += 1
    used.add(candidate)
    return candidate


def section_header(ws, row: int, title: str, end_column: int = 10) -> int:
    """Write a section header row and return the next row index."""
    ws.cell(row=row, column=1, value=title)
    ws.cell(row=row, column=1).fill = HEADER_FILL
    ws.cell(row=row, column=1).font = HEADER_FONT
    ws.cell(row=row, column=1).alignment = WRAP_ALIGNMENT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_column)
    return row + 1


def key_value_rows(ws, row: int, pairs: list[tuple[str, Any]], columns: int = 2) -> int:
    """Write compact key/value pairs across the sheet."""
    index = 0
    while index < len(pairs):
        for block in range(columns):
            if index >= len(pairs):
                break
            label, value = pairs[index]
            base_col = (block * 2) + 1
            ws.cell(row=row, column=base_col, value=label)
            ws.cell(row=row, column=base_col).font = SUBHEADER_FONT
            ws.cell(row=row, column=base_col).fill = SUBHEADER_FILL
            ws.cell(row=row, column=base_col + 1, value=display_value_for_label(label, value))
            ws.cell(row=row, column=base_col + 1).alignment = WRAP_ALIGNMENT
            index += 1
        row += 1
    return row


def gate_snapshot(gate_readiness: dict[str, Any], channel: str, gate: str) -> dict[str, Any]:
    """Return the matching gate snapshot when present."""
    target_channel = channel.strip().lower()
    target_gate = gate.strip().upper()
    for snapshot in as_list(gate_readiness.get("snapshots")):
        snapshot = as_dict(snapshot)
        if (
            normalize_text(snapshot.get("channel")).strip().lower() == target_channel
            and normalize_text(snapshot.get("gate")).strip().upper() == target_gate
        ):
            return snapshot
    return {}


def issue_sentence(parts: list[str]) -> str:
    """Join short issue fragments into one readable sentence."""
    parts = [part for part in parts if part]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return f"{parts[0]} and {parts[1]}"
    return f"{', '.join(parts[:-1])}, and {parts[-1]}"


def blocker_action_summary(packet: dict[str, Any], analysis: dict[str, Any]) -> dict[str, Any]:
    """Build a compact top-of-report blocker/action summary."""
    identity = as_dict(packet.get("identity"))
    performance = as_dict(analysis.get("performance_estimation"))
    pricing = as_dict(analysis.get("pricing_analysis"))
    spec_coverage = as_dict(analysis.get("spec_coverage"))
    gate_readiness = as_dict(analysis.get("gate_readiness"))

    suggested = as_dict(pricing.get("suggested_msrp_range"))
    margin_targets = as_dict(pricing.get("margin_targets"))
    amazon_margin = as_dict(margin_targets.get("amazon"))
    amazon_snapshot = gate_snapshot(gate_readiness, "amazon", "G2")

    outlook = normalize_text(performance.get("launch_outlook")).strip().lower() or "mixed"
    confidence = normalize_text(performance.get("confidence")).strip().lower() or "medium"
    posture = normalize_text(performance.get("posture")).strip().lower() or "undetermined"
    positioning = normalize_text(suggested.get("positioning")).strip().lower() or "undetermined"
    evidence_label = normalize_text(amazon_snapshot.get("evidence_label")).strip().lower() or "unknown"

    target_msrp = pricing.get("target_msrp")
    target_vendor_cost = pricing.get("target_vendor_cost")
    margin_conflict = bool(suggested.get("margin_conflict"))
    minimum_margin_safe_price = suggested.get("minimum_margin_safe_price")
    recommended_ceiling = suggested.get("recommended_ceiling")
    notable_gaps = [normalize_text(item) for item in as_list(spec_coverage.get("notable_gaps")) if normalize_text(item)]
    gap_count = len(notable_gaps)
    feature_coverage = [as_dict(item) for item in as_list(spec_coverage.get("feature_coverage"))]
    certification_coverage = [as_dict(item) for item in as_list(spec_coverage.get("certification_coverage"))]
    zero_match_features = [entry for entry in feature_coverage if entry.get("matched_count") == 0]
    zero_match_certs = [entry for entry in certification_coverage if entry.get("matched_count") == 0]
    whitespace_labels = [
        normalize_text(entry.get("label"))
        for entry in zero_match_features
        if normalize_text(entry.get("signal")).strip().lower() == "whitespace"
    ]
    weak_signal_labels = [
        normalize_text(entry.get("label"))
        for entry in feature_coverage
        if entry.get("matched_count", 0) > 0 and normalize_text(entry.get("evidence_strength")).strip().lower() == "weak"
    ]
    zero_match_cert_labels = [normalize_text(entry.get("label")) for entry in zero_match_certs]

    actions: list[dict[str, str]] = []

    if target_msrp is None or target_vendor_cost is None:
        missing_bits = []
        if target_msrp is None:
            missing_bits.append("target MSRP")
        if target_vendor_cost is None:
            missing_bits.append("landed vendor cost")
        actions.append(
            {
                "category": "commercial",
                "issue": "Pricing inputs are incomplete",
                "why": f"Without {issue_sentence(missing_bits)}, the model has to infer key commercial inputs instead of scoring them directly.",
                "action": "Enter a provisional MSRP and landed vendor cost for this row, then rerun the report.",
                "impact": "Price Fit; Margin Viability; Outlook",
            }
        )

    if margin_conflict:
        why = "Current target pricing conflicts with the minimum margin-safe price."
        if minimum_margin_safe_price and recommended_ceiling:
            why = (
                f"The current minimum margin-safe price ({normalize_text(minimum_margin_safe_price)}) "
                f"sits above the recommended market ceiling ({normalize_text(recommended_ceiling)})."
            )
        actions.append(
            {
                "category": "commercial",
                "issue": "Margin conflict is still unresolved",
                "why": why,
                "action": "Lower vendor cost, raise MSRP, or trim feature scope until the target clears the safety floor.",
                "impact": "Margin Viability; Price Fit; Outlook",
            }
        )
    elif positioning == "premium":
        actions.append(
            {
                "category": "commercial",
                "issue": "Target pricing is above the market band",
                "why": "The current MSRP is positioned above the observed comparable range and needs stronger feature justification.",
                "action": "Either lower MSRP or strengthen the differentiator and certification story enough to defend a premium position.",
                "impact": "Price Fit; Outlook",
            }
        )

    if posture == "no_stackline_context":
        actions.append(
            {
                "category": "evidence",
                "issue": "Market growth context is missing",
                "why": "This row is missing usable Stackline market-growth context, which caps how strongly the outlook can score.",
                "action": "Confirm the row is mapped to the correct Stackline segment and refresh the run before using Outlook as a go / no-go signal.",
                "impact": "Market Support; Outlook",
            }
        )

    if gap_count >= 3:
        highlighted_labels = [label for label in whitespace_labels + zero_match_cert_labels if label][:3]
        gap_preview = issue_sentence(highlighted_labels[:2]) or issue_sentence(notable_gaps[:2])
        actions.append(
            {
                "category": "evidence",
                "issue": "Competitor evidence does not yet validate several proposed attributes",
                "why": (
                    f"The current caution is being driven by evidence blind spots, not proof the concept is wrong. "
                    f"Right now competitor records do not clearly confirm attributes such as {gap_preview}."
                ),
                "action": (
                    "Treat these as validation items, not automatic negatives: keep intentional innovations, "
                    "confirm whether each item is a real customer requirement or compliance need, and only cut it if it adds cost without demand support."
                ),
                "impact": "Confidence; Market Support; Outlook",
            }
        )
    elif gap_count > 0 and outlook != "favorable":
        highlighted_labels = [label for label in whitespace_labels + zero_match_cert_labels if label][:3]
        gap_preview = issue_sentence(highlighted_labels[:2]) or issue_sentence(notable_gaps[:2])
        actions.append(
            {
                "category": "evidence",
                "issue": "A few proposed attributes still need market validation",
                "why": (
                    f"The concept includes attributes that are not yet well confirmed in competitor evidence, including {gap_preview}."
                ),
                "action": (
                    "Decide whether these are intentional innovations, documentation-only blind spots, or true must-have requirements before treating them as blockers."
                ),
                "impact": "Confidence; Market Support",
            }
        )
    elif weak_signal_labels and outlook != "favorable":
        weak_preview = issue_sentence(weak_signal_labels[:2])
        actions.append(
            {
                "category": "evidence",
                "issue": "Some differentiator claims still rest on thin evidence",
                "why": f"The market set only weakly supports items such as {weak_preview}, so the report should treat them as directional rather than proven gaps or advantages.",
                "action": "Use these as hypotheses to validate, not reasons to downscore the concept unless stronger competitor evidence later contradicts them.",
                "impact": "Confidence; Outlook",
            }
        )

    if evidence_label in {"low", "none"}:
        actions.append(
            {
                "category": "evidence",
                "issue": "Amazon evidence is still thin",
                "why": "The current Amazon Gate 2 score is supported by limited direct evidence, so the score is less trustworthy than the research confidence alone suggests.",
                "action": "Treat the Amazon G2 score as directional until more direct listing / competitor evidence is collected.",
                "impact": "Amazon Evidence; Confidence",
            }
        )

    if not actions:
        actions.append(
            {
                "category": "execution",
                "issue": "Protect the current favorable read",
                "why": "Current market, pricing, and competitor signals are supportive enough that the concept now reads cleanly.",
                "action": "Lock the vendor quote, confirm the key certifications, and preserve the current feature set unless cost changes materially.",
                "impact": "Outlook; Margin Viability; Execution Readiness",
            }
        )

    actions = actions[:3]
    action_categories = [entry.get("category", "") for entry in actions]
    evidence_heavy = bool(actions) and all(category in {"evidence", "market_context"} for category in action_categories)

    subcategory_label = normalize_text(identity.get("subcategory")).strip().lower()

    if outlook == "favorable":
        overall_read = (
            f"This {'concept in ' + subcategory_label if subcategory_label else 'concept'} currently looks favorable to pursue "
            f"with {confidence} research confidence."
        )
        remaining_watchouts = [entry["issue"].lower() for entry in actions if entry.get("category") != "execution"]
        if remaining_watchouts:
            why_not_stronger = (
                "The concept is already favorable, but the remaining watchouts are "
                f"{issue_sentence(remaining_watchouts)}."
            )
        else:
            why_not_stronger = "The concept is already favorable; the main task now is to protect that position as quotes and compliance details firm up."
    elif outlook == "mixed":
        if evidence_heavy:
            overall_read = (
                "This concept looks viable, and the main constraint right now is incomplete market confirmation rather than a clearly weak product position."
            )
            why_not_stronger = (
                "It is not stronger yet because the report still needs clearer evidence around "
                f"{issue_sentence([entry['issue'].lower() for entry in actions])}."
            )
        else:
            overall_read = (
                "This concept looks viable, but it still needs targeted commercial or evidence cleanup before it becomes a strong go-forward candidate."
            )
            why_not_stronger = (
                "It is not stronger yet because "
                f"{issue_sentence([entry['issue'].lower() for entry in actions])}."
            )
    elif outlook == "cautious":
        if evidence_heavy:
            overall_read = (
                "This concept has enough support to analyze, and the current caution is being driven more by evidence ambiguity than by proof the product is weak."
            )
            why_not_stronger = (
                "The main watchouts right now are "
                f"{issue_sentence([entry['issue'].lower() for entry in actions])}."
            )
        else:
            overall_read = (
                "This concept has enough support to analyze, but the current configuration is still being held back by one or more commercial or evidence blockers."
            )
            why_not_stronger = (
                "The main blockers right now are "
                f"{issue_sentence([entry['issue'].lower() for entry in actions])}."
            )
    else:
        overall_read = "This concept still needs more complete evidence before the report can support a strong directional call."
        why_not_stronger = (
            "The current read is still limited because "
            f"{issue_sentence([entry['issue'].lower() for entry in actions])}."
        )

    table_rows = [
        [f"P{index}", entry["issue"], entry["why"], entry["action"], entry["impact"]]
        for index, entry in enumerate(actions, start=1)
    ]

    return {
        "overall_read": overall_read,
        "why_not_stronger": why_not_stronger,
        "actions": table_rows,
    }


def first_sentence(text: Any) -> str:
    """Return a compact first sentence for executive display."""
    clean = normalize_text(text).strip()
    if not clean:
        return ""
    parts = re.split(r"(?<=[.!?])\s+", clean, maxsplit=1)
    return parts[0]


def extract_research_note_evidence(packet: dict[str, Any]) -> dict[str, str]:
    """Extract structured Step 1B evidence from the research notes block."""
    target_profile = as_dict(packet.get("target_profile"))
    notes = normalize_text(target_profile.get("research_notes"))
    evidence: dict[str, str] = {}

    def clean_fragment(value: str) -> str:
        return value.strip().rstrip(".")

    movement_match = re.search(
        r"Ecommerce movement:\s*([^\n]+)",
        notes,
        flags=re.IGNORECASE,
    )
    if movement_match:
        evidence["ecommerce_movement"] = clean_fragment(movement_match.group(1))

    competitor_match = re.search(
        r"Redshift ecommerce competitor evidence found\s*([^.\n]+)",
        notes,
        flags=re.IGNORECASE,
    )
    if competitor_match:
        evidence["competitor_evidence"] = clean_fragment(competitor_match.group(1))

    overlay_match = re.search(r"Product Demand overlay score:\s*([0-9.]+/100)[^\n.]*", notes, flags=re.IGNORECASE)
    if overlay_match:
        evidence["overlay_score"] = clean_fragment(overlay_match.group(0))

    inventory_match = re.search(r"Inventory support:\s*([^\n]+)", notes, flags=re.IGNORECASE)
    if inventory_match:
        evidence["inventory_support"] = clean_fragment(inventory_match.group(1))

    sunco_match = re.search(r"Sunco check:\s*([^\n]+)", notes, flags=re.IGNORECASE)
    if sunco_match:
        evidence["sunco_coverage"] = clean_fragment(sunco_match.group(1))

    recommended_match = re.search(r"Recommended action:\s*([^\n]+)", notes, flags=re.IGNORECASE)
    if recommended_match:
        evidence["recommended_action"] = clean_fragment(recommended_match.group(1))

    review_link_match = re.search(r"Review link:\s*(https?://\S+)", notes, flags=re.IGNORECASE)
    if review_link_match:
        evidence["review_link"] = review_link_match.group(1).strip()

    bias_match = re.search(r"Vendor bias risk:\s*([^.\n]+)", notes, flags=re.IGNORECASE)
    if bias_match:
        evidence["vendor_bias_risk"] = clean_fragment(bias_match.group(1))

    return evidence


def source_link_cell(url: Any, label: str = "Open source") -> dict[str, str] | str:
    """Return a short, clickable source-link payload."""
    text = optional_text(url)
    if not text or text.startswith("stackline://"):
        return ""
    return {"value": label, "hyperlink": text}


def source_link_from_item(item: dict[str, Any], label: str = "Open PDP") -> dict[str, str] | str:
    """Return a short hyperlink payload for a competitor item."""
    url = optional_text(item.get("url"))
    if not url:
        return ""
    return source_link_cell(url, label)


def normalize_opportunity_type(value: Any) -> str:
    """Normalize arbitrary opportunity language into the Step 1-3 shared vocabulary."""
    text = normalize_text(value).lower()
    for opportunity_type in OPPORTUNITY_TYPES:
        if opportunity_type.lower() in text:
            return opportunity_type
    if "missing feature" in text or "existing sunco coverage" in text:
        return "Existing Sunco coverage, but missing feature"
    if "feature gap" in text or "possible feature" in text:
        return "Possible feature gap"
    if "revision" in text or "merchandising" in text or "coverage exists" in text:
        return "Product Revision or merchandising review"
    if "variant" in text or "true gap" in text or "new product" in text:
        return "New variant opportunity"
    return ""


def opportunity_type_for(packet: dict[str, Any], analysis: dict[str, Any] | None = None) -> str:
    """Return the Step 1-3 shared opportunity type for a row."""
    identity = as_dict(packet.get("identity"))
    target_profile = as_dict(packet.get("target_profile"))
    research_notes = normalize_text(target_profile.get("research_notes"))
    evidence = extract_research_note_evidence(packet)
    context_values = [
        identity.get("ideation_name"),
        identity.get("strategy"),
        research_notes,
        evidence.get("recommended_action"),
        evidence.get("sunco_coverage"),
    ]
    for value in context_values:
        normalized = normalize_opportunity_type(value)
        if normalized:
            return normalized

    sunco_coverage = normalize_text(evidence.get("sunco_coverage")).lower()
    has_anchor = bool(normalize_text(identity.get("sunco_reference_sku")))
    if "0 strong" in sunco_coverage or "assortment gap" in sunco_coverage:
        return "New variant opportunity"
    if has_anchor:
        return "Product Revision or merchandising review"
    return "New variant opportunity"


def evidence_strength_for(packet: dict[str, Any], analysis: dict[str, Any]) -> str:
    """Classify the row using the shared evidence-strength language."""
    performance = as_dict(analysis.get("performance_estimation"))
    confidence = normalize_text(performance.get("confidence")).lower()
    outlook = normalize_text(performance.get("launch_outlook")).lower()
    evidence = extract_research_note_evidence(packet)
    overlay_match = re.search(r"([0-9]+(?:\.[0-9]+)?)/100", normalize_text(evidence.get("overlay_score")))
    overlay_score = decimal_from_value(overlay_match.group(1)) if overlay_match else None

    if overlay_score is not None and overlay_score >= Decimal("80"):
        return EVIDENCE_STRENGTH_STRONG
    if any(token in confidence for token in ["high", "strong"]) and outlook == "favorable":
        return EVIDENCE_STRENGTH_STRONG
    if any(token in confidence for token in ["low", "limited", "weak"]) or outlook == "cautious":
        return EVIDENCE_STRENGTH_REVIEW
    return EVIDENCE_STRENGTH_DIRECTIONAL


def match_quality_for_item(item: dict[str, Any]) -> str:
    """Return the shared match-quality label for a competitor item."""
    if not item:
        return MATCH_QUALITY_NOT_COMPARABLE
    confidence = decimal_from_value(item.get("match_confidence"))
    if verification_status(item) == "verified_listing" and confidence is not None and confidence >= Decimal("0.80"):
        return MATCH_QUALITY_APPLES
    if verification_status(item) == "verified_listing":
        return MATCH_QUALITY_SIMILAR
    return MATCH_QUALITY_NOT_COMPARABLE


def link_status_for_item(item: dict[str, Any]) -> str:
    """Return a visible link-status label for manual PM verification."""
    if not item or not optional_text(item.get("url")):
        return LINK_STATUS_MISSING
    if verification_status(item) == "verified_listing":
        return LINK_STATUS_VERIFIED
    return LINK_STATUS_NEEDS_CHECK


def link_status_for_url(url: Any, verified: bool = False) -> str:
    """Return the link-status label for a standalone source URL."""
    if not optional_text(url):
        return LINK_STATUS_MISSING
    return LINK_STATUS_VERIFIED if verified else LINK_STATUS_NEEDS_CHECK


def primary_competitor_item(normalized: dict[str, Any]) -> dict[str, Any]:
    """Return the best source-backed competitor record across preferred channels."""
    return best_verified_item(normalized, AMAZON_CHANNELS) or best_verified_item(normalized, BM_DIRECT_CHANNELS)


def primary_source_link(packet: dict[str, Any], normalized: dict[str, Any]) -> dict[str, str] | str:
    """Return the most useful source link for a top-level PM snapshot row."""
    item = primary_competitor_item(normalized)
    if item:
        return source_link_from_item(item)
    evidence = extract_research_note_evidence(packet)
    return source_link_cell(evidence.get("review_link"), "Open PDP")


def msrp_basis_summary(packet: dict[str, Any], analysis: dict[str, Any]) -> str:
    """Return a compact explanation for the recommended MSRP."""
    pricing = as_dict(analysis.get("pricing_analysis"))
    target_position = as_dict(pricing.get("target_price_position"))
    suggested = as_dict(pricing.get("suggested_msrp_range"))
    pieces: list[str] = []
    percentile = target_position.get("percentile")
    if percentile is not None:
        pieces.append(f"{display_value_for_label('Percentile', percentile)} market price percentile")
    if suggested.get("positioning"):
        pieces.append(normalize_text(suggested.get("positioning")))
    if suggested.get("margin_conflict"):
        pieces.append(f"Margin read: {normalize_text(suggested.get('margin_conflict'))}")
    if pieces:
        return "; ".join(pieces)
    evidence = extract_research_note_evidence(packet)
    if evidence.get("review_link"):
        return "Market MSRP is anchored to the verified Step 1 competitor listing until broader comps are refreshed."
    return "Recommended MSRP is based on Step 2/3 pricing analysis."


def money_or_blank(value: Any) -> str:
    """Format a money-like value for compact text."""
    if value is None:
        return ""
    return display_value_for_label("Price", value)


def number_from_text(text: Any, unit_pattern: str) -> Decimal | None:
    """Find the first numeric value before a unit pattern."""
    match = re.search(rf"\b([0-9][0-9,]*(?:\.\d+)?)\s*{unit_pattern}\b", normalize_text(text), flags=re.IGNORECASE)
    if not match:
        return None
    return decimal_from_value(match.group(1))


def cct_display_from_context(context: str) -> str:
    """Return a readable CCT summary from concept context."""
    values = sorted(set(re.findall(r"\b(?:27|30|35|40|50|60|65|70)00\s*K\b", context, flags=re.IGNORECASE)))
    if values:
        normalized_values = [value.upper().replace(" ", "") if value.upper().endswith("K") else f"{value}K" for value in values]
        return "/".join(normalized_values)
    short = re.search(r"\b([234567])K\b", context, flags=re.IGNORECASE)
    if short:
        return f"{short.group(1)}000K"
    return ""


def product_difference_summary(packet: dict[str, Any], analysis: dict[str, Any]) -> str:
    """Summarize how the new concept differs from the closest Sunco anchor."""
    identity = as_dict(packet.get("identity"))
    reference = as_dict(packet.get("reference_baseline"))
    target_profile = as_dict(packet.get("target_profile"))
    electrical = as_dict(target_profile.get("electrical"))
    physical = as_dict(target_profile.get("physical"))
    reference_anchor = as_dict(analysis.get("reference_anchor_context"))

    anchor_context = " ".join(
        normalize_text(value)
        for value in [
            identity.get("sunco_reference_sku"),
            reference.get("title"),
            reference.get("product_type"),
        ]
        if normalize_text(value)
    )
    target_context = context_text_for_packet(packet)

    differences: list[str] = []
    target_lumens = number_from_text(electrical.get("lumens_target"), r"(?:lm|lumens?)")
    anchor_lumens = number_from_text(anchor_context, r"(?:lm|lumens?)")
    if target_lumens and anchor_lumens and target_lumens != anchor_lumens:
        direction = "higher" if target_lumens > anchor_lumens else "lower"
        differences.append(f"{direction} light output ({format_decimal(target_lumens, 0)}lm vs {format_decimal(anchor_lumens, 0)}lm anchor)")

    target_wattage = first_wattage_token(target_context)
    anchor_wattage = first_wattage_token(anchor_context)
    if target_wattage and anchor_wattage and target_wattage != anchor_wattage:
        differences.append(f"different load class ({target_wattage} target vs {anchor_wattage} anchor)")

    target_cct = cct_display_from_context(target_context)
    anchor_cct = cct_display_from_context(anchor_context)
    if target_cct and anchor_cct and target_cct != anchor_cct:
        cct_phrase = "selectable CCT" if as_dict(target_profile.get("electrical")).get("selectable_cct") else target_cct
        differences.append(f"{cct_phrase} vs {anchor_cct} anchor")

    feature_watchlist = [normalize_text(item) for item in as_list(target_profile.get("feature_watchlist")) if normalize_text(item)]
    if feature_watchlist:
        differences.append(f"feature watchlist: {', '.join(feature_watchlist[:3])}")

    target_price = as_dict(analysis.get("pricing_analysis")).get("target_msrp")
    anchor_price = reference.get("listing_price")
    if target_price and anchor_price:
        target_number = decimal_from_value(target_price)
        anchor_number = decimal_from_value(anchor_price)
        if target_number and anchor_number and target_number != anchor_number:
            direction = "premium" if target_number > anchor_number else "value"
            differences.append(f"{direction} price position ({money_or_blank(target_price)} target vs {money_or_blank(anchor_price)} anchor)")

    if differences:
        return "; ".join(differences[:4]) + "."

    role = reference_anchor.get("primary_use")
    if role:
        return f"Closest Sunco item is mainly a schema anchor; exact spec delta is limited in the current packet. {normalize_text(role)}"
    return "No close Sunco/NSL anchor was available in the current packet."


def product_fit_classification(packet: dict[str, Any]) -> str:
    """Return the Step 1-3 shared opportunity type."""
    return opportunity_type_for(packet)


def related_sunco_sales_trend(reference: dict[str, Any]) -> str:
    """Return related Sunco sales trend when the packet has trend fields."""
    for key in ["sales_growth_pct", "revenue_growth_pct", "units_growth_pct", "shopify_revenue_growth_pct", "amazon_revenue_growth_pct"]:
        value = reference.get(key)
        number = decimal_from_value(value)
        if number is None:
            continue
        if number > Decimal("5"):
            return f"Growing ({display_value_for_label('Growth %', value)})"
        if number < Decimal("-5"):
            return f"Declining ({display_value_for_label('Growth %', value)})"
        return f"Flat ({display_value_for_label('Growth %', value)})"

    shopify = reference.get("shopify_revenue_12mo")
    amazon = reference.get("amazon_revenue_12mo")
    if shopify is not None or amazon is not None:
        return (
            "Trend not available in current packet; use 12mo anchor revenue only "
            f"(Shopify {money_or_blank(shopify) or 'n/a'}, Amazon {money_or_blank(amazon) or 'n/a'})."
        )
    return "Trend not available in current packet."


def related_sales_caution(reference: dict[str, Any]) -> str:
    """Explain how to treat weak or missing related-Sunco sales support."""
    trend = related_sunco_sales_trend(reference)
    lowered = trend.lower()
    if "declining" in lowered:
        return "Related anchor appears declining; use it as category context, not as proof of upside, until the decline driver is known."
    if "trend not available" in lowered:
        return "The report has 12mo anchor volume but no monthly trend field, so it should not claim growing or declining related Sunco sales yet."
    return "No weak/declining related-sales caution surfaced from available packet fields."


def related_sales_slide_note(reference: dict[str, Any]) -> str:
    """Return a short related-sales note that fits in the slide table."""
    trend = related_sunco_sales_trend(reference)
    lowered = trend.lower()
    if "trend not available" in lowered:
        return "Trend unavailable; 12mo context only."
    if "declining" in lowered:
        return "Declining; explain cause before using as support."
    return trend


def market_momentum_summary(packet: dict[str, Any], analysis: dict[str, Any]) -> str:
    """Build a concise market-momentum statement."""
    performance = as_dict(analysis.get("performance_estimation"))
    snapshot = as_dict(performance.get("market_snapshot"))
    evidence = extract_research_note_evidence(packet)
    pieces: list[str] = []

    segment = normalize_text(snapshot.get("segment_name"))
    sales = snapshot.get("segment_retail_sales")
    sales_growth = snapshot.get("segment_sales_growth_pct")
    traffic_growth = snapshot.get("segment_traffic_growth_pct")
    if traffic_growth is None:
        for channel_data in as_dict(as_dict(performance.get("channel_comparison")).get("channels")).values():
            channel_data = as_dict(channel_data)
            if channel_data.get("traffic_growth_pct") is not None:
                traffic_growth = channel_data.get("traffic_growth_pct")
                break
    if segment:
        market_text = f"{segment}"
        if sales is not None:
            market_text += f" {display_value_for_label('Sales', sales)} retail sales"
        if sales_growth is not None:
            market_text += f", {display_value_for_label('Growth %', sales_growth)} sales growth"
        pieces.append(market_text)

    movement = evidence.get("ecommerce_movement")
    if movement:
        pieces.append(f"Redshift ecommerce movement: {movement}")
    if traffic_growth is not None:
        pieces.append(f"Stackline traffic growth {display_value_for_label('Growth %', traffic_growth)}")
    return "; ".join(pieces) + "." if pieces else "Market momentum is directional; no structured movement summary was available."


def channel_strategy(packet: dict[str, Any], analysis: dict[str, Any]) -> str:
    """Recommend a channel posture from available evidence."""
    identity = as_dict(packet.get("identity"))
    performance = as_dict(analysis.get("performance_estimation"))
    evidence = extract_research_note_evidence(packet)
    name = normalize_text(identity.get("ideation_name")).lower()
    outlook = normalize_text(performance.get("launch_outlook")).lower()
    notes = normalize_text(as_dict(packet.get("target_profile")).get("research_notes")).lower()
    amazon_snapshot = gate_snapshot(as_dict(analysis.get("gate_readiness")), "amazon", "G2")
    amazon_score = decimal_from_value(amazon_snapshot.get("weighted_score"))

    if "shopify/front-end launch review first" in notes:
        return "Shopify/front-end first; Amazon follow-up if Stackline, pack economics, and margin support."
    if "amazon stackline" in name:
        return "Amazon-first; add Shopify if assortment and margin story remain clean."
    if outlook == "favorable" and amazon_score is not None and amazon_score >= Decimal("7") and evidence.get("ecommerce_movement"):
        return "Amazon + Shopify"
    if outlook in {"mixed", "cautious"}:
        return "Staged test"
    return "Both channels if pricing and certification requirements stay clean."


def certification_summary(packet: dict[str, Any], analysis: dict[str, Any]) -> str:
    """Return required certifications from Step 2 or observed Section F coverage."""
    business = as_dict(as_dict(packet.get("target_profile")).get("business_case"))
    explicit = normalize_text(business.get("certifications"))
    if explicit:
        return explicit
    cert_rows = as_list(as_dict(analysis.get("spec_coverage")).get("certification_coverage"))
    labels = [normalize_text(as_dict(row).get("label")) for row in cert_rows if normalize_text(as_dict(row).get("label"))]
    if labels:
        return ", ".join(labels[:4])
    return "No required certification captured in current Step 2 row."


def step1b_evidence_rows(packet: dict[str, Any], analysis: dict[str, Any]) -> list[list[Any]]:
    """Build source-backed rows for the Step 1B demand evidence block."""
    performance = as_dict(analysis.get("performance_estimation"))
    snapshot = as_dict(performance.get("market_snapshot"))
    evidence = extract_research_note_evidence(packet)
    rows: list[list[Any]] = []

    if evidence.get("ecommerce_movement"):
        rows.append(["Ecommerce inventory movement", evidence["ecommerce_movement"], "Redshift ecommerce competitor snapshot", ""])
    if evidence.get("inventory_support"):
        rows.append(["Matched inventory support", evidence["inventory_support"], "Redshift ecommerce competitor snapshot", ""])
    if evidence.get("competitor_evidence"):
        rows.append(["Competitor PDP coverage", evidence["competitor_evidence"], "Redshift ecommerce competitor snapshot", ""])
    if evidence.get("review_link"):
        rows.append(["Primary competitor PDP", evidence["review_link"], "Step 1B verified review link", {"value": "Open listing", "hyperlink": evidence["review_link"]}])
    if evidence.get("sunco_coverage"):
        rows.append(["Sunco active-catalog coverage", evidence["sunco_coverage"], "Postgres product/catalog snapshot", ""])
    if snapshot:
        rows.append(
            [
                "Stackline market segment",
                market_momentum_summary(packet, analysis),
                normalize_text(snapshot.get("segment_name")) or "Stackline segment",
                "",
            ]
        )
    if evidence.get("vendor_bias_risk"):
        rows.append(["Evidence caveat", f"Vendor bias risk: {evidence['vendor_bias_risk']}", "Step 1B demand overlay", ""])
    return rows


def leadership_brief_rows(
    row_number: int,
    packet: dict[str, Any],
    analysis: dict[str, Any],
    normalized: dict[str, Any],
    action_summary: dict[str, Any],
) -> list[tuple[str, Any]]:
    """Build the screenshot-friendly top-of-sheet leadership brief."""
    identity = as_dict(packet.get("identity"))
    reference = as_dict(packet.get("reference_baseline"))
    performance = as_dict(analysis.get("performance_estimation"))
    pricing = as_dict(analysis.get("pricing_analysis"))
    normalized_summary = as_dict(normalized.get("summary"))
    amazon_snapshot = gate_snapshot(as_dict(analysis.get("gate_readiness")), "amazon", "G2")
    decision = "Pursue" if normalize_text(performance.get("launch_outlook")).lower() == "favorable" else "Stage test / refine"
    if normalize_text(performance.get("launch_outlook")).lower() == "cautious":
        decision = "Revise before pursuit"

    return [
        ("Decision Read", f"{decision} - {normalize_text(action_summary.get('overall_read'))}"),
        ("Concept Tracking Name", concept_tracking_name(packet)),
        ("Naming Basis", "SKU Decoder product type + Step 2 target specs; tracking name only, not a final SKU."),
        ("Opportunity Type", product_fit_classification(packet)),
        ("Channel Strategy", channel_strategy(packet, analysis)),
        ("Closest Sunco / NSL Anchor", f"{normalize_text(identity.get('sunco_reference_sku'))} - {normalize_text(reference.get('title'))}"),
        ("Difference vs Current Sunco", product_difference_summary(packet, analysis)),
        ("Related Sunco Sales Trend", related_sunco_sales_trend(reference)),
        ("Related Sales Caution", related_sales_caution(reference)),
        ("Why Worth Reviewing Now", market_momentum_summary(packet, analysis)),
        ("Target MSRP", pricing.get("target_msrp")),
        ("Target Vendor Cost", pricing.get("target_vendor_cost")),
        ("Amazon G2 / Evidence", f"{display_value_for_label('Score', amazon_snapshot.get('weighted_score')) or 'n/a'} / {normalize_text(as_dict(amazon_snapshot.get('evidence_confidence')).get('label')) or 'n/a'}"),
        ("Competitor Records", f"{normalized_summary.get('verified_listing_count', 0)} verified listings; {normalized_summary.get('inferred_competitor_count', 0)} inferred competitors"),
        ("Required Certifications", certification_summary(packet, analysis)),
        ("Leadership Watchout", action_summary.get("why_not_stronger")),
    ]


def compact_competitor_read(item: dict[str, Any]) -> str:
    """Return a short competitor proof statement for PM review."""
    if not item:
        return "No direct competitor PDP was selected in this report."
    pieces = [
        normalize_text(item.get("brand")),
        listing_identifier(item),
        clean_slide_text(item.get("product_title"), 80),
    ]
    price = display_value_for_label("Price", item.get("price"))
    if price:
        pieces.append(price)
    return " | ".join(piece for piece in pieces if piece)


def pm_decision_snapshot_rows(
    row_number: int,
    packet: dict[str, Any],
    analysis: dict[str, Any],
    normalized: dict[str, Any],
    action_summary: dict[str, Any],
) -> list[list[Any]]:
    """Build the compact PM-facing decision table at the top of each row sheet."""
    identity = as_dict(packet.get("identity"))
    reference = as_dict(packet.get("reference_baseline"))
    pricing = as_dict(analysis.get("pricing_analysis"))
    opportunity_type = opportunity_type_for(packet, analysis)
    evidence_strength = evidence_strength_for(packet, analysis)
    competitor = primary_competitor_item(normalized)
    competitor_match = match_quality_for_item(competitor)
    competitor_link = source_link_from_item(competitor) if competitor else primary_source_link(packet, normalized)
    evidence = extract_research_note_evidence(packet)
    pricing_link = source_link_cell(evidence.get("review_link"), "Open price source") or competitor_link
    decision = "Pursue" if normalize_text(as_dict(analysis.get("performance_estimation")).get("launch_outlook")).lower() == "favorable" else "Stage test / refine"
    if normalize_text(as_dict(analysis.get("performance_estimation")).get("launch_outlook")).lower() == "cautious":
        decision = "Revise before pursuit"

    return [
        [
            "Concept Tracking Name",
            concept_tracking_name(packet),
            evidence_strength,
            f"Row {row_number}; tracking name is generated from SKU Decoder cache and Step 2 target specs.",
            "",
        ],
        [
            "Opportunity Type",
            opportunity_type,
            evidence_strength,
            "Uses the same Step 1-3 opportunity vocabulary so PMs do not need to relearn the classification at each step.",
            primary_source_link(packet, normalized),
        ],
        [
            "Decision Read",
            decision,
            evidence_strength,
            first_sentence(action_summary.get("overall_read")),
            "",
        ],
        [
            "Recommended MSRP",
            pricing.get("target_msrp"),
            evidence_strength,
            msrp_basis_summary(packet, analysis),
            pricing_link,
        ],
        [
            "Target Vendor Cost",
            pricing.get("target_vendor_cost"),
            evidence_strength,
            "Landed cost ceiling from the 50-55% pre-ads gross-margin policy midpoint.",
            "",
        ],
        [
            "Sunco Gap",
            opportunity_type,
            evidence_strength,
            evidence.get("sunco_coverage") or product_difference_summary(packet, analysis),
            primary_source_link(packet, normalized),
        ],
        [
            "Closest Sunco / NSL Anchor",
            f"{normalize_text(identity.get('sunco_reference_sku'))} - {normalize_text(reference.get('title'))}".strip(" -"),
            evidence_strength,
            related_sunco_sales_trend(reference),
            "",
        ],
        [
            "Competitor Proof",
            compact_competitor_read(competitor),
            evidence_strength,
            f"{competitor_match}; {item_demand_signal(competitor) if competitor else normalize_text(evidence.get('ecommerce_movement'))}",
            competitor_link,
        ],
        [
            "Link Status",
            link_status_for_item(competitor) if competitor else link_status_for_url(evidence.get("review_link"), verified=True),
            evidence_strength,
            "PMs should use this visible link to manually confirm product, price, and attributes before presenting.",
            competitor_link,
        ],
        [
            "Channel Strategy",
            channel_strategy(packet, analysis),
            evidence_strength,
            market_momentum_summary(packet, analysis),
            primary_source_link(packet, normalized),
        ],
        [
            "Main Watchout",
            first_sentence(action_summary.get("why_not_stronger")),
            EVIDENCE_STRENGTH_REVIEW if action_summary.get("why_not_stronger") else evidence_strength,
            "Use this as the first follow-up question if the concept moves toward vendor quote or sample work.",
            "",
        ],
    ]


def write_pm_decision_snapshot(
    ws,
    row: int,
    row_number: int,
    packet: dict[str, Any],
    analysis: dict[str, Any],
    normalized: dict[str, Any],
    action_summary: dict[str, Any],
) -> int:
    """Write the compact PM decision snapshot at the top of the row sheet."""
    return write_table(
        ws,
        row,
        "PM Decision Snapshot",
        ["Field", "Recommendation", "Evidence Strength", "PM Read / Evidence", "Source Link"],
        pm_decision_snapshot_rows(row_number, packet, analysis, normalized, action_summary),
    )


DECISION_EVIDENCE_HEADERS = [
    "Opportunity Type",
    "Evidence Strength",
    "Match Quality",
    "Evidence",
    "Source Link",
    "PM Read",
]


def top_verified_items(normalized: dict[str, Any], channels: set[str], limit: int) -> list[dict[str, Any]]:
    """Return a small, PM-readable set of verified competitor items."""
    rows: list[dict[str, Any]] = []
    for item in sort_candidates(as_list(normalized.get("items"))):
        if resolved_source_channel(item) not in channels:
            continue
        if verification_status(item) != "verified_listing":
            continue
        rows.append(item)
        if len(rows) >= limit:
            break
    return rows


def evidence_row(
    packet: dict[str, Any],
    analysis: dict[str, Any],
    match_quality: str,
    evidence: Any,
    source_link: Any,
    pm_read: Any,
) -> list[Any]:
    """Build a standardized evidence row using shared Step 1-3 vocabulary."""
    return [
        opportunity_type_for(packet, analysis),
        evidence_strength_for(packet, analysis),
        match_quality,
        evidence,
        source_link,
        pm_read,
    ]


def demand_proof_rows(packet: dict[str, Any], analysis: dict[str, Any], normalized: dict[str, Any]) -> list[list[Any]]:
    """Summarize the strongest demand evidence without exposing every raw field."""
    evidence = extract_research_note_evidence(packet)
    performance = as_dict(analysis.get("performance_estimation"))
    snapshot = as_dict(performance.get("market_snapshot"))
    rows: list[list[Any]] = []
    if evidence.get("ecommerce_movement"):
        rows.append(
            evidence_row(
                packet,
                analysis,
                MATCH_QUALITY_SIMILAR,
                evidence["ecommerce_movement"],
                primary_source_link(packet, normalized),
                "Competitor inventory movement supports a demand read, but PM should confirm the linked PDP matches the exact target spec.",
            )
        )
    if snapshot.get("segment_name"):
        rows.append(
            evidence_row(
                packet,
                analysis,
                MATCH_QUALITY_SIMILAR,
                market_momentum_summary(packet, analysis),
                "",
                "Use Stackline as category/segment demand support, not as a replacement for exact PDP matching.",
            )
        )
    if evidence.get("inventory_support"):
        rows.append(
            evidence_row(
                packet,
                analysis,
                MATCH_QUALITY_SIMILAR,
                evidence["inventory_support"],
                primary_source_link(packet, normalized),
                "Inventory support is a directional movement signal for the normalized spec cluster.",
            )
        )
    if not rows:
        rows.append(
            evidence_row(
                packet,
                analysis,
                MATCH_QUALITY_NOT_COMPARABLE,
                "No structured demand proof was captured for this row.",
                "",
                "Treat as a PM review item before using demand as the main selling point.",
            )
        )
    return rows[:3]


def sunco_coverage_decision_rows(packet: dict[str, Any], analysis: dict[str, Any], normalized: dict[str, Any]) -> list[list[Any]]:
    """Summarize Sunco coverage and gap evidence."""
    identity = as_dict(packet.get("identity"))
    reference = as_dict(packet.get("reference_baseline"))
    evidence = extract_research_note_evidence(packet)
    anchor = f"{normalize_text(identity.get('sunco_reference_sku'))} - {normalize_text(reference.get('title'))}".strip(" -")
    return [
        evidence_row(
            packet,
            analysis,
            MATCH_QUALITY_SIMILAR if anchor else MATCH_QUALITY_NOT_COMPARABLE,
            evidence.get("sunco_coverage") or product_difference_summary(packet, analysis),
            primary_source_link(packet, normalized),
            "This is the main line-review read: confirm whether the gap is a new SKU, variant depth, or merchandising/revision work.",
        ),
        evidence_row(
            packet,
            analysis,
            MATCH_QUALITY_SIMILAR if anchor else MATCH_QUALITY_NOT_COMPARABLE,
            anchor or "No close Sunco/NSL anchor captured.",
            "",
            related_sales_caution(reference),
        ),
    ]


def competitor_decision_rows(packet: dict[str, Any], analysis: dict[str, Any], normalized: dict[str, Any]) -> list[list[Any]]:
    """Summarize the top competitor comparables with visible links."""
    rows: list[list[Any]] = []
    for item in top_verified_items(normalized, AMAZON_CHANNELS, 2) + top_verified_items(normalized, BM_DIRECT_CHANNELS, 2):
        rows.append(
            evidence_row(
                packet,
                analysis,
                match_quality_for_item(item),
                compact_competitor_read(item),
                source_link_from_item(item),
                f"{link_status_for_item(item)}; {item_demand_signal(item) or 'Use the linked PDP to verify product attributes.'}",
            )
        )
    if not rows:
        rows.append(
            evidence_row(
                packet,
                analysis,
                MATCH_QUALITY_NOT_COMPARABLE,
                "No verified competitor comparables were captured.",
                primary_source_link(packet, normalized),
                "Do not present this as apples-to-apples until a comparable PDP is verified.",
            )
        )
    return rows


def pricing_channel_decision_rows(packet: dict[str, Any], analysis: dict[str, Any], normalized: dict[str, Any]) -> list[list[Any]]:
    """Summarize MSRP, cost, and channel positioning."""
    pricing = as_dict(analysis.get("pricing_analysis"))
    evidence = extract_research_note_evidence(packet)
    price_link = source_link_cell(evidence.get("review_link"), "Open price source") or primary_source_link(packet, normalized)
    return [
        evidence_row(
            packet,
            analysis,
            MATCH_QUALITY_SIMILAR,
            f"Recommended MSRP: {display_value_for_label('Recommended MSRP', pricing.get('target_msrp'))}",
            price_link,
            msrp_basis_summary(packet, analysis),
        ),
        evidence_row(
            packet,
            analysis,
            MATCH_QUALITY_SIMILAR,
            f"Target vendor cost: {display_value_for_label('Target Vendor Cost', pricing.get('target_vendor_cost'))}",
            "",
            "Backsolved as landed-cost ceiling from the 50-55% pre-ads gross-margin policy midpoint.",
        ),
        evidence_row(
            packet,
            analysis,
            MATCH_QUALITY_SIMILAR,
            channel_strategy(packet, analysis),
            primary_source_link(packet, normalized),
            "Use this as the launch-channel recommendation unless PM review finds a better channel constraint.",
        ),
    ]


def risk_watchout_rows(packet: dict[str, Any], analysis: dict[str, Any], action_summary: dict[str, Any]) -> list[list[Any]]:
    """Summarize the highest-value PM watchouts."""
    rows: list[list[Any]] = []
    for action in as_list(action_summary.get("actions"))[:4]:
        action_row = as_list(action)
        issue = action_row[1] if len(action_row) > 1 else ""
        why = action_row[2] if len(action_row) > 2 else ""
        recommendation = action_row[3] if len(action_row) > 3 else ""
        rows.append(
            evidence_row(
                packet,
                analysis,
                MATCH_QUALITY_SIMILAR,
                f"{issue}: {why}",
                "",
                recommendation,
            )
        )
    if not rows:
        rows.append(
            evidence_row(
                packet,
                analysis,
                MATCH_QUALITY_SIMILAR,
                "No major watchout was generated from the current analysis.",
                "",
                "Protect the assumptions as quotes, samples, certifications, and channel details firm up.",
            )
        )
    return rows


def write_leadership_brief(
    ws,
    row: int,
    row_number: int,
    packet: dict[str, Any],
    analysis: dict[str, Any],
    normalized: dict[str, Any],
    action_summary: dict[str, Any],
) -> int:
    """Write a compact leadership brief designed for screenshotting."""
    row = section_header(ws, row, "Leadership Brief", end_column=10)
    row = key_value_rows(
        ws,
        row,
        leadership_brief_rows(row_number, packet, analysis, normalized, action_summary),
        columns=2,
    )
    return row + 1


def merged_text_row(ws, row: int, label: str, value: Any) -> int:
    """Write one labeled merged text row."""
    ws.cell(row=row, column=1, value=label)
    ws.cell(row=row, column=1).font = SUBHEADER_FONT
    ws.cell(row=row, column=1).fill = SUBHEADER_FILL
    ws.cell(row=row, column=2, value=normalize_text(value))
    ws.cell(row=row, column=2).alignment = WRAP_ALIGNMENT
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=10)
    return row + 1


def write_list_section(ws, row: int, title: str, values: list[str]) -> int:
    """Write a simple bulleted list section."""
    row = section_header(ws, row, title)
    if not values:
        ws.cell(row=row, column=1, value="No applicable items for this section.")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        return row + 2
    for value in values:
        ws.cell(row=row, column=1, value=f"- {normalize_text(value)}")
        ws.cell(row=row, column=1).alignment = WRAP_ALIGNMENT
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        row += 1
    return row + 1


def write_table(ws, row: int, title: str, headers: list[str], rows: list[list[Any]]) -> int:
    """Write a basic table and return the next row index."""
    row = section_header(ws, row, title, end_column=max(10, len(headers)))
    for col_index, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col_index, value=header)
        cell.font = SUBHEADER_FONT
        cell.fill = SUBHEADER_FILL
        cell.alignment = WRAP_ALIGNMENT
    row += 1

    if not rows:
        ws.cell(row=row, column=1, value="No applicable rows for this section.")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
        return row + 2

    for row_offset, values in enumerate(rows):
        for col_index, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col_index)
            header = headers[col_index - 1] if col_index <= len(headers) else ""
            if len(headers) >= 2 and headers[0] == "Metric" and headers[1] == "Value" and col_index == 2 and values:
                header = normalize_text(values[0])
            elif len(headers) >= 2 and headers[0] == "Field" and headers[1] == "Recommendation" and col_index == 2 and values:
                header = normalize_text(values[0])
            elif header:
                header = f"{title} {header}"
            apply_table_cell(cell, value, header)
            cell.alignment = WRAP_ALIGNMENT
            if row_offset % 2 == 1:
                cell.fill = ALTERNATE_ROW_FILL
        row += 1
    return row + 1


def apply_table_cell(cell, value: Any, label: str = "") -> None:
    """Write a table cell, optionally attaching a hyperlink."""
    hyperlink = None
    text_value = value
    if isinstance(value, dict):
        text_value = value.get("value")
        hyperlink = value.get("hyperlink")

    cell.value = display_value_for_label(label, text_value)
    if hyperlink:
        cell.hyperlink = hyperlink
        cell.font = HYPERLINK_FONT


def optional_text(value: Any) -> str | None:
    """Return a stripped string or None."""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def looks_like_amazon_asin(value: str | None) -> bool:
    """Return whether a value looks like an Amazon ASIN."""
    return bool(value and re.fullmatch(r"[A-Z0-9]{10}", value.upper()))


def source_channel_label(item: dict[str, Any]) -> str:
    """Render a human-friendly source channel label."""
    channel = resolved_source_channel(item)
    return CHANNEL_DISPLAY_NAMES.get(channel, channel.replace("_", " ").title() or "Unknown")


def discovery_source_label(item: dict[str, Any]) -> str:
    """Render the original discovery source for inferred competitors."""
    domain = optional_text(item.get("discovery_source_domain"))
    if domain:
        host = urlparse(domain if "://" in domain else f"//{domain}").netloc or domain
        return host.replace("www.", "")
    channel = optional_text(item.get("discovery_source_channel"))
    if channel:
        lowered = channel.lower()
        return CHANNEL_DISPLAY_NAMES.get(lowered, lowered.replace("_", " ").title())
    domain = optional_text(item.get("source_domain"))
    if domain:
        host = urlparse(domain if "://" in domain else f"//{domain}").netloc or domain
        return host.replace("www.", "")
    return source_channel_label(item)


def listing_identifier_from_url(url: str | None, source_channel: str | None) -> str | None:
    """Best-effort identifier fallback derived from the listing URL."""
    if not url:
        return None

    channel = (source_channel or "").lower()
    patterns = []
    if channel == "amazon":
        patterns = [(r"/dp/([A-Z0-9]{10})(?:[/?]|$)", "ASIN {match}")]
    elif channel == "home_depot":
        patterns = [(r"/p/(?:[^/]+/)?(\d+)(?:[/?]|$)", "Item {match}")]
    elif channel == "walmart":
        patterns = [(r"/ip/(?:[^/]+/)?(\d+)(?:[/?]|$)", "Item {match}")]
    elif channel == "lowes":
        patterns = [(r"/pd/[^/]+/(\d+)(?:[/?]|$)", "Item {match}")]

    for pattern, template in patterns:
        match = re.search(pattern, url, flags=re.IGNORECASE)
        if match:
            return template.format(match=match.group(1))
    return None


def listing_identifier(item: dict[str, Any]) -> str:
    """Return the best channel-aware identifier label for one competitor listing."""
    source_channel = resolved_source_channel(item)
    sku = optional_text(item.get("sku"))
    model_number = optional_text(item.get("model_number"))
    url = optional_text(item.get("url"))

    if source_channel == "amazon":
        asin = sku if looks_like_amazon_asin(sku) else None
        if not asin and looks_like_amazon_asin(model_number):
            asin = model_number
        if not asin and url:
            derived = listing_identifier_from_url(url, source_channel)
            if derived and derived.startswith("ASIN "):
                asin = derived.replace("ASIN ", "", 1)
        if asin:
            return f"ASIN {asin}"
        if model_number:
            return f"Model {model_number}"

    if model_number:
        return f"Model {model_number}"

    if sku:
        if source_channel == "brand_site":
            return f"Part {sku}"
        if source_channel in {"home_depot", "walmart", "lowes"}:
            return f"Item {sku}"
        return f"SKU {sku}"

    derived = listing_identifier_from_url(url, source_channel)
    return derived or ""


def listing_link_cell(item: dict[str, Any]) -> dict[str, str] | str:
    """Return a hyperlink cell payload for a competitor source listing."""
    url = optional_text(item.get("url"))
    if not url or url.startswith("stackline://"):
        return ""
    return {
        "value": source_channel_label(item),
        "hyperlink": url,
    }


def verification_status(item: dict[str, Any]) -> str:
    """Return the normalized verification status for one competitor record."""
    status = (optional_text(item.get("verification_status")) or "").lower()
    if status in {"verified_listing", "inferred_competitor"}:
        return status
    return "verified_listing"


def verification_label(item: dict[str, Any]) -> str:
    """Render a user-facing verification label."""
    return "Verified Listing" if verification_status(item) == "verified_listing" else "Inferred Competitor"


def sort_candidates(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Sort candidate records by confidence, then by price."""
    def sort_key(item: dict[str, Any]) -> tuple[float, float, str]:
        confidence = item.get("match_confidence")
        if not isinstance(confidence, (int, float)):
            confidence = 0
        price = item.get("price")
        if not isinstance(price, (int, float)):
            price = 0
        title = normalize_text(item.get("product_title"))
        return (-confidence, price, title)

    return sorted(items, key=sort_key)


def candidate_rows(
    normalized_items: list[dict[str, Any]],
    channels: set[str],
    limit: int = 10,
    verification_filter: str | None = "verified_listing",
) -> list[list[Any]]:
    """Build a compact competitor table filtered by source channel."""
    rows = []
    for item in sort_candidates(normalized_items):
        if resolved_source_channel(item) not in channels:
            continue
        if verification_filter and verification_status(item) != verification_filter:
            continue
        rows.append(
            [
                item.get("brand"),
                item.get("product_title"),
                listing_identifier(item),
                source_channel_label(item),
                item.get("price"),
                item.get("wattage"),
                item.get("lumens"),
                item.get("cct"),
                item.get("cri"),
                listing_link_cell(item),
                item.get("match_confidence"),
            ]
        )
        if len(rows) >= limit:
            break
    return rows


def item_demand_signal(item: dict[str, Any]) -> str:
    """Return movement or demand evidence attached to one competitor item."""
    observations = [normalize_text(value) for value in as_list(item.get("raw_observations")) if normalize_text(value)]
    if observations:
        return "; ".join(observations[:3])

    signal_keys = [
        "inventory_movement",
        "stock_decrease_units",
        "stock_decrease_events",
        "velocity_units_per_week",
        "sales_rank",
        "units_sold",
        "sales_share_pct",
        "review_count",
        "rating",
    ]
    pieces = []
    for key in signal_keys:
        value = item.get(key)
        if value is not None and value != "":
            label = key.replace("_", " ").title()
            pieces.append(f"{label}: {normalize_text(value)}")
    return "; ".join(pieces[:3])


def item_attribute_summary(item: dict[str, Any]) -> str:
    """Return compact attribute evidence for a competitor item."""
    parts: list[str] = []
    for label, key in [
        ("Voltage", "voltage"),
        ("Pack", "pack_quantity"),
        ("Certs", "certifications"),
        ("Features", "features"),
        ("Dimming", "dimming"),
        ("Mount", "mounting_type"),
        ("Finish", "finish_color"),
    ]:
        value = item.get(key)
        if value is not None and value != "" and value != []:
            parts.append(f"{label}: {normalize_text(value)}")
    return "; ".join(parts[:4])


def detailed_candidate_rows(
    normalized_items: list[dict[str, Any]],
    channels: set[str],
    limit: int = 10,
    verification_filter: str | None = "verified_listing",
) -> list[list[Any]]:
    """Build a competitor table that includes movement and attribute evidence."""
    rows = []
    for item in sort_candidates(normalized_items):
        if resolved_source_channel(item) not in channels:
            continue
        if verification_filter and verification_status(item) != verification_filter:
            continue
        rows.append(
            [
                item.get("brand"),
                item.get("product_title"),
                listing_identifier(item),
                source_channel_label(item),
                item.get("price"),
                item_attribute_summary(item),
                item_demand_signal(item),
                match_quality_for_item(item),
                listing_link_cell(item),
                link_status_for_item(item),
                discovery_source_label(item),
            ]
        )
        if len(rows) >= limit:
            break
    return rows


def inferred_competitor_rows(
    normalized_items: list[dict[str, Any]],
    limit: int = 12,
) -> list[list[Any]]:
    """Build a table of inferred competitors that need listing verification."""
    rows = []
    for item in sort_candidates(normalized_items):
        if verification_status(item) != "inferred_competitor":
            continue
        rows.append(
            [
                item.get("brand"),
                item.get("product_title"),
                source_channel_label(item),
                discovery_source_label(item),
                listing_link_cell(item),
                link_status_for_item(item),
                item.get("verification_reason") or item.get("extraction_notes") or item.get("match_notes"),
                item.get("match_confidence"),
            ]
        )
        if len(rows) >= limit:
            break
    return rows


def top_brand_rows(brands: list[dict[str, Any]]) -> list[list[Any]]:
    """Format summarized brand rows."""
    rows = []
    for brand in brands:
        rows.append(
            [
                brand.get("brand"),
                brand.get("candidate_count"),
                normalize_text(brand.get("source_channels")),
                brand.get("median_unit_price"),
            ]
        )
    return rows


def coverage_rows(entries: list[dict[str, Any]]) -> list[list[Any]]:
    """Format feature/certification coverage rows."""
    rows = []
    for entry in entries:
        rows.append([entry.get("label"), entry.get("matched_count"), entry.get("coverage_pct")])
    return rows


def benchmark_rows(pricing: dict[str, Any]) -> list[list[Any]]:
    """Format multi-metric pricing benchmark rows."""
    metrics = [
        ("Raw Price", as_dict(pricing.get("price_benchmarks"))),
        ("Unit Price", as_dict(pricing.get("unit_price_benchmarks"))),
        ("Unit Price / Watt", as_dict(pricing.get("unit_price_per_watt_benchmarks"))),
        ("Unit Price / Lumen", as_dict(pricing.get("unit_price_per_lumen_benchmarks"))),
    ]
    rows = []
    for label, metric in metrics:
        rows.append(
            [
                label,
                metric.get("sample_size"),
                metric.get("min"),
                metric.get("p25"),
                metric.get("median"),
                metric.get("mean"),
                metric.get("p75"),
                metric.get("max"),
            ]
        )
    return rows


def pricing_position_rows(pricing: dict[str, Any]) -> list[list[Any]]:
    """Format pricing positioning rows."""
    suggested = as_dict(pricing.get("suggested_msrp_range"))
    target_position = as_dict(pricing.get("target_price_position"))
    return [
        ["Target MSRP", pricing.get("target_msrp")],
        ["Evaluated Price", target_position.get("evaluated_value")],
        ["Evaluated Price Source", target_position.get("evaluated_value_source")],
        ["Target Price Percentile", target_position.get("percentile")],
        ["Target Price Bucket", target_position.get("bucket")],
        ["Target vs Market Median %", target_position.get("vs_median_pct")],
        ["Observed Unit Price Floor (P25)", suggested.get("observed_unit_price_floor")],
        ["Observed Unit Price Ceiling (P75)", suggested.get("observed_unit_price_ceiling")],
        ["Recommended Floor", suggested.get("recommended_floor")],
        ["Recommended Ceiling", suggested.get("recommended_ceiling")],
        ["Minimum Margin-Safe MSRP", suggested.get("minimum_margin_safe_price")],
        ["Suggested Positioning", suggested.get("positioning")],
        ["Margin Conflict", suggested.get("margin_conflict")],
    ]


def margin_rows(pricing: dict[str, Any]) -> list[list[Any]]:
    """Format channel-specific margin guidance rows."""
    rows = []
    for channel in ["shopify", "amazon"]:
        entry = as_dict(as_dict(pricing.get("margin_targets")).get(channel))
        if not entry:
            continue
        rows.append(
            [
                channel.title(),
                entry.get("target_margin_pct"),
                entry.get("minimum_viable_msrp"),
                entry.get("vs_target_msrp_pct"),
                entry.get("vs_market_median_pct"),
            ]
        )
    return rows


def value_position_rows(pricing: dict[str, Any]) -> list[list[Any]]:
    """Format value-ranking rows for unit price, price per watt, and price per lumen."""
    rows = []
    metrics = [
        ("Unit Price", as_dict(pricing.get("target_price_position"))),
        ("Unit Price / Watt", as_dict(pricing.get("target_price_per_watt_position"))),
        ("Unit Price / Lumen", as_dict(pricing.get("target_price_per_lumen_position"))),
    ]
    for label, metric in metrics:
        if not metric:
            continue
        rows.append(
            [
                label,
                metric.get("evaluated_value"),
                metric.get("percentile"),
                metric.get("bucket"),
                metric.get("vs_median_pct"),
            ]
        )
    return rows


def channel_comparison_rows(performance: dict[str, Any]) -> list[list[Any]]:
    """Format Stackline channel comparison rows for Section A."""
    comparison = as_dict(performance.get("channel_comparison"))
    channels = as_dict(comparison.get("channels"))
    rows = []
    for channel_name, channel in channels.items():
        channel = as_dict(channel)
        rows.append(
            [
                channel_name,
                channel.get("retail_sales"),
                channel.get("units_sold"),
                channel.get("avg_retail_price"),
                channel.get("retail_sales_growth_pct"),
                channel.get("sunco_sales_share_pct"),
            ]
        )
    return rows


def spec_action_rows(spec_coverage: dict[str, Any]) -> list[list[Any]]:
    """Format actionable feature/certification coverage rows."""
    rows = []
    for entry in as_list(spec_coverage.get("feature_coverage")):
        rows.append(
            [
                "Feature",
                entry.get("label"),
                entry.get("signal"),
                entry.get("evidence_strength"),
                entry.get("coverage_pct"),
                entry.get("matched_count"),
                entry.get("recommended_action"),
            ]
        )
    for entry in as_list(spec_coverage.get("certification_coverage")):
        rows.append(
            [
                "Certification",
                entry.get("label"),
                entry.get("signal"),
                entry.get("evidence_strength"),
                entry.get("coverage_pct"),
                entry.get("matched_count"),
                entry.get("recommended_action"),
            ]
        )
    return rows


def numeric_guidance_rows(spec_coverage: dict[str, Any]) -> list[list[Any]]:
    """Format numeric target-positioning rows."""
    rows = []
    for entry in as_list(spec_coverage.get("numeric_guidance")):
        rows.append(
            [
                entry.get("label"),
                entry.get("target_value"),
                entry.get("median"),
                entry.get("p75"),
                entry.get("target_percentile"),
                entry.get("recommended_action"),
            ]
        )
    return rows


def gate_readiness_snapshot_rows(gate_readiness: dict[str, Any]) -> list[list[Any]]:
    """Format gate/channel readiness rows for the report."""
    rows = []
    for snapshot in as_list(gate_readiness.get("snapshots")):
        evidence = as_dict(snapshot.get("evidence_confidence"))
        rows.append(
            [
                snapshot.get("channel"),
                snapshot.get("gate"),
                snapshot.get("family_state"),
                snapshot.get("weighted_score"),
                evidence.get("score"),
                evidence.get("label"),
                f"{evidence.get('implemented_questions')}/{evidence.get('methodology_active_questions')}",
            ]
        )
    return rows


def gate_readiness_pillar_rows(gate_readiness: dict[str, Any]) -> list[list[Any]]:
    """Format pillar-level rollups from the primary G2 channel snapshot."""
    primary_channel = normalize_text(gate_readiness.get("primary_channel"))
    for snapshot in as_list(gate_readiness.get("snapshots")):
        if snapshot.get("channel") == primary_channel and snapshot.get("gate") == "G2":
            rows = []
            for pillar in as_list(snapshot.get("pillar_scores")):
                rows.append(
                    [
                        pillar.get("label"),
                        pillar.get("base_weight"),
                        pillar.get("effective_weight"),
                        pillar.get("average_score"),
                        f"{pillar.get('scored_question_count')}/{pillar.get('question_count')}",
                        pillar.get("status"),
                    ]
                )
            return rows
    return []


def vendor_request_rows(items: list[dict[str, Any]]) -> list[list[Any]]:
    """Format vendor optimization requests."""
    rows = []
    for item in items:
        rows.append(
            [
                item.get("priority"),
                item.get("linked_metric"),
                item.get("request"),
                item.get("reason"),
            ]
        )
    return rows


def optimization_driver_rows(items: list[dict[str, Any]]) -> list[list[Any]]:
    """Format category-aware decision drivers for the report."""
    rows = []
    for item in items:
        rows.append(
            [
                item.get("tier"),
                item.get("label"),
                item.get("driver_type"),
                item.get("signal"),
                item.get("reason"),
            ]
        )
    return rows


def low_signal_rows(items: list[dict[str, Any]]) -> list[list[Any]]:
    """Format lower-signal attributes that should be validated before over-weighting."""
    rows = []
    for item in items:
        rows.append(
            [
                item.get("label"),
                item.get("driver_type"),
                item.get("signal"),
                item.get("reason"),
            ]
        )
    return rows


def optimization_modifier_rows(items: list[dict[str, Any]]) -> list[list[Any]]:
    """Format active optimization modifiers for the report."""
    rows = []
    for item in items:
        rows.append(
            [
                item.get("label"),
                ", ".join(as_list(item.get("matched_keywords"))),
                item.get("match_score"),
            ]
        )
    return rows


def optimization_scorecard_rows(scorecard: dict[str, Any]) -> list[list[Any]]:
    """Format optimization score components for the report."""
    rows = []
    for item in as_list(scorecard.get("components")):
        rows.append(
            [
                item.get("component"),
                item.get("score"),
                item.get("weight"),
                item.get("reason"),
            ]
        )
    return rows


def report_filename(row_number: int) -> str:
    """Return the stable report filename for a row."""
    return f"row_{row_number:03d}_research_report.xlsx"


def prd_prefill_pairs(packet: dict[str, Any], analysis: dict[str, Any]) -> list[tuple[str, Any]]:
    """Build the PRD-oriented prefill block from packet targets."""
    identity = as_dict(packet.get("identity"))
    target_profile = as_dict(packet.get("target_profile"))
    electrical = as_dict(target_profile.get("electrical"))
    physical = as_dict(target_profile.get("physical"))
    business = as_dict(target_profile.get("business_case"))
    reference = as_dict(packet.get("reference_baseline"))

    return [
        ("Ideation Name", normalize_text(identity.get("ideation_name"))),
        ("Reference Image URL", reference.get("image_url")),
        ("Voltage", electrical.get("voltage")),
        ("Wattage Primary", electrical.get("wattage_primary")),
        ("Wattage Max", electrical.get("wattage_max")),
        ("Selectable Wattage", electrical.get("selectable_wattage")),
        ("CCT Primary", electrical.get("cct_primary")),
        ("CCT Max", electrical.get("cct_max")),
        ("Selectable CCT", electrical.get("selectable_cct")),
        ("CRI", electrical.get("cri")),
        ("Lumens Target", electrical.get("lumens_target")),
        ("Dimmable", electrical.get("dimmable")),
        ("Dimming Type", electrical.get("dimming_type")),
        ("Size / Form Factor", physical.get("size_form_factor")),
        ("Mounting Type", physical.get("mounting_type")),
        ("Material", physical.get("material")),
        ("Finish / Color", physical.get("finish_color")),
        ("IP Rating", physical.get("ip_rating")),
        ("Moisture Rating", physical.get("moisture_rating")),
        ("Target MSRP", as_dict(analysis.get("pricing_analysis")).get("target_msrp")),
        ("Target Vendor Cost", as_dict(analysis.get("pricing_analysis")).get("target_vendor_cost")),
        ("Certifications", business.get("certifications")),
        ("Lifetime Hours", business.get("lifetime_hours")),
        ("Warranty", business.get("warranty")),
    ]


def render_row_sheet(
    ws,
    row_number: int,
    packet: dict[str, Any],
    analysis: dict[str, Any],
    normalized: dict[str, Any],
) -> None:
    """Render one ideation sheet into an existing worksheet."""
    set_default_layout(ws)

    identity = as_dict(packet.get("identity"))
    reference = as_dict(packet.get("reference_baseline"))
    performance = as_dict(analysis.get("performance_estimation"))
    pricing = as_dict(analysis.get("pricing_analysis"))
    summary = as_dict(analysis.get("summary"))
    spec_coverage = as_dict(analysis.get("spec_coverage"))
    reference_anchor = as_dict(analysis.get("reference_anchor_context"))
    gate_readiness = as_dict(analysis.get("gate_readiness"))
    vendor_requests = as_list(analysis.get("highest_impact_vendor_requests"))
    ideation_optimization = as_dict(analysis.get("ideation_optimization"))
    optimization_scorecard = as_dict(ideation_optimization.get("optimization_scorecard"))
    optimization_modifiers = as_list(ideation_optimization.get("active_modifiers"))
    normalized_summary = as_dict(normalized.get("summary"))
    action_summary = blocker_action_summary(packet, analysis)
    tracking_name = concept_tracking_name(packet)

    ws.cell(row=1, column=1, value=f"{tracking_name} | {normalize_text(identity.get('ideation_name'))}")
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).alignment = WRAP_ALIGNMENT
    ws.row_dimensions[1].height = 48
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    ws.cell(row=2, column=1, value=f"Row {row_number} Research Report | Concept Tracking Name generated from SKU Decoder cache")
    ws.cell(row=2, column=1).fill = ACCENT_FILL
    ws.cell(row=2, column=1).alignment = WRAP_ALIGNMENT
    ws.row_dimensions[2].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)

    row = 4
    row = write_pm_decision_snapshot(ws, row, row_number, packet, analysis, normalized, action_summary)
    row = write_table(
        ws,
        row,
        "PM Action Checklist",
        ["Priority", "Issue", "Why It Matters", "Recommended Action", "Expected Impact"],
        as_list(action_summary.get("actions")),
    )
    row = write_table(
        ws,
        row,
        "Demand Proof",
        DECISION_EVIDENCE_HEADERS,
        demand_proof_rows(packet, analysis, normalized),
    )
    row = write_table(
        ws,
        row,
        "Sunco Coverage",
        DECISION_EVIDENCE_HEADERS,
        sunco_coverage_decision_rows(packet, analysis, normalized),
    )
    row = write_table(
        ws,
        row,
        "Competitor Comparables",
        DECISION_EVIDENCE_HEADERS,
        competitor_decision_rows(packet, analysis, normalized),
    )
    row = write_table(
        ws,
        row,
        "Pricing / Channel Read",
        DECISION_EVIDENCE_HEADERS,
        pricing_channel_decision_rows(packet, analysis, normalized),
    )
    row = write_table(
        ws,
        row,
        "Risks and Watchouts",
        DECISION_EVIDENCE_HEADERS,
        risk_watchout_rows(packet, analysis, action_summary),
    )
    row = section_header(ws, row, "Section A - Ideation + Reference Anchor Context")
    row = key_value_rows(
        ws,
        row,
        [
            ("Concept Tracking Name", tracking_name),
            ("Opportunity Type", product_fit_classification(packet)),
            ("Recommended Channel", channel_strategy(packet, analysis)),
            ("Category Owner", identity.get("category_owner")),
            ("Category", identity.get("category")),
            ("Subcategory", identity.get("subcategory")),
            ("Strategy", identity.get("strategy")),
            ("Reference Anchor SKU", identity.get("sunco_reference_sku")),
            ("Launch Outlook", performance.get("launch_outlook")),
            ("Confidence", performance.get("confidence")),
            ("Anchor Data Quality", reference_anchor.get("data_quality")),
            ("Anchor Title", reference.get("title")),
            ("Anchor Title Source", reference.get("title_source")),
            ("Anchor Listing Price", reference.get("listing_price")),
            ("Listing Price Source", reference.get("listing_price_source")),
            ("Listing Price Note", reference.get("listing_price_note")),
            ("Anchor Shopify Revenue 12mo", reference.get("shopify_revenue_12mo")),
            ("Anchor Shopify Units 12mo", reference.get("shopify_units_12mo")),
            ("Shopify Data Source", reference.get("shopify_data_source")),
            ("Anchor Amazon Revenue 12mo", reference.get("amazon_revenue_12mo")),
            ("Anchor Amazon Units 12mo", reference.get("amazon_units_12mo")),
            ("Amazon Data Source", reference.get("amazon_data_source")),
            ("Anchor Data Source", reference.get("reference_data_source")),
            ("Anchor Sales Period", reference.get("sales_period_label")),
        ],
    )
    row = merged_text_row(ws, row, "Reference Anchor Role", reference_anchor.get("primary_use"))
    row = merged_text_row(ws, row, "Reference Anchor Secondary Use", reference_anchor.get("secondary_use"))
    row = merged_text_row(ws, row, "Reference Anchor Caution", reference_anchor.get("caution"))
    row = merged_text_row(ws, row, "Reference Anchor Guardrail", reference_anchor.get("do_not_overweight"))
    row = merged_text_row(ws, row, "Reference Anchor Image URL", reference.get("image_url"))
    row = merged_text_row(
        ws,
        row,
        "Performance Rationale",
        " | ".join(as_list(performance.get("rationale"))),
    )
    row = key_value_rows(
        ws,
        row,
        [
            ("Verified Listings", normalized_summary.get("verified_listing_count")),
            ("Inferred Competitors", normalized_summary.get("inferred_competitor_count")),
        ],
        columns=2,
    )
    row = write_table(
        ws,
        row,
        "Step 1B Demand + Coverage Evidence",
        ["Signal", "Evidence", "Source", "Link"],
        step1b_evidence_rows(packet, analysis),
    )
    row = write_table(
        ws,
        row,
        "Stackline Channel Comparison",
        ["Channel", "Retail Sales", "Units Sold", "Avg Price", "Sales Growth %", "Sunco Share %"],
        channel_comparison_rows(performance),
    )
    row = merged_text_row(ws, row, "Gate Readiness Summary", gate_readiness.get("summary"))
    row = write_table(
        ws,
        row,
        "Gate Readiness by Channel / Gate",
        ["Channel", "Gate", "Family State", "Score", "Evidence Score", "Evidence Label", "Implemented Questions"],
        gate_readiness_snapshot_rows(gate_readiness),
    )
    row = write_table(
        ws,
        row,
        "Primary G2 Pillar Rollup",
        ["Pillar", "Base Weight", "Effective Weight", "Avg Score", "Scored Questions", "Status"],
        gate_readiness_pillar_rows(gate_readiness),
    )

    row = write_table(
        ws,
        row,
        "Section B - Amazon Competitors (Verified Listings)",
        ["Brand", "Product", "Identifier", "Channel", "Price", "Attributes", "Movement / Demand Signal", "Match Quality", "Source Link", "Link Status", "Source Basis"],
        detailed_candidate_rows(as_list(normalized.get("items")), AMAZON_CHANNELS, limit=6, verification_filter="verified_listing"),
    )

    row = write_table(
        ws,
        row,
        "Section C - Brick-and-Mortar / Direct Competitors (Verified Listings)",
        ["Brand", "Product", "Identifier", "Channel", "Price", "Attributes", "Movement / Demand Signal", "Match Quality", "Source Link", "Link Status", "Source Basis"],
        detailed_candidate_rows(as_list(normalized.get("items")), BM_DIRECT_CHANNELS, limit=6, verification_filter="verified_listing"),
    )

    row = write_table(
        ws,
        row,
        "Section D - Inferred Competitors / Directional Only",
        ["Brand", "Product", "Likely Channel", "Supporting Source", "Source Link", "Link Status", "Why Inferred", "Confidence"],
        inferred_competitor_rows(as_list(normalized.get("items")), limit=6),
    )

    row = write_table(
        ws,
        row,
        "Section E - Pricing Position",
        ["Metric", "Value"],
        pricing_position_rows(pricing),
    )
    row = write_table(
        ws,
        row,
        "Pricing Benchmarks",
        ["Metric", "Samples", "Min", "P25", "Median", "Mean", "P75", "Max"],
        benchmark_rows(pricing),
    )
    row = write_table(
        ws,
        row,
        "Margin Targets",
        ["Channel", "Target Margin %", "Min MSRP", "Vs Target MSRP %", "Vs Market Median %"],
        margin_rows(pricing),
    )
    row = write_table(
        ws,
        row,
        "Value Ranking",
        ["Metric", "Target", "Percentile", "Bucket", "Vs Median %"],
        value_position_rows(pricing),
    )
    row = write_table(
        ws,
        row,
        "Top Brands",
        ["Brand", "Candidate Count", "Channels", "Median Unit Price"],
        top_brand_rows(as_list(summary.get("top_brands"))),
    )

    row = write_table(
        ws,
        row,
        "Section F - Feature / Certification Signals",
        ["Type", "Label", "Signal", "Evidence", "Coverage %", "Matched", "Recommendation"],
        spec_action_rows(spec_coverage),
    )
    row = write_table(
        ws,
        row,
        "Numeric Target Positioning",
        ["Metric", "Target", "Median", "P75", "Percentile", "Recommendation"],
        numeric_guidance_rows(spec_coverage),
    )
    row = merged_text_row(ws, row, "Category Optimization Summary", ideation_optimization.get("summary"))
    row = key_value_rows(
        ws,
        row,
        [
            ("Optimization Profile", ideation_optimization.get("profile_label")),
            ("Profile Match Basis", ", ".join(as_list(ideation_optimization.get("matched_taxonomy"))) or ", ".join(as_list(ideation_optimization.get("matched_keywords")))),
            ("Active Modifiers", ", ".join(item.get("label") for item in optimization_modifiers if item.get("label"))),
            ("Optimization Score", optimization_scorecard.get("score")),
            ("Optimization Label", optimization_scorecard.get("label")),
            ("Optimization Confidence", optimization_scorecard.get("confidence")),
        ],
    )
    row = write_table(
        ws,
        row,
        "Active Variant Modifiers",
        ["Modifier", "Matched Keywords", "Match Score"],
        optimization_modifier_rows(optimization_modifiers),
    )
    row = write_table(
        ws,
        row,
        "Optimization Scorecard",
        ["Component", "Score", "Weight", "Reason"],
        optimization_scorecard_rows(optimization_scorecard),
    )
    row = write_table(
        ws,
        row,
        "Primary Decision Drivers",
        ["Tier", "Driver", "Type", "Signal", "Why It Matters"],
        optimization_driver_rows(as_list(ideation_optimization.get("primary_decision_drivers"))),
    )
    row = write_table(
        ws,
        row,
        "Secondary Decision Drivers",
        ["Tier", "Driver", "Type", "Signal", "Why It Matters"],
        optimization_driver_rows(as_list(ideation_optimization.get("secondary_decision_drivers"))),
    )
    row = write_table(
        ws,
        row,
        "Do Not Over-Weight Yet",
        ["Driver", "Type", "Signal", "Reason"],
        low_signal_rows(as_list(ideation_optimization.get("low_signal_attributes"))),
    )
    row = write_table(
        ws,
        row,
        "Highest-Impact Vendor Requests",
        ["Priority", "Area", "Request", "Reason"],
        vendor_request_rows(vendor_requests),
    )
    row = write_list_section(ws, row, "Recommendations", as_list(analysis.get("recommendations")))
    row = write_list_section(ws, row, "Audit Notes", as_list(analysis.get("notes")))

    row = section_header(ws, row, "Section G - PRD Generator Pre-Fill")
    row = key_value_rows(ws, row, prd_prefill_pairs(packet, analysis))


def clean_slide_text(value: Any, limit: int | None = None) -> str:
    """Return compact text suitable for one-page slide summary cells."""
    text = normalize_text(value).strip()
    text = re.sub(r"\s+", " ", text)
    if limit and len(text) > limit:
        return text[: max(0, limit - 1)].rstrip(" ,.;") + "..."
    return text


def slide_title(packet: dict[str, Any]) -> str:
    """Create a human-readable slide title from the ideation packet."""
    identity = as_dict(packet.get("identity"))
    target_profile = as_dict(packet.get("target_profile"))
    physical = as_dict(target_profile.get("physical"))
    context = context_text_for_packet(packet)
    category = normalize_text(identity.get("subcategory") or identity.get("category") or "Product")
    lumen_match = re.search(r"\b([0-9][0-9,]*)\s*(?:lm|lumens?)\b", context, flags=re.IGNORECASE)
    lumens = f"{lumen_match.group(1).replace(',', ',')}lm" if lumen_match else ""
    size = normalize_text(physical.get("size_form_factor"))
    size_match = re.search(r"\b(\d+\s*(?:ft|foot|feet)|\d+x\d+|1\s*x\s*4|2\s*x\s*2|2\s*x\s*4)\b", size, flags=re.IGNORECASE)
    size_label = size_match.group(1).replace(" ", "") if size_match else ""
    pieces = [category]
    if lumens:
        pieces.append(lumens)
    if size_label:
        pieces.append(size_label)
    title = ", ".join(pieces)
    if len(title) < 8:
        title = normalize_text(identity.get("ideation_name")) or concept_tracking_name(packet)
    return clean_slide_text(title, 95)


def slide_sheet_title_for_row(row_number: int, packet: dict[str, Any], used: set[str]) -> str:
    """Return a workbook-safe slide summary sheet name."""
    return safe_sheet_title(f"S{row_number:02d} {concept_tracking_name(packet)}", used)


def set_slide_layout(ws) -> None:
    """Apply one-page slide-style layout defaults."""
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = None
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.sheet_view.zoomScale = 85

    widths = {
        "A": 13,
        "B": 13,
        "C": 13,
        "D": 13,
        "E": 12,
        "F": 13,
        "G": 13,
        "H": 13,
        "I": 12,
        "J": 12,
        "K": 12,
        "L": 12,
        "M": 12,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width
    for row_index in range(1, 39):
        ws.row_dimensions[row_index].height = 18
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 48
    ws.row_dimensions[6].height = 46
    ws.row_dimensions[7].height = 46
    ws.row_dimensions[9].height = 20
    ws.row_dimensions[19].height = 24
    ws.row_dimensions[24].height = 24
    ws.row_dimensions[29].height = 24
    ws.row_dimensions[34].height = 24
    ws.row_dimensions[35].height = 16
    ws.print_area = "A1:M37"


def style_range(
    ws,
    cell_range: str,
    *,
    fill: PatternFill | None = None,
    font: Font | None = None,
    border: Border | None = None,
    alignment: Alignment | None = None,
) -> None:
    """Apply style fields to every cell in a range."""
    for row in ws[cell_range]:
        for cell in row:
            if fill is not None:
                cell.fill = fill
            if font is not None:
                cell.font = font
            if border is not None:
                cell.border = border
            if alignment is not None:
                cell.alignment = alignment


def merge_write(
    ws,
    cell_range: str,
    value: Any,
    *,
    fill: PatternFill | None = None,
    font: Font | None = None,
    border: Border | None = SLIDE_TABLE_BORDER,
    alignment: Alignment | None = None,
) -> None:
    """Merge a range and write a styled value into the top-left cell."""
    ws.merge_cells(cell_range)
    start = cell_range.split(":", 1)[0]
    cell = ws[start]
    cell.value = normalize_text(value)
    style_range(ws, cell_range, fill=fill, font=font, border=border, alignment=alignment or WRAP_ALIGNMENT)


def section_bar(ws, row: int, start_col: int, end_col: int, title: str, *, fill: PatternFill = SLIDE_TEAL_FILL) -> None:
    """Write a colored slide section bar across columns."""
    start = f"{get_column_letter(start_col)}{row}"
    end = f"{get_column_letter(end_col)}{row}"
    merge_write(
        ws,
        f"{start}:{end}",
        title,
        fill=fill,
        font=SLIDE_SECTION_FONT,
        border=SLIDE_TABLE_BORDER,
        alignment=Alignment(horizontal="left", vertical="center", wrap_text=True),
    )


def write_slide_table(
    ws,
    row: int,
    start_col: int,
    headers: list[str],
    rows: list[list[Any]],
    *,
    header_fill: PatternFill = SLIDE_TEAL_FILL,
    max_rows: int | None = None,
) -> int:
    """Write a compact slide-ready table."""
    if max_rows is not None:
        rows = rows[:max_rows]
    for offset, header in enumerate(headers):
        cell = ws.cell(row=row, column=start_col + offset, value=header)
        cell.fill = header_fill
        cell.font = SLIDE_SECTION_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = SLIDE_TABLE_BORDER
    row += 1
    body_rows = rows or [["No available source-backed row."] + [""] * (len(headers) - 1)]
    for row_values in body_rows:
        for offset in range(len(headers)):
            value = row_values[offset] if offset < len(row_values) else ""
            cell = ws.cell(row=row, column=start_col + offset)
            apply_table_cell(cell, value, headers[offset])
            cell.font = SLIDE_SMALL_FONT if not cell.hyperlink else SLIDE_LINK_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = SLIDE_TABLE_BORDER
            cell.fill = SLIDE_WHITE_FILL
        row += 1
    return row


def write_span_table(
    ws,
    row: int,
    sections: list[tuple[str, int, int]],
    rows: list[list[Any]],
    *,
    header_fill: PatternFill = SLIDE_TEAL_FILL,
    max_rows: int | None = None,
) -> int:
    """Write a compact table where logical columns can span multiple Excel columns."""
    if max_rows is not None:
        rows = rows[:max_rows]
    for header, start_col, end_col in sections:
        section_bar(ws, row, start_col, end_col, header, fill=header_fill)
    row += 1
    body_rows = rows or [["No available source-backed row."] + [""] * (len(sections) - 1)]
    for row_values in body_rows:
        for index, (_, start_col, end_col) in enumerate(sections):
            value = row_values[index] if index < len(row_values) else ""
            display_value = value.get("value") if isinstance(value, dict) else value
            display_text = display_value_for_label(sections[index][0], display_value)
            start = f"{get_column_letter(start_col)}{row}"
            end = f"{get_column_letter(end_col)}{row}"
            merge_write(
                ws,
                f"{start}:{end}",
                display_text,
                fill=SLIDE_WHITE_FILL,
                font=SLIDE_SMALL_FONT,
                border=SLIDE_TABLE_BORDER,
                alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
            )
            if isinstance(value, dict) and value.get("hyperlink"):
                ws[start].hyperlink = value.get("hyperlink")
                ws[start].font = SLIDE_LINK_FONT
        row += 1
    return row


def target_spec_lines(packet: dict[str, Any], analysis: dict[str, Any]) -> list[str]:
    """Build concise target-spec bullets for a slide card."""
    target_profile = as_dict(packet.get("target_profile"))
    electrical = as_dict(target_profile.get("electrical"))
    physical = as_dict(target_profile.get("physical"))
    pricing = as_dict(analysis.get("pricing_analysis"))
    specs = [
        ("Wattage", electrical.get("wattage_primary") or electrical.get("wattage_max")),
        ("Lumens", electrical.get("lumens_target")),
        ("CCTs", electrical.get("cct_primary") or electrical.get("cct_max")),
        ("Voltage", electrical.get("voltage")),
        ("Dimming", electrical.get("dimming_type") or ("Dimmable" if electrical.get("dimmable") else "")),
        ("CRI", electrical.get("cri")),
        ("Moisture", physical.get("moisture_rating") or physical.get("ip_rating")),
        ("Certs", certification_summary(packet, analysis)),
        ("Target MSRP", pricing.get("target_msrp")),
    ]
    lines = []
    for label, value in specs:
        text = display_value_for_label(label, value)
        if text:
            lines.append(f"- {label}: {text}")
    return lines[:9]


def primary_feature_summary(packet: dict[str, Any], analysis: dict[str, Any]) -> str:
    """Return a short key-differentiator statement."""
    target_profile = as_dict(packet.get("target_profile"))
    feature_watchlist = [normalize_text(item) for item in as_list(target_profile.get("feature_watchlist")) if normalize_text(item)]
    spec_coverage = as_dict(analysis.get("spec_coverage"))
    feature_rows = as_list(spec_coverage.get("feature_coverage"))
    supported = [
        normalize_text(as_dict(row).get("label"))
        for row in feature_rows
        if normalize_text(as_dict(row).get("signal")).lower() in {"competitive", "differentiator", "table_stakes"}
    ]
    merged = list(dict.fromkeys(feature_watchlist + supported))
    return clean_slide_text(", ".join(merged[:4]) or product_difference_summary(packet, analysis), 130)


def slide_strategy_text(packet: dict[str, Any], analysis: dict[str, Any]) -> str:
    """Build the slide strategy narrative from Step 3 evidence."""
    identity = as_dict(packet.get("identity"))
    evidence = extract_research_note_evidence(packet)
    performance = as_dict(analysis.get("performance_estimation"))
    snapshot = as_dict(performance.get("market_snapshot"))
    movement = evidence.get("ecommerce_movement") or ""
    movement_short = "; ".join(movement.split(";")[:3]).strip()
    channel = channel_strategy(packet, analysis).replace(
        "Shopify/front-end first; Amazon follow-up if Stackline, pack economics, and margin support.",
        "Shopify first; Amazon follow-up if economics support.",
    )
    parts = [
        f"{concept_tracking_name(packet)} is recommended as a {product_fit_classification(packet).lower()} in {normalize_text(identity.get('subcategory') or identity.get('category'))}.",
        f"Gap vs current line: {clean_slide_text(product_difference_summary(packet, analysis), 140)}",
    ]
    market = []
    if snapshot.get("segment_retail_sales") is not None:
        market.append(f"{display_value_for_label('Sales', snapshot.get('segment_retail_sales'))} Stackline sales")
    if snapshot.get("segment_sales_growth_pct") is not None:
        market.append(f"{display_value_for_label('Growth %', snapshot.get('segment_sales_growth_pct'))} growth")
    if movement_short:
        market.append(f"Redshift movement: {movement_short}")
    if market:
        parts.append("; ".join(market) + ".")
    if evidence.get("sunco_coverage"):
        coverage = evidence["sunco_coverage"].split(".", 1)[0]
        if re.search(r"\b0\s+strong\b", coverage, flags=re.IGNORECASE):
            coverage = "0 exact Sunco active-catalog matches"
        parts.append(f"Coverage check: {coverage}.")
    parts.append(f"Channel: {channel}")
    return clean_slide_text(" ".join(part for part in parts if part), 500)


def best_verified_item(normalized: dict[str, Any], channels: set[str]) -> dict[str, Any]:
    """Return the best verified competitor item for a channel set."""
    for item in sort_candidates(as_list(normalized.get("items"))):
        if resolved_source_channel(item) in channels and verification_status(item) == "verified_listing":
            return item
    return {}


def competitor_table_row(item: dict[str, Any], *, direct: bool = False) -> list[Any]:
    """Format the key competitor row for the slide summary."""
    if not item:
        return []
    if direct:
        identifier = listing_identifier(item)
        url = optional_text(item.get("url"))
        identifier_cell: Any = {"value": identifier, "hyperlink": url} if url else identifier
        return [
            identifier_cell,
            item.get("brand"),
            item.get("price"),
            clean_slide_text(item.get("product_title"), 100),
        ]
    return [
        listing_identifier(item),
        item.get("brand"),
        item.get("price"),
        item_demand_signal(item),
        item.get("rating") or "",
        item.get("review_count") or "",
        listing_link_cell(item),
    ]


def source_footer(packet: dict[str, Any], analysis: dict[str, Any], normalized: dict[str, Any]) -> str:
    """Build a compact source footer for the slide summary."""
    evidence = extract_research_note_evidence(packet)
    performance = as_dict(analysis.get("performance_estimation"))
    snapshot = as_dict(performance.get("market_snapshot"))
    sources = ["Sources: Step 3 packet/analysis"]
    if snapshot.get("segment_name"):
        sources.append(f"Stackline segment: {snapshot.get('segment_name')}")
    if evidence.get("ecommerce_movement") or evidence.get("inventory_support"):
        sources.append("Redshift ecommerce competitor snapshot")
    if evidence.get("sunco_coverage"):
        sources.append("Postgres/Sunco catalog coverage")
    if evidence.get("review_link"):
        sources.append(evidence["review_link"])
    direct_item = best_verified_item(normalized, BM_DIRECT_CHANNELS)
    if direct_item and optional_text(direct_item.get("url")):
        sources.append(optional_text(direct_item.get("url")) or "")
    return clean_slide_text(" | ".join(sources), 420)


def best_slide_image_url(packet: dict[str, Any], normalized: dict[str, Any]) -> str:
    """Return the best available product image URL for a slide summary."""
    reference = as_dict(packet.get("reference_baseline"))
    if optional_text(reference.get("image_url")):
        return optional_text(reference.get("image_url")) or ""
    for item in sort_candidates(as_list(normalized.get("items"))):
        image_url = optional_text(item.get("image_url") or item.get("image"))
        if image_url:
            return image_url
    return ""


def cached_slide_image(session_dir: Path, url: str) -> Path | None:
    """Download and cache an image URL for slide workbook embedding."""
    if not url or not url.startswith(("http://", "https://")):
        return None
    parsed = urlparse(url)
    suffix = Path(parsed.path).suffix.lower()
    if suffix not in {".png", ".jpg", ".jpeg"}:
        suffix = ".jpg"
    cache_dir = session_dir / "reports" / "_slide_image_cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    image_path = cache_dir / f"{hashlib.sha1(url.encode('utf-8')).hexdigest()[:16]}{suffix}"
    if image_path.exists() and image_path.stat().st_size:
        return image_path
    try:
        request = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(request, timeout=8) as response:
            data = response.read(5_000_000)
        if not data:
            return None
        image_path.write_bytes(data)
        return image_path
    except Exception:
        return None


def add_slide_image(ws, session_dir: Path, image_url: str) -> None:
    """Add a product/reference image to the slide summary when possible."""
    merge_write(
        ws,
        "A4:D15",
        "",
        fill=SLIDE_LIGHT_FILL,
        border=Border(left=SLIDE_THIN_SIDE, right=SLIDE_THIN_SIDE, top=SLIDE_THIN_SIDE, bottom=SLIDE_THIN_SIDE),
    )
    image_path = cached_slide_image(session_dir, image_url)
    if image_path:
        try:
            from openpyxl.drawing.image import Image as XLImage

            image = XLImage(str(image_path))
            image.width = 260
            image.height = 185
            ws.add_image(image, "B5")
            return
        except Exception:
            pass
    if image_url:
        ws["A4"].value = f"Image URL:\n{image_url}"
        ws["A4"].hyperlink = image_url
        ws["A4"].font = SLIDE_LINK_FONT
    else:
        ws["A4"].value = "No image available in Step 3 packet."
        ws["A4"].font = SLIDE_SMALL_FONT


def render_slide_summary_sheet(
    ws,
    session_dir: Path,
    row_number: int,
    packet: dict[str, Any],
    analysis: dict[str, Any],
    normalized: dict[str, Any],
) -> None:
    """Render one slide-ready SKU summary sheet."""
    set_slide_layout(ws)
    identity = as_dict(packet.get("identity"))
    reference = as_dict(packet.get("reference_baseline"))
    pricing = as_dict(analysis.get("pricing_analysis"))
    amazon_snapshot = gate_snapshot(as_dict(analysis.get("gate_readiness")), "amazon", "G2")
    image_url = best_slide_image_url(packet, normalized)

    merge_write(
        ws,
        "A1:J2",
        slide_title(packet),
        fill=SLIDE_WHITE_FILL,
        font=SLIDE_TITLE_FONT,
        border=Border(bottom=Side(style="dotted", color="7F8DA0")),
        alignment=Alignment(horizontal="left", vertical="center", wrap_text=True),
    )
    merge_write(
        ws,
        "K1:M2",
        "Total SKU Count: 1 variation",
        fill=SLIDE_WHITE_FILL,
        font=Font(color=SLIDE_MUTED, size=9),
        border=Border(bottom=Side(style="dotted", color="7F8DA0")),
        alignment=Alignment(horizontal="right", vertical="center", wrap_text=True),
    )

    add_slide_image(ws, session_dir, image_url)

    ws["E4"].value = "Strategy:"
    ws["E4"].font = Font(color=SLIDE_TEAL, bold=True, underline="single", size=10)
    ws["E4"].alignment = WRAP_ALIGNMENT
    merge_write(
        ws,
        "E5:M8",
        slide_strategy_text(packet, analysis),
        fill=SLIDE_WHITE_FILL,
        font=SLIDE_BODY_FONT,
        border=None,
        alignment=Alignment(horizontal="left", vertical="top", wrap_text=True),
    )

    ws["E10"].value = "Target Specs:"
    ws["E10"].font = Font(color=SLIDE_TEAL, bold=True, underline="single", size=10)
    merge_write(
        ws,
        "E11:H16",
        "\n".join(target_spec_lines(packet, analysis)),
        fill=SLIDE_WHITE_FILL,
        font=SLIDE_SMALL_FONT,
        border=None,
        alignment=Alignment(horizontal="left", vertical="top", wrap_text=True),
    )

    section_bar(ws, 10, 10, 13, "Sunco Anchor Rev (12mo)")
    write_slide_table(
        ws,
        11,
        10,
        ["Comparable SKU", "Shopify Rev", "Amazon Rev", "Price"],
        [
            [
                identity.get("sunco_reference_sku"),
                reference.get("shopify_revenue_12mo"),
                reference.get("amazon_revenue_12mo"),
                reference.get("listing_price"),
            ]
        ],
        max_rows=1,
    )

    ws["A17"].value = "Product Variants:"
    ws["A17"].font = Font(color=SLIDE_TEAL, bold=True, underline="single", size=10)
    write_span_table(
        ws,
        18,
        [
            ("Concept Tracking Name", 1, 4),
            ("Description", 5, 9),
            ("Key Differentiators", 10, 13),
        ],
        [
            [
                concept_tracking_name(packet),
                clean_slide_text(identity.get("ideation_name"), 120),
                primary_feature_summary(packet, analysis),
            ]
        ],
        max_rows=1,
    )

    ws["A22"].value = "Comparable SKU Performance (12mo):"
    ws["A22"].font = Font(color=SLIDE_TEAL, bold=True, underline="single", size=10)
    write_slide_table(
        ws,
        23,
        1,
        ["Comparable SKU", "Shopify Rev", "Amazon Rev", "Shopify Units", "Amazon Units", "Sale Price", "Trend / Caveat"],
        [
            [
                identity.get("sunco_reference_sku"),
                reference.get("shopify_revenue_12mo"),
                reference.get("amazon_revenue_12mo"),
                reference.get("shopify_units_12mo"),
                reference.get("amazon_units_12mo"),
                reference.get("listing_price"),
                related_sales_slide_note(reference),
            ]
        ],
        max_rows=1,
    )

    amazon_item = best_verified_item(normalized, AMAZON_CHANNELS)
    ws["A27"].value = "Amazon Segment Competitor"
    ws["A27"].font = Font(color=SLIDE_DARK_NAVY, bold=True, size=8)
    ws["C27"].value = f"Amazon G2 {display_value_for_label('Score', amazon_snapshot.get('weighted_score')) or 'n/a'} | evidence {normalize_text(as_dict(amazon_snapshot.get('evidence_confidence')).get('label')) or 'n/a'}"
    ws["C27"].font = SLIDE_MUTED_FONT
    write_slide_table(
        ws,
        28,
        1,
        ["ASIN / ID", "Brand", "Price", "Demand Signal", "Rating", "Reviews", "Source"],
        [competitor_table_row(amazon_item)],
        header_fill=SLIDE_NAVY_FILL,
        max_rows=1,
    )

    direct_item = best_verified_item(normalized, BM_DIRECT_CHANNELS)
    ws["A32"].value = "Direct Segment Competitor"
    ws["A32"].font = Font(color=SLIDE_DARK_NAVY, bold=True, size=8)
    write_span_table(
        ws,
        33,
        [
            ("Product Code", 1, 3),
            ("Brand", 4, 5),
            ("Price", 6, 7),
            ("Product Name", 8, 13),
        ],
        [competitor_table_row(direct_item, direct=True)],
        header_fill=SLIDE_NAVY_FILL,
        max_rows=1,
    )

    merge_write(
        ws,
        "A36:M36",
        source_footer(packet, analysis, normalized),
        fill=SLIDE_WHITE_FILL,
        font=SLIDE_MUTED_FONT,
        border=None,
        alignment=Alignment(horizontal="left", vertical="center", wrap_text=True),
    )
    ws["M37"].value = "Sunco"
    ws["M37"].font = Font(color=SLIDE_TEAL, bold=True, size=10)
    ws["M37"].alignment = Alignment(horizontal="right", vertical="bottom")


def build_slide_index_sheet(ws, payloads: list[tuple[int, dict[str, Any], dict[str, Any], dict[str, Any]]]) -> None:
    """Build a clean index sheet for the slide summary workbook."""
    set_default_layout(ws)
    ws.title = "Index"
    ws.cell(row=1, column=1, value="Step 3 Slide Summary Index")
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    rows = []
    used_titles = {ws.title}
    for row_number, packet, analysis, normalized in payloads:
        rows.append(
            [
                row_number,
                concept_tracking_name(packet),
                slide_title(packet),
                product_fit_classification(packet),
                channel_strategy(packet, analysis),
                as_dict(analysis.get("performance_estimation")).get("launch_outlook"),
                as_dict(analysis.get("performance_estimation")).get("confidence"),
                slide_sheet_title_for_row(row_number, packet, used_titles),
            ]
        )
    write_table(
        ws,
        3,
        "Slide Summary Sheets",
        ["Row", "Concept Tracking Name", "Slide Title", "Opportunity Type", "Channel Strategy", "Outlook", "Confidence", "Worksheet"],
        rows,
    )


def slide_summary_path_for(destination: Path, session_dir: Path) -> Path:
    """Resolve the companion slide-summary workbook path for a detailed workbook."""
    if destination.name.endswith("_completed_rows.xlsx"):
        return destination.with_name(destination.name.replace("_completed_rows.xlsx", "_slide_summaries.xlsx"))
    return destination.with_name(f"{destination.stem}_slide_summaries{destination.suffix}")


def build_slide_summary_workbook(
    session_dir: Path,
    payloads: list[tuple[int, dict[str, Any], dict[str, Any], dict[str, Any]]],
    output_path: Path,
) -> Path | None:
    """Build the slide-ready companion workbook with one sheet per concept."""
    if not payloads:
        return None

    wb = Workbook()
    index_ws = wb.active
    build_slide_index_sheet(index_ws, payloads)

    used_titles = {index_ws.title}
    for row_number, packet, analysis, normalized in payloads:
        title = slide_sheet_title_for_row(row_number, packet, used_titles)
        ws = wb.create_sheet(title=title)
        render_slide_summary_sheet(ws, session_dir, row_number, packet, analysis, normalized)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def completed_row_payloads(
    session_dir: Path,
    rows: list[int] | None = None,
) -> list[tuple[int, dict[str, Any], dict[str, Any], dict[str, Any]]]:
    """Load completed packet/analysis/normalized payloads for selected rows."""
    manifest = read_json(session_dir / "manifest.json")
    target_rows = set(rows or [row["row_number"] for row in manifest.get("rows", [])])
    payloads = []

    for row in manifest.get("rows", []):
        row_number = row["row_number"]
        if row_number not in target_rows:
            continue
        if row.get("stages", {}).get("analyzed") != "complete":
            continue
        packet = read_json(packet_path_for(session_dir, row_number))
        analysis = read_json(artifact_path_for(session_dir, row_number, "analyzed"))
        normalized = read_json(artifact_path_for(session_dir, row_number, "normalized"))
        payloads.append((row_number, packet, analysis, normalized))

    return payloads


def build_report(session_dir: Path, row_number: int) -> Path | None:
    """Build one Excel report for a completed analyzed row."""
    payloads = completed_row_payloads(session_dir, rows=[row_number])
    if not payloads:
        return None

    _, packet, analysis, normalized = payloads[0]
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title_for_row(row_number, packet, set())
    render_row_sheet(ws, row_number, packet, analysis, normalized)

    report_path = artifact_path_for(session_dir, row_number, "reported")
    report_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(report_path)
    return report_path


def build_summary_sheet(ws, payloads: list[tuple[int, dict[str, Any], dict[str, Any], dict[str, Any]]]) -> None:
    """Build a summary sheet for the combined workbook."""
    set_default_layout(ws)
    ws.title = "Summary"
    ws.cell(row=1, column=1, value="Step 3 Research Summary")
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    row = key_value_rows(
        ws,
        3,
        [
            ("Completed Row Count", len(payloads)),
            ("Workbook Scope", "One summary sheet plus one sheet per completed ideation row"),
        ],
        columns=1,
    )
    row = write_table(
        ws,
        row + 1,
        "Summary Metric Guide",
        ["Metric", "What It Means", "Question It Answers"],
        SUMMARY_METRIC_GUIDE_ROWS,
    )
    used_titles = {ws.title}
    completed_rows = []
    for row_number, packet, analysis, _ in payloads:
        identity = as_dict(packet.get("identity"))
        amazon_g2 = next(
            (
                snapshot.get("weighted_score")
                for snapshot in as_list(as_dict(analysis.get("gate_readiness")).get("snapshots"))
                if snapshot.get("channel") == "amazon" and snapshot.get("gate") == "G2"
            ),
            None,
        )
        amazon_evidence = next(
            (
                as_dict(snapshot.get("evidence_confidence")).get("label")
                for snapshot in as_list(as_dict(analysis.get("gate_readiness")).get("snapshots"))
                if snapshot.get("channel") == "amazon" and snapshot.get("gate") == "G2"
            ),
            None,
        )
        completed_rows.append(
            [
                row_number,
                concept_tracking_name(packet),
                normalize_text(identity.get("ideation_name")),
                product_fit_classification(packet),
                channel_strategy(packet, analysis),
                as_dict(analysis.get("performance_estimation")).get("launch_outlook"),
                as_dict(analysis.get("performance_estimation")).get("confidence"),
                amazon_g2,
                amazon_evidence,
                sheet_title_for_row(row_number, packet, used_titles),
            ]
        )
    write_table(
        ws,
        row,
        "Completed Ideation Rows",
        ["Row", "Concept Tracking Name", "Ideation", "Opportunity Type", "Channel Strategy", "Outlook", "Confidence", "Amazon G2", "Amazon Evidence", "Worksheet"],
        completed_rows,
    )


def build_combined_workbook(
    session_dir: Path,
    rows: list[int] | None = None,
    output_path: str | None = None,
) -> Path | None:
    """Build one workbook with a summary sheet plus one sheet per completed row."""
    payloads = completed_row_payloads(session_dir, rows=rows)
    if not payloads:
        return None

    wb = Workbook()
    summary_ws = wb.active
    build_summary_sheet(summary_ws, payloads)

    used_titles = {summary_ws.title}
    for row_number, packet, analysis, normalized in payloads:
        title = sheet_title_for_row(row_number, packet, used_titles)
        ws = wb.create_sheet(title=title)
        render_row_sheet(ws, row_number, packet, analysis, normalized)

    if output_path:
        destination = Path(output_path).resolve()
    else:
        destination = session_dir / "reports" / f"{session_dir.name}_completed_rows.xlsx"
    destination.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destination)
    return destination


def parse_rows_argument(value: str | None) -> list[int] | None:
    """Parse an optional comma-separated row list."""
    if not value:
        return None
    rows = []
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        rows.append(int(part))
    return rows or None


def build_reports(
    session_root: str,
    rows: list[int] | None = None,
    combined: bool = False,
    output_path: str | None = None,
) -> dict[str, Any]:
    """Build report artifacts for selected session rows."""
    session_dir = Path(session_root).resolve()

    if combined:
        combined_path = build_combined_workbook(session_dir, rows=rows, output_path=output_path)
        update_result = update_session(str(session_dir))
        return {
            "session_root": str(session_dir),
            "rows_requested": sorted(rows or []),
            "combined_workbook": str(combined_path) if combined_path else None,
            "manifest_summary": update_result["summary"],
        }

    manifest = read_json(session_dir / "manifest.json")
    target_rows = set(rows or [row["row_number"] for row in manifest.get("rows", [])])

    written_rows = []
    skipped_rows = []
    report_files = []

    for row in manifest.get("rows", []):
        row_number = row["row_number"]
        if row_number not in target_rows:
            continue
        report_path = build_report(session_dir, row_number)
        if report_path is None:
            skipped_rows.append(row_number)
            continue
        written_rows.append(row_number)
        report_files.append(str(report_path))

    update_result = update_session(str(session_dir))
    return {
        "session_root": str(session_dir),
        "rows_requested": sorted(target_rows),
        "rows_written": written_rows,
        "rows_skipped": skipped_rows,
        "report_files": report_files,
        "manifest_summary": update_result["summary"],
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build Excel report artifacts for completed analyzed rows."
    )
    parser.add_argument("session_root", help="Path to an initialized research session.")
    parser.add_argument(
        "--rows",
        default=None,
        help="Optional comma-separated row numbers to report.",
    )
    parser.add_argument(
        "--combined",
        action="store_true",
        help="Build one combined workbook with one sheet per completed row.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Optional explicit output path for the combined workbook.",
    )
    args = parser.parse_args()

    result = build_reports(
        args.session_root,
        rows=parse_rows_argument(args.rows),
        combined=args.combined,
        output_path=args.output,
    )
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
