"""
Shared utility functions for the GST Reconciliation Tool.
Covers:
  - Column keyword mapping (per spec)
  - Month normalization (fiscal year: Apr → Mar)
  - Currency formatting (INR)
  - Difference calculation helpers
"""

import pandas as pd
import numpy as np
from datetime import datetime

# ── Fiscal year month order (India: April to March) ──────────────────────────
FISCAL_MONTH_ORDER = [
    "Apr", "May", "Jun", "Jul", "Aug", "Sep",
    "Oct", "Nov", "Dec", "Jan", "Feb", "Mar",
]

# ── Month string → abbreviation map ─────────────────────────────────────────
MONTH_ABBREV: dict[str, str] = {
    # Full names
    "january": "Jan", "february": "Feb", "march": "Mar", "april": "Apr",
    "may": "May", "june": "Jun", "july": "Jul", "august": "Aug",
    "september": "Sep", "october": "Oct", "november": "Nov", "december": "Dec",
    # Short names
    "jan": "Jan", "feb": "Feb", "mar": "Mar", "apr": "Apr",
    "jun": "Jun", "jul": "Jul", "aug": "Aug", "sep": "Sep",
    "oct": "Oct", "nov": "Nov", "dec": "Dec",
    # Numeric (zero-padded and plain)
    "01": "Jan", "1": "Jan",
    "02": "Feb", "2": "Feb",
    "03": "Mar", "3": "Mar",
    "04": "Apr", "4": "Apr",
    "05": "May", "5": "May",
    "06": "Jun", "6": "Jun",
    "07": "Jul", "7": "Jul",
    "08": "Aug", "8": "Aug",
    "09": "Sep", "9": "Sep",
    "10": "Oct",
    "11": "Nov",
    "12": "Dec",
}

# ── Column keyword mapping (per spec, case-insensitive substring match) ──────
COLUMN_KEYWORDS: dict[str, list[str]] = {
    "sales_value":   ["sale", "job work", "taxable value", "taxable amount"],
    "export_value":  ["export"],
    "sez_value":     ["sez"],
    "igst":          ["igst", "integrated tax", "gst-integrated", "gst integrated"],
    "cgst":          ["cgst", "central tax", "gst- central", "gst central"],
    "sgst":          ["sgst", "state tax", "gst- state", "gst state", "utgst"],
    "date":          ["date", "invoice date", "bill date", "voucher date", "doc date"],
    "invoice_no":    ["invoice no", "invoice number", "bill no", "voucher no",
                      "doc no", "document no", "inv no"],
    "gstin":         ["gstin", "gst no", "gst number", "supplier gstin",
                      "party gstin", "receiver gstin"],
    "note_type":     ["note type", "type", "debit/credit", "cr/dr", "transaction type"],
    "total_value":   ["total value", "invoice value", "total amount", "gross amount"],
}

# ── Tax field names (convenience) ────────────────────────────────────────────
TAX_FIELDS = ["igst", "cgst", "sgst"]
VALUE_FIELDS = ["sales_value", "export_value", "sez_value"] + TAX_FIELDS


# ════════════════════════════════════════════════════════════════════════════
#  Column mapping
# ════════════════════════════════════════════════════════════════════════════

def map_columns(df: pd.DataFrame) -> dict[str, str]:
    """
    Scan DataFrame column names and return a mapping:
        standard_field_name → actual_column_name
    Uses case-insensitive substring matching per the keyword spec.
    """
    col_map: dict[str, str] = {}
    # Build lowercase → original lookup (handles duplicate lowers gracefully)
    col_lower_map = {col.lower().strip(): col for col in df.columns if isinstance(col, str)}

    for field, keywords in COLUMN_KEYWORDS.items():
        for col_lower, col_orig in col_lower_map.items():
            if any(kw.lower() in col_lower for kw in keywords):
                col_map.setdefault(field, col_orig)   # first match wins
    return col_map


# ════════════════════════════════════════════════════════════════════════════
#  Month normalisation
# ════════════════════════════════════════════════════════════════════════════

def normalize_month(date_val) -> str | None:
    """
    Accept any date-like value and return a 3-letter month abbreviation, e.g. 'Apr'.
    Returns None if the value cannot be parsed.
    """
    if date_val is None or (isinstance(date_val, float) and np.isnan(date_val)):
        return None

    # Already a datetime / Timestamp
    if isinstance(date_val, (datetime, pd.Timestamp)):
        return MONTH_ABBREV.get(date_val.strftime("%b").lower())

    date_str = str(date_val).strip()

    # Try pandas parser first (handles DD-MM-YYYY, YYYY-MM-DD, etc.)
    try:
        dt = pd.to_datetime(date_str, dayfirst=True, errors="raise")
        return MONTH_ABBREV.get(dt.strftime("%b").lower())
    except Exception:
        pass

    # Fallback: substring search for month name
    lower = date_str.lower()
    # Prefer longer keys to avoid "mar" matching inside "march" twice
    for key in sorted(MONTH_ABBREV, key=len, reverse=True):
        if len(key) >= 3 and key in lower:
            return MONTH_ABBREV[key]

    return None


def sort_months_fiscal(months: list[str]) -> list[str]:
    """Sort a list of month abbreviations in Indian fiscal-year order."""
    def _key(m: str) -> int:
        try:
            return FISCAL_MONTH_ORDER.index(m)
        except ValueError:
            return 99   # unknown months go last
    return sorted(months, key=_key)


# ════════════════════════════════════════════════════════════════════════════
#  Numeric helpers
# ════════════════════════════════════════════════════════════════════════════

def safe_float(val) -> float:
    """Convert any value to float, returning 0.0 on failure / NaN."""
    if val is None:
        return 0.0
    try:
        if pd.isna(val):
            return 0.0
    except (TypeError, ValueError):
        pass
    try:
        # Strip commas/spaces that appear in Indian-formatted numbers
        return float(str(val).replace(",", "").replace(" ", "").strip() or 0)
    except (ValueError, TypeError):
        return 0.0


# ════════════════════════════════════════════════════════════════════════════
#  Formatting
# ════════════════════════════════════════════════════════════════════════════

def format_inr(val) -> str:
    """Format a number as Indian Rupees with ₹ prefix and comma separation."""
    try:
        v = float(val)
    except (TypeError, ValueError):
        return "₹0.00"
    if v == 0:
        return "₹0.00"
    neg = v < 0
    v = abs(v)
    # Indian number formatting: last 3 digits, then groups of 2
    s = f"{v:,.2f}"
    # Re-apply Indian grouping
    parts = s.split(".")
    integer_part = parts[0].replace(",", "")
    if len(integer_part) > 3:
        last3 = integer_part[-3:]
        rest = integer_part[:-3]
        groups = []
        while len(rest) > 2:
            groups.append(rest[-2:])
            rest = rest[:-2]
        if rest:
            groups.append(rest)
        groups.reverse()
        integer_part = ",".join(groups) + "," + last3
    formatted = f"₹{integer_part}.{parts[1]}"
    return f"-{formatted}" if neg else formatted


def format_diff(val: float) -> str:
    """Return a diff value string with arrow indicator."""
    v = safe_float(val)
    if abs(v) < 0.01:
        return "✅ Nil"
    arrow = "▲" if v > 0 else "▼"
    return f"{arrow} {format_inr(abs(v))}"


def diff_style(val: float) -> str:
    """Return HTML colour tag for a difference value."""
    v = safe_float(val)
    if abs(v) < 0.01:
        return "green"
    return "red" if v > 0 else "blue"


# ════════════════════════════════════════════════════════════════════════════
#  Month-dict arithmetic
# ════════════════════════════════════════════════════════════════════════════

_ZERO_ROW = {f: 0.0 for f in VALUE_FIELDS}


def merge_monthwise(*dicts) -> dict:
    """Sum multiple month→{field: value} dicts together."""
    all_months: set[str] = set()
    for d in dicts:
        all_months.update(d.keys())

    result: dict = {}
    for month in all_months:
        result[month] = dict(_ZERO_ROW)
        for d in dicts:
            row = d.get(month, {})
            for f in VALUE_FIELDS:
                result[month][f] += safe_float(row.get(f, 0))
    return result


def subtract_monthwise(base: dict, to_sub: dict) -> dict:
    """Return base − to_sub for every field in every month."""
    all_months = set(list(base.keys()) + list(to_sub.keys()))
    result: dict = {}
    for month in all_months:
        b = base.get(month, {})
        s = to_sub.get(month, {})
        result[month] = {
            f: safe_float(b.get(f, 0)) - safe_float(s.get(f, 0))
            for f in VALUE_FIELDS
        }
    return result
