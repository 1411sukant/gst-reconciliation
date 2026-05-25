"""
file_parser.py — Excel & PDF ingestion layer.

Handles:
  - Generic Excel / PDF books (sales/purchase/journal registers)
  - GSTR-1 PDF  (outward supply summary)
  - Electronic Credit Ledger Excel/PDF (ITC availment)
  - GSTR-2B Excel (B2B, B2B-CDNR, IMPZ sheets)
"""

from __future__ import annotations

import io
import re
import logging

import pandas as pd
import pdfplumber

from .utils import (
    map_columns,
    normalize_month,
    safe_float,
    VALUE_FIELDS,
    TAX_FIELDS,
    MONTH_ABBREV,
)

logger = logging.getLogger(__name__)


# ════════════════════════════════════════════════════════════════════════════
#  Internal helpers
# ════════════════════════════════════════════════════════════════════════════

def _find_header_row(df_raw: pd.DataFrame) -> int:
    """Scan rows to find the first 'real' header (≥3 non-null cells)."""
    for i, row in df_raw.iterrows():
        non_null = row.dropna()
        if len(non_null) >= 3:
            return int(i)
    return 0


def _clean_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Promote the detected header row and strip whitespace."""
    h = _find_header_row(df_raw)
    df = df_raw.copy()
    df.columns = [
        str(c).strip() if not pd.isna(c) else f"_Col{i}"
        for i, c in enumerate(df_raw.iloc[h])
    ]
    df = df.iloc[h + 1:].reset_index(drop=True)
    # Drop rows that are entirely null
    df = df.dropna(how="all")
    return df


def _add_month_col(df: pd.DataFrame, col_map: dict) -> pd.DataFrame | None:
    """Add '_month' column; return None if no date column found."""
    date_col = col_map.get("date")
    if not date_col:
        # Heuristic: try every column for date-like content
        for col in df.columns:
            sample = df[col].dropna().head(10)
            hits = sum(1 for v in sample if normalize_month(v) is not None)
            if hits >= max(2, len(sample) // 2):
                date_col = col
                break
    if not date_col:
        return None
    df = df.copy()
    df["_month"] = df[date_col].apply(normalize_month)
    return df.dropna(subset=["_month"])


# ════════════════════════════════════════════════════════════════════════════
#  Generic Excel parser
# ════════════════════════════════════════════════════════════════════════════

def parse_excel_generic(
    uploaded_file, sheet_name: int | str = 0
) -> pd.DataFrame:
    """
    Parse any Excel file, auto-detect header row, return clean DataFrame.
    Pass sheet_name=0 (first sheet) by default; use sheet name or index for others.
    """
    try:
        uploaded_file.seek(0)
        raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, dtype=str)
        return _clean_df(raw)
    except Exception as exc:
        logger.warning("Excel parse error (%s): %s", sheet_name, exc)
        return pd.DataFrame()


# ════════════════════════════════════════════════════════════════════════════
#  Books: month-wise aggregation (sales / purchase / credit-note / debit-note)
# ════════════════════════════════════════════════════════════════════════════

def extract_books_monthwise(df: pd.DataFrame) -> dict:
    """
    Given a Books DataFrame (Sales / Purchase / CN / DN register),
    return a dict:  { month_abbrev: {sales_value, export_value, sez_value,
                                     igst, cgst, sgst} }
    """
    if df.empty:
        return {}

    col_map = map_columns(df)
    df_m = _add_month_col(df, col_map)
    if df_m is None or df_m.empty:
        return {}

    result: dict = {}
    for month, group in df_m.groupby("_month"):
        row: dict = {}
        for field in VALUE_FIELDS:
            col = col_map.get(field)
            row[field] = group[col].apply(safe_float).sum() if col else 0.0
        result[month] = row

    return result


# ════════════════════════════════════════════════════════════════════════════
#  GSTR-1 PDF parser
# ════════════════════════════════════════════════════════════════════════════

def _detect_filing_month(text: str) -> str:
    """Extract filing period from the first page of a GSTR-1 PDF."""
    pattern = re.search(
        r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
        r"jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
        r"[\s\-/]*(\d{4})",
        text.lower(),
    )
    if pattern:
        key = pattern.group(1)[:3]
        return MONTH_ABBREV.get(key, key.capitalize())
    return "Unknown"


def parse_gstr1_pdf(uploaded_file) -> dict:
    """
    Parse a GSTR-1 PDF downloaded from the GST portal.

    Returns: { month: {sales_value, export_value, sez_value, igst, cgst, sgst} }

    Strategy:
      1. Detect filing month from page-1 text.
      2. Walk every table on every page.
      3. Detect IGST / CGST / SGST columns by keyword; accumulate.
      4. Detect export / SEZ flags from table headers.
    """
    totals: dict = {
        "sales_value": 0.0,
        "export_value": 0.0,
        "sez_value": 0.0,
        "igst": 0.0,
        "cgst": 0.0,
        "sgst": 0.0,
    }
    filing_month = "Unknown"

    def _col_idx(headers: list[str], keywords: list[str]) -> int | None:
        for i, h in enumerate(headers):
            h_l = str(h).lower().strip()
            if any(kw in h_l for kw in keywords):
                return i
        return None

    def _val(row: list, idx: int | None) -> float:
        if idx is None or idx >= len(row) or row[idx] is None:
            return 0.0
        return safe_float(row[idx])

    try:
        uploaded_file.seek(0)
        with pdfplumber.open(uploaded_file) as pdf:
            # Detect month from first page
            first_text = pdf.pages[0].extract_text() or ""
            filing_month = _detect_filing_month(first_text)

            for page in pdf.pages:
                tables = page.extract_tables() or []
                for table in tables:
                    if len(table) < 2:
                        continue
                    headers = [str(c or "").lower().strip() for c in table[0]]

                    # Column indices
                    igst_i  = _col_idx(headers, ["igst", "integrated"])
                    cgst_i  = _col_idx(headers, ["cgst", "central tax"])
                    sgst_i  = _col_idx(headers, ["sgst", "state tax", "utgst"])
                    val_i   = _col_idx(headers, ["taxable value", "taxable amount",
                                                  "value of supply", "total value"])
                    exp_i   = _col_idx(headers, ["export", "zero rated"])
                    sez_i   = _col_idx(headers, ["sez"])

                    is_export_table = any("export" in h for h in headers)
                    is_sez_table    = any("sez" in h for h in headers)

                    for row in table[1:]:
                        if not any(row):
                            continue
                        totals["igst"]        += _val(row, igst_i)
                        totals["cgst"]        += _val(row, cgst_i)
                        totals["sgst"]        += _val(row, sgst_i)
                        base_val               = _val(row, val_i)
                        totals["sales_value"] += base_val
                        if is_export_table or exp_i:
                            totals["export_value"] += _val(row, exp_i) or base_val
                        if is_sez_table or sez_i:
                            totals["sez_value"]    += _val(row, sez_i) or base_val

    except Exception as exc:
        logger.error("GSTR-1 PDF parse error: %s", exc)

    return {filing_month: totals}


# ════════════════════════════════════════════════════════════════════════════
#  Electronic Credit Ledger parser (Module 2)
# ════════════════════════════════════════════════════════════════════════════

def parse_credit_ledger(uploaded_file) -> dict:
    """
    Parse the Electronic Credit Ledger (ITC Ledger) downloaded from GST portal.

    As per spec: Column F contains 'Credit' (ITC Availed) or 'Debit' (ITC Utilized).
    However, the parser also handles named 'Credit'/'Debit' columns.

    Returns:
        { month: { igst_credit, cgst_credit, sgst_credit,
                   igst_debit,  cgst_debit,  sgst_debit } }
    """
    result: dict = {}

    try:
        # ── Try Excel ──────────────────────────────────────────────────────
        uploaded_file.seek(0)
        raw = pd.read_excel(uploaded_file, header=None, dtype=str)
        df = _clean_df(raw)
        col_map = map_columns(df)

        # ── Locate Credit/Debit type column ───────────────────────────────
        type_col: str | None = None

        # 1. Named column: "Credit", "Debit", "Transaction Type", "Type", "Cr/Dr"
        for col in df.columns:
            col_l = col.lower().strip()
            if col_l in {"credit", "debit", "type", "cr/dr", "transaction type",
                         "nature of transaction"}:
                type_col = col
                break

        # 2. Column F (index 5) per spec
        if type_col is None and len(df.columns) > 5:
            col_f = df.columns[5]
            uniq = df[col_f].dropna().astype(str).str.strip().str.lower().unique()
            if set(uniq) & {"credit", "debit", "cr", "dr"}:
                type_col = col_f

        # 3. Any column whose unique values are only credit/debit-like
        if type_col is None:
            for col in df.columns:
                uniq = df[col].dropna().astype(str).str.strip().str.lower().unique()
                if len(uniq) <= 4 and set(uniq) & {"credit", "debit", "cr", "dr"}:
                    type_col = col
                    break

        if type_col is None or "date" not in col_map:
            logger.warning("Credit Ledger: could not detect type or date column.")
            return {}

        df["_month"] = df[col_map["date"]].apply(normalize_month)
        df = df.dropna(subset=["_month"])

        credit_like = {"credit", "cr"}
        debit_like  = {"debit",  "dr"}

        for month, group in df.groupby("_month"):
            cred_rows = group[
                group[type_col].astype(str).str.strip().str.lower().isin(credit_like)
            ]
            deb_rows = group[
                group[type_col].astype(str).str.strip().str.lower().isin(debit_like)
            ]
            month_data: dict = {}
            for tax in TAX_FIELDS:
                col = col_map.get(tax)
                month_data[f"{tax}_credit"] = (
                    cred_rows[col].apply(safe_float).sum() if col else 0.0
                )
                month_data[f"{tax}_debit"] = (
                    deb_rows[col].apply(safe_float).sum() if col else 0.0
                )
            result[month] = month_data

    except Exception as exc:
        logger.error("Credit Ledger parse error: %s", exc)

    return result


# ════════════════════════════════════════════════════════════════════════════
#  GSTR-2B Excel parser (Module 3 & 4)
# ════════════════════════════════════════════════════════════════════════════

def _sheet_name_lookup(xl: pd.ExcelFile, targets: list[str]) -> str | None:
    """Return the first sheet name whose lowercase form contains any target."""
    for sheet in xl.sheet_names:
        sheet_l = sheet.lower().replace(" ", "").replace("-", "")
        for t in targets:
            if t.lower().replace("-", "").replace(" ", "") in sheet_l:
                return sheet
    return None


def parse_gstr2b_excel(uploaded_file) -> dict:
    """
    Parse a GSTR-2B Excel file (standard GST portal MIS download).

    Looks for sheets:
      • B2B       → Add IGST, CGST, SGST
      • B2B-CDNR  → Debit Notes: Add  /  Credit Notes: Subtract
      • IMPZ      → Add IGST, CGST, SGST

    Returns:
        {
          "b2b":  { month: {igst, cgst, sgst} },
          "cdnr": { month: {igst, cgst, sgst} },   # already sign-adjusted
          "impz": { month: {igst, cgst, sgst} },
        }
    Also cached raw line items for Module 4 invoice matching.
    """
    result = {"b2b": {}, "cdnr": {}, "impz": {}, "_b2b_lines": pd.DataFrame()}

    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)

        b2b_sheet  = _sheet_name_lookup(xl, ["b2b"])
        cdnr_sheet = _sheet_name_lookup(xl, ["b2bcdnr", "cdnr"])
        impz_sheet = _sheet_name_lookup(xl, ["impz", "impg"])

        # ── Helper: parse a simple additive sheet ─────────────────────────
        def _agg_sheet(sheet_name: str | None) -> dict:
            if not sheet_name:
                return {}
            uploaded_file.seek(0)
            df = parse_excel_generic(uploaded_file, sheet_name=sheet_name)
            if df.empty:
                return {}
            col_map = map_columns(df)
            df_m = _add_month_col(df, col_map)
            if df_m is None or df_m.empty:
                return {}
            out: dict = {}
            for month, grp in df_m.groupby("_month"):
                row: dict = {}
                for tax in TAX_FIELDS:
                    col = col_map.get(tax)
                    row[tax] = grp[col].apply(safe_float).sum() if col else 0.0
                out[month] = row
            return out

        # ── B2B ────────────────────────────────────────────────────────────
        result["b2b"] = _agg_sheet(b2b_sheet)

        # Cache raw B2B lines for Module 4
        if b2b_sheet:
            uploaded_file.seek(0)
            result["_b2b_lines"] = parse_excel_generic(uploaded_file, sheet_name=b2b_sheet)

        # ── IMPZ ───────────────────────────────────────────────────────────
        result["impz"] = _agg_sheet(impz_sheet)

        # ── B2B-CDNR (sign-aware) ──────────────────────────────────────────
        if cdnr_sheet:
            uploaded_file.seek(0)
            df_cdnr = parse_excel_generic(uploaded_file, sheet_name=cdnr_sheet)
            if not df_cdnr.empty:
                col_map = map_columns(df_cdnr)
                df_m = _add_month_col(df_cdnr, col_map)
                note_type_col = col_map.get("note_type")

                # Fallback: find any column with only 'debit'/'credit' values
                if not note_type_col:
                    for col in df_cdnr.columns:
                        uniq = (
                            df_cdnr[col]
                            .dropna()
                            .astype(str)
                            .str.strip()
                            .str.lower()
                            .unique()
                        )
                        if len(uniq) <= 4 and set(uniq) & {"debit", "credit", "d", "c"}:
                            note_type_col = col
                            break

                if df_m is not None and not df_m.empty:
                    cdnr_out: dict = {}
                    for month, grp in df_m.groupby("_month"):
                        row: dict = {t: 0.0 for t in TAX_FIELDS}
                        for _, line in grp.iterrows():
                            sign = 1  # default: debit → add
                            if note_type_col:
                                nt = str(line.get(note_type_col, "")).strip().lower()
                                sign = -1 if "credit" in nt else 1
                            for tax in TAX_FIELDS:
                                col = col_map.get(tax)
                                if col:
                                    row[tax] += sign * safe_float(line.get(col, 0))
                        cdnr_out[month] = row
                    result["cdnr"] = cdnr_out

    except Exception as exc:
        logger.error("GSTR-2B parse error: %s", exc)

    return result


# ════════════════════════════════════════════════════════════════════════════
#  Line-item extractor for Module 4 (invoice-level matching)
# ════════════════════════════════════════════════════════════════════════════

def extract_invoice_lines(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise a register DataFrame to a standard line-item frame with:
        gstin | invoice_no | igst | cgst | sgst | total_value | source
    Used by Module 4 for matching.
    """
    if df.empty:
        return pd.DataFrame(
            columns=["gstin", "invoice_no", "igst", "cgst", "sgst", "total_value"]
        )
    col_map = map_columns(df)
    out = pd.DataFrame()
    out["gstin"]       = df[col_map["gstin"]].astype(str).str.strip().str.upper() \
        if "gstin" in col_map else ""
    out["invoice_no"]  = df[col_map["invoice_no"]].astype(str).str.strip().str.upper() \
        if "invoice_no" in col_map else ""
    for tax in TAX_FIELDS:
        out[tax] = df[col_map[tax]].apply(safe_float) if tax in col_map else 0.0
    out["total_value"] = df[col_map["total_value"]].apply(safe_float) \
        if "total_value" in col_map else (out["igst"] + out["cgst"] + out["sgst"])
    return out.reset_index(drop=True)
