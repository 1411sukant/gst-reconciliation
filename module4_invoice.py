"""Module 4 — Invoice-Level Reconciliation Report
Match Purchase/Journal Register vs GSTR-2B line-by-line
Key: GSTIN + Invoice Number
Buckets: Matched | Not in Books | Not in GSTR-2B | Amount Mismatch
"""
from __future__ import annotations
import io
import pandas as pd
import streamlit as st
from .file_parser import parse_excel_generic, extract_invoice_lines, parse_gstr2b_excel
from .utils import safe_float, TAX_FIELDS, format_inr


AMOUNT_TOLERANCE = 1.0   # ₹1 rounding tolerance


def _match_invoices(books_df: pd.DataFrame, gstr2b_df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Match books vs gstr2b on (gstin + invoice_no).
    Returns dict with keys: matched, not_in_books, not_in_gstr2b, amount_mismatch
    """
    empty = pd.DataFrame(columns=[
        "GSTIN","Invoice No","Books IGST","Books CGST","Books SGST",
        "GSTR-2B IGST","GSTR-2B CGST","GSTR-2B SGST",
        "IGST Diff","CGST Diff","SGST Diff",
    ])

    if books_df.empty and gstr2b_df.empty:
        return {k: empty.copy() for k in ("matched","not_in_books","not_in_gstr2b","amount_mismatch")}

    # Build key on both sides
    books_df   = books_df.copy()
    gstr2b_df  = gstr2b_df.copy()
    books_df["_key"]  = books_df["gstin"].str.upper().str.strip()  + "||" + \
                        books_df["invoice_no"].str.upper().str.strip()
    gstr2b_df["_key"] = gstr2b_df["gstin"].str.upper().str.strip() + "||" + \
                        gstr2b_df["invoice_no"].str.upper().str.strip()

    books_idx  = books_df.set_index("_key")
    gstr2b_idx = gstr2b_df.set_index("_key")

    books_keys  = set(books_idx.index)
    gstr2b_keys = set(gstr2b_idx.index)

    common_keys     = books_keys & gstr2b_keys
    only_in_books   = books_keys - gstr2b_keys
    only_in_gstr2b  = gstr2b_keys - books_keys

    matched_rows    = []
    mismatch_rows   = []

    for key in common_keys:
        b = books_idx.loc[key]
        g = gstr2b_idx.loc[key]
        # Handle duplicate keys (take first occurrence)
        if isinstance(b, pd.DataFrame): b = b.iloc[0]
        if isinstance(g, pd.DataFrame): g = g.iloc[0]

        diffs = {t: safe_float(b.get(t,0)) - safe_float(g.get(t,0)) for t in TAX_FIELDS}
        parts = key.split("||")
        row = {
            "GSTIN":       parts[0] if len(parts) > 0 else "",
            "Invoice No":  parts[1] if len(parts) > 1 else "",
            "Books IGST":  safe_float(b.get("igst",0)),
            "Books CGST":  safe_float(b.get("cgst",0)),
            "Books SGST":  safe_float(b.get("sgst",0)),
            "GSTR-2B IGST":safe_float(g.get("igst",0)),
            "GSTR-2B CGST":safe_float(g.get("cgst",0)),
            "GSTR-2B SGST":safe_float(g.get("sgst",0)),
            "IGST Diff":   diffs["igst"],
            "CGST Diff":   diffs["cgst"],
            "SGST Diff":   diffs["sgst"],
        }
        if all(abs(diffs[t]) <= AMOUNT_TOLERANCE for t in TAX_FIELDS):
            matched_rows.append(row)
        else:
            mismatch_rows.append(row)

    def _side_rows(keys, source_idx, label):
        rows = []
        for key in keys:
            r = source_idx.loc[key]
            if isinstance(r, pd.DataFrame): r = r.iloc[0]
            parts = key.split("||")
            is_books = (label == "books")
            rows.append({
                "GSTIN":        parts[0] if len(parts)>0 else "",
                "Invoice No":   parts[1] if len(parts)>1 else "",
                "Books IGST":   safe_float(r.get("igst",0)) if is_books else 0.0,
                "Books CGST":   safe_float(r.get("cgst",0)) if is_books else 0.0,
                "Books SGST":   safe_float(r.get("sgst",0)) if is_books else 0.0,
                "GSTR-2B IGST": safe_float(r.get("igst",0)) if not is_books else 0.0,
                "GSTR-2B CGST": safe_float(r.get("cgst",0)) if not is_books else 0.0,
                "GSTR-2B SGST": safe_float(r.get("sgst",0)) if not is_books else 0.0,
                "IGST Diff":0.0,"CGST Diff":0.0,"SGST Diff":0.0,
            })
        return rows

    def _to_df(rows):
        if rows:
            return pd.DataFrame(rows)
        return empty.copy()

    return {
        "matched":        _to_df(matched_rows),
        "not_in_books":   _to_df(_side_rows(only_in_gstr2b, gstr2b_idx, "gstr2b")),
        "not_in_gstr2b":  _to_df(_side_rows(only_in_books,  books_idx,  "books")),
        "amount_mismatch":_to_df(mismatch_rows),
    }


def _fmt_df(df: pd.DataFrame) -> pd.DataFrame:
    """Format numeric columns as INR strings for display."""
    df = df.copy()
    money_cols = [c for c in df.columns if any(x in c for x in ["IGST","CGST","SGST"])]
    for col in money_cols:
        df[col] = df[col].apply(lambda v: format_inr(safe_float(v)))
    return df


def _export_all(buckets: dict, fname: str):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        sheet_map = {
            "matched":        "Matched",
            "not_in_books":   "Not_In_Books",
            "not_in_gstr2b":  "Not_In_GSTR2B",
            "amount_mismatch":"Amount_Mismatch",
        }
        for key, sheet in sheet_map.items():
            df = buckets.get(key, pd.DataFrame())
            if df.empty:
                df = pd.DataFrame(columns=["No records"])
            df.to_excel(w, index=False, sheet_name=sheet)
    buf.seek(0)
    st.download_button("⬇️ Export All Buckets to Excel", data=buf, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_module4():
    st.header("🔍 Module 4 — Invoice-Level Reconciliation")
    st.markdown("Granular **line-by-line matching** of Books vs GSTR-2B. Match key: **GSTIN + Invoice Number**")
    st.divider()

    cb, cp = st.columns(2, gap="large")
    with cb:
        st.subheader("📚 Books Data")
        pur_f = st.file_uploader("Purchase Register (Excel)", type=["xlsx","xls"], key="m4_pur")
        jnl_f = st.file_uploader("Journal Register (Excel)",  type=["xlsx","xls"], key="m4_jnl")
    with cp:
        st.subheader("🌐 GSTR-2B")
        gstr2b_f = st.file_uploader("GSTR-2B Excel", type=["xlsx","xls"], key="m4_2b")
        st.info("💡 Or run Module 3 first — the parsed GSTR-2B lines are reused automatically.")
        reuse_2b = st.checkbox("♻️ Reuse GSTR-2B lines from Module 3", value=True, key="m4_reuse")

    st.divider()
    if st.button("🔍  Run Invoice Matching — Module 4", type="primary", key="m4_run"):
        # Books
        books_dfs = []
        for f, lbl in [(pur_f,"Purchase"),(jnl_f,"Journal")]:
            if f:
                with st.spinner(f"Parsing {lbl} Register…"):
                    books_dfs.append(extract_invoice_lines(parse_excel_generic(f)))
        if not books_dfs:
            st.error("Upload at least one Books register.")
            return
        books_lines = pd.concat(books_dfs, ignore_index=True)

        # GSTR-2B
        gstr2b_lines = pd.DataFrame()
        if reuse_2b and "gstr2b_lines" in st.session_state:
            gstr2b_lines = st.session_state["gstr2b_lines"]
            st.info("♻️ Reusing GSTR-2B lines from Module 3.")
        elif gstr2b_f:
            with st.spinner("Parsing GSTR-2B…"):
                raw = parse_gstr2b_excel(gstr2b_f)
                gstr2b_lines = raw.get("_b2b_lines", pd.DataFrame())
        if gstr2b_lines.empty:
            st.error("GSTR-2B data not available. Upload the file or run Module 3 first.")
            return

        gstr2b_inv = extract_invoice_lines(gstr2b_lines)

        with st.spinner("Matching invoices…"):
            buckets = _match_invoices(books_lines, gstr2b_inv)

        st.session_state["m4_buckets"] = buckets
        st.success("✅ Matching complete!")
        _show(buckets)

    elif "m4_buckets" in st.session_state:
        _show(st.session_state["m4_buckets"])


def _show(buckets: dict):
    matched    = buckets.get("matched",        pd.DataFrame())
    not_books  = buckets.get("not_in_books",   pd.DataFrame())
    not_2b     = buckets.get("not_in_gstr2b",  pd.DataFrame())
    mismatch   = buckets.get("amount_mismatch",pd.DataFrame())

    # Summary KPIs
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("✅ Matched",        len(matched))
    c2.metric("📋 Not in Books",   len(not_books))
    c3.metric("❌ Not in GSTR-2B", len(not_2b))
    c4.metric("⚠️ Amount Mismatch",len(mismatch))
    st.divider()

    tabs = st.tabs(["✅ Matched","📋 Not in Books","❌ Not in GSTR-2B","⚠️ Amount Mismatch"])

    with tabs[0]:
        st.subheader(f"✅ Matched Invoices ({len(matched)})")
        if matched.empty:
            st.info("No matched invoices.")
        else:
            st.dataframe(_fmt_df(matched), use_container_width=True, hide_index=True)

    with tabs[1]:
        st.subheader(f"📋 Not in Books — Present in GSTR-2B only ({len(not_books)})")
        st.caption("These invoices were uploaded by your suppliers in GSTR-2B but are missing from your registers.")
        if not_books.empty:
            st.info("None found.")
        else:
            st.dataframe(_fmt_df(not_books[["GSTIN","Invoice No","GSTR-2B IGST","GSTR-2B CGST","GSTR-2B SGST"]]),
                         use_container_width=True, hide_index=True)

    with tabs[2]:
        st.subheader(f"❌ Not in GSTR-2B — Present in Books only ({len(not_2b)})")
        st.caption("Your supplier has NOT uploaded these invoices. ITC may not be available.")
        if not_2b.empty:
            st.info("None found.")
        else:
            st.dataframe(_fmt_df(not_2b[["GSTIN","Invoice No","Books IGST","Books CGST","Books SGST"]]),
                         use_container_width=True, hide_index=True)

    with tabs[3]:
        st.subheader(f"⚠️ Amount Mismatch ({len(mismatch)})")
        st.caption("Invoice and GSTIN match, but the tax amounts differ between Books and GSTR-2B.")
        if mismatch.empty:
            st.info("No mismatches found.")
        else:
            st.dataframe(_fmt_df(mismatch), use_container_width=True, hide_index=True)

    st.divider()
    _export_all(buckets, "module4_invoice_level_recon.xlsx")
