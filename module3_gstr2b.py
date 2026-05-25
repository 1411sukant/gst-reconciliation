"""Module 3 — ITC Reconciliation vs GSTR-2B
Books (Purchase + Journal - Debit Notes) vs GSTR-2B
B2B: Add | B2B-CDNR: Debit=Add, Credit=Subtract | IMPZ: Add
"""
from __future__ import annotations
import io
import pandas as pd
import streamlit as st
from .file_parser import parse_excel_generic, extract_books_monthwise, parse_gstr2b_excel
from .utils import merge_monthwise, subtract_monthwise, sort_months_fiscal, format_inr, format_diff, safe_float, TAX_FIELDS
from .ui_helpers import render_summary_metrics


def _agg_gstr2b(gstr2b: dict) -> dict:
    """Combine B2B + CDNR (sign-adjusted) + IMPZ into one month→{igst,cgst,sgst} dict."""
    all_months = set()
    for key in ("b2b", "cdnr", "impz"):
        all_months.update(gstr2b.get(key, {}).keys())
    result = {}
    for m in all_months:
        row = {t: 0.0 for t in TAX_FIELDS}
        for key in ("b2b", "cdnr", "impz"):
            src = gstr2b.get(key, {}).get(m, {})
            for t in TAX_FIELDS:
                row[t] += safe_float(src.get(t, 0))
        result[m] = row
    return result


def _render_table(net_books, portal):
    all_months = sort_months_fiscal(list(set(list(net_books)+list(portal))))
    cols = ["Description","IGST","CGST","SGST","Total Tax"]

    def row(lbl, d):
        i,c,s = safe_float(d.get("igst",0)),safe_float(d.get("cgst",0)),safe_float(d.get("sgst",0))
        return [lbl, format_inr(i), format_inr(c), format_inr(s), format_inr(i+c+s)]

    def diff_row(b, p):
        diffs = [safe_float(b.get(t,0))-safe_float(p.get(t,0)) for t in TAX_FIELDS]
        return ["🔺 Difference"] + [format_diff(d) for d in diffs] + [format_diff(sum(diffs))]

    for month in all_months:
        b, p = net_books.get(month,{}), portal.get(month,{})
        with st.expander(f"📅  **{month}**", expanded=False):
            st.dataframe(pd.DataFrame([
                row("📚 Books ITC (Net of DN)", b),
                row("🌐 GSTR-2B (B2B+CDNR+IMPZ)", p),
                diff_row(b, p),
            ], columns=cols), use_container_width=True, hide_index=True)

    st.divider(); st.subheader("📊 Grand Total")
    tb = {t: sum(safe_float(net_books.get(m,{}).get(t,0)) for m in all_months) for t in TAX_FIELDS}
    tp = {t: sum(safe_float(portal.get(m,{}).get(t,0))    for m in all_months) for t in TAX_FIELDS}
    st.dataframe(pd.DataFrame([
        row("📚 Total Books ITC", tb),
        row("🌐 Total GSTR-2B",   tp),
        ["🔺 Net Diff"] + [format_diff(safe_float(tb[t])-safe_float(tp[t])) for t in TAX_FIELDS]
          + [format_diff(sum(safe_float(tb[t])-safe_float(tp[t]) for t in TAX_FIELDS))],
    ], columns=cols), use_container_width=True, hide_index=True)


def _export(nb, portal, fname):
    rows = []
    for m in sort_months_fiscal(list(set(list(nb)+list(portal)))):
        for lbl, d in [("Books ITC",nb.get(m,{})),("GSTR-2B",portal.get(m,{}))]:
            rows.append({"Month":m,"Source":lbl,**{t.upper():safe_float(d.get(t,0)) for t in TAX_FIELDS},
                         "Total":sum(safe_float(d.get(t,0)) for t in TAX_FIELDS)})
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as w:
        pd.DataFrame(rows).to_excel(w,index=False,sheet_name="ITC_vs_GSTR2B")
    buf.seek(0)
    st.download_button("⬇️ Export to Excel",data=buf,file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_module3():
    st.header("🔄 Module 3 — ITC Reconciliation vs GSTR-2B")
    st.markdown("Reconcile **Books ITC** (net of Debit Notes) against **GSTR-2B** supplier-uploaded data.")
    st.info("📌 **GSTR-2B Sheet Logic** — B2B: ➕ Add | B2B-CDNR: Debit ➕ Add, Credit ➖ Subtract | IMPZ: ➕ Add")
    st.divider()

    cb, cp = st.columns(2, gap="large")
    with cb:
        st.subheader("📚 Books Data")
        pur_f = st.file_uploader("Purchase Register (Excel)", type=["xlsx","xls"], key="m3_pur")
        jnl_f = st.file_uploader("Journal Register (Excel)",  type=["xlsx","xls"], key="m3_jnl")
        dn_f  = st.file_uploader("Debit Note Register *(optional)*", type=["xlsx","xls"], key="m3_dn")
        st.info("💡 Or re-use Module 2 data — if already processed, tick below.")
        reuse = st.checkbox("♻️ Reuse Books data from Module 2", value=True, key="m3_reuse")
    with cp:
        st.subheader("🌐 Portal Data")
        gstr2b_f = st.file_uploader("GSTR-2B Excel *(standard MIS download)*", type=["xlsx","xls"], key="m3_2b")
        st.info("💡 Download from: GST Portal → Returns → GSTR-2B → Download")

    st.divider()
    if st.button("🔄  Run Reconciliation — Module 3", type="primary", key="m3_run"):
        # Books ITC
        if reuse and "books_itc_net" in st.session_state:
            net_books = st.session_state["books_itc_net"]
            st.info("♻️ Reusing Books ITC data from Module 2.")
        else:
            if not pur_f and not jnl_f:
                st.error("Upload Purchase/Journal Register or run Module 2 first.")
                return
            raws = []
            for f, lbl in [(pur_f,"Purchase"),(jnl_f,"Journal")]:
                if f:
                    with st.spinner(f"Parsing {lbl}…"):
                        raws.append(extract_books_monthwise(parse_excel_generic(f)))
            gross = merge_monthwise(*raws) if raws else {}
            dn_data = {}
            if dn_f:
                with st.spinner("Parsing Debit Notes…"):
                    dn_data = extract_books_monthwise(parse_excel_generic(dn_f))
            net_books = subtract_monthwise(gross, dn_data)
            st.session_state["books_itc_net"] = net_books

        if not gstr2b_f:
            st.error("Please upload the GSTR-2B Excel file.")
            return
        with st.spinner("Parsing GSTR-2B…"):
            gstr2b_raw = parse_gstr2b_excel(gstr2b_f)
        portal = _agg_gstr2b(gstr2b_raw)

        # Cache raw lines for Module 4
        st.session_state["gstr2b_lines"] = gstr2b_raw.get("_b2b_lines", pd.DataFrame())
        st.session_state.update({"m3_nb":net_books,"m3_portal":portal})
        st.success("✅ Processing complete!")
        _show(net_books, portal)

    elif "m3_nb" in st.session_state:
        _show(st.session_state["m3_nb"], st.session_state["m3_portal"])


def _show(nb, portal):
    if not nb and not portal:
        st.info("No data available."); return
    st.subheader("📊 ITC vs GSTR-2B Results")
    render_summary_metrics(nb, portal)
    st.divider()
    _render_table(nb, portal)
    st.divider()
    _export(nb, portal, "module3_itc_vs_gstr2b.xlsx")
