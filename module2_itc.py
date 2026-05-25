"""Module 2 — ITC Availment Reconciliation
Books (Purchase + Journal - Debit Notes) vs Electronic Credit Ledger
Column F: Credit = ITC Availed | Debit = ITC Utilized
"""
from __future__ import annotations
import io
import pandas as pd
import streamlit as st
from .file_parser import parse_excel_generic, extract_books_monthwise, parse_credit_ledger
from .utils import merge_monthwise, subtract_monthwise, sort_months_fiscal, format_inr, format_diff, safe_float, TAX_FIELDS
from .ui_helpers import render_summary_metrics


def _split_ledger(ledger: dict) -> tuple[dict, dict]:
    availed  = {m: {t: safe_float(d.get(f"{t}_credit", 0)) for t in TAX_FIELDS} for m, d in ledger.items()}
    utilized = {m: {t: safe_float(d.get(f"{t}_debit",  0)) for t in TAX_FIELDS} for m, d in ledger.items()}
    return availed, utilized


def _render_table(net_books, availed, utilized):
    all_months = sort_months_fiscal(list(set(list(net_books) + list(availed))))
    cols = ["Description", "IGST", "CGST", "SGST", "Total Tax"]

    def row(lbl, d):
        i,c,s = safe_float(d.get("igst",0)), safe_float(d.get("cgst",0)), safe_float(d.get("sgst",0))
        return [lbl, format_inr(i), format_inr(c), format_inr(s), format_inr(i+c+s)]

    def diff_row(b, p):
        diffs = [safe_float(b.get(t,0)) - safe_float(p.get(t,0)) for t in TAX_FIELDS]
        return ["🔺 Difference"] + [format_diff(d) for d in diffs] + [format_diff(sum(diffs))]

    for month in all_months:
        b, a, u = net_books.get(month,{}), availed.get(month,{}), utilized.get(month,{})
        with st.expander(f"📅  **{month}**", expanded=False):
            rows = [
                row("📚 Books ITC (Net of DN)", b),
                row("🌐 Credit Ledger — Availed", a),
                diff_row(b, a),
            ]
            if any(safe_float(u.get(t,0)) > 0 for t in TAX_FIELDS):
                rows.append(row("ℹ️ ITC Utilized (Debit entries)", u))
            st.dataframe(pd.DataFrame(rows, columns=cols), use_container_width=True, hide_index=True)

    st.divider(); st.subheader("📊 Grand Total")
    tb = {t: sum(safe_float(net_books.get(m,{}).get(t,0)) for m in all_months) for t in TAX_FIELDS}
    ta = {t: sum(safe_float(availed.get(m,{}).get(t,0))   for m in all_months) for t in TAX_FIELDS}
    tu = {t: sum(safe_float(utilized.get(m,{}).get(t,0))  for m in all_months) for t in TAX_FIELDS}
    st.dataframe(pd.DataFrame([
        row("📚 Total Books ITC", tb),
        row("🌐 Total ITC Availed", ta),
        ["🔺 Net Diff"] + [format_diff(safe_float(tb[t])-safe_float(ta[t])) for t in TAX_FIELDS]
          + [format_diff(sum(safe_float(tb[t])-safe_float(ta[t]) for t in TAX_FIELDS))],
        row("ℹ️ Total ITC Utilized", tu),
    ], columns=cols), use_container_width=True, hide_index=True)


def _export(nb, av, ut, fname):
    rows = []
    for m in sort_months_fiscal(list(set(list(nb)+list(av)))):
        for lbl, d in [("Books ITC",nb.get(m,{})),("ITC Availed",av.get(m,{})),("ITC Utilized",ut.get(m,{}))]:
            rows.append({"Month":m,"Source":lbl,**{t.upper():safe_float(d.get(t,0)) for t in TAX_FIELDS},
                         "Total":sum(safe_float(d.get(t,0)) for t in TAX_FIELDS)})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(rows).to_excel(w, index=False, sheet_name="ITC_Availment")
    buf.seek(0)
    st.download_button("⬇️ Export to Excel", data=buf, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_module2():
    st.header("📥 Module 2 — ITC Availment Reconciliation")
    st.markdown("Reconcile **ITC in Books** (net of Debit Notes) against the **Electronic Credit Ledger**.")
    st.info("📌 **Column F Logic** — Credit = ITC Availed | Debit = ITC Utilized (shown informally)")
    st.divider()

    cb, cp = st.columns(2, gap="large")
    with cb:
        st.subheader("📚 Books Data")
        pur_f = st.file_uploader("Purchase Register (Excel)", type=["xlsx","xls"], key="m2_pur")
        jnl_f = st.file_uploader("Journal Register (Excel)",  type=["xlsx","xls"], key="m2_jnl")
        dn_f  = st.file_uploader("Debit Note Register *(optional)*", type=["xlsx","xls"], key="m2_dn")
    with cp:
        st.subheader("🌐 Portal Data")
        led_f = st.file_uploader("Electronic Credit Ledger (Excel/PDF)", type=["xlsx","xls","pdf"], key="m2_led")
        st.info("💡 Download from: GST Portal → Services → Ledgers → Electronic Credit Ledger")

    st.divider()
    if st.button("🔄  Run Reconciliation — Module 2", type="primary", key="m2_run"):
        if not pur_f and not jnl_f:
            st.error("Upload at least the Purchase or Journal Register.")
            return
        raws = []
        for f, lbl in [(pur_f,"Purchase"),(jnl_f,"Journal")]:
            if f:
                with st.spinner(f"Parsing {lbl} Register…"):
                    raws.append(extract_books_monthwise(parse_excel_generic(f)))
        gross = merge_monthwise(*raws) if raws else {}
        dn_data = {}
        if dn_f:
            with st.spinner("Parsing Debit Notes…"):
                dn_data = extract_books_monthwise(parse_excel_generic(dn_f))
        net_books = subtract_monthwise(gross, dn_data)
        ledger = {}
        if led_f:
            with st.spinner("Parsing Credit Ledger…"):
                ledger = parse_credit_ledger(led_f)
        availed, utilized = _split_ledger(ledger)
        st.session_state.update({"m2_nb":net_books,"m2_av":availed,"m2_ut":utilized})
        # Also share net_books for modules 3 & 4
        st.session_state["books_itc_net"] = net_books
        st.success("✅ Processing complete!")
        _show(net_books, availed, utilized)
    elif "m2_nb" in st.session_state:
        _show(st.session_state["m2_nb"], st.session_state["m2_av"], st.session_state["m2_ut"])


def _show(nb, av, ut):
    if not nb and not av:
        st.info("No data available.")
        return
    st.subheader("📊 ITC Reconciliation Results")
    render_summary_metrics(nb, av)
    st.divider()
    _render_table(nb, av, ut)
    st.divider()
    _export(nb, av, ut, "module2_itc_availment.xlsx")
