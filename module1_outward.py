"""
Module 1 — Outward Supplies Reconciliation
Books (Sales Register - Credit Notes) vs GSTR-1 (Portal PDF)
"""
from __future__ import annotations
import streamlit as st
from .file_parser import parse_excel_generic, extract_books_monthwise, parse_gstr1_pdf
from .utils import subtract_monthwise
from .ui_helpers import render_month_recon_table, render_summary_metrics, render_export_button


def render_module1() -> None:
    st.header("📤 Module 1 — Outward Supplies Reconciliation")
    st.markdown("Reconcile **GST as per Books** (net of Credit Notes) against **GSTR-1** filed on the portal.")
    st.divider()

    col_books, col_portal = st.columns(2, gap="large")
    with col_books:
        st.subheader("📚 Books Data")
        sales_file = st.file_uploader("Sales Register (Excel or PDF)", type=["xlsx","xls","pdf"], key="m1_sales")
        cn_file    = st.file_uploader("Credit Note Register *(optional)*", type=["xlsx","xls","pdf"], key="m1_cn")
    with col_portal:
        st.subheader("🌐 Portal Data")
        gstr1_file = st.file_uploader("GSTR-1 PDF *(from GST portal)*", type=["pdf"], key="m1_gstr1")
        st.info("💡 Upload the GSTR-1 PDF for each month separately, or a combined PDF. "
                "The tool auto-detects the filing period.")

    st.divider()

    if st.button("🔄  Run Reconciliation — Module 1", type="primary", key="m1_run"):
        if not sales_file:
            st.error("Please upload at least the Sales Register.")
            return

        with st.spinner("Parsing Sales Register…"):
            if sales_file.name.lower().endswith(".pdf"):
                books_raw = parse_gstr1_pdf(sales_file)
            else:
                books_raw = extract_books_monthwise(parse_excel_generic(sales_file))

        if not books_raw:
            st.warning("⚠️ Could not extract data from Sales Register. Check column names.")

        cn_data: dict = {}
        if cn_file:
            with st.spinner("Parsing Credit Note Register…"):
                if cn_file.name.lower().endswith(".pdf"):
                    cn_data = parse_gstr1_pdf(cn_file)
                else:
                    cn_data = extract_books_monthwise(parse_excel_generic(cn_file))

        net_books = subtract_monthwise(books_raw, cn_data)

        gstr1_data: dict = {}
        if gstr1_file:
            with st.spinner("Parsing GSTR-1 PDF…"):
                gstr1_data = parse_gstr1_pdf(gstr1_file)

        st.session_state["m1_net_books"] = net_books
        st.session_state["m1_gstr1"]     = gstr1_data
        st.success("✅ Processing complete!")
        _display_results(net_books, gstr1_data)

    elif "m1_net_books" in st.session_state:
        _display_results(st.session_state["m1_net_books"], st.session_state.get("m1_gstr1", {}))


def _display_results(net_books: dict, gstr1_data: dict) -> None:
    if not net_books and not gstr1_data:
        st.info("No data available. Please upload files and process.")
        return
    st.subheader("📊 Reconciliation Results")
    render_summary_metrics(net_books, gstr1_data)
    st.divider()
    render_month_recon_table(
        net_books, gstr1_data,
        label_books="Books (Net of CN)", label_portal="GSTR-1", show_value_cols=True,
    )
    st.divider()
    render_export_button(net_books, gstr1_data, "Books (Net of CN)", "GSTR-1", "module1_outward_recon.xlsx")
