"""
ui_helpers.py — Shared Streamlit rendering utilities.

Provides:
  - render_month_recon_table()  : vertical month-wise Books vs Portal view
  - render_summary_metrics()    : top-level KPI cards
  - render_export_button()      : download as Excel
"""

from __future__ import annotations

import io
import pandas as pd
import streamlit as st

from .utils import (
    format_inr,
    format_diff,
    safe_float,
    sort_months_fiscal,
    TAX_FIELDS,
    VALUE_FIELDS,
)


# ════════════════════════════════════════════════════════════════════════════
#  Month-wise reconciliation table
# ════════════════════════════════════════════════════════════════════════════

def render_month_recon_table(
    books_data: dict,
    portal_data: dict,
    label_books: str = "Books",
    label_portal: str = "Portal",
    show_value_cols: bool = True,
    extra_rows: dict | None = None,   # e.g. {"ITC Utilized": {...}}
) -> None:
    """
    Render a vertical, month-wise, expandable reconciliation table.

    Parameters
    ----------
    books_data   : { month: {field: value} }
    portal_data  : { month: {field: value} }
    label_books  : display name for books row
    label_portal : display name for portal row
    show_value_cols : whether to include Sales/Export/SEZ columns
    extra_rows   : additional informational rows to append (month-keyed)
    """
    all_months = set(list(books_data.keys()) + list(portal_data.keys()))
    if not all_months:
        st.info("No data to display. Please upload and process files first.")
        return

    # Build column header list
    tax_cols  = ["IGST", "CGST", "SGST", "Total Tax"]
    val_cols  = ["Sales Value", "Export Value", "SEZ Value"] if show_value_cols else []
    all_cols  = ["Description"] + val_cols + tax_cols

    sorted_months = sort_months_fiscal(list(all_months))

    # Aggregated totals for summary
    total_books  = {f: 0.0 for f in VALUE_FIELDS}
    total_portal = {f: 0.0 for f in VALUE_FIELDS}
    total_diff   = {f: 0.0 for f in VALUE_FIELDS}

    for month in sorted_months:
        b = books_data.get(month, {})
        p = portal_data.get(month, {})

        # Accumulate totals
        for f in VALUE_FIELDS:
            total_books[f]  += safe_float(b.get(f, 0))
            total_portal[f] += safe_float(p.get(f, 0))
            total_diff[f]   += safe_float(b.get(f, 0)) - safe_float(p.get(f, 0))

        with st.expander(f"📅  **{month}**", expanded=False):
            rows = []

            def _build_row(label: str, data: dict) -> list:
                row = [label]
                if show_value_cols:
                    row += [
                        format_inr(data.get("sales_value", 0)),
                        format_inr(data.get("export_value", 0)),
                        format_inr(data.get("sez_value", 0)),
                    ]
                i = safe_float(data.get("igst", 0))
                c = safe_float(data.get("cgst", 0))
                s = safe_float(data.get("sgst", 0))
                row += [format_inr(i), format_inr(c), format_inr(s), format_inr(i+c+s)]
                return row

            def _build_diff_row(b_data: dict, p_data: dict) -> list:
                row = ["🔺 Difference"]
                if show_value_cols:
                    for f in ["sales_value", "export_value", "sez_value"]:
                        row.append(format_diff(
                            safe_float(b_data.get(f, 0)) - safe_float(p_data.get(f, 0))
                        ))
                for tax in TAX_FIELDS:
                    row.append(format_diff(
                        safe_float(b_data.get(tax, 0)) - safe_float(p_data.get(tax, 0))
                    ))
                b_total = sum(safe_float(b_data.get(t, 0)) for t in TAX_FIELDS)
                p_total = sum(safe_float(p_data.get(t, 0)) for t in TAX_FIELDS)
                row.append(format_diff(b_total - p_total))
                return row

            rows.append(_build_row(f"📚 {label_books}", b))
            rows.append(_build_row(f"🌐 {label_portal}", p))
            rows.append(_build_diff_row(b, p))

            # Extra informational rows (e.g. ITC Utilized)
            if extra_rows and month in extra_rows:
                for row_label, row_data in extra_rows[month].items():
                    rows.append(_build_row(f"ℹ️ {row_label}", row_data))

            df_display = pd.DataFrame(rows, columns=all_cols)
            st.dataframe(df_display, use_container_width=True, hide_index=True)

    # ── Grand total summary row ──────────────────────────────────────────
    st.divider()
    st.subheader("📊 Grand Total Summary")

    summary_rows = []

    def _total_row(label: str, data: dict) -> list:
        row = [label]
        if show_value_cols:
            row += [format_inr(data.get(f, 0)) for f in
                    ["sales_value", "export_value", "sez_value"]]
        i = safe_float(data.get("igst", 0))
        c = safe_float(data.get("cgst", 0))
        s = safe_float(data.get("sgst", 0))
        row += [format_inr(i), format_inr(c), format_inr(s), format_inr(i+c+s)]
        return row

    def _total_diff_row(data: dict) -> list:
        row = ["🔺 Net Difference"]
        if show_value_cols:
            row += [format_diff(data.get(f, 0)) for f in
                    ["sales_value", "export_value", "sez_value"]]
        for t in TAX_FIELDS:
            row.append(format_diff(data.get(t, 0)))
        row.append(format_diff(sum(data.get(t, 0) for t in TAX_FIELDS)))
        return row

    summary_rows.append(_total_row(f"📚 Total {label_books}", total_books))
    summary_rows.append(_total_row(f"🌐 Total {label_portal}", total_portal))
    summary_rows.append(_total_diff_row(total_diff))

    df_summary = pd.DataFrame(summary_rows, columns=all_cols)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)


# ════════════════════════════════════════════════════════════════════════════
#  Summary metric cards
# ════════════════════════════════════════════════════════════════════════════

def render_summary_metrics(
    books_data: dict,
    portal_data: dict,
) -> None:
    """Display four KPI metric cards at the top of a module."""
    def _total_tax(d: dict) -> float:
        return sum(
            sum(safe_float(v.get(t, 0)) for t in TAX_FIELDS)
            for v in d.values()
        )

    books_tax  = _total_tax(books_data)
    portal_tax = _total_tax(portal_data)
    diff_tax   = books_tax - portal_tax
    diff_pct   = (diff_tax / books_tax * 100) if books_tax else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📚 Books Total Tax",   format_inr(books_tax))
    c2.metric("🌐 Portal Total Tax",  format_inr(portal_tax))
    c3.metric("🔺 Difference",        format_inr(abs(diff_tax)),
              delta=f"{diff_pct:+.2f}%",
              delta_color="inverse")
    c4.metric("📅 Months Covered", len(set(list(books_data.keys()) + list(portal_data.keys()))))


# ════════════════════════════════════════════════════════════════════════════
#  Excel export
# ════════════════════════════════════════════════════════════════════════════

def render_export_button(
    books_data: dict,
    portal_data: dict,
    label_books: str = "Books",
    label_portal: str = "Portal",
    filename: str = "gst_reconciliation.xlsx",
) -> None:
    """Offer a download button to export reconciliation as Excel."""

    sorted_months = sort_months_fiscal(
        list(set(list(books_data.keys()) + list(portal_data.keys())))
    )
    rows = []
    for month in sorted_months:
        b = books_data.get(month, {})
        p = portal_data.get(month, {})
        for label, data in [(label_books, b), (label_portal, p)]:
            rows.append({
                "Month":       month,
                "Source":      label,
                "Sales Value": safe_float(data.get("sales_value", 0)),
                "Export":      safe_float(data.get("export_value", 0)),
                "SEZ":         safe_float(data.get("sez_value", 0)),
                "IGST":        safe_float(data.get("igst", 0)),
                "CGST":        safe_float(data.get("cgst", 0)),
                "SGST":        safe_float(data.get("sgst", 0)),
                "Total Tax":   sum(safe_float(data.get(t, 0)) for t in TAX_FIELDS),
            })
        # Diff row
        diff = {
            f: safe_float(b.get(f, 0)) - safe_float(p.get(f, 0))
            for f in VALUE_FIELDS
        }
        rows.append({
            "Month":       month,
            "Source":      "Difference",
            "Sales Value": diff.get("sales_value", 0),
            "Export":      diff.get("export_value", 0),
            "SEZ":         diff.get("sez_value", 0),
            "IGST":        diff.get("igst", 0),
            "CGST":        diff.get("cgst", 0),
            "SGST":        diff.get("sgst", 0),
            "Total Tax":   sum(diff.get(t, 0) for t in TAX_FIELDS),
        })

    df_out = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Reconciliation")
    buf.seek(0)

    st.download_button(
        label="⬇️ Export Reconciliation to Excel",
        data=buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
