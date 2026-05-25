"""Home page for the GST Reconciliation Tool."""
import streamlit as st


def render_home():
    st.title("Welcome to the GST Reconciliation Tool 🇮🇳")
    st.markdown("""
    This tool automates the reconciliation of GST data across your Books and the GST Portal.
    Use the **left sidebar** to navigate between modules.

    ---
    """)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        ### 📤 Module 1 — Outward Supplies
        Reconcile **Sales** and **Credit Notes** from your books against the **GSTR-1** filed on the portal.
        - Month-wise view: Books (net of CN) vs GSTR-1
        - Bifurcation: IGST, CGST, SGST, Total Tax
        - Highlights Export and SEZ supplies separately

        ---

        ### 🔄 Module 3 — ITC vs GSTR-2B
        Reconcile **Input Tax Credit** in books against what suppliers have uploaded in **GSTR-2B**.
        - Processes B2B, B2B-CDNR (Debit/Credit note logic), and IMPZ sheets
        - Net of Debit Notes from Purchase/Journal registers
        """)

    with col2:
        st.markdown("""
        ### 📥 Module 2 — ITC Availment
        Reconcile ITC in books (Purchase + Journal Register) against the **Electronic Credit Ledger**.
        - Column F logic: Credit = ITC Availed | Debit = ITC Utilized
        - Shows Utilized amounts separately for reference

        ---

        ### 🔍 Module 4 — Invoice-Level Match
        Granular line-by-line matching of **Purchase/Journal Register** vs **GSTR-2B**.
        - Match key: GSTIN + Invoice Number
        - Four output buckets:
          - ✅ **Matched** — both present, amounts agree
          - 📋 **Not in Books** — in GSTR-2B but not in your registers
          - ❌ **Not in GSTR-2B** — in Books but supplier hasn't uploaded
          - ⚠️ **Amount Mismatch** — matched key, different amounts
        """)

    st.divider()
    st.markdown("""
    ### 📁 Supported File Formats
    | Input | Format |
    |-------|--------|
    | Sales / Purchase / Journal Register | `.xlsx`, `.xls`, `.pdf` |
    | Credit Note / Debit Note Register | `.xlsx`, `.xls`, `.pdf` |
    | GSTR-1 | `.pdf` |
    | Electronic Credit Ledger | `.xlsx`, `.xls`, `.pdf` |
    | GSTR-2B | `.xlsx`, `.xls` |

    ### 🔑 Column Keyword Mapping (auto-detected)
    | Field | Recognised Keywords |
    |-------|---------------------|
    | Sales Value | `Sale`, `Job work` |
    | Export Value | `Export` |
    | SEZ Value | `SEZ` |
    | IGST | `IGST`, `Integrated Tax`, `GST-Integrated` |
    | CGST | `CGST`, `Central Tax`, `GST- Central` |
    | SGST | `SGST`, `State Tax`, `GST- State` |
    """)
