"""
GST Reconciliation Tool — Main Application
Run with: streamlit run app.py
"""
import streamlit as st

st.set_page_config(
    page_title="GST Reconciliation Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a3c6b 0%, #2563a8 100%);
        padding: 1.5rem 2rem; border-radius: 10px;
        color: white; margin-bottom: 1.5rem;
    }
    thead tr th { background-color: #1a3c6b !important; color: white !important; }
    .stButton > button { background-color: #2563a8; color: white; border-radius: 6px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📊 GST Reconciliation Tool</h1>
    <p style="margin:0;opacity:0.85;">
        Automated reconciliation across Books, GSTR-1, GSTR-2B, GSTR-3B &amp; Credit Ledger
    </p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.title("Navigation")
    page = st.radio(
        "Select Module",
        options=[
            "🏠 Home",
            "📤 Module 1 — Outward Supplies",
            "📥 Module 2 — ITC Availment",
            "🔄 Module 3 — ITC vs GSTR-2B",
            "🔍 Module 4 — Invoice-Level Match",
        ],
        index=0,
    )
    st.divider()
    st.caption("📁 Supported: .xlsx, .xls, .pdf")
    st.caption("📅 FY Order: April → March")
    st.caption("v1.0 | Streamlit + Pandas")

# ── HOME PAGE (inlined — no external import needed) ──────────────────────────
if page == "🏠 Home":
    st.title("Welcome to the GST Reconciliation Tool 🇮🇳")
    st.markdown("Use the **left sidebar** to navigate between modules.")
    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
### 📤 Module 1 — Outward Supplies
Reconcile **Sales** and **Credit Notes** from your books against **GSTR-1**.
- Month-wise: Books (net of CN) vs GSTR-1
- IGST, CGST, SGST, Total Tax bifurcation
- Export and SEZ supplies shown separately

---
### 🔄 Module 3 — ITC vs GSTR-2B
Reconcile **ITC in Books** against supplier-uploaded **GSTR-2B**.
- B2B ➕ | B2B-CDNR Debit ➕ / Credit ➖ | IMPZ ➕
- Net of Debit Notes from registers
        """)
    with c2:
        st.markdown("""
### 📥 Module 2 — ITC Availment
Reconcile **ITC** (Purchase + Journal) against **Electronic Credit Ledger**.
- Column F: Credit = ITC Availed | Debit = ITC Utilized
- Utilized amounts shown separately

---
### 🔍 Module 4 — Invoice-Level Match
Line-by-line matching of Books vs GSTR-2B.
- Match key: **GSTIN + Invoice Number**
- ✅ Matched | 📋 Not in Books | ❌ Not in GSTR-2B | ⚠️ Amount Mismatch
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

elif page == "📤 Module 1 — Outward Supplies":
    from modules.module1_outward import render_module1
    render_module1()

elif page == "📥 Module 2 — ITC Availment":
    from modules.module2_itc import render_module2
    render_module2()

elif page == "🔄 Module 3 — ITC vs GSTR-2B":
    from modules.module3_gstr2b import render_module3
    render_module3()

elif page == "🔍 Module 4 — Invoice-Level Match":
    from modules.module4_invoice import render_module4
    render_module4()
