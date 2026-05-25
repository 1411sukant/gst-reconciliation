"""
GST Reconciliation Tool — Main Application
Run with: streamlit run app.py
"""

import streamlit as st

# ── Page configuration (must be first Streamlit call) ──────────────────────
st.set_page_config(
    page_title="GST Reconciliation Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a3c6b 0%, #2563a8 100%);
        padding: 1.5rem 2rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 1.5rem;
    }
    .diff-positive { color: #dc2626; font-weight: 700; }
    .diff-negative { color: #2563a8; font-weight: 700; }
    .diff-zero     { color: #16a34a; font-weight: 700; }
    thead tr th { background-color: #1a3c6b !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 GST Reconciliation Tool</h1>
    <p style="margin:0; opacity:0.85;">
        Automated reconciliation across Books, GSTR-1, GSTR-2B, GSTR-3B &amp; Credit Ledger
    </p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar navigation ───────────────────────────────────────────────────────
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

# ── Route to modules ─────────────────────────────────────────────────────────
if page == "🏠 Home":
    from modules.home import render_home
    render_home()
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
