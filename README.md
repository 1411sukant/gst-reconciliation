# 📊 GST Reconciliation Tool

An automated, multi-module GST reconciliation web application built with **Streamlit** and **Pandas**.

---

## 🚀 Modules

| Module | Purpose | Inputs |
|--------|---------|--------|
| **Module 1** | Outward Supplies — Books vs GSTR-1 | Sales Register, Credit Notes, GSTR-1 PDF |
| **Module 2** | ITC Availment — Books vs Credit Ledger | Purchase + Journal Register, Debit Notes, Electronic Credit Ledger |
| **Module 3** | ITC vs GSTR-2B | Same as Module 2 + GSTR-2B Excel |
| **Module 4** | Invoice-Level Match | Purchase/Journal Register + GSTR-2B |

---

## ⚙️ Setup & Installation

### Prerequisites
- Python 3.9 or higher
- pip
- Git

### Step 1 — Clone the repository
```bash
git clone https://github.com/YOUR_USERNAME/gst-reconciliation.git
cd gst-reconciliation
```

### Step 2 — Create a virtual environment
```bash
python -m venv venv

# Activate on Windows
venv\Scripts\activate

# Activate on Mac/Linux
source venv/bin/activate
```

### Step 3 — Install dependencies
```bash
pip install -r requirements.txt
```

### Step 4 — Run the app
```bash
streamlit run app.py
```
The app opens automatically at **http://localhost:8501**

---

## 📁 File Format Guide

### Sales / Purchase / Journal Register (Excel)
Columns auto-detected by keywords (case-insensitive):

| Field | Accepted Column Names |
|-------|-----------------------|
| Sales Value | `Sale`, `Job work`, `Taxable Value` |
| Export Value | `Export` |
| SEZ Value | `SEZ` |
| IGST | `IGST`, `Integrated Tax`, `GST-Integrated` |
| CGST | `CGST`, `Central Tax`, `GST- Central` |
| SGST | `SGST`, `State Tax`, `GST- State` |
| Date | `Date`, `Invoice Date`, `Bill Date` |
| Invoice No | `Invoice No`, `Bill No`, `Voucher No` |
| GSTIN | `GSTIN`, `GST No`, `Supplier GSTIN` |

### GSTR-2B Excel
Standard MIS download from GST portal. Must contain sheets named **B2B**, **B2B-CDNR**, **IMPZ**.

### Electronic Credit Ledger
Standard download from GST Portal → Services → Ledgers.  
Column F: `Credit` = ITC Availed | `Debit` = ITC Utilized.

---

## 🌐 Deploy on Streamlit Cloud (Free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New App** → select your repo → set main file to `app.py`
4. Click **Deploy** — your app will be live in ~2 minutes

---

## 📦 Project Structure
```
gst-reconciliation/
├── app.py                      # Main Streamlit entry point
├── requirements.txt
├── README.md
└── modules/
    ├── __init__.py
    ├── home.py                 # Home/landing page
    ├── utils.py                # Keyword mapping, month normalisation, formatting
    ├── file_parser.py          # Excel + PDF parsing for all file types
    ├── ui_helpers.py           # Shared table/chart/export components
    ├── module1_outward.py      # Outward Supplies reconciliation
    ├── module2_itc.py          # ITC Availment reconciliation
    ├── module3_gstr2b.py       # ITC vs GSTR-2B reconciliation
    └── module4_invoice.py      # Invoice-level matching (4 buckets)
```
