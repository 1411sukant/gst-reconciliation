"""
GST Reconciliation Tool — Single File Version
Run with: streamlit run app.py
"""
import io, re
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# ════════════════════════════════ UTILS ═════════════════════════════════════
FISCAL_MONTH_ORDER = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
MONTH_ABBREV = {
    "january":"Jan","february":"Feb","march":"Mar","april":"Apr","may":"May",
    "june":"Jun","july":"Jul","august":"Aug","september":"Sep","october":"Oct",
    "november":"Nov","december":"Dec","jan":"Jan","feb":"Feb","mar":"Mar",
    "apr":"Apr","jun":"Jun","jul":"Jul","aug":"Aug","sep":"Sep","oct":"Oct",
    "nov":"Nov","dec":"Dec","01":"Jan","02":"Feb","03":"Mar","04":"Apr",
    "05":"May","06":"Jun","07":"Jul","08":"Aug","09":"Sep","10":"Oct","11":"Nov","12":"Dec",
}
COLUMN_KEYWORDS = {
    "sales_value":  ["sale","job work","taxable value","taxable amount"],
    "export_value": ["export"],
    "sez_value":    ["sez"],
    "igst":         ["igst","integrated tax","gst-integrated","gst integrated"],
    "cgst":         ["cgst","central tax","gst- central","gst central"],
    "sgst":         ["sgst","state tax","gst- state","gst state","utgst"],
    "date":         ["date","invoice date","bill date","voucher date","doc date"],
    "invoice_no":   ["invoice no","invoice number","bill no","voucher no","doc no","inv no"],
    "gstin":        ["gstin","gst no","gst number","supplier gstin","party gstin"],
    "note_type":    ["note type","type","debit/credit","cr/dr","transaction type"],
    "total_value":  ["total value","invoice value","total amount","gross amount"],
}
TAX_FIELDS   = ["igst","cgst","sgst"]
VALUE_FIELDS = ["sales_value","export_value","sez_value","igst","cgst","sgst"]

def normalize_month(val):
    if val is None: return None
    try:
        if isinstance(val, float) and np.isnan(val): return None
    except: pass
    if isinstance(val, (datetime, pd.Timestamp)):
        return MONTH_ABBREV.get(val.strftime("%b").lower())
    s = str(val).strip()
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="raise")
        return MONTH_ABBREV.get(dt.strftime("%b").lower())
    except: pass
    low = s.lower()
    for k in sorted(MONTH_ABBREV, key=len, reverse=True):
        if len(k) >= 3 and k in low: return MONTH_ABBREV[k]
    return None

def map_columns(df):
    col_map = {}
    lowers = {c.lower().strip(): c for c in df.columns if isinstance(c, str)}
    for field, kws in COLUMN_KEYWORDS.items():
        for cl, co in lowers.items():
            if any(kw.lower() in cl for kw in kws): col_map.setdefault(field, co)
    return col_map

def safe_float(v):
    if v is None: return 0.0
    try:
        if pd.isna(v): return 0.0
    except: pass
    try: return float(str(v).replace(",","").replace(" ","").strip() or 0)
    except: return 0.0

def sort_months_fiscal(months):
    def k(m):
        try: return FISCAL_MONTH_ORDER.index(m)
        except: return 99
    return sorted(months, key=k)

def format_inr(v):
    try: v = float(v)
    except: return "0.00"
    if v == 0: return "0.00"
    neg = v < 0; v = abs(v)
    parts = f"{v:,.2f}".split(".")
    ip = parts[0].replace(",","")
    if len(ip) > 3:
        last3=ip[-3:]; rest=ip[:-3]; groups=[]
        while len(rest)>2: groups.append(rest[-2:]); rest=rest[:-2]
        if rest: groups.append(rest)
        groups.reverse(); ip=",".join(groups)+","+last3
    r = f"Rs.{ip}.{parts[1]}"
    return f"-{r}" if neg else r

def format_diff(v):
    v = safe_float(v)
    if abs(v) < 0.01: return "NIL"
    return f"{'(+)' if v>0 else '(-)'} {format_inr(abs(v))}"

def merge_monthwise(*dicts):
    all_months = set()
    for d in dicts: all_months.update(d.keys())
    result = {}
    for m in all_months:
        result[m] = {f:0.0 for f in VALUE_FIELDS}
        for d in dicts:
            for f in VALUE_FIELDS: result[m][f] += safe_float(d.get(m,{}).get(f,0))
    return result

def subtract_monthwise(base, sub):
    all_months = set(list(base)+list(sub))
    return {m:{f:safe_float(base.get(m,{}).get(f,0))-safe_float(sub.get(m,{}).get(f,0))
               for f in VALUE_FIELDS} for m in all_months}

# ════════════════════════════════ PARSERS ═══════════════════════════════════
def _find_header(raw):
    for i,row in raw.iterrows():
        if len(row.dropna())>=3: return int(i)
    return 0

def _clean_df(raw):
    h=_find_header(raw)
    cols=[str(c).strip() if not pd.isna(c) else f"_C{i}" for i,c in enumerate(raw.iloc[h])]
    df=raw.copy(); df.columns=cols
    return df.iloc[h+1:].reset_index(drop=True).dropna(how="all")

def parse_excel(f, sheet=0):
    try:
        f.seek(0)
        return _clean_df(pd.read_excel(f,sheet_name=sheet,header=None,dtype=str))
    except Exception as e:
        return pd.DataFrame()

def _add_month(df, col_map):
    dc=col_map.get("date")
    if not dc:
        for col in df.columns:
            s=df[col].dropna().head(10)
            if sum(1 for v in s if normalize_month(v))>=max(2,len(s)//2): dc=col; break
    if not dc: return None
    df=df.copy(); df["_month"]=df[dc].apply(normalize_month)
    return df.dropna(subset=["_month"])

def extract_books(df):
    if df.empty: return {}
    cm=map_columns(df); dfm=_add_month(df,cm)
    if dfm is None or dfm.empty: return {}
    result={}
    for m,g in dfm.groupby("_month"):
        result[m]={f:g[cm[f]].apply(safe_float).sum() if f in cm else 0.0 for f in VALUE_FIELDS}
    return result

def parse_gstr1_pdf(f):
    totals={f:0.0 for f in VALUE_FIELDS}; month="Unknown"
    try:
        import pdfplumber; f.seek(0)
        with pdfplumber.open(f) as pdf:
            txt=pdf.pages[0].extract_text() or ""
            m=re.search(r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s\-]*(\d{4})",txt.lower())
            if m: month=MONTH_ABBREV.get(m.group(1)[:3],"Unknown")
            for page in pdf.pages:
                for tbl in (page.extract_tables() or []):
                    if len(tbl)<2: continue
                    hdrs=[str(h or "").lower() for h in tbl[0]]
                    def ci(kws): return next((i for i,h in enumerate(hdrs) if any(k in h for k in kws)),None)
                    ii=ci(["igst","integrated"]); ci2=ci(["cgst","central tax"])
                    si=ci(["sgst","state","utgst"]); vi=ci(["taxable","value of supply"])
                    for row in tbl[1:]:
                        def gv(idx): return 0.0 if idx is None or idx>=len(row) or row[idx] is None else safe_float(row[idx])
                        totals["igst"]+=gv(ii); totals["cgst"]+=gv(ci2); totals["sgst"]+=gv(si)
                        totals["sales_value"]+=gv(vi)
    except Exception as e:
        st.warning(f"PDF parse note: {e}")
    return {month:totals}

def parse_credit_ledger(f):
    result={}
    try:
        f.seek(0); df=_clean_df(pd.read_excel(f,header=None,dtype=str))
        cm=map_columns(df); type_col=None
        for col in df.columns:
            cl=col.lower().strip()
            if cl in {"credit","debit","type","cr/dr","transaction type"}: type_col=col; break
        if not type_col and len(df.columns)>5:
            cf=df.columns[5]; uniq=df[cf].dropna().astype(str).str.lower().unique()
            if set(uniq)&{"credit","debit","cr","dr"}: type_col=cf
        if not type_col or "date" not in cm: return {}
        df["_month"]=df[cm["date"]].apply(normalize_month); df=df.dropna(subset=["_month"])
        for m,g in df.groupby("_month"):
            cr=g[g[type_col].astype(str).str.lower().isin({"credit","cr"})]
            db=g[g[type_col].astype(str).str.lower().isin({"debit","dr"})]
            result[m]={}
            for t in TAX_FIELDS:
                col=cm.get(t)
                result[m][f"{t}_credit"]=cr[col].apply(safe_float).sum() if col else 0.0
                result[m][f"{t}_debit"] =db[col].apply(safe_float).sum() if col else 0.0
    except Exception as e:
        st.warning(f"Credit Ledger note: {e}")
    return result

def parse_gstr2b(f):
    result={"b2b":{},"cdnr":{},"impz":{},"_lines":pd.DataFrame()}
    try:
        f.seek(0); xl=pd.ExcelFile(f)
        def find_sheet(targets):
            for s in xl.sheet_names:
                sl=s.lower().replace(" ","").replace("-","")
                for t in targets:
                    if t.lower().replace("-","") in sl: return s
            return None
        b2b_s=find_sheet(["b2b"]); cdnr_s=find_sheet(["b2bcdnr","cdnr"]); impz_s=find_sheet(["impz","impg"])
        def agg(sname):
            if not sname: return {}
            f.seek(0); df=parse_excel(f,sheet=sname)
            if df.empty: return {}
            cm=map_columns(df); dfm=_add_month(df,cm)
            if dfm is None: return {}
            out={}
            for m,g in dfm.groupby("_month"):
                out[m]={t:g[cm[t]].apply(safe_float).sum() if t in cm else 0.0 for t in TAX_FIELDS}
            return out
        result["b2b"]=agg(b2b_s); result["impz"]=agg(impz_s)
        if b2b_s: f.seek(0); result["_lines"]=parse_excel(f,sheet=b2b_s)
        if cdnr_s:
            f.seek(0); df=parse_excel(f,sheet=cdnr_s)
            if not df.empty:
                cm=map_columns(df); dfm=_add_month(df,cm); ntc=cm.get("note_type")
                if not ntc:
                    for col in df.columns:
                        uniq=df[col].dropna().astype(str).str.lower().unique()
                        if len(uniq)<=4 and set(uniq)&{"debit","credit","d","c"}: ntc=col; break
                if dfm is not None:
                    co={}
                    for m,g in dfm.groupby("_month"):
                        row={t:0.0 for t in TAX_FIELDS}
                        for _,ln in g.iterrows():
                            sign=-1 if ntc and "credit" in str(ln.get(ntc,"")).lower() else 1
                            for t in TAX_FIELDS:
                                col=cm.get(t)
                                if col: row[t]+=sign*safe_float(ln.get(col,0))
                        co[m]=row
                    result["cdnr"]=co
    except Exception as e:
        st.warning(f"GSTR-2B note: {e}")
    return result

def extract_invoice_lines(df):
    if df.empty: return pd.DataFrame(columns=["gstin","invoice_no","igst","cgst","sgst"])
    cm=map_columns(df); out=pd.DataFrame()
    out["gstin"]     =df[cm["gstin"]].astype(str).str.strip().str.upper() if "gstin" in cm else ""
    out["invoice_no"]=df[cm["invoice_no"]].astype(str).str.strip().str.upper() if "invoice_no" in cm else ""
    for t in TAX_FIELDS: out[t]=df[cm[t]].apply(safe_float) if t in cm else 0.0
    return out.reset_index(drop=True)

# ════════════════════════════════ UI HELPERS ═════════════════════════════════
def recon_table(books, portal, lb="Books", lp="Portal", show_val=True):
    all_months=set(list(books)+list(portal))
    if not all_months: st.info("No data to display."); return
    vcols=["Sales Value","Export Value","SEZ Value"] if show_val else []
    cols=["Description"]+vcols+["IGST","CGST","SGST","Total Tax"]
    def mk(lbl,d):
        r=[lbl]
        if show_val: r+=[format_inr(d.get(f,0)) for f in ["sales_value","export_value","sez_value"]]
        i,c,s=safe_float(d.get("igst",0)),safe_float(d.get("cgst",0)),safe_float(d.get("sgst",0))
        r+=[format_inr(i),format_inr(c),format_inr(s),format_inr(i+c+s)]; return r
    def mkd(b,p):
        r=["Difference"]
        if show_val: r+=[format_diff(safe_float(b.get(f,0))-safe_float(p.get(f,0))) for f in ["sales_value","export_value","sez_value"]]
        for t in TAX_FIELDS: r.append(format_diff(safe_float(b.get(t,0))-safe_float(p.get(t,0))))
        bt=sum(safe_float(b.get(t,0)) for t in TAX_FIELDS); pt=sum(safe_float(p.get(t,0)) for t in TAX_FIELDS)
        r.append(format_diff(bt-pt)); return r
    for m in sort_months_fiscal(list(all_months)):
        b,p=books.get(m,{}),portal.get(m,{})
        with st.expander(f"Month: {m}", expanded=False):
            st.dataframe(pd.DataFrame([mk(f"Books: {lb}",b),mk(f"Portal: {lp}",p),mkd(b,p)],columns=cols),
                         use_container_width=True,hide_index=True)
    st.divider(); st.subheader("Grand Total")
    sm=sort_months_fiscal(list(all_months))
    tb={f:sum(safe_float(books.get(m,{}).get(f,0)) for m in sm) for f in VALUE_FIELDS}
    tp={f:sum(safe_float(portal.get(m,{}).get(f,0)) for m in sm) for f in VALUE_FIELDS}
    st.dataframe(pd.DataFrame([mk(f"Total {lb}",tb),mk(f"Total {lp}",tp),mkd(tb,tp)],columns=cols),
                 use_container_width=True,hide_index=True)

def metrics(books,portal):
    def tt(d): return sum(sum(safe_float(v.get(t,0)) for t in TAX_FIELDS) for v in d.values())
    bt=tt(books); pt=tt(portal); diff=bt-pt
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Books Total Tax",format_inr(bt)); c2.metric("Portal Total Tax",format_inr(pt))
    c3.metric("Difference",format_inr(abs(diff))); c4.metric("Months Covered",len(set(list(books)+list(portal))))

def export_xls(books,portal,lb,lp,fname):
    rows=[]
    for m in sort_months_fiscal(list(set(list(books)+list(portal)))):
        for lbl,d in [(lb,books.get(m,{})),(lp,portal.get(m,{}))]:
            rows.append({"Month":m,"Source":lbl,**{t.upper():safe_float(d.get(t,0)) for t in TAX_FIELDS},
                         "Total":sum(safe_float(d.get(t,0)) for t in TAX_FIELDS)})
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as w: pd.DataFrame(rows).to_excel(w,index=False,sheet_name="Recon")
    buf.seek(0)
    st.download_button("Download Excel",data=buf,file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ════════════════════════════════ PAGE CONFIG ════════════════════════════════
st.set_page_config(page_title="GST Reconciliation Tool",page_icon="📊",layout="wide")
st.markdown("""
<style>
.hdr{background:linear-gradient(135deg,#1a3c6b,#2563a8);padding:1.2rem 2rem;border-radius:10px;color:white;margin-bottom:1.2rem}
</style>""",unsafe_allow_html=True)
st.markdown('<div class="hdr"><h1>📊 GST Reconciliation Tool</h1><p style="margin:0;opacity:.85">Automated reconciliation — Books vs GSTR-1 / GSTR-2B / Credit Ledger</p></div>',unsafe_allow_html=True)

with st.sidebar:
    st.title("Navigation")
    page=st.radio("Select Module",["🏠 Home","📤 Module 1","📥 Module 2","🔄 Module 3","🔍 Module 4"])
    st.divider()
    st.caption("Formats: .xlsx .xls .pdf"); st.caption("FY: April to March")

# ════════════════════════════════ MODULE 1 ═══════════════════════════════════
def module1():
    st.header("📤 Module 1 — Outward Supplies")
    st.markdown("Books (Sales minus Credit Notes) vs **GSTR-1**")
    st.divider()
    cb,cp=st.columns(2,gap="large")
    with cb:
        st.subheader("Books Data")
        sf=st.file_uploader("Sales Register (Excel/PDF)",type=["xlsx","xls","pdf"],key="m1s")
        cnf=st.file_uploader("Credit Note Register (optional)",type=["xlsx","xls","pdf"],key="m1cn")
    with cp:
        st.subheader("Portal Data")
        gf=st.file_uploader("GSTR-1 PDF",type=["pdf"],key="m1g")
    st.divider()
    if st.button("Run Module 1",type="primary",key="m1run"):
        if not sf: st.error("Upload the Sales Register."); return
        with st.spinner("Parsing Sales Register..."):
            books=parse_gstr1_pdf(sf) if sf.name.lower().endswith(".pdf") else extract_books(parse_excel(sf))
        cn={}
        if cnf:
            with st.spinner("Parsing Credit Notes..."):
                cn=parse_gstr1_pdf(cnf) if cnf.name.lower().endswith(".pdf") else extract_books(parse_excel(cnf))
        net=subtract_monthwise(books,cn)
        portal={}
        if gf:
            with st.spinner("Parsing GSTR-1 PDF..."):
                portal=parse_gstr1_pdf(gf)
        st.session_state.update({"m1_net":net,"m1_portal":portal})
        st.success("Done!"); _s1(net,portal)
    elif "m1_net" in st.session_state: _s1(st.session_state["m1_net"],st.session_state["m1_portal"])

def _s1(net,portal):
    if not net and not portal: st.info("No data."); return
    metrics(net,portal); st.divider()
    recon_table(net,portal,"Books (Net of CN)","GSTR-1",show_val=True)
    st.divider(); export_xls(net,portal,"Books","GSTR-1","module1_outward.xlsx")

# ════════════════════════════════ MODULE 2 ═══════════════════════════════════
def module2():
    st.header("📥 Module 2 — ITC Availment")
    st.markdown("Books ITC (minus Debit Notes) vs **Electronic Credit Ledger**")
    st.info("Column F: Credit = ITC Availed | Debit = ITC Utilized")
    st.divider()
    cb,cp=st.columns(2,gap="large")
    with cb:
        st.subheader("Books Data")
        pf=st.file_uploader("Purchase Register",type=["xlsx","xls"],key="m2p")
        jf=st.file_uploader("Journal Register", type=["xlsx","xls"],key="m2j")
        df2=st.file_uploader("Debit Note Register (optional)",type=["xlsx","xls"],key="m2dn")
    with cp:
        st.subheader("Portal Data")
        lf=st.file_uploader("Electronic Credit Ledger",type=["xlsx","xls","pdf"],key="m2l")
    st.divider()
    if st.button("Run Module 2",type="primary",key="m2run"):
        if not pf and not jf: st.error("Upload Purchase or Journal Register."); return
        raws=[extract_books(parse_excel(fi)) for fi in [pf,jf] if fi]
        gross=merge_monthwise(*raws) if raws else {}
        dn=extract_books(parse_excel(df2)) if df2 else {}
        net=subtract_monthwise(gross,dn)
        ledger=parse_credit_ledger(lf) if lf else {}
        av={m:{t:safe_float(d.get(f"{t}_credit",0)) for t in TAX_FIELDS} for m,d in ledger.items()}
        ut={m:{t:safe_float(d.get(f"{t}_debit",0))  for t in TAX_FIELDS} for m,d in ledger.items()}
        st.session_state.update({"m2_nb":net,"m2_av":av,"m2_ut":ut,"books_itc_net":net})
        st.success("Done!"); _s2(net,av,ut)
    elif "m2_nb" in st.session_state: _s2(st.session_state["m2_nb"],st.session_state["m2_av"],st.session_state["m2_ut"])

def _s2(nb,av,ut):
    if not nb and not av: st.info("No data."); return
    metrics(nb,av); st.divider()
    cols=["Description","IGST","CGST","SGST","Total Tax"]
    def r(lbl,d):
        i,c,s=safe_float(d.get("igst",0)),safe_float(d.get("cgst",0)),safe_float(d.get("sgst",0))
        return [lbl,format_inr(i),format_inr(c),format_inr(s),format_inr(i+c+s)]
    def dr(b,p):
        diffs=[safe_float(b.get(t,0))-safe_float(p.get(t,0)) for t in TAX_FIELDS]
        return ["Difference"]+[format_diff(d) for d in diffs]+[format_diff(sum(diffs))]
    all_months=sort_months_fiscal(list(set(list(nb)+list(av))))
    for m in all_months:
        b,a,u=nb.get(m,{}),av.get(m,{}),ut.get(m,{})
        with st.expander(f"Month: {m}",expanded=False):
            rows=[r("Books ITC (Net of DN)",b),r("ITC Availed (Portal)",a),dr(b,a)]
            if any(safe_float(u.get(t,0))>0 for t in TAX_FIELDS): rows.append(r("ITC Utilized (info)",u))
            st.dataframe(pd.DataFrame(rows,columns=cols),use_container_width=True,hide_index=True)
    st.divider(); st.subheader("Grand Total")
    tb={t:sum(safe_float(nb.get(m,{}).get(t,0)) for m in all_months) for t in TAX_FIELDS}
    ta={t:sum(safe_float(av.get(m,{}).get(t,0)) for m in all_months) for t in TAX_FIELDS}
    tu={t:sum(safe_float(ut.get(m,{}).get(t,0)) for m in all_months) for t in TAX_FIELDS}
    st.dataframe(pd.DataFrame([r("Total Books ITC",tb),r("Total ITC Availed",ta),
        ["Net Diff"]+[format_diff(safe_float(tb[t])-safe_float(ta[t])) for t in TAX_FIELDS]
        +[format_diff(sum(safe_float(tb[t])-safe_float(ta[t]) for t in TAX_FIELDS))],
        r("Total ITC Utilized",tu)],columns=cols),use_container_width=True,hide_index=True)
    st.divider(); export_xls(nb,av,"Books ITC","ITC Availed","module2_itc.xlsx")

# ════════════════════════════════ MODULE 3 ═══════════════════════════════════
def module3():
    st.header("🔄 Module 3 — ITC vs GSTR-2B")
    st.markdown("Books ITC vs **GSTR-2B** (B2B + CDNR + IMPZ)")
    st.info("B2B: Add | B2B-CDNR Debit: Add / Credit: Subtract | IMPZ: Add")
    st.divider()
    cb,cp=st.columns(2,gap="large")
    with cb:
        st.subheader("Books Data")
        pf=st.file_uploader("Purchase Register",type=["xlsx","xls"],key="m3p")
        jf=st.file_uploader("Journal Register", type=["xlsx","xls"],key="m3j")
        df3=st.file_uploader("Debit Note Register (optional)",type=["xlsx","xls"],key="m3dn")
        reuse=st.checkbox("Reuse Books data from Module 2",value=True,key="m3r")
    with cp:
        st.subheader("Portal Data")
        gf=st.file_uploader("GSTR-2B Excel",type=["xlsx","xls"],key="m3g")
    st.divider()
    if st.button("Run Module 3",type="primary",key="m3run"):
        if reuse and "books_itc_net" in st.session_state:
            net=st.session_state["books_itc_net"]; st.info("Reusing Module 2 Books data.")
        else:
            if not pf and not jf: st.error("Upload registers or run Module 2 first."); return
            raws=[extract_books(parse_excel(fi)) for fi in [pf,jf] if fi]
            gross=merge_monthwise(*raws) if raws else {}
            dn=extract_books(parse_excel(df3)) if df3 else {}
            net=subtract_monthwise(gross,dn); st.session_state["books_itc_net"]=net
        if not gf: st.error("Upload the GSTR-2B file."); return
        with st.spinner("Parsing GSTR-2B..."):
            raw2b=parse_gstr2b(gf)
        all_months=set()
        for k in ("b2b","cdnr","impz"): all_months.update(raw2b.get(k,{}).keys())
        portal={m:{t:sum(safe_float(raw2b.get(k,{}).get(m,{}).get(t,0)) for k in ("b2b","cdnr","impz"))
                   for t in TAX_FIELDS} for m in all_months}
        st.session_state.update({"m3_nb":net,"m3_portal":portal,"gstr2b_lines":raw2b.get("_lines",pd.DataFrame())})
        st.success("Done!"); _s3(net,portal)
    elif "m3_nb" in st.session_state: _s3(st.session_state["m3_nb"],st.session_state["m3_portal"])

def _s3(nb,portal):
    if not nb and not portal: st.info("No data."); return
    metrics(nb,portal); st.divider()
    recon_table(nb,portal,"Books ITC (Net of DN)","GSTR-2B",show_val=False)
    st.divider(); export_xls(nb,portal,"Books ITC","GSTR-2B","module3_gstr2b.xlsx")

# ════════════════════════════════ MODULE 4 ═══════════════════════════════════
def module4():
    st.header("🔍 Module 4 — Invoice-Level Reconciliation")
    st.markdown("Line-by-line match on **GSTIN + Invoice Number**")
    st.divider()
    cb,cp=st.columns(2,gap="large")
    with cb:
        st.subheader("Books Data")
        pf=st.file_uploader("Purchase Register",type=["xlsx","xls"],key="m4p")
        jf=st.file_uploader("Journal Register", type=["xlsx","xls"],key="m4j")
    with cp:
        st.subheader("GSTR-2B")
        gf=st.file_uploader("GSTR-2B Excel",type=["xlsx","xls"],key="m4g")
        reuse=st.checkbox("Reuse GSTR-2B from Module 3",value=True,key="m4r")
    st.divider()
    if st.button("Run Module 4",type="primary",key="m4run"):
        dfs=[extract_invoice_lines(parse_excel(fi)) for fi in [pf,jf] if fi]
        if not dfs: st.error("Upload at least one register."); return
        books_lines=pd.concat(dfs,ignore_index=True)
        if reuse and "gstr2b_lines" in st.session_state:
            g2b_lines=st.session_state["gstr2b_lines"]; st.info("Reusing GSTR-2B from Module 3.")
        elif gf:
            with st.spinner("Parsing GSTR-2B..."):
                raw=parse_gstr2b(gf); g2b_lines=raw.get("_lines",pd.DataFrame())
        else:
            st.error("Upload GSTR-2B or run Module 3 first."); return
        g2b_inv=extract_invoice_lines(g2b_lines)
        with st.spinner("Matching invoices..."):
            buckets=_match(books_lines,g2b_inv)
        st.session_state["m4_buckets"]=buckets; st.success("Done!"); _s4(buckets)
    elif "m4_buckets" in st.session_state: _s4(st.session_state["m4_buckets"])

def _match(books,g2b):
    TOL=1.0
    ecols=["GSTIN","Invoice No","Books IGST","Books CGST","Books SGST","GSTR-2B IGST","GSTR-2B CGST","GSTR-2B SGST","IGST Diff","CGST Diff","SGST Diff"]
    empty=pd.DataFrame(columns=ecols)
    if books.empty and g2b.empty: return {k:empty.copy() for k in ("matched","not_in_books","not_in_gstr2b","amount_mismatch")}
    books=books.copy(); g2b=g2b.copy()
    books["_key"]=books["gstin"].str.upper().str.strip()+"||"+books["invoice_no"].str.upper().str.strip()
    g2b["_key"]  =g2b["gstin"].str.upper().str.strip()  +"||"+g2b["invoice_no"].str.upper().str.strip()
    bi=books.set_index("_key"); gi=g2b.set_index("_key")
    bk,gk=set(bi.index),set(gi.index)
    mr,mis=[],[]
    for key in bk&gk:
        b=bi.loc[key]; g=gi.loc[key]
        if isinstance(b,pd.DataFrame): b=b.iloc[0]
        if isinstance(g,pd.DataFrame): g=g.iloc[0]
        diffs={t:safe_float(b.get(t,0))-safe_float(g.get(t,0)) for t in TAX_FIELDS}
        p=key.split("||")
        row={"GSTIN":p[0],"Invoice No":p[1] if len(p)>1 else "",
             "Books IGST":safe_float(b.get("igst",0)),"Books CGST":safe_float(b.get("cgst",0)),"Books SGST":safe_float(b.get("sgst",0)),
             "GSTR-2B IGST":safe_float(g.get("igst",0)),"GSTR-2B CGST":safe_float(g.get("cgst",0)),"GSTR-2B SGST":safe_float(g.get("sgst",0)),
             "IGST Diff":diffs["igst"],"CGST Diff":diffs["cgst"],"SGST Diff":diffs["sgst"]}
        (mr if all(abs(diffs[t])<=TOL for t in TAX_FIELDS) else mis).append(row)
    def side(keys,idx,ib):
        rows=[]
        for key in keys:
            r=idx.loc[key]
            if isinstance(r,pd.DataFrame): r=r.iloc[0]
            p=key.split("||")
            rows.append({"GSTIN":p[0],"Invoice No":p[1] if len(p)>1 else "",
                "Books IGST":safe_float(r.get("igst",0)) if ib else 0,"Books CGST":safe_float(r.get("cgst",0)) if ib else 0,
                "Books SGST":safe_float(r.get("sgst",0)) if ib else 0,
                "GSTR-2B IGST":safe_float(r.get("igst",0)) if not ib else 0,
                "GSTR-2B CGST":safe_float(r.get("cgst",0)) if not ib else 0,
                "GSTR-2B SGST":safe_float(r.get("sgst",0)) if not ib else 0,
                "IGST Diff":0,"CGST Diff":0,"SGST Diff":0})
        return rows
    def to_df(rows): return pd.DataFrame(rows) if rows else empty.copy()
    return {"matched":to_df(mr),"not_in_books":to_df(side(gk-bk,gi,False)),
            "not_in_gstr2b":to_df(side(bk-gk,bi,True)),"amount_mismatch":to_df(mis)}

def _s4(buckets):
    ma=buckets.get("matched",pd.DataFrame()); nb=buckets.get("not_in_books",pd.DataFrame())
    ng=buckets.get("not_in_gstr2b",pd.DataFrame()); mm=buckets.get("amount_mismatch",pd.DataFrame())
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Matched",len(ma)); c2.metric("Not in Books",len(nb))
    c3.metric("Not in GSTR-2B",len(ng)); c4.metric("Amount Mismatch",len(mm))
    st.divider()
    t1,t2,t3,t4=st.tabs(["Matched","Not in Books","Not in GSTR-2B","Amount Mismatch"])
    with t1: st.subheader(f"Matched ({len(ma)})"); st.dataframe(ma,use_container_width=True,hide_index=True) if not ma.empty else st.info("None.")
    with t2: st.subheader(f"Not in Books ({len(nb)})"); st.caption("In GSTR-2B but missing from registers."); st.dataframe(nb,use_container_width=True,hide_index=True) if not nb.empty else st.info("None.")
    with t3: st.subheader(f"Not in GSTR-2B ({len(ng)})"); st.caption("In Books but supplier hasn't uploaded. ITC risk."); st.dataframe(ng,use_container_width=True,hide_index=True) if not ng.empty else st.info("None.")
    with t4: st.subheader(f"Amount Mismatch ({len(mm)})"); st.caption("GSTIN+Invoice match but amounts differ."); st.dataframe(mm,use_container_width=True,hide_index=True) if not mm.empty else st.info("None.")
    st.divider()
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as w:
        for key,sname in [("matched","Matched"),("not_in_books","Not_In_Books"),("not_in_gstr2b","Not_In_GSTR2B"),("amount_mismatch","Mismatch")]:
            df=buckets.get(key,pd.DataFrame())
            (df if not df.empty else pd.DataFrame(columns=["No records"])).to_excel(w,index=False,sheet_name=sname)
    buf.seek(0)
    st.download_button("Download All Buckets Excel",data=buf,file_name="module4_invoice_match.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ════════════════════════════════ HOME ═══════════════════════════════════════
def home():
    st.title("Welcome to GST Reconciliation Tool")
    c1,c2=st.columns(2)
    with c1:
        st.markdown("""
**Module 1 — Outward Supplies**
Sales Register minus Credit Notes vs GSTR-1. Month-wise IGST/CGST/SGST + Export/SEZ.

---
**Module 3 — ITC vs GSTR-2B**
Books ITC vs supplier-uploaded GSTR-2B. Handles B2B, B2B-CDNR, IMPZ sheets.
        """)
    with c2:
        st.markdown("""
**Module 2 — ITC Availment**
Purchase + Journal Register vs Electronic Credit Ledger (Column F logic).

---
**Module 4 — Invoice-Level Match**
Line-by-line on GSTIN + Invoice No.
4 buckets: Matched / Not in Books / Not in GSTR-2B / Amount Mismatch
        """)
    st.divider()
    st.markdown("""
| Column Keyword | Auto-Detected From |
|---|---|
| Sales Value | Sale, Job work |
| IGST | IGST, Integrated Tax |
| CGST | CGST, Central Tax |
| SGST | SGST, State Tax |
| Export | Export |
| SEZ | SEZ |
    """)

# ════════════════════════════════ ROUTER ═════════════════════════════════════
if   page=="🏠 Home":       home()
elif page=="📤 Module 1":   module1()
elif page=="📥 Module 2":   module2()
elif page=="🔄 Module 3":   module3()
elif page=="🔍 Module 4":   module4()
