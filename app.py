from __future__ import annotations
import io, re, json
from pathlib import Path
from typing import Dict, List, Tuple
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Company X | Slow-stock review", page_icon="📦", layout="wide")

# Responsive CSS: Streamlit is the application shell; HTML/CSS/JS are used for the mobile-friendly UI polish.
st.markdown("""
<style>
:root { --ink:#152d38; --teal:#0d5964; --mint:#eaf5f3; --line:#dce5e8; }
.block-container {padding-top:1.4rem; max-width:1440px;}
[data-testid="stMetric"] {background:var(--mint); border:1px solid var(--line); padding:14px 16px; border-radius:12px;}
section[data-testid="stSidebar"] {background:#f5f8f9;}
.stButton>button, .stDownloadButton>button {border-radius:9px; font-weight:600;}
[data-testid="stDataFrame"] {border:1px solid var(--line); border-radius:10px;}
.mobile-call {border:1px solid var(--line);border-radius:12px;padding:12px;margin:8px 0;background:#fff;}
.small-muted {color:#61737b;font-size:.88rem;}
@media(max-width:700px){.block-container{padding:1rem .7rem}.mobile-call{font-size:1rem}}
</style>
""", unsafe_allow_html=True)

ALIASES = {
    "sku": ["sku", "item sku", "item code", "product code", "stock code"],
    "name": ["name", "item name", "sku name", "product name", "description", "item description"],
    "category": ["category", "item category", "product category"],
    "pack": ["master pack quantity", "master pack", "pack quantity", "units per case", "case pack", "pack size"],
    "unit_cost": ["unit cost", "cost per unit", "unit price", "cost"],
    "alternate": ["alternate item", "alternate sku name", "alternate product", "replacement item"],
    "alternate_flag": ["alternate flag", "alternate status", "switch flag"],
    "status": ["status", "active status", "item status"],
    "stock_sku": ["sku", "item sku", "item code", "product code", "stock code"],
    "b0": ["0-30", "0–30", "0-30 days", "0–30 days", "0 to 30", "0_30"],
    "b1": ["31-60", "31–60", "31-60 days", "31–60 days", "31 to 60", "31_60"],
    "b2": ["61-90", "61–90", "61-90 days", "61–90 days", "61 to 90", "61_90"],
    "b3": ["91-180", "91–180", "91-180 days", "91–180 days", "91 to 180", "91_180"],
    "b4": ["180+", "180 +", "180+ days", "180 plus", "over 180", "180_plus"],
    "month": ["month", "sales month", "invoice month", "date"],
    "customer": ["customer", "customer name", "account name", "client name"],
    "industry": ["industry", "customer industry", "sector"],
    "quantity": ["quantity", "qty", "units sold", "sales quantity"],
    "value": ["value", "sales value", "revenue", "line value", "net sales"],
    "customer_id": ["customer id", "customer code", "account id", "customer number"],
    "area": ["area", "region", "location", "territory"],
}

def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()

def clean_cols(df):
    df = df.copy()
    df.columns = [norm(c) for c in df.columns]
    return df

def find_col(df, key, required=True):
    cols = list(df.columns)
    for alias in ALIASES[key]:
        a = norm(alias)
        if a in cols: return a
    # permissive normalized substring match
    for alias in ALIASES[key]:
        a = norm(alias)
        for c in cols:
            if a and (a in c or c in a): return c
    if required: raise ValueError(f"Could not identify the '{key}' column. Available columns: {', '.join(cols)}")
    return None

def standardize(df, mapping):
    d = clean_cols(df)
    out = pd.DataFrame(index=d.index)
    for target, key in mapping.items():
        c = find_col(d, key, required=target in mapping.get("required", [])) if False else find_col(d, key, required=False)
        if c is not None: out[target] = d[c]
    return out

def read_file(upload):
    name = upload.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(upload)
    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(upload)
    raise ValueError(f"Unsupported file: {upload.name}. Upload .xlsx, .xls, or .csv.")

def load_inputs(files):
    tables = {}
    for f in files:
        base = norm(Path(f.name).stem).replace(" ", "_")
        raw = read_file(f)
        # Workbook with named sheets
        if f.name.lower().endswith((".xlsx", ".xls")):
            try:
                f.seek(0)
                xls = pd.ExcelFile(f)
                if len(xls.sheet_names) > 1:
                    for sh in xls.sheet_names:
                        key = norm(sh).replace(" ", "_")
                        tables[key] = clean_cols(pd.read_excel(xls, sheet_name=sh))
                    continue
            except Exception:
                pass
        tables[base] = clean_cols(raw)
    aliases = {
        "item_master": ["item_master", "item master", "itemmaster", "items"],
        "stock_aging": ["stock_aging", "stock aging", "stockaging", "aging", "stock"],
        "sales_lines": ["sales_lines", "sales lines", "saleslines", "sales"],
        "customers": ["customers", "customer"]
    }
    result = {}
    for wanted, names in aliases.items():
        found = next((v for k,v in tables.items() if any(norm(n).replace(" ", "_") == k or norm(n) == k.replace("_", " ") for n in names)), None)
        if found is None:
            # Try matching workbook sheet/filename by tokens
            found = next((v for k,v in tables.items() if wanted.split("_")[0] in k), None)
        if found is not None: result[wanted] = found
    missing = [k for k in aliases if k not in result]
    if missing: raise ValueError("Missing required table(s): " + ", ".join(missing) + ". Upload the workbook with all four sheets, or four files named Item_Master, Stock_Aging, Sales_Lines, and Customers.")
    return result

def standardize_tables(tables):
    issues=[]
    # Master
    raw=tables["item_master"]
    def get(d,key,req=True):
        c=find_col(d,key,required=req)
        return d[c] if c else pd.Series([None]*len(d),index=d.index)
    m=pd.DataFrame({
        "sku":get(raw,"sku").astype(str).str.strip(), "name":get(raw,"name").astype(str).str.strip(),
        "category":get(raw,"category",False).fillna("Uncategorized").astype(str).str.strip(),
        "pack":pd.to_numeric(get(raw,"pack",False),errors="coerce"), "unit_cost":pd.to_numeric(get(raw,"unit_cost",False),errors="coerce"),
        "alternate":get(raw,"alternate",False).fillna("").astype(str).str.strip(),
        "alternate_flag":get(raw,"alternate_flag",False).fillna("").astype(str).str.upper().str.strip(),
        "status":get(raw,"status",False).fillna("ACTIVE").astype(str).str.upper().str.strip()})
    if m.sku.duplicated().any():
        dup=m.loc[m.sku.duplicated(False),"sku"].unique().tolist(); issues.append(f"Duplicate SKU(s) in Item_Master: {', '.join(dup)}. First record retained."); m=m.drop_duplicates("sku",keep="first")
    for c in ["pack","unit_cost"]:
        if m[c].isna().any(): issues.append(f"Missing or non-numeric {c} for {int(m[c].isna().sum())} item(s); treated as 0 for calculations and flagged for source correction."); m[c]=m[c].fillna(0)
    m["pack"]=m.pack.where(m.pack>0,0)
    # stock
    raw=tables["stock_aging"]
    stock=pd.DataFrame({"sku":get(raw,"stock_sku").astype(str).str.strip()})
    bucket_keys=["b0","b1","b2","b3","b4"]
    for k in bucket_keys:
        c=find_col(raw,k,required=False)
        if c is None:
            stock[k]=0
            issues.append(f"Stock_Aging is missing the {k} aging bucket; treated as 0. Confirm source export format.")
        else:
            stock[k]=pd.to_numeric(raw[c],errors="coerce")
            if stock[k].isna().any(): issues.append(f"Non-numeric/missing stock values in {k}: {int(stock[k].isna().sum())} row(s); treated as 0.")
            stock[k]=stock[k].fillna(0)
    if stock.sku.duplicated().any():
        issues.append("Duplicate SKU rows in Stock_Aging were summed by SKU.")
        stock=stock.groupby("sku",as_index=False)[bucket_keys].sum()
    # sales
    raw=tables["sales_lines"]
    sales=pd.DataFrame({"month":get(raw,"month",False),"customer":get(raw,"customer").astype(str).str.strip(),"industry":get(raw,"industry",False).fillna("").astype(str).str.strip(),"sku":get(raw,"sku").astype(str).str.strip(),"quantity":pd.to_numeric(get(raw,"quantity"),errors="coerce"),"value":pd.to_numeric(get(raw,"value"),errors="coerce")})
    for c in ["quantity","value"]:
        if sales[c].isna().any(): issues.append(f"Sales_Lines has {int(sales[c].isna().sum())} missing/non-numeric {c} value(s); treated as 0."); sales[c]=sales[c].fillna(0)
    if sales.sku.isna().any(): issues.append("Sales rows with missing SKU exist; investigate source export.")
    # customers
    raw=tables["customers"]
    customers=pd.DataFrame({"customer":get(raw,"customer").astype(str).str.strip(),"customer_id":get(raw,"customer_id",False).fillna("").astype(str).str.strip(),"industry":get(raw,"industry",False).fillna("").astype(str).str.strip(),"area":get(raw,"area",False).fillna("").astype(str).str.strip()})
    if customers.customer.duplicated().any(): issues.append("Duplicate customer names in Customers; first record retained for lookup."); customers=customers.drop_duplicates("customer",keep="first")
    # normalize status
    m["active"] = ~m.status.str.contains("inactive|in active|discontinued",case=False,regex=True) & m.status.ne("INACTIVE")
    m.loc[m.status.isin(["N","NO","0","FALSE"]),"active"]=False
    return m,stock,sales,customers,issues

def analyze(tables):
    m,stock,sales,customers,issues=standardize_tables(tables)
    # name-based SWITCH resolution to new/alternate SKU
    name_to_sku={norm(n):s for s,n in zip(m.sku,m.name) if n and n.lower()!="nan"}
    switch_map={}
    for _,r in m.iterrows():
        if r.alternate_flag=="SWITCH" and r.alternate:
            dest=name_to_sku.get(norm(r.alternate))
            if dest:
                switch_map[r.sku]=dest
            else: issues.append(f"SWITCH item {r.sku} references alternate name '{r.alternate}' that did not match an Item_Master name; kept as its own SKU. Fix alternate naming at source.")
    def root(s):
        seen=set(); cur=str(s)
        while cur in switch_map and cur not in seen:
            seen.add(cur); cur=switch_map[cur]
        return cur
    # chain root maps old SKU to current SKU; apply to stock and sales
    stock["root_sku"]=stock.sku.map(root)
    sales["root_sku"]=sales.sku.map(root)
    if not stock.sku.isin(m.sku).all():
        missing=sorted(set(stock.loc[~stock.sku.isin(m.sku),"sku"]))
        issues.append(f"Stock rows reference SKU(s) absent from Item_Master: {', '.join(missing)}. They are excluded from active item metrics but shown in data issues.")
    if not sales.sku.isin(m.sku).all():
        missing=sorted(set(sales.loc[~sales.sku.isin(m.sku),"sku"]))
        issues.append(f"Sales rows reference SKU(s) absent from Item_Master: {', '.join(missing)}. They remain in customer spend totals but cannot be classified as items.")
    buckets=["b0","b1","b2","b3","b4"]
    st_agg=stock.groupby("root_sku",as_index=False)[buckets].sum()
    sales_agg=sales.groupby("root_sku",as_index=False).agg(total_sold=("quantity","sum"),item_sales_value=("value","sum"))
    # active item rows only, but switch children should be folded into destination; omit switched-away item row
    m["root_sku"]=m.sku.map(root)
    active=m[m.active].copy()
    # root row determines final name/category/pack/cost. Exclude switch source if its root is different.
    active=active[active.sku.eq(active.root_sku)]
    items=active.merge(st_agg,left_on="sku",right_on="root_sku",how="left").merge(sales_agg,left_on="sku",right_on="root_sku",how="left",suffixes=("","_sales"))
    for c in buckets+["total_sold","item_sales_value"]: items[c]=items[c].fillna(0)
    items["total_stock"]=items[buckets].sum(axis=1)
    items["aged_qty"]=items["b3"]+items["b4"]
    items["aged_value"]=items.aged_qty*items.unit_cost
    items["avg_monthly_sales"]=items.total_sold/6
    items["months_to_clear"]=items.apply(lambda r: (r.total_stock/r.avg_monthly_sales) if r.avg_monthly_sales>0 else (float("inf") if r.total_stock>0 else 0),axis=1)
    def classify(r):
        if r.total_sold < r.pack: return "Red List"
        if r.aged_qty<=0 or r.months_to_clear<3: return "No Action"
        return "Push"
    items["class"]=items.apply(classify,axis=1)
    # Inactive items with stock warnings; include even if switch alias
    inactive=m[~m.active].merge(stock.groupby("sku",as_index=False)[buckets].sum(),on="sku",how="left")
    for c in buckets: inactive[c]=inactive[c].fillna(0)
    inactive["stock_qty"]=inactive[buckets].sum(axis=1)
    inactive_stock=inactive[inactive.stock_qty>0][["sku","name","stock_qty"]]
    if len(inactive_stock): issues.append(f"{len(inactive_stock)} inactive item(s) still hold stock; excluded from active analysis and listed separately.")
    # customer spend from all sales lines, using customer name as join key
    customer_spend=sales.groupby("customer",as_index=False).agg(customer_spend=("value","sum"))
    cust=customers.merge(customer_spend,on="customer",how="outer")
    cust["customer_spend"]=cust.customer_spend.fillna(0)
    cust["industry"]=cust.industry.fillna("")
    # sales customer industry may be populated where customer master is blank
    sales_ind=sales.groupby("customer",as_index=False).agg(sales_industry=("industry",lambda x: next((v for v in x if str(v).strip()),"")))
    cust=cust.merge(sales_ind,on="customer",how="left")
    cust["industry_final"]=cust.industry.where(cust.industry.str.strip().ne(""),cust.sales_industry.fillna(""))
    # item category by root; set of original and switch partner SKUs
    root_to_members={}
    for s in m.sku: root_to_members.setdefault(root(s),set()).add(s)
    target_rows=[]
    sales["root_sku"]=sales.sku.map(root)
    # customer spend eligibility strictly greater than 100k
    eligible=set(cust.loc[cust.customer_spend>100000,"customer"])
    customer_industry=dict(zip(cust.customer,cust.industry_final))
    for _,it in items[items["class"].eq("Push")].iterrows():
        sku=it.sku; members=root_to_members.get(sku,{sku})
        item_sales=sales[sales.root_sku.eq(sku)]
        bought=set(item_sales.loc[item_sales.quantity.ne(0),"customer"])
        bought_industries={customer_industry.get(c,"").strip().lower() for c in bought if customer_industry.get(c,"").strip()}
        # Bought other item in same category, excluding this item and its switch partner
        category_skus=set(m.loc[m.category.eq(it.category),"sku"].map(root))
        category_skus.discard(sku)
        cat_buyers=set(sales.loc[sales.root_sku.isin(category_skus),"customer"])
        for c in eligible:
            if c in bought: bucket=1; label="Bought before"
            elif customer_industry.get(c,"").strip() and customer_industry[c].strip().lower() in bought_industries: bucket=2; label="Similar business"
            elif c in cat_buyers: bucket=3; label="Same category"
            else: continue
            target_rows.append({"sku":sku,"item_name":it["name"],"customer":c,"bucket":bucket,"bucket_label":label,"customer_purchase_value":float(cust.loc[cust.customer.eq(c),"customer_spend"].iloc[0]),"industry":customer_industry.get(c,""),"area":(cust.loc[cust.customer.eq(c),"area"].iloc[0] if (cust.customer.eq(c)).any() and "area" in cust else "")})
    targets=pd.DataFrame(target_rows,columns=["sku","item_name","customer","bucket","bucket_label","customer_purchase_value","industry","area"])
    if not targets.empty: targets=targets.sort_values(["sku","bucket","customer_purchase_value"],ascending=[True,True,False])
    # duplicate SKU rows can create merged suffix columns; drop helper root columns
    keep=["sku","name","category","pack","unit_cost"]+buckets+["total_stock","aged_qty","aged_value","total_sold","avg_monthly_sales","months_to_clear","class","item_sales_value"]
    items=items[[c for c in keep if c in items.columns]].copy()
    return {"items":items,"targets":targets,"issues":issues,"inactive_stock":inactive_stock,"customers":cust,"switch_map":switch_map,"master":m,"stock":stock,"sales":sales}

def excel_bytes(result):
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as w:
        result["items"].to_excel(w,index=False,sheet_name="Item_Review")
        result["targets"].to_excel(w,index=False,sheet_name="Customer_Targets")
        pd.DataFrame({"Data_Quality_Issue":result["issues"]}).to_excel(w,index=False,sheet_name="Data_Quality")
        result["inactive_stock"].to_excel(w,index=False,sheet_name="Inactive_Stock")
    out.seek(0); return out.getvalue()

st.title("📦 Company X — Slow-stock review")
st.caption("Upload the monthly system exports, run the rules, and export an actionable call list. Review data-quality warnings before acting.")
with st.sidebar:
    st.header("1. Upload monthly data")
    uploads=st.file_uploader("Workbook (.xlsx) or four separate files (.xlsx/.csv)",type=["xlsx","xls","csv"],accept_multiple_files=True)
    st.markdown("**Required tables**")
    st.markdown("- Item_Master\n- Stock_Aging\n- Sales_Lines\n- Customers")
    st.markdown("**Business rules**")
    st.caption("Uses the case rules: switch consolidation, inactive-item exclusion, aged-stock valuation, Red List / No Action / Push classification, and target buckets.")

if not uploads:
    st.info("Upload the Company X workbook (all four sheets) or the four source files to calculate the real results. No sample numbers are fabricated.")
    st.markdown("### What you will get")
    c1,c2,c3=st.columns(3)
    c1.markdown("**Review summary**\n\nClass counts, Push aged value, and top 10 Push items.")
    c2.markdown("**Sales call list**\n\nCustomer targets ranked by bucket and six-month spend, designed to work on mobile.")
    c3.markdown("**Excel export**\n\nItem review, customer targets, data-quality log, and inactive stock.")
else:
    try:
        tables=load_inputs(uploads)
        result=analyze(tables)
        items=result["items"]
        st.success("Review completed. Check the data-quality notes before using the results.")
        st.subheader("Review summary")
        counts=items["class"].value_counts()
        red=int(counts.get("Red List",0)); no=int(counts.get("No Action",0)); push=int(counts.get("Push",0))
        aged=float(items.loc[items["class"].eq("Push"),"aged_value"].sum())
        a,b,c,d=st.columns(4)
        a.metric("Red List items",red); b.metric("No Action items",no); c.metric("Push items",push); d.metric("Push aged value",f"₹{aged:,.2f}")
        st.subheader("Top 10 Push items by aged value")
        top=items[items["class"].eq("Push")].sort_values("aged_value",ascending=False).head(10)
        st.dataframe(top[["sku","name","category","total_stock","aged_qty","unit_cost","aged_value","total_sold","months_to_clear"]],use_container_width=True,hide_index=True)
        st.subheader("Item review")
        fc,fg=st.columns(2)
        selected_class=fc.selectbox("Filter by class",["All","Red List","No Action","Push"])
        cats=["All"]+sorted(items.category.dropna().astype(str).unique().tolist())
        selected_cat=fg.selectbox("Filter by category",cats)
        view=items.copy()
        if selected_class!="All": view=view[view["class"]==selected_class]
        if selected_cat!="All": view=view[view.category.astype(str)==selected_cat]
        st.dataframe(view,use_container_width=True,hide_index=True)
        st.subheader("Customer call list")
        pushes=items.loc[items["class"].eq("Push"),["sku","name","category","aged_value"]].sort_values("aged_value",ascending=False)
        if pushes.empty: st.info("No Push items were identified from this upload.")
        else:
            chosen=st.selectbox("Choose a Push item",pushes.sku.tolist(),format_func=lambda s:f"{s} — {pushes.loc[pushes.sku.eq(s),'name'].iloc[0]}")
            itemrow=pushes[pushes.sku.eq(chosen)].iloc[0]
            st.markdown(f"**{itemrow['name']}** · {itemrow['category']} · aged value ₹{itemrow['aged_value']:,.2f}")
            t=result["targets"]
            t=t[t.sku.eq(chosen)].sort_values(["bucket","customer_purchase_value"],ascending=[True,False]) if not t.empty else t
            if t.empty: st.warning("No qualifying customer targets matched the case rules for this item.")
            else:
                for i,(_,r) in enumerate(t.iterrows(),1):
                    st.markdown(f"<div class='mobile-call'><b>{i}. {r['customer']}</b><br>Bucket {r['bucket']} — {r['bucket_label']}<br><span class='small-muted'>6-month spend: ₹{r['customer_purchase_value']:,.2f} · Industry: {r['industry'] or 'Not recorded'} · Area: {r['area'] or 'Not recorded'}</span></div>",unsafe_allow_html=True)
                st.download_button("Download this call list (CSV)",t.to_csv(index=False).encode("utf-8-sig"),file_name=f"{chosen}_customer_targets.csv",mime="text/csv")
        st.subheader("Data-quality review")
        if result["issues"]:
            for issue in result["issues"]: st.warning(issue)
        else: st.success("No issues were detected by the configured checks. This does not guarantee source data is error-free.")
        if not result["inactive_stock"].empty:
            st.markdown("**Inactive items with stock (excluded from active analysis)**")
            st.dataframe(result["inactive_stock"],use_container_width=True,hide_index=True)
        st.download_button("Export complete review to Excel",excel_bytes(result),file_name="CompanyX_Slow_Stock_Review.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with st.expander("Method and assumptions"):
            st.markdown("- Six-month sales are Apr–Sep 2026; stock date is 30-Sep-2026, as specified in the case.\n- A SWITCH item is mapped to the SKU whose Item_Master name matches its Alternate Item field.\n- The destination/current SKU's category, pack and cost are used after consolidation.\n- Missing numeric fields are treated as zero only to keep the run operational and are flagged for correction.\n- Customer eligibility is strictly greater than ₹1,00,000 total sales value. Bucket precedence is 1, then 2, then 3.")
    except Exception as e:
        st.error(f"Could not run the review: {e}")
        st.markdown("Check that all four tables are uploaded and their column headers correspond to the case fields. The app does not silently create missing tables or fake results.")

st.divider()
st.caption("Prototype for the Company X case study. Validate outputs with Supply Chain before operational use; upload data is processed for the current session only.")
