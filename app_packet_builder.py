import streamlit as st
import pandas as pd
import numpy as np
from datetime import date

st.set_page_config(page_title="FSW Validation — Wide View", layout="wide")
st.title("FSW Validation — Wide View (Metrics as Columns)")

st.markdown("""
Upload **FSW_Master** and **metrics_map.csv**. Filter by Month/Area/Campus/FSW/Department.  
This view shows one row per FSW with **metrics as columns**, plus editable **- Validated** columns on the right.
""")

def _norm(s):
    return s.replace("\u2013","-").replace("\u2014","-").strip() if isinstance(s,str) else s

def load_any(uploaded):
    if uploaded is None: return None
    name = getattr(uploaded, "name", "").lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded, engine="openpyxl")
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(_norm)
    return df

# Uploads
left, right = st.columns(2)
with left:
    fsw_up = st.file_uploader("FSW_Master (xlsx/csv)", type=["xlsx","csv"])
with right:
    map_up = st.file_uploader("metrics_map.csv", type=["csv"])

fsw = load_any(fsw_up)
mmap = load_any(map_up)

if fsw is None or mmap is None:
    st.info("Upload both files to continue.")
    st.stop()

need_fsw = {"Month","Area","Campus","FSW","Metric","Value"}
if missing := (need_fsw - set(fsw.columns)):
    st.error(f"FSW_Master missing columns: {sorted(missing)}"); st.stop()
if not {"Department","Metric"}.issubset(mmap.columns):
    st.error("metrics_map.csv must have columns: Department, Metric"); st.stop()

fsw = fsw.rename(columns={"Value":"FSW_Value"})
mmap["Metric"] = mmap["Metric"].map(_norm)
mmap["Department"] = mmap["Department"].map(_norm)
master = fsw.merge(mmap[["Metric","Department"]], on="Metric", how="left").dropna(subset=["Department"])

# Filters
st.sidebar.header("Filters")
months = [m for m in ["Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May"] if m in master["Month"].unique().tolist()]
depts  = sorted(master["Department"].unique().tolist())
areas  = sorted(master["Area"].dropna().unique().tolist())
campi  = sorted(master["Campus"].dropna().unique().tolist())
fsws   = sorted(master["FSW"].dropna().unique().tolist())

m_sel = st.sidebar.multiselect("Month(s)", months, default=months)
d_sel = st.sidebar.multiselect("Department(s)", depts, default=depts)
a_sel = st.sidebar.multiselect("Area(s)", areas, default=areas)
c_sel = st.sidebar.multiselect("Campus(es)", campi, default=campi)
f_sel = st.sidebar.multiselect("FSW(s)", fsws, default=fsws)

f = master.copy()
if m_sel: f = f[f["Month"].isin(m_sel)]
if d_sel: f = f[f["Department"].isin(d_sel)]
if a_sel: f = f[f["Area"].isin(a_sel)]
if c_sel: f = f[f["Campus"].isin(c_sel)]
if f_sel: f = f[f["FSW"].isin(f_sel)]

# Choose one department for a consistent metric set (recommended)
dept_one = st.sidebar.selectbox("Focus Department (optional)", ["(All)"] + depts, index=0)
if dept_one != "(All)":
    f = f[f["Department"] == dept_one]
    metrics_order = mmap[mmap["Department"] == dept_one]["Metric"].tolist()
else:
    metrics_order = sorted(f["Metric"].dropna().unique().tolist())

# Wide table
show_month = st.sidebar.checkbox("Include Month as a column", value=True)
id_cols = ["Area","Campus","FSW"] + (["Month"] if show_month else [])
if f.empty:
    st.warning("No rows after filters."); st.stop()

pv = f.pivot_table(index=id_cols, columns="Metric", values="FSW_Value", aggfunc="first")

# Ensure all chosen metrics exist as columns
for m in metrics_order:
    if m not in pv.columns:
        pv[m] = np.nan
pv = pv[metrics_order]  # order
wide = pv.reset_index()

# Add editable validated columns
for m in metrics_order:
    wide[f"{m} - Validated"] = np.nan

st.subheader("Wide Validation Table")
edited = st.data_editor(
    wide,
    use_container_width=True,
    height=620,
    column_config={
        **{m: st.column_config.NumberColumn(help="FSW value (read-only here)") for m in metrics_order},
        **{f"{m} - Validated": st.column_config.NumberColumn(help="Enter validated value") for m in metrics_order},
    },
)

# Quick per-row match %
def row_match_status(row):
    flags = []
    for m in metrics_order:
        v = row.get(m)
        vv = row.get(f"{m} - Validated")
        if pd.notna(vv):
            flags.append(int(pd.notna(v) and (vv == v)))
    if len(flags)==0:
        return np.nan
    return round(100.0 * sum(flags)/len(flags), 1)

out = edited.copy()
out["% Metrics Matched (row)"] = out.apply(row_match_status, axis=1)

with st.expander("Preview Output", expanded=False):
    st.dataframe(out, use_container_width=True)

# Export CSV
fname = "Dept_Validated_WIDE.csv" if dept_one in ("", "(All)") else f"Dept_Validated_WIDE_{dept_one}.csv"
st.download_button(
    "⬇️ Export Wide CSV",
    out.to_csv(index=False).encode("utf-8"),
    file_name=fname,
    mime="text/csv"
)

st.caption("Tip: Pick one Department to keep the metric set consistent across columns.")
