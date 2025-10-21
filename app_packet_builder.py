
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

VALIDATED_LIST = ["Validated", "Mismatch", "Unable to Validate"]
READONLY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

st.set_page_config(page_title="Packet Builder (Prefilled Excel)", layout="wide")
st.title("Prefilled Department Packets (Excel)")

st.markdown("""
Upload **FSW_Master** and **metrics_map.csv**, choose filters, then click **Build ZIP**.  
The ZIP contains one workbook per **Month × Department**, with one sheet per **Campus** and built-in:
- Validated dropdown (Validated / Mismatch / Unable to Validate)  
- Validation_Date column (date)  
- Red/green highlighting when Validated_Value ≠ / = FSW_Value
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

def write_packet_excel(filename: str, df: pd.DataFrame) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    dv_status = DataValidation(type="list", formula1=f'"{",".join(VALIDATED_LIST)}"', allow_blank=True)
    dv_date = DataValidation(type="date", allow_blank=True)

    for campus, cdf in df.groupby("Campus", dropna=False):
        title = str(campus).strip()[:31] if pd.notna(campus) and str(campus).strip() else "Unknown"
        ws = wb.create_sheet(title=title)

        view = cdf[["FSW","Metric","FSW_Value"]].copy()
        view["Validated_Value"] = np.nan
        view["Validated"] = ""
        view["Validation_Date"] = ""
        view["Issues"] = ""
        view["Services"] = ""
        view["Referrals"] = ""
        view["Notes"] = ""

        headers = ["FSW","Metric","FSW_Value","Validated_Value","Validated","Validation_Date","Issues","Services","Referrals","Notes"]
        ws.append(headers)
        for r in dataframe_to_rows(view, index=False, header=False):
            ws.append(r)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:J{ws.max_row}"

        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(vertical="center")

        for cell in ws["A"][1:]: cell.fill = READONLY_FILL
        for cell in ws["B"][1:]: cell.fill = READONLY_FILL
        for cell in ws["C"][1:]: cell.fill = READONLY_FILL

        ws.add_data_validation(dv_status); dv_status.add(f"E2:E{ws.max_row}")
        ws.add_data_validation(dv_date);   dv_date.add(f"F2:F{ws.max_row}")

        for r in range(2, ws.max_row + 1):
            d_cell, c_cell = f"D{r}", f"C{r}"
            ws.conditional_formatting.add(
                d_cell,
                FormulaRule(formula=[f'AND(NOT(ISBLANK({d_cell})), {d_cell}<>${c_cell})'],
                            stopIfTrue=False, fill=RED_FILL)
            )
            ws.conditional_formatting.add(
                d_cell,
                FormulaRule(formula=[f'AND(NOT(ISBLANK({d_cell})), {d_cell}=${c_cell})'],
                            stopIfTrue=False, fill=GREEN_FILL)
            )

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

left, right = st.columns(2)
with left:
    fsw_up = st.file_uploader("FSW_Master (xlsx/csv)", type=["xlsx","csv"])
with right:
    mm_up = st.file_uploader("metrics_map.csv", type=["csv"])

if fsw_up is None or mm_up is None:
    st.info("Upload both files to continue.")
    st.stop()

fsw = load_any(fsw_up)
mmap = load_any(mm_up)

need_fsw = {"Month","Area","Campus","FSW","Metric","Value"}
if missing := (need_fsw - set(fsw.columns)):
    st.error(f"FSW_Master missing columns: {sorted(missing)}")
    st.stop()
if not {"Department","Metric"}.issubset(mmap.columns):
    st.error("metrics_map.csv must have columns: Department, Metric")
    st.stop()

fsw = fsw.rename(columns={"Value":"FSW_Value"})
master = fsw.merge(mmap[["Metric","Department"]], on="Metric", how="left")
master = master.dropna(subset=["Department"])

st.subheader("Filters")
c1,c2,c3 = st.columns(3)
months_all = [m for m in ["Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May"] if m in master["Month"].unique().tolist()]
depts_all = sorted(master["Department"].dropna().unique().tolist())
campuses_all = sorted(master["Campus"].dropna().unique().tolist())

months = c1.multiselect("Months", months_all, default=months_all)
depts = c2.multiselect("Departments", depts_all, default=depts_all)
campuses = c3.multiselect("Campuses", campuses_all, default=campuses_all)

f = master.copy()
if months:  f = f[f["Month"].isin(months)]
if depts:   f = f[f["Department"].isin(depts)]
if campuses:f = f[f["Campus"].isin(campuses)]

st.write(f"Rows selected: **{len(f)}** | FSWs: **{f['FSW'].nunique()}** | Metrics: **{f['Metric'].nunique()}**")

if st.button("Build Packet ZIP"):
    if f.empty:
        st.warning("No rows match the filters.")
    else:
        memzip = BytesIO()
        with ZipFile(memzip, "w", ZIP_DEFLATED) as zf:
            for (month, dept), chunk in f.groupby(["Month","Department"], dropna=False):
                chunk = chunk.sort_values(["Campus","FSW","Metric"]).copy()
                xbytes = write_packet_excel(f"{dept}.xlsx", chunk)
                arcname = f"{month}/{dept}.xlsx"
                zf.writestr(arcname, xbytes)
        memzip.seek(0)
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            "⬇️ Download Packets ZIP",
            data=memzip,
            file_name=f"DeptPackets_{stamp}.zip",
            mime="application/zip"
        )
