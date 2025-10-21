import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import datetime

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import FormulaRule

st.set_page_config(page_title='Department Packets - Wide Excel ZIP', layout='wide')
st.title('Department Packets - Wide Excel ZIP')

st.markdown("""
Build prefilled Excel packets for departments to enter validations.

- Upload **FSW_Master.xlsx/.csv** and **metrics_map.csv**
- Generates a ZIP: one **Excel workbook** per **Month x Department**
- Sheets inside each workbook: **Central, West, East**
- Columns: **Area | Campus | FSW | [Metric] | Department - [Metric]** for every metric
- Conditional formatting: green if Department equals FSW value, red if it differs
""")

AREA_SHEETS = ['Central','West','East']
READONLY_FILL = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
GREEN_FILL = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
RED_FILL   = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

def _norm(s):
    return s.replace('\u2013','-').replace('\u2014','-').strip() if isinstance(s,str) else s

def load_any(uploaded):
    if uploaded is None: return None
    name = getattr(uploaded, 'name', '').lower()
    if name.endswith('.csv'):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded, engine='openpyxl')
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(_norm)
    return df

def excel_col(n:int)->str:
    s=''
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def add_styles_filters(ws):
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical='center')
    ws.freeze_panes = 'D2'
    last_col = excel_col(ws.max_column)
    ws.auto_filter.ref = f'A1:{last_col}{ws.max_row}'
    for i in range(1, ws.max_column+1):
        ws.column_dimensions[excel_col(i)].width = 16
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 22
    for cell in ws['A'][1:]: cell.fill = READONLY_FILL
    for cell in ws['B'][1:]: cell.fill = READONLY_FILL
    for cell in ws['C'][1:]: cell.fill = READONLY_FILL

def add_match_colors(ws, metrics_order):
    num_metrics = len(metrics_order)
    start_fsw = 4
    start_dept = 4 + num_metrics
    for i in range(num_metrics):
        fsw_col = excel_col(start_fsw + i)
        dept_col = excel_col(start_dept + i)
        for r in range(2, ws.max_row+1):
            d_cell = f'{dept_col}{r}'
            c_cell = f'{fsw_col}{r}'
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

def build_wide_sheet(wb, sheet_name, df_area, metrics_order):
    ws = wb.create_sheet(sheet_name)
    headers = ['Area','Campus','FSW'] + metrics_order + [f'Department - {m}' for m in metrics_order]
    ws.append(headers)

    if not df_area.empty:
        pv = df_area.pivot_table(index=['Area','Campus','FSW'], columns='Metric', values='FSW_Value', aggfunc='first')
        for m in metrics_order:
            if m not in pv.columns:
                pv[m] = np.nan
        pv = pv[metrics_order]
        wide = pv.reset_index()
        for m in metrics_order:
            wide[f'Department - {m}'] = np.nan
        for r in dataframe_to_rows(wide, index=False, header=False):
            ws.append(r)

    add_styles_filters(ws)
    add_match_colors(ws, metrics_order)
    return ws

def build_workbook_bytes(df_md, metrics_order):
    wb = Workbook()
    wb.remove(wb.active)

    for area in AREA_SHEETS:
        sub = df_md[df_md['Area'].astype(str).str.strip().str.casefold() == area.lower()]
        build_wide_sheet(wb, area, sub, metrics_order)

    ws = wb.create_sheet('Summary', 0)
    ws['A1'].value = 'Dashboard / Summary'
    ws['A1'].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(['Area','Total Rows','Department Cells Entered'])
    for cell in ws[3]:
        cell.font = Font(bold=True)
    row = 4
    for area in AREA_SHEETS:
        sname = area
        ws[f'A{row}'] = area
        ws[f'B{row}'] = f"=MAX(ROWS('{sname}'!A2:A1048576),0)"
        ws[f'C{row}'] = f"=SUMPRODUCT(--(LEN('{sname}'!D2:ZZ1048576)>0))"
        row += 1
    for col in range(1, 4):
        ws.column_dimensions[excel_col(col)].width = 26
    ws.freeze_panes = 'A4'

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

# UI
left, right = st.columns(2)
with left:
    fsw_up = st.file_uploader('FSW_Master (xlsx/csv)', type=['xlsx','csv'])
with right:
    mm_up = st.file_uploader('metrics_map.csv', type=['csv'])

if fsw_up is None or mm_up is None:
    st.info('Upload both files to continue.')
    st.stop()

fsw = load_any(fsw_up)
mmap = load_any(mm_up)

need_fsw = {'Month','Area','Campus','FSW','Metric','Value'}
if missing := (need_fsw - set(fsw.columns)):
    st.error(f'FSW_Master missing columns: {sorted(missing)}'); st.stop()
if not {'Department','Metric'}.issubset(mmap.columns):
    st.error('metrics_map.csv must have columns: Department, Metric'); st.stop()

fsw = fsw.rename(columns={'Value':'FSW_Value'})
master = fsw.merge(mmap[['Metric','Department']], on='Metric', how='left').dropna(subset=['Department'])

st.subheader('Filters')
c1,c2,c3 = st.columns(3)
months_all = [m for m in ['Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May'] if m in master['Month'].unique().tolist()]
depts_all  = sorted(master['Department'].unique().tolist())
camp_all   = sorted(master['Campus'].dropna().unique().tolist())

months = c1.multiselect('Months', months_all, default=months_all)
depts  = c2.multiselect('Departments', depts_all, default=depts_all)
camp   = c3.multiselect('Campuses', camp_all, default=camp_all)

f = master.copy()
if months: f = f[f['Month'].isin(months)]
if depts:  f = f[f['Department'].isin(depts)]
if camp:   f = f[f['Campus'].isin(camp)]

st.write(f"Rows selected: **{len(f)}** | FSWs: **{f['FSW'].nunique()}** | Metrics: **{f['Metric'].nunique()}**")

if st.button('Build ZIP of Department Packets (Excel)'):
    if f.empty:
        st.warning('No rows after filters.')
    else:
        memzip = BytesIO()
        with ZipFile(memzip, 'w', ZIP_DEFLATED) as zf:
            for (month, dept), chunk in f.groupby(['Month','Department'], dropna=False):
                metrics_order = mmap[mmap['Department'] == dept]['Metric'].tolist()
                chunk = chunk[chunk['Metric'].isin(metrics_order)].copy()
                xbytes = build_workbook_bytes(chunk, metrics_order)
                zf.writestr(f"{month}/{dept}_DEPT.xlsx", xbytes)
        memzip.seek(0)
        stamp = datetime.now().strftime('%Y%m%d_%H%M')
        st.download_button('Download Department Packets ZIP', data=memzip, file_name=f'DeptPackets_DEPT_{stamp}.zip', mime='application/zip')
