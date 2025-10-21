
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
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(page_title='Department Packets - Wide Excel ZIP (v3.1)', layout='wide')
st.title('Department Packets - Wide Excel ZIP (v3.1)')

st.markdown('''
This version fills **0** for missing FSW metric values (so departments see 0 instead of blank).
One Excel per Month x Department, with sheets: Central, West, East.
''')

AREA_SHEETS = ['Central','West','East']
READONLY_FILL = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
GREEN_FILL = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
RED_FILL   = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
VALID_STATUSES = ['Validated','Mismatch','Unable to Validate']

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
    ws.freeze_panes = 'E2'
    last_col = excel_col(ws.max_column)
    ws.auto_filter.ref = f'A1:{last_col}{ws.max_row}'
    for i in range(1, ws.max_column+1):
        ws.column_dimensions[excel_col(i)].width = 16
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 22
    for col_letter in ['A','B','C','D']:
        for cell in ws[col_letter][1:]:
            cell.fill = READONLY_FILL

def add_match_colors(ws, metric_pairs, start_row=2, end_row=None):
    if end_row is None:
        end_row = ws.max_row
    for fsw_col, dept_col in metric_pairs:
        for r in range(start_row, end_row+1):
            d_cell = f'{dept_col}{r}'
            c_cell = f'{fsw_col}{r}'
            ws.conditional_formatting.add(
                d_cell,
                FormulaRule(formula=[f'AND({d_cell}<>"", {c_cell}<>"", {d_cell}<>${c_cell})'],
                            stopIfTrue=False, fill=RED_FILL)
            )
            ws.conditional_formatting.add(
                d_cell,
                FormulaRule(formula=[f'AND({d_cell}<>"", {c_cell}<>"", {d_cell}=${c_cell})'],
                            stopIfTrue=False, fill=GREEN_FILL)
            )

def add_validations(ws, col_letter, max_row):
    dv = DataValidation(type='list', formula1='"' + ','.join(VALID_STATUSES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f'{col_letter}2:{col_letter}{max_row}')

def build_roster_for_month(master_month_df):
    return master_month_df[['Area','Campus','FSW']].drop_duplicates().sort_values(['Area','Campus','FSW'])

def build_wide_sheet(wb, sheet_name, month_value, roster_area_df, df_dept_area, metrics_order):
    ws = wb.create_sheet(sheet_name)

    # Interleaved headers
    interleaved = []
    for m in metrics_order:
        interleaved += [m, f'Department - {m}']

    headers = ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']
    ws.append(headers)

    # Pivot FSW values by metric
    if not df_dept_area.empty:
        pv = df_dept_area.pivot_table(index=['Area','Campus','FSW'], columns='Metric',
                                      values='FSW_Value', aggfunc='first')
        for m in metrics_order:
            if m not in pv.columns:
                pv[m] = np.nan
        pv = pv[metrics_order].reset_index()
    else:
        pv = pd.DataFrame(columns=['Area','Campus','FSW'] + metrics_order)

    base = roster_area_df.copy()
    if not pv.empty:
        base = base.merge(pv, on=['Area','Campus','FSW'], how='left')
    else:
        for m in metrics_order:
            base[m] = np.nan

    # >>> NEW: fill missing FSW metric values with 0 <<<
    if metrics_order:
        base[metrics_order] = base[metrics_order].fillna(0)

    # Add department columns and trailing columns
    for m in metrics_order:
        base[f'Department - {m}'] = ''  # keep blank for departments to fill

    base['Validated'] = ''
    base['Validation_Date'] = ''
    base['Issues'] = ''
    base['Services'] = ''
    base['Referrals'] = ''
    base['Notes'] = ''

    # Order + Month
    base.insert(0, 'Month', month_value)
    ordered_cols = ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']
    base = base[ordered_cols]

    # Write rows
    for r in dataframe_to_rows(base, index=False, header=False):
        ws.append(r)

    add_styles_filters(ws)

    # Conditional format pairs
    metric_pairs = []
    start_col_idx = 5  # Month..FSW = 4
    for i, m in enumerate(metrics_order):
        fsw_col_idx = start_col_idx + 2*i
        dept_col_idx = start_col_idx + 2*i + 1
        metric_pairs.append((excel_col(fsw_col_idx), excel_col(dept_col_idx)))
    add_match_colors(ws, metric_pairs)

    # Validated dropdown
    validated_col_idx = 4 + 2*len(metrics_order) + 1
    add_validations(ws, excel_col(validated_col_idx), ws.max_row)

    return ws

def build_workbook_bytes(master_month_df, dept_df, month_value, metrics_order):
    wb = Workbook()
    wb.remove(wb.active)

    roster = build_roster_for_month(master_month_df)

    for area in AREA_SHEETS:
        roster_area = roster[roster['Area'].astype(str).str.strip().str.casefold() == area.lower()].copy()
        dept_area = dept_df[dept_df['Area'].astype(str).str.strip().str.casefold() == area.lower()].copy()
        build_wide_sheet(wb, area, month_value, roster_area, dept_area, metrics_order)

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

# ---- UI ----
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
                master_month = master[master['Month'] == month][['Month','Area','Campus','FSW','Metric','FSW_Value','Department']].copy()
                dept_only = chunk[chunk['Metric'].isin(metrics_order)].copy()
                xbytes = build_workbook_bytes(master_month, dept_only, month, metrics_order)
                zf.writestr(f"{month}/{dept}_DEPT.xlsx", xbytes)
        memzip.seek(0)
        stamp = datetime.now().strftime('%Y%m%d_%H%M')
        st.download_button('Download Department Packets ZIP', data=memzip,
                           file_name=f'DeptPackets_DEPT_{stamp}.zip', mime='application/zip')
