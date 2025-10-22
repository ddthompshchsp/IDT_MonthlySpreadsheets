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

# ---------------- UI SETUP ----------------
st.set_page_config(page_title='Department Packets - Wide Excel ZIP (v3.5)', layout='wide')
st.title('Department Packets - Wide Excel ZIP (v3.5)')

st.markdown('''
**FSW values come directly from your FSW_Master for the selected month(s).**  
Each department workbook shows the FSW’s true values for that department’s metrics, with a side-by-side column for the department to enter/validate.  
Toggle the **"Fill missing FSW metric values with 0"** to prefill blanks as 0 (off by default).
''')

# ---------------- CONSTANTS ----------------
AREA_SHEETS = ['Central','West','East']
READONLY_FILL = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
GREEN_FILL    = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
RED_FILL      = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
VALID_STATUSES = ['Validated','Mismatch','Unable to Validate']

# ---------------- HELPERS ----------------
def _norm(s):
    return s.replace('\u2013','-').replace('\u2014','-').strip() if isinstance(s,str) else s

def load_any(uploaded):
    if uploaded is None:
        return None
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
    s = ''
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def add_styles_filters(ws):
    # Header
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical='center')
    ws.freeze_panes = 'E2'
    last_col = excel_col(ws.max_column)
    ws.auto_filter.ref = f'A1:{last_col}{ws.max_row}'
    # Column widths
    for i in range(1, ws.max_column+1):
        ws.column_dimensions[excel_col(i)].width = 16
    ws.column_dimensions['A'].width = 8   # Month
    ws.column_dimensions['B'].width = 12  # Area
    ws.column_dimensions['C'].width = 18  # Campus
    ws.column_dimensions['D'].width = 22  # FSW
    # Read-only background for first 4 (IDs)
    for col_letter in ['A','B','C','D']:
        for cell in ws[col_letter][1:]:
            cell.fill = READONLY_FILL

def add_match_colors(ws, metric_pairs, start_row=2, end_row=None):
    if end_row is None:
        end_row = ws.max_row
    # Add conditional red/green fill comparing dept entry vs FSW value
    for fsw_col, dept_col in metric_pairs:
        for r in range(start_row, end_row+1):
            d_cell = f'{dept_col}{r}'
            c_cell = f'{fsw_col}{r}'
            # RED when both present and not equal
            ws.conditional_formatting.add(
                d_cell,
                FormulaRule(formula=[f'AND({d_cell}<>"", {c_cell}<>"", {d_cell}<> {c_cell})'],
                            stopIfTrue=False, fill=RED_FILL)
            )
            # GREEN when both present and equal
            ws.conditional_formatting.add(
                d_cell,
                FormulaRule(formula=[f'AND({d_cell}<>"", {c_cell}<>"", {d_cell}= {c_cell})'],
                            stopIfTrue=False, fill=GREEN_FILL)
            )

def add_validations(ws, col_letter, max_row):
    dv = DataValidation(type='list', formula1='"' + ','.join(VALID_STATUSES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f'{col_letter}2:{col_letter}{max_row}')

def ensure_roster(master_month_df, roster_df=None):
    """Guarantee we have full FSW list per area/campus (roster preferred)."""
    if roster_df is not None and set(['Area','Campus','FSW']).issubset(roster_df.columns):
        roster = roster_df[['Area','Campus','FSW']].drop_duplicates()
    else:
        roster = master_month_df[['Area','Campus','FSW']].drop_duplicates()
    roster = roster.assign(
        Area   = roster['Area'].astype(str).str.strip(),
        Campus = roster['Campus'].astype(str).str.strip(),
        FSW    = roster['FSW'].astype(str).str.strip()
    )
    return roster.sort_values(['Area','Campus','FSW'])

def build_wide_sheet(wb, sheet_name, month_value, roster_area_df, fsw_month_area_df, metrics_order, fill_missing_zero=False):
    ws = wb.create_sheet(sheet_name)

    # Interleaved headers: Metric, Department - Metric, ...
    interleaved = []
    for m in metrics_order:
        interleaved += [m, f'Department - {m}']

    headers = ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']
    ws.append(headers)

    # Pivot FSW values (from master) for this area + month
    if not fsw_month_area_df.empty:
        pv = fsw_month_area_df.pivot_table(index=['Area','Campus','FSW'],
                                           columns='Metric', values='FSW_Value', aggfunc='first')
        for m in metrics_order:
            if m not in pv.columns:
                pv[m] = np.nan
        pv = pv[metrics_order].reset_index()
    else:
        pv = pd.DataFrame(columns=['Area','Campus','FSW'] + metrics_order)

    # Left-join: ensure every FSW appears
    base = roster_area_df.copy()
    if not pv.empty:
        base = base.merge(pv, on=['Area','Campus','FSW'], how='left')
    else:
        for m in metrics_order:
            base[m] = np.nan

    # Optional: fill missing FSW values as zero
    if fill_missing_zero and metrics_order:
        base[metrics_order] = base[metrics_order].fillna(0)

    # Department entry columns + trailing admin cols
    for m in metrics_order:
        base[f'Department - {m}'] = ''  # department fills these
    base['Validated'] = ''
    base['Validation_Date'] = ''
    base['Issues'] = ''
    base['Services'] = ''
    base['Referrals'] = ''
    base['Notes'] = ''

    base.insert(0, 'Month', month_value)
    ordered = ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']
    base = base[ordered]

    for r in dataframe_to_rows(base, index=False, header=False):
        ws.append(r)

    add_styles_filters(ws)

    # Conditional formatting column pairs
    metric_pairs = []
    start_idx = 5  # Month..FSW = 4, so first metric value starts at col 5
    for i, m in enumerate(metrics_order):
        fsw_idx  = start_idx + 2*i       # FSW value col
        dept_idx = start_idx + 2*i + 1   # Dept entry col
        metric_pairs.append((excel_col(fsw_idx), excel_col(dept_idx)))
    add_match_colors(ws, metric_pairs)

    # Validated dropdown
    validated_idx = 4 + 2*len(metrics_order) + 1
    add_validations(ws, excel_col(validated_idx), ws.max_row)

    return ws

def build_workbook_bytes(master_month_df, month_value, metrics_order, roster_df=None, fill_missing_zero=False):
    wb = Workbook()
    wb.remove(wb.active)

    roster_all = ensure_roster(master_month_df, roster_df)

    # For each Area sheet, slice month data by area to pivot correct FSW values
    for area in AREA_SHEETS:
        roster_area = roster_all[roster_all['Area'].str.casefold() == area.lower()].copy()
        fsw_area    = master_month_df[master_month_df['Area'].astype(str).str.strip().str.casefold() == area.lower()].copy()
        build_wide_sheet(
            wb, area, month_value,
            roster_area_df=roster_area,
            fsw_month_area_df=fsw_area,
            metrics_order=metrics_order,
            fill_missing_zero=fill_missing_zero
        )

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

# ---------------- INPUTS -----------------
left, right = st.columns(2)
with left:
    fsw_up = st.file_uploader('FSW_Master (xlsx/csv)', type=['xlsx','csv'])
with right:
    mm_up = st.file_uploader('metrics_map.csv', type=['csv'])

roster_up = st.file_uploader('Optional: Roster (xlsx/csv) with Area, Campus, FSW', type=['xlsx','csv'])

if fsw_up is None or mm_up is None:
    st.info('Upload FSW_Master and metrics_map to continue.')
    st.stop()

fsw = load_any(fsw_up)
mmap = load_any(mm_up)
roster_df = load_any(roster_up) if roster_up is not None else None

# Basic checks
need_fsw = {'Month','Area','Campus','FSW','Metric','Value'}
missing = need_fsw - set(fsw.columns)
if missing:
    st.error(f'FSW_Master missing columns: {sorted(missing)}')
    st.stop()

if not {'Department','Metric'}.issubset(mmap.columns):
    st.error('metrics_map.csv must have columns: Department, Metric')
    st.stop()

# Normalize and join master to department map
fsw = fsw.rename(columns={'Value':'FSW_Value'})
# keep order of metrics within each department as in the CSV
mmap['Department'] = mmap['Department'].astype(str).str.strip()
mmap['Metric']     = mmap['Metric'].astype(str).str.strip()
master = fsw.merge(mmap[['Metric','Department']], on='Metric', how='left')

# ---------------- FILTERS ----------------
st.subheader('Filters')
c1, c2, c3 = st.columns(3)

months_all = [m for m in ['Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May']
              if m in master['Month'].dropna().unique().tolist()]
depts_all  = [d for d in mmap['Department'].dropna().unique().tolist()]
camp_all   = sorted(master['Campus'].dropna().unique().tolist())

months = c1.multiselect('Months', months_all, default=months_all)
depts  = c2.multiselect('Departments', sorted(depts_all), default=sorted(depts_all))
camp   = c3.multiselect('Campuses', camp_all, default=camp_all)

fill_zero = st.checkbox('Fill missing FSW metric values with 0', value=False)

mf = master.copy()
if months: mf = mf[mf['Month'].isin(months)]
if camp:   mf = mf[mf['Campus'].isin(camp)]

# ---------------- BUILD ZIP ----------------
if st.button('Build ZIP of Department Packets (Excel)'):
    if mf.empty:
        st.warning('No rows after filters.')
    else:
        memzip = BytesIO()
        with ZipFile(memzip, 'w', ZIP_DEFLATED) as zf:
            for month in sorted(mf['Month'].dropna().unique().tolist()):
                m_month = mf[mf['Month'] == month].copy()
                for dept in (depts if depts else sorted(mmap['Department'].dropna().unique())):
                    # Metrics for this department, preserving order from CSV
                    metrics_order = mmap.loc[mmap['Department'] == dept, 'Metric'].tolist()

                    # Slice FSW master to only this dept's metrics (so FSW values are correct columns)
                    month_for_metrics = m_month[m_month['Metric'].isin(metrics_order)][
                        ['Month','Area','Campus','FSW','Metric','FSW_Value']
                    ].copy()

                    xbytes = build_workbook_bytes(
                        master_month_df=month_for_metrics,
                        month_value=month,
                        metrics_order=metrics_order,
                        roster_df=roster_df,
                        fill_missing_zero=fill_zero
                    )
                    zf.writestr(f"{month}/{dept}_DEPT.xlsx", xbytes)

        memzip.seek(0)
        stamp = datetime.now().strftime('%Y%m%d_%H%M')
        st.download_button('Download Department Packets ZIP', data=memzip,
                           file_name=f'DeptPackets_DEPT_{stamp}.zip', mime='application/zip')
