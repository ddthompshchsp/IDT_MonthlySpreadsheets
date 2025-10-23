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
st.set_page_config(page_title='Department Packets - Wide Excel ZIP', layout='wide')
st.title('Department Packets - Wide Excel ZIP')

st.markdown('''
Uploads:
1) **FSW_Master** (xlsx/csv) with columns: Month, Area, Campus, FSW, Metric, Value (or FSW_Value)  
2) **metrics_map.csv** with columns: Department, Metric (defines which metrics each department owns)

Output:
- For each selected **Month × Department**, an Excel workbook with **Central**, **West**, **East** sheets.
- Columns: Month, Area, Campus, FSW, then **each Metric** (prefilled from FSW master), then **Department – Metric** (blank for staff input), then Validated, Validation_Date, Issues, Services, Referrals, Notes.
- Cells in **Department – Metric** turn **green** if they match the FSW value, **red** if not.
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

def normalize_master_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize flexible headers to: Month, Area, Campus, FSW, Metric, FSW_Value
    Accepts case-insensitive synonyms for Area/Campus/FSW/Value.
    """
    lower = {str(c).strip().lower(): c for c in df.columns}

    def pick(*cands):
        for c in cands:
            if c in lower:
                return lower[c]
        return None

    col_month  = pick('month')
    col_area   = pick('area','region','cluster')
    col_campus = pick('campus','center','center/campus','site','school','location')
    col_fsw    = pick('fsw','fsw name','family service worker','family advocate','staff name')
    col_metric = pick('metric','metrics')
    col_value  = pick('value','fsw_value','fsw value')

    needed = [col_month, col_area, col_campus, col_fsw, col_metric, col_value]
    if any(c is None for c in needed):
        return df  # caller will validate and show error

    out = df.rename(columns={
        col_month:  'Month',
        col_area:   'Area',
        col_campus: 'Campus',
        col_fsw:    'FSW',
        col_metric: 'Metric',
        col_value:  'FSW_Value',
    })
    for c in ['Month','Area','Campus','FSW','Metric']:
        out[c] = out[c].astype(str).str.strip()
    return out

def excel_col(n:int)->str:
    s = ''
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def add_styles_filters(ws):
    # Header styling
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical='center')
    ws.freeze_panes = 'E2'

    # Auto-filter
    last_col = excel_col(ws.max_column)
    ws.auto_filter.ref = f'A1:{last_col}{ws.max_row}'

    # Column widths
    for i in range(1, ws.max_column+1):
        ws.column_dimensions[excel_col(i)].width = 16
    ws.column_dimensions['A'].width = 8   # Month
    ws.column_dimensions['B'].width = 12  # Area
    ws.column_dimensions['C'].width = 18  # Campus
    ws.column_dimensions['D'].width = 22  # FSW

    # Gray read-only for Month/Area/Campus/FSW rows
    for col_letter in ['A','B','C','D']:
        for cell in ws[col_letter][1:]:
            cell.fill = READONLY_FILL

def add_match_colors(ws, metric_pairs, start_row=2, end_row=None):
    if end_row is None:
        end_row = ws.max_row
    if end_row < start_row:
        return
    # For each Metric: compare Department cell vs FSW cell
    for fsw_col, dept_col in metric_pairs:
        # Apply over the data range; Excel will adjust the row in the formula
        ws.conditional_formatting.add(
            f'{dept_col}{start_row}:{dept_col}{end_row}',
            FormulaRule(formula=[f'AND({dept_col}{start_row}<>"", {fsw_col}{start_row}<>"", {dept_col}{start_row}<> {fsw_col}{start_row})'],
                        stopIfTrue=False, fill=RED_FILL)
        )
        ws.conditional_formatting.add(
            f'{dept_col}{start_row}:{dept_col}{end_row}',
            FormulaRule(formula=[f'AND({dept_col}{start_row}<>"", {fsw_col}{start_row}<>"", {dept_col}{start_row}= {fsw_col}{start_row})'],
                        stopIfTrue=False, fill=GREEN_FILL)
        )

def add_validations(ws, col_letter, max_row):
    if max_row < 2:
        return  # avoid invalid ranges like K2:K1
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

def build_wide_sheet(
    wb,
    sheet_name,
    month_value,
    roster_area_df,
    fsw_month_area_df,
    metrics_order,
    fill_missing_zero=False,          # affects FSW value columns only
):
    ws = wb.create_sheet(sheet_name)

    # Build headers: interleaved (FSW metric, Department metric)
    interleaved = []
    for m in metrics_order:
        interleaved += [m, f'Department - {m}']

    headers = ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']
    ws.append(headers)

    # Pivot FSW values for this area+month -> one row per FSW with metric columns
    if not fsw_month_area_df.empty:
        pv = fsw_month_area_df.pivot_table(index=['Area','Campus','FSW'],
                                           columns='Metric', values='FSW_Value', aggfunc='first')
        # Ensure all requested metrics exist as columns
        for m in metrics_order:
            if m not in pv.columns:
                pv[m] = np.nan
        pv = pv[metrics_order].reset_index()
    else:
        pv = pd.DataFrame(columns=['Area','Campus','FSW'] + metrics_order)

    # Left-join roster so every FSW appears (even if they have no values)
    base = roster_area_df.copy() if (roster_area_df is not None and not roster_area_df.empty) else pd.DataFrame(columns=['Area','Campus','FSW'])
    if not pv.empty:
        base = base.merge(pv, on=['Area','Campus','FSW'], how='left')
    else:
        for m in metrics_order:
            base[m] = np.nan

    # Optionally fill missing FSW values with zero (does NOT touch Department columns)
    if fill_missing_zero and metrics_order:
        present = [m for m in metrics_order if m in base.columns]
        if present:
            base[present] = base[present].fillna(0)

    # Department entry columns should be BLANK for staff input
    for m in metrics_order:
        base[f'Department - {m}'] = ''  # <-- BLANK by design

    # Admin columns
    base['Validated'] = ''
    base['Validation_Date'] = ''
    base['Issues'] = ''
    base['Services'] = ''
    base['Referrals'] = ''
    base['Notes'] = ''

    # Month column first
    base.insert(0, 'Month', month_value)

    # Ensure all header columns exist before ordering
    for col in ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']:
        if col not in base.columns:
            base[col] = []

    # Reorder columns
    ordered = ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']
    base = base.loc[:, ordered]

    # Write rows
    if not base.empty:
        for r in dataframe_to_rows(base, index=False, header=False):
            ws.append(r)

    # Styling, filters
    add_styles_filters(ws)

    # Conditional formatting pairs: compare Department vs FSW metric
    metric_pairs = []
    start_idx = 5  # Month..FSW = 4 columns
    for i, m in enumerate(metrics_order):
        fsw_idx  = start_idx + 2*i
        dept_idx = start_idx + 2*i + 1
        metric_pairs.append((excel_col(fsw_idx), excel_col(dept_idx)))
    if ws.max_row >= 2 and metric_pairs:
        add_match_colors(ws, metric_pairs, start_row=2, end_row=ws.max_row)

    # Validated dropdown
    validated_idx = 4 + 2*len(metrics_order) + 1
    if ws.max_row >= 2:
        add_validations(ws, excel_col(validated_idx), ws.max_row)

    return ws

def build_workbook_bytes(
    master_month_df,
    month_value,
    metrics_order,
    roster_df=None,
    fill_missing_zero=False,
):
    wb = Workbook()
    wb.remove(wb.active)

    roster_all = ensure_roster(master_month_df, roster_df)

    for area in AREA_SHEETS:
        roster_area = roster_all[roster_all['Area'].str.casefold() == area.lower()].copy()
        fsw_area    = master_month_df[master_month_df['Area'].astype(str).str.strip().str.casefold() == area.lower()].copy()
        build_wide_sheet(
            wb, area, month_value,
            roster_area_df=roster_area,
            fsw_month_area_df=fsw_area,
            metrics_order=metrics_order,
            fill_missing_zero=fill_missing_zero,
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

fsw_raw = load_any(fsw_up)
mmap = load_any(mm_up)
roster_df = load_any(roster_up) if roster_up is not None else None

# Normalize master headers to: Month, Area, Campus, FSW, Metric, FSW_Value
fsw = normalize_master_columns(fsw_raw)

# Validate required columns after normalization
need_cols = {'Month','Area','Campus','FSW','Metric','FSW_Value'}
missing = need_cols - set(fsw.columns)
if missing:
    st.error(f"FSW_Master is missing columns: {sorted(missing)}")
    st.stop()

# metrics_map check
if not {'Department','Metric'}.issubset(mmap.columns):
    st.error('metrics_map.csv must have columns: Department, Metric')
    st.stop()

# Clean map values
mmap['Department'] = mmap['Department'].astype(str).str.strip()
mmap['Metric']     = mmap['Metric'].astype(str).str.strip()

# Attach Department to master rows (keeps all master metrics)
master = fsw.merge(mmap[['Metric','Department']], on='Metric', how='left')

# ---------------- FILTERS ----------------
st.subheader('Filters & Options')
c1, c2, c3 = st.columns(3)

months_all = [m for m in ['Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May']
              if m in master['Month'].dropna().unique().tolist()]
depts_all  = sorted(mmap['Department'].dropna().unique().tolist())
camp_all   = sorted(master['Campus'].dropna().unique().tolist())

months = c1.multiselect('Months', months_all, default=months_all)
depts  = c2.multiselect('Departments', depts_all, default=depts_all)
camp   = c3.multiselect('Campuses', camp_all, default=camp_all)

fill_zero = st.checkbox('Fill missing **FSW** metric values with 0 (Department columns remain blank)', value=False)

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
                # Departments to generate (use selection if provided)
                for dept in (depts if depts else sorted(mmap['Department'].dropna().unique())):
                    # Ordered metrics for this department
                    metrics_order = mmap.loc[mmap['Department'] == dept, 'Metric'].tolist()

                    # Slice the FSW master to only this dept's metrics
                    month_for_metrics = m_month[m_month['Metric'].isin(metrics_order)][
                        ['Month','Area','Campus','FSW','Metric','FSW_Value']
                    ].copy()

                    xbytes = build_workbook_bytes(
                        master_month_df=month_for_metrics,
                        month_value=month,
                        metrics_order=metrics_order,
                        roster_df=roster_df,
                        fill_missing_zero=fill_zero,  # FSW columns only
                    )

                    # Safe filename
                    safe_dept = "".join(ch for ch in dept if ch.isalnum() or ch in (" ","-","_")).strip().replace(" ","_")
                    zf.writestr(f"{month}/{safe_dept}_DEPT.xlsx", xbytes)

        memzip.seek(0)
        stamp = datetime.now().strftime('%Y%m%d_%H%M')
        st.download_button('Download Department Packets ZIP', data=memzip,
                           file_name=f'DeptPackets_{stamp}.zip', mime='application/zip')
