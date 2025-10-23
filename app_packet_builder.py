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

# ---------- UI SETUP ----------
st.set_page_config(page_title='Department Packets - Wide Excel ZIP (hardened)', layout='wide')
st.title('Department Packets - Wide Excel ZIP (hardened)')

st.markdown("""
**Upload:**
1) `FSW_Master` (xlsx/csv) with columns: **Month, Area, Campus, FSW, Metric, Value** (or **FSW_Value**)
2) `metrics_map.csv` with columns: **Department, Metric** (the exact metric names to appear for each department)

**Output per Month × Department:**
- Sheets: **Central**, **West**, **East**
- Columns: **Month, Area, Campus, FSW** (prefilled from master, gray)  
  Then for each metric: **[FSW metric column (prefilled)]**, **[Department - metric (blank)]**  
  Then **Validated, Validation_Date, Issues, Services, Referrals, Notes**
- Conditional formatting: **green** if Dept value = FSW value, **red** otherwise.
""")

AREA_SHEETS = ['Central','West','East']
READONLY_FILL = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
GREEN_FILL    = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
RED_FILL      = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
VALID_STATUSES = ['Validated','Mismatch','Unable to Validate']

# ---------- Helpers ----------
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
    """Normalize flexible headers to: Month, Area, Campus, FSW, Metric, FSW_Value."""
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
        return df  # caller will error

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
    if end_row < start_row:
        return
    for fsw_col, dept_col in metric_pairs:
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
        return
    dv = DataValidation(type='list', formula1='"' + ','.join(VALID_STATUSES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f'{col_letter}2:{col_letter}{max_row}')

def ensure_roster_union(master_month_df, roster_df=None):
    """
    Return the union of FSWs found in master_month_df and in roster (if provided),
    so we never drop rows because roster is incomplete.
    """
    from_master = master_month_df[['Area','Campus','FSW']].drop_duplicates()
    if roster_df is not None and set(['Area','Campus','FSW']).issubset(roster_df.columns):
        from_roster = roster_df[['Area','Campus','FSW']].drop_duplicates()
        all_rows = pd.concat([from_master, from_roster], ignore_index=True).drop_duplicates()
    else:
        all_rows = from_master
    all_rows = all_rows.assign(
        Area   = all_rows['Area'].astype(str).str.strip(),
        Campus = all_rows['Campus'].astype(str).str.strip(),
        FSW    = all_rows['FSW'].astype(str).str.strip(),
    )
    return all_rows.sort_values(['Area','Campus','FSW'])

def build_wide_sheet(
    wb,
    sheet_name,
    month_value,
    roster_area_df,
    fsw_month_area_df,
    metrics_order,
    fill_missing_zero=False,  # affects FSW metric columns only
):
    ws = wb.create_sheet(sheet_name)

    # Interleaved header list: [Metric, Department - Metric] × N
    interleaved = []
    for m in metrics_order:
        interleaved += [m, f'Department - {m}']

    headers = ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']
    ws.append(headers)

    # Pivot FSW values for this area/month → one row per FSW, metric columns
    if not fsw_month_area_df.empty:
        pv = fsw_month_area_df.pivot_table(index=['Area','Campus','FSW'],
                                           columns='Metric', values='FSW_Value', aggfunc='first')
        # Ensure all target metrics exist
        for m in metrics_order:
            if m not in pv.columns:
                pv[m] = np.nan
        pv = pv[metrics_order].reset_index()
    else:
        pv = pd.DataFrame(columns=['Area','Campus','FSW'] + metrics_order)

    # Base rows: everyone in the area (from union of roster & master)
    base = roster_area_df.copy() if (roster_area_df is not None and not roster_area_df.empty) else pd.DataFrame(columns=['Area','Campus','FSW'])
    if not pv.empty:
        base = base.merge(pv, on=['Area','Campus','FSW'], how='left')
    else:
        # still create empty FSW metric columns so the Department columns appear
        for m in metrics_order:
            base[m] = np.nan

    # Optionally zero-fill missing FSW metric values
    if fill_missing_zero and metrics_order:
        present = [m for m in metrics_order if m in base.columns]
        if present:
            base[present] = base[present].fillna(0)

    # Department entry columns (blank by design)
    for m in metrics_order:
        base[f'Department - {m}'] = ''

    # Trailing admin columns
    base['Validated'] = ''
    base['Validation_Date'] = ''
    base['Issues'] = ''
    base['Services'] = ''
    base['Referrals'] = ''
    base['Notes'] = ''

    # Add Month column
    base.insert(0, 'Month', month_value)

    # Ensure all required header columns exist before ordering
    for col in ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']:
        if col not in base.columns:
            base[col] = []

    # Reorder
    ordered = ['Month','Area','Campus','FSW'] + interleaved + ['Validated','Validation_Date','Issues','Services','Referrals','Notes']
    base = base.loc[:, ordered]

    # Write rows (if any)
    if not base.empty:
        for r in dataframe_to_rows(base, index=False, header=False):
            ws.append(r)

    # Styles & filters
    add_styles_filters(ws)

    # Conditional formatting pairs (Dept vs FSW)
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

    # Union of FSWs from master + roster (prevents empty sheets)
    roster_all = ensure_roster_union(master_month_df, roster_df)

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

# ---------- INPUTS ----------
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

# Standardize FSW master headers
fsw = normalize_master_columns(fsw_raw)

need_cols = {'Month','Area','Campus','FSW','Metric','FSW_Value'}
missing = need_cols - set(fsw.columns)
if missing:
    st.error(f"FSW_Master is missing columns: {sorted(missing)}")
    st.stop()

if not {'Department','Metric'}.issubset(mmap.columns):
    st.error('metrics_map.csv must have columns: Department, Metric')
    st.stop()

# Clean values
for c in ['Month','Area','Campus','FSW','Metric']:
    fsw[c] = fsw[c].astype(str).str.strip()
mmap['Department'] = mmap['Department'].astype(str).str.strip()
mmap['Metric']     = mmap['Metric'].astype(str).str.strip()

# Attach Department to master (left join keeps all master rows)
master = fsw.merge(mmap[['Metric','Department']], on='Metric', how='left')

# ---------- Debug panel ----------
with st.expander("Debug / Data Health", expanded=False):
    st.write("Rows in FSW_Master:", len(master))
    st.write("Unique Months:", sorted(master['Month'].dropna().unique().tolist()))
    st.write("Unique Campuses:", len(master['Campus'].dropna().unique()))
    st.write("Unique FSWs:", master['FSW'].nunique())
    # How many metrics map to a department
    mm_counts = mmap.groupby('Department')['Metric'].count().rename('Metric_Count').reset_index()
    st.dataframe(mm_counts)

# ---------- Filters ----------
st.subheader('Filters & Options')
c1, c2, c3 = st.columns(3)

months_all = [m for m in ['Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May']
              if m in master['Month'].dropna().unique().tolist()]
depts_all  = sorted(mmap['Department'].dropna().unique().tolist())
camp_all   = sorted(master['Campus'].dropna().unique().tolist())

months = c1.multiselect('Months', months_all, default=months_all)
depts  = c2.multiselect('Departments', depts_all, default=depts_all)
camp   = c3.multiselect('Campuses', camp_all, default=camp_all)

fill_zero = st.checkbox('Fill missing FSW metric values with 0 (Department columns stay blank)', value=False)

mf = master.copy()
if months: mf = mf[mf['Month'].isin(months)]
if camp:   mf = mf[mf['Campus'].isin(camp)]

# ---------- Build ZIP ----------
if st.button('Build ZIP of Department Packets (Excel)'):
    if mf.empty:
        st.warning('No rows after filters. Check Months/Campuses.')
    else:
        memzip = BytesIO()
        with ZipFile(memzip, 'w', ZIP_DEFLATED) as zf:
            for month in sorted(mf['Month'].dropna().unique().tolist()):
                m_month = mf[mf['Month'] == month].copy()
                # If user selected departments, use those; otherwise all mapped depts
                target_depts = depts if depts else sorted(mmap['Department'].dropna().unique())
                for dept in target_depts:
                    # Metrics for this department (order preserved from CSV)
                    metrics_order = mmap.loc[mmap['Department'] == dept, 'Metric'].tolist()

                    if len(metrics_order) == 0:
                        # Still make an empty template with Month/Area/Campus/FSW + trailing admin cols
                        xbytes = build_workbook_bytes(
                            master_month_df=m_month[['Month','Area','Campus','FSW','Metric','FSW_Value']].copy(),
                            month_value=month,
                            metrics_order=[],
                            roster_df=roster_df,
                            fill_missing_zero=fill_zero,
                        )
                        safe_dept = "".join(ch for ch in dept if ch.isalnum() or ch in (" ","-","_")).strip().replace(" ","_") or "Unmapped"
                        zf.writestr(f"{month}/{safe_dept}_DEPT.xlsx", xbytes)
                        continue

                    # Use only rows that belong to this department's metrics
                    month_for_metrics = m_month[m_month['Metric'].isin(metrics_order)][
                        ['Month','Area','Campus','FSW','Metric','FSW_Value']
                    ].copy()

                    # If no FSW rows match (e.g., master is missing this dept’s metrics this month),
                    # still create columns for those metrics so teams have places to enter numbers.
                    if month_for_metrics.empty:
                        # fabricate an empty frame that still has the required columns
                        month_for_metrics = m_month[['Month','Area','Campus','FSW']].drop_duplicates().copy()
                        # add empty metric/value columns so pivot formatting works
                        month_for_metrics['Metric'] = pd.Series(dtype=str)
                        month_for_metrics['FSW_Value'] = pd.Series(dtype=float)

                    xbytes = build_workbook_bytes(
                        master_month_df=month_for_metrics,
                        month_value=month,
                        metrics_order=metrics_order,
                        roster_df=roster_df,
                        fill_missing_zero=fill_zero,
                    )

                    safe_dept = "".join(ch for ch in dept if ch.isalnum() or ch in (" ","-","_")).strip().replace(" ","_")
                    zf.writestr(f"{month}/{safe_dept}_DEPT.xlsx", xbytes)

        memzip.seek(0)
        stamp = datetime.now().strftime('%Y%m%d_%H%M')
        st.download_button('Download Department Packets ZIP', data=memzip,
                           file_name=f'DeptPackets_{stamp}.zip', mime='application/zip')
