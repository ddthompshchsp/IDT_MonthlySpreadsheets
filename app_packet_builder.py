
import argparse
from pathlib import Path
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

VALIDATED_LIST = ["Validated", "Mismatch", "Unable to Validate"]
READONLY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

AREA_SHEETS = ["Central", "West", "East"]

def norm(s):
    return s.replace("\u2013","-").replace("\u2014","-").strip() if isinstance(s,str) else s

def load_master(master_path: Path) -> pd.DataFrame:
    df = pd.read_excel(master_path, sheet_name="FSW_Master", engine="openpyxl")
    need = {"Month","Area","Campus","FSW","Metric","Value"}
    miss = need - set(df.columns)
    if miss:
        raise SystemExit(f"FSW_Master is missing columns: {sorted(miss)}")
    for c in ["Month","Area","Campus","FSW","Metric"]:
        if df[c].dtype == object:
            df[c] = df[c].map(norm)
    df = df.rename(columns={"Value":"FSW_Value"})
    return df

def load_metrics_map(path: Path) -> pd.DataFrame:
    mm = pd.read_csv(path)
    need = {"Department","Metric"}
    if not need.issubset(mm.columns):
        raise SystemExit("metrics_map.csv must have columns: Department, Metric")
    mm["Department"] = mm["Department"].map(norm)
    mm["Metric"] = mm["Metric"].map(norm)
    if "Display_Order" in mm.columns:
        mm = mm.sort_values(["Department","Display_Order","Metric"])
    else:
        mm = mm.sort_values(["Department","Metric"])
    return mm

def excel_col(n):
    s = ""
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65 + r) + s
    return s

def build_wide_table(df_dept: pd.DataFrame, metrics_order: list[str]) -> pd.DataFrame:
    id_cols = ["Area","Campus","FSW"]
    pv = df_dept.pivot_table(index=id_cols, columns="Metric", values="FSW_Value", aggfunc="first")
    for m in metrics_order:
        if m not in pv.columns:
            pv[m] = np.nan
    pv = pv[metrics_order]
    validated_cols = [f"{m} - Validated" for m in metrics_order]
    out = pv.reset_index()
    for col in validated_cols:
        out[col] = np.nan
    return out, validated_cols

def stylesheet(ws):
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")
    last_col = excel_col(ws.max_column)
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{last_col}{ws.max_row}"

def size_columns(ws):
    for col in range(1, ws.max_column+1):
        ws.column_dimensions[excel_col(col)].width = 16
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22

def add_dropdowns_and_colors(ws, metrics_order):
    num_metrics = len(metrics_order)
    start_fsw = 4
    start_val = 4 + num_metrics
    for i in range(num_metrics):
        fsw_col = excel_col(start_fsw + i)
        val_col = excel_col(start_val + i)
        for r in range(2, ws.max_row+1):
            d_cell = f"{val_col}{r}"
            c_cell = f"{fsw_col}{r}"
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
    for cell in ws["A"][1:]: cell.fill = READONLY_FILL
    for cell in ws["B"][1:]: cell.fill = READONLY_FILL
    for cell in ws["C"][1:]: cell.fill = READONLY_FILL

def write_summary(wb, sheet_map, metrics_order):
    ws = wb.create_sheet("Summary", 0)
    ws["A1"].value = "Dashboard / Summary"
    ws["A1"].font = Font(bold=True, size=14)

    ws.append([])
    ws.append(["Area","Metric","Total Rows","Validated Entered","Exact Matches","Mismatches"])
    for cell in ws[3]:
        cell.font = Font(bold=True)

    for area, (sname, last_row, fsw_start_col, val_start_col) in sheet_map.items():
        if last_row < 2:
            continue
        for i, m in enumerate(metrics_order):
            fsw_col = excel_col(fsw_start_col + i)
            val_col = excel_col(val_start_col + i)
            rng_fsw = f"'{sname}'!{fsw_col}2:{fsw_col}{last_row}"
            rng_val = f"'{sname}'!{val_col}2:{val_col}{last_row}"
            total_rows = f"=ROWS({rng_fsw})"
            validated_entered = f"=COUNTIF({rng_val},\"<>\")"
            exact_matches = f"=SUMPRODUCT(--({rng_val}={rng_fsw}),--({rng_val}<>\"\"))"
            mismatches = f"=SUMPRODUCT(--({rng_val}<>{rng_fsw}),--({rng_val}<>\"\"))"
            ws.append([area, m, total_rows, validated_entered, exact_matches, mismatches])

    last_col = excel_col(ws.max_column)
    ws.auto_filter.ref = f"A3:{last_col}{ws.max_row}"
    for col in range(1, ws.max_column+1):
        ws.column_dimensions[excel_col(col)].width = 22
    ws.freeze_panes = "A4"

def build_workbook(out_path: Path, month: str, dept: str, df: pd.DataFrame, metrics_order: list[str]):
    wb = Workbook()
    wb.remove(wb.active)

    sheet_meta = {}

    for area in ["Central","West","East"]:
        sub = df[df["Area"].astype(str).str.strip().str.casefold() == area.lower()]
        ws = wb.create_sheet(area)
        headers = ["Area","Campus","FSW"] + metrics_order + [f"{m} - Validated" for m in metrics_order]
        ws.append(headers)
        if not sub.empty:
            wide, validated_cols = build_wide_table(sub, metrics_order)
            for r in dataframe_to_rows(wide, index=False, header=False):
                ws.append(r)

        stylesheet(ws)
        size_columns(ws)
        add_dropdowns_and_colors(ws, metrics_order)

        sheet_meta[area] = (area, ws.max_row, 4, 4+len(metrics_order))

    write_summary(wb, sheet_meta, metrics_order)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--master", required=True, help="FSW_Master.xlsx path")
    ap.add_argument("--metrics_map", required=True, help="metrics_map.csv path")
    ap.add_argument("--out_dir", required=True, help="Output folder")
    args = ap.parse_args()

    master = load_master(Path(args.master))
    mmap = load_metrics_map(Path(args.metrics_map))

    master = master.merge(mmap[["Metric","Department"]], on="Metric", how="left")
    master = master.dropna(subset=["Department"])

    for (month, dept), chunk in master.groupby(["Month","Department"], dropna=False):
        metrics_order = mmap[mmap["Department"] == dept]["Metric"].tolist()
        out_path = Path(args.out_dir) / str(month) / f"{dept}_WIDE.xlsx"
        print(f"Writing {out_path} ...")
        build_workbook(out_path, month, dept, chunk, metrics_order)

    print(f"Done. Files in: {Path(args.out_dir).resolve()}")

if __name__ == "__main__":
    main()
