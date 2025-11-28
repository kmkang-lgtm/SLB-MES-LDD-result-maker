import pandas as pd
import numpy as np
import openpyxl
import re
import os

HOUR_ORDER  = list(range(8, 24)) + [0] + list(range(1, 8))
HOUR_LABELS = list(range(8, 24)) + [24] + list(range(1, 8))

HEADER_ROW = 3

LANE_SHEETS = {
    "1Lane": ["1Lane_frt", "1Lane_rr", "1Lane_frt side", "1Lane_rr side"],
    "2Lane": ["2Lane_frt", "2Lane_rr", "2Lane_frt side", "2Lane_rr side"],
}

def detect_dtype(item_name: str) -> str:
    if "KHD" in item_name: return "KHD"
    if "WPH" in item_name: return "WPH"
    return "UNKNOWN"

def item_to_sheetname(item_name: str) -> str:
    name = item_name.replace("버스바 ", "")
    name = name.replace(" KHD AVG ", "-").replace(" WPH AVG ", "-")
    name = name.replace("FRT Side", "FS 1").replace("RR Side", "RS 1")
    return name

def safe_to_datetime(col):
    dt = pd.to_datetime(col, errors="coerce")
    if dt.isna().mean() > 0.5:
        num = pd.to_numeric(col, errors="coerce")
        dt2 = pd.to_datetime(num, unit="d", origin="1899-12-30", errors="coerce")
        if dt2.notna().sum() > dt.notna().sum():
            dt = dt2
    return dt

def load_lane_raw(xl: pd.ExcelFile, sheets):
    dfs = []
    for sh in sheets:
        df = xl.parse(sh)
        dfs.append(df)

    raw = pd.concat(dfs, ignore_index=True)
    raw["항목명"] = raw["항목명"].astype(str)
    raw["dtype"] = raw["항목명"].apply(detect_dtype)

    raw["측정일시"] = safe_to_datetime(raw["측정일시"])
    raw = raw.dropna(subset=["측정일시"])

    raw["hour"] = raw["측정일시"].dt.hour
    raw["val"] = pd.to_numeric(raw["측정값"], errors="coerce")
    return raw

def build_dynamic_hours(df_lane: pd.DataFrame, selected_hours=None):
    hours = sorted(df_lane["hour"].dropna().unique().tolist())
    if not hours:
        hours = HOUR_ORDER[:]

    # ✅ 유저가 선택한 시간만 남기기
    if selected_hours:
        hours = [h for h in hours if h in selected_hours]

    hours = hours[:len(HOUR_ORDER)]  # 템플릿 컬럼 초과 방지
    labels = [24 if h == 0 else h for h in hours]
    return hours, labels

def compute_hour_lists(df_item: pd.DataFrame, hour_order):
    hour_lists = {
        h: df_item.loc[df_item["hour"] == h, "val"].dropna().tolist()
        for h in hour_order
    }
    mins = [min(hour_lists[h]) if hour_lists[h] else 0 for h in hour_order]
    maxs = [max(hour_lists[h]) if hour_lists[h] else 0 for h in hour_order]
    avgs = [
        (sum(hour_lists[h]) / len(hour_lists[h])) if hour_lists[h] else np.nan
        for h in hour_order
    ]
    return hour_lists, mins, maxs, avgs

def update_summary_lane_title_from_template(out_wb, template_wb, lane_key):
    tpl_b2 = template_wb["summary"]["B2"].value or ""
    lane_no = "1" if lane_key.startswith("1") else "2"
    new_b2 = re.sub(r"^[12]", lane_no, str(tpl_b2))
    out_wb["summary"]["B2"].value = new_b2

def fill_data_into_ws(
    ws, dtype, sheet_name,
    hour_order, hour_labels,
    hour_lists, mins, maxs, avgs,
    raw_start_row=7, raw_end_row=100, raw_start_col=2,
    header_row=HEADER_ROW
):
    max_cols = len(HOUR_ORDER)

    ws.cell(row=2, column=2).value = f"{dtype} {sheet_name}"

    for i, lab in enumerate(hour_labels):
        c = raw_start_col + i
        ws.cell(row=header_row, column=c).value = lab
    for i in range(len(hour_labels), max_cols):
        c = raw_start_col + i
        ws.cell(row=header_row, column=c).value = None

    for i in range(len(hour_order)):
        c = raw_start_col + i
        ws.cell(row=4, column=c).value = mins[i]
        ws.cell(row=5, column=c).value = maxs[i]
        ws.cell(row=6, column=c).value = avgs[i]

    for i in range(len(hour_order), max_cols):
        c = raw_start_col + i
        ws.cell(row=4, column=c).value = None
        ws.cell(row=5, column=c).value = None
        ws.cell(row=6, column=c).value = None

    for r in range(raw_start_row, raw_end_row + 1):
        for i in range(max_cols):
            c = raw_start_col + i
            ws.cell(row=r, column=c).value = None

    max_len = max((len(v) for v in hour_lists.values()), default=0)
    for row_i in range(max_len):
        r = raw_start_row + row_i
        if r > raw_end_row:
            break
        for i, h in enumerate(hour_order):
            c = raw_start_col + i
            vals = hour_lists[h]
            ws.cell(row=r, column=c).value = vals[row_i] if row_i < len(vals) else None

def make_results_for_input(
    input_path: str,
    templates: dict,
    output_dir: str,
    raw_end_row: int = 100,
    selected_hours=None  # ✅ 추가
):
    os.makedirs(output_dir, exist_ok=True)
    created = []

    with pd.ExcelFile(input_path) as xl:
        for lane_key, sheets in LANE_SHEETS.items():
            raw_lane = load_lane_raw(xl, sheets)

            hour_order, hour_labels = build_dynamic_hours(raw_lane, selected_hours)

            for dtype, df_dtype in raw_lane.groupby("dtype"):
                if dtype == "UNKNOWN":
                    continue
                if dtype not in templates:
                    raise KeyError(f"템플릿이 없습니다: {dtype}")

                template_path = templates[dtype]
                template_wb = openpyxl.load_workbook(template_path)
                out_wb = openpyxl.load_workbook(template_path)

                update_summary_lane_title_from_template(out_wb, template_wb, lane_key)

                for item_name, df_item in df_dtype.groupby("항목명"):
                    sheet_name = item_to_sheetname(item_name)
                    if sheet_name not in out_wb.sheetnames:
                        raise KeyError(f"[{dtype}] 템플릿에 시트가 없음: {sheet_name}")

                    ws = out_wb[sheet_name]
                    hour_lists, mins, maxs, avgs = compute_hour_lists(df_item, hour_order)

                    fill_data_into_ws(
                        ws, dtype, sheet_name,
                        hour_order, hour_labels,
                        hour_lists, mins, maxs, avgs,
                        raw_end_row=raw_end_row
                    )

                out_path = os.path.join(output_dir, f"SLB_MES_{dtype}_Result_{lane_key}.xlsx")
                out_wb.save(out_path)

                out_wb.close()
                template_wb.close()
                created.append(out_path)

    return created
