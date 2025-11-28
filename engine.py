import pandas as pd
import numpy as np
import openpyxl
import re
import os

# 시간 순서 기본(템플릿 최대 컬럼 수 기준)
HOUR_ORDER  = list(range(8, 24)) + [0] + list(range(1, 8))
HOUR_LABELS = list(range(8, 24)) + [24] + list(range(1, 8))

# 템플릿에서 시간 라벨이 적혀있는 행 (기본: B3부터)
HEADER_ROW = 3

# Lane별 원본 시트 구성
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
    """
    측정일시 컬럼에 문자열/빈값/엑셀시리얼 등이 섞여 있어도 안전하게 변환.
    - 1차: 일반 to_datetime(errors='coerce')
    - NaT가 너무 많으면(절반 이상) 2차: 엑셀 시리얼(origin)로 재시도
    """
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

def build_dynamic_hours(df_lane: pd.DataFrame):
    """
    Lane 데이터에서 실제 존재하는 hour만 추출해 사용.
    - 없으면 기본 HOUR_ORDER/HOUR_LABELS fallback
    - 0시는 라벨만 24로 표시
    - ✅ 템플릿 최대 컬럼 수(len(HOUR_ORDER)) 초과 방지
    """
    hours = sorted(df_lane["hour"].dropna().unique().tolist())
    if not hours:
        return HOUR_ORDER, HOUR_LABELS

    # ✅ 템플릿 최대 시간 컬럼 수(24개) 초과 시 잘라냄
    hours = hours[:len(HOUR_ORDER)]

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
    """
    템플릿 summary B2 스타일(1Line/1Lain/1Lane 등)을 유지하고
    Lane 번호만 lane_key에 맞게 교체
    """
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
    """
    템플릿 시트에 값만 덮어쓰기 (서식/그래프/도형 유지)
    + 실제 측정 시간(hour_labels)을 헤더 행에 동적으로 반영
    """
    max_cols = len(HOUR_ORDER)  # 템플릿이 가진 최대 시간 컬럼 수(고정)

    # 제목(B2)
    ws.cell(row=2, column=2).value = f"{dtype} {sheet_name}"

    # ✅ 시간 라벨(B3~) 덮어쓰기
    for i, lab in enumerate(hour_labels):
        c = raw_start_col + i
        ws.cell(row=header_row, column=c).value = lab

    # 남는 라벨 칸은 비우기
    for i in range(len(hour_labels), max_cols):
        c = raw_start_col + i
        ws.cell(row=header_row, column=c).value = None

    # MIN/MAX/AVG (4/5/6행) — 실제 hour 개수만큼만
    for i in range(len(hour_order)):
        c = raw_start_col + i
        ws.cell(row=4, column=c).value = mins[i]
        ws.cell(row=5, column=c).value = maxs[i]
        ws.cell(row=6, column=c).value = avgs[i]

    # 남는 MIN/MAX/AVG 칸 비우기
    for i in range(len(hour_order), max_cols):
        c = raw_start_col + i
        ws.cell(row=4, column=c).value = None
        ws.cell(row=5, column=c).value = None
        ws.cell(row=6, column=c).value = None

    # Raw 영역 전체 클리어(템플릿 최대 컬럼 기준으로)
    for r in range(raw_start_row, raw_end_row + 1):
        for i in range(max_cols):
            c = raw_start_col + i
            ws.cell(row=r, column=c).value = None

    # Raw 값 채우기 — 실제 hour_order 기준
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
    raw_end_row: int = 100
):
    """
    input_path: 원본(KHD/WPH) 엑셀
    templates: {"KHD": khd_template_path, "WPH": wph_template_path}
    output_dir: 결과 저장 폴더
    raw_end_row: 템플릿 차트 참조 끝행

    return: 생성된 결과 파일 경로 리스트
    """
    os.makedirs(output_dir, exist_ok=True)
    created = []

    with pd.ExcelFile(input_path) as xl:
        for lane_key, sheets in LANE_SHEETS.items():
            raw_lane = load_lane_raw(xl, sheets)

            # ✅ Lane 단위로 실제 시간 목록 생성
            hour_order, hour_labels = build_dynamic_hours(raw_lane)

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
