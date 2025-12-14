import pandas as pd
import numpy as np
import zipfile
import tempfile
import re
import shutil
import gc
from pathlib import Path
from datetime import datetime


# ---------------------------
# 0) zip 파일명에서 날짜 추출 (YY.MM.DD 우선)
# ---------------------------
def extract_date_from_zipname(zip_filename: str) -> str:
    """
    기대하는 ZIP 이름 예:
      SLB_MES_Result_Package_25.11.18.zip  -> 25.11.18
    하위호환:
      SLB_MES_Result_Package_11.18.zip     -> 현재 연도(YY) 붙여 25.11.18
    """
    name = Path(zip_filename).stem

    m = re.search(r"SLB_MES_Result_Package_(\d{2}\.\d{2}\.\d{2})", name)
    if m:
        return m.group(1)

    m2 = re.search(r"SLB_MES_Result_Package_(\d{1,2})\.(\d{1,2})", name)
    if m2:
        yy = datetime.now().strftime("%y")
        mm = int(m2.group(1))
        dd = int(m2.group(2))
        return f"{yy}.{mm:02d}.{dd:02d}"

    # fallback: 마지막 토큰이라도 사용
    tail = name.split("_")[-1]
    # 만약 25.11 형태면 day는 오늘 날짜로 보강
    m3 = re.fullmatch(r"(\d{2})\.(\d{2})", tail)
    if m3:
        yy = m3.group(1)
        mm = m3.group(2)
        dd = datetime.now().strftime("%d")
        return f"{yy}.{mm}.{dd}"
    return tail


# ---------------------------
# 1) Result 엑셀에서 Position별 값 수집 → 통계
# ---------------------------
def parse_result_xlsx(file_path: str) -> pd.DataFrame:
    import openpyxl

    wb = openpyxl.load_workbook(file_path, data_only=True)

    # ✅ 1) prefix에 "-2"가 존재하는지 먼저 훑어서 AL 그룹 판정(안전하게 suffix 기준)
    prefixes_with_2 = set()
    for ws in wb.worksheets:
        if ws.title.lower() == "summary":
            continue
        name = str(ws.title).strip()
        m = re.match(r"^(.*)-2$", name)
        if m:
            prefixes_with_2.add(m.group(1).strip())

    records = []
    for ws in wb.worksheets:
        if ws.title.lower() == "summary":
            continue

        vals = []
        # Raw 영역: 7행~, 2열~ (기존 규칙 유지)
        for c in range(2, ws.max_column + 1):
            for r in range(7, ws.max_row + 1):
                v = ws.cell(row=r, column=c).value
                if v is None:
                    continue
                try:
                    vals.append(float(v))
                except Exception:
                    continue

        if not vals:
            continue

        arr = np.array(vals, dtype=float)
        position = str(ws.title).strip()

        # ✅ Material: 같은 prefix에 -2가 있으면 AL 그룹
        material = "CU"
        m1 = re.match(r"^(.*)-1$", position)
        m2 = re.match(r"^(.*)-2$", position)
        if m2:
            material = "AL2"
        elif m1 and m1.group(1).strip() in prefixes_with_2:
            material = "AL1"
        else:
            # 하위호환: 이름에 -2가 포함되기만 해도 AL2로(혹시 suffix가 깨진 케이스)
            if "-2" in position:
                material = "AL2"

        records.append(
            {
                "Position": position,
                "Material": material,
                "Mean": float(np.mean(arr)),
                "Std": float(np.std(arr, ddof=1)) if arr.size > 1 else 0.0,
                "Min": float(np.min(arr)),
                "Max": float(np.max(arr)),
                "Range": float(np.max(arr) - np.min(arr)),
            }
        )

    wb.close()

    df = pd.DataFrame(records, columns=["Position", "Material", "Mean", "Std", "Min", "Max", "Range"])
    if df.empty:
        return df

    # Material 정렬(AL1 → AL2 → CU)
    order = {"AL1": 0, "AL2": 1, "AL": 1, "CU": 2}
    df["_mord"] = df["Material"].map(order).fillna(9).astype(int)
    df = df.sort_values(["_mord", "Position"]).drop(columns=["_mord"]).reset_index(drop=True)
    return df


# ---------------------------
# 2) Lane별 WPH/KHD 결합
# ---------------------------
def combine_lane_df(wph_df: pd.DataFrame, khd_df: pd.DataFrame, lane_label: str) -> pd.DataFrame:
    merged = pd.merge(
        wph_df, khd_df, on=["Position", "Material"], how="outer", suffixes=("_WPH", "_KHD")
    )

    cols = ["Lane", "Position", "Material"]
    for k in ["Mean", "Std", "Min", "Max", "Range"]:
        cols += [f"{k}_WPH", f"{k}_KHD"]

    for c in cols:
        if c not in merged.columns:
            merged[c] = np.nan

    merged.insert(0, "Lane", lane_label)
    merged = merged[cols]
    return merged


# ---------------------------
# 3) 엑셀 저장(양식 유지 + 날짜 표기)
# ---------------------------
def write_summary_excel(output_path: str, lane1_df: pd.DataFrame, lane2_df: pd.DataFrame, date_str: str):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    import openpyxl

    sheet_name = "Summary"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Lane1 표: startrow=2 → header는 3행(row=3)
        lane1_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=2)

        wb = writer.book
        ws = writer.sheets[sheet_name]

        # ---------------------------
        # 스타일 정의
        # ---------------------------
        title_font = Font(bold=True, size=14)
        header_font = Font(bold=True, size=11, color="FFFFFF")
        center = Alignment(horizontal="center", vertical="center")
        left = Alignment(horizontal="left", vertical="center")

        fill_title = PatternFill("solid", fgColor="1F4E79")
        fill_header = PatternFill("solid", fgColor="2F5597")
        fill_alt = PatternFill("solid", fgColor="F7F7F7")
        fill_meta = PatternFill("solid", fgColor="D9E1F2")

        thin = Side(style="thin", color="BFBFBF")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        max_col = lane1_df.shape[1]

        # 1행: 타이틀(날짜 포함)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        tcell = ws.cell(row=1, column=1, value=f"Deviation Summary ({date_str})")
        tcell.font = title_font
        tcell.fill = fill_title
        tcell.alignment = center

        # 2행: Summary | YY.MM.DD
        ws.cell(row=2, column=1, value="Summary").fill = fill_meta
        ws.cell(row=2, column=1).font = Font(bold=True)
        ws.cell(row=2, column=1).alignment = center
        ws.cell(row=2, column=1).border = border

        ws.cell(row=2, column=2, value=date_str).fill = fill_meta
        ws.cell(row=2, column=2).font = Font(bold=True)
        ws.cell(row=2, column=2).alignment = center
        ws.cell(row=2, column=2).border = border

        if max_col >= 3:
            ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=max_col)
            mcell = ws.cell(row=2, column=3, value="")
            mcell.fill = fill_meta
            mcell.border = border

        # Lane1 헤더는 3행
        header_row_1 = 3
        for c in range(1, max_col + 1):
            cell = ws.cell(row=header_row_1, column=c)
            cell.font = header_font
            cell.fill = fill_header
            cell.alignment = center
            cell.border = border

        # Lane1 데이터 영역
        start_data_row_1 = header_row_1 + 1
        end_data_row_1 = start_data_row_1 + len(lane1_df) - 1

        for r in range(start_data_row_1, end_data_row_1 + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                cell.alignment = left if c in (1, 2, 3) else center
                if (r - start_data_row_1) % 2 == 1:
                    cell.fill = fill_alt

        # Lane2 표: title 행 없이 한 줄 위로 당김
        gap = 3
        startrow2 = end_data_row_1 + gap  # ✅ +1 제거 (한 줄 위로)
        lane2_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=startrow2)

        # Lane2 헤더는 startrow2+1
        header_row_2 = startrow2 + 1
        for c in range(1, max_col + 1):
            cell = ws.cell(row=header_row_2, column=c)
            cell.font = header_font
            cell.fill = fill_header
            cell.alignment = center
            cell.border = border

        start_data_row_2 = header_row_2 + 1
        end_data_row_2 = start_data_row_2 + len(lane2_df) - 1

        for r in range(start_data_row_2, end_data_row_2 + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                cell.alignment = left if c in (1, 2, 3) else center
                if (r - start_data_row_2) % 2 == 1:
                    cell.fill = fill_alt

        # 열 너비
        widths = {1: 10, 2: 20, 3: 10}
        for c in range(4, max_col + 1):
            widths[c] = 12
        for c, w in widths.items():
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

        # 숫자 포맷(소수점 3자리)
        for r in range(start_data_row_1, end_data_row_2 + 1):
            for c in range(4, max_col + 1):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell.value, (int, float, np.number)):
                    cell.number_format = "0.000"

        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A4"


# ---------------------------
# 4) ZIP bytes → Summary 생성
# ---------------------------
def build_from_zip_bytes(zip_bytes: bytes, zip_name: str):
    date_str = extract_date_from_zipname(zip_name)

    tmpdir = tempfile.mkdtemp(prefix="mes_summary_")
    try:
        zip_path = Path(tmpdir) / "input.zip"
        with open(zip_path, "wb") as f:
            f.write(zip_bytes)

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(tmpdir)

        files = list(Path(tmpdir).glob("*.xlsx"))

        def pick(dtype: str, lane: str):
            dtype = dtype.lower()
            lane = lane.lower()

            lane_aliases = {lane, lane.replace("lane", "_lane"), lane.replace("lane", "-lane")}
            if lane == "1lane":
                lane_aliases |= {"lane1", "1_lane", "1-lane", "lane_1", "lane-1"}
            if lane == "2lane":
                lane_aliases |= {"lane2", "2_lane", "2-lane", "lane_2", "lane-2"}

            for f in files:
                name = f.name.lower().replace(" ", "")
                if dtype in name:
                    for a in lane_aliases:
                        if a.replace(" ", "") in name:
                            return str(f)
            return None

        wph_1 = pick("wph", "1lane")
        wph_2 = pick("wph", "2lane")
        khd_1 = pick("khd", "1lane")
        khd_2 = pick("khd", "2lane")

        missing = [k for k, v in {
            "WPH 1Lane": wph_1, "WPH 2Lane": wph_2,
            "KHD 1Lane": khd_1, "KHD 2Lane": khd_2,
        }.items() if v is None]

        if missing:
            raise Exception(
                f"Result 파일이 부족합니다: {missing}\n"
                f"ZIP 안 파일명에 WPH/KHD, 1Lane/2Lane 포함 여부를 확인하세요."
            )

        df_wph_1 = parse_result_xlsx(wph_1)
        df_wph_2 = parse_result_xlsx(wph_2)
        df_khd_1 = parse_result_xlsx(khd_1)
        df_khd_2 = parse_result_xlsx(khd_2)

        lane1_df = combine_lane_df(df_wph_1, df_khd_1, "Lane1")
        lane2_df = combine_lane_df(df_wph_2, df_khd_2, "Lane2")

        out_name = f"SLB_MES_Deviation_Summary_{date_str}.xlsx"
        out_path = Path(tmpdir) / out_name

        write_summary_excel(str(out_path), lane1_df, lane2_df, date_str=date_str)

        with open(out_path, "rb") as f:
            out_bytes = f.read()

        return out_name, out_bytes

    finally:
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass
        gc.collect()
