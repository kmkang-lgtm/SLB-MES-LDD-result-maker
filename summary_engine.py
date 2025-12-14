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
    name = Path(zip_filename).stem

    # 1) YY.MM.DD
    m = re.search(r"(\d{2}\.\d{2}\.\d{2})", name)
    if m:
        return m.group(1)

    # 2) YYYY.MM.DD -> YY.MM.DD로 축약
    m = re.search(r"(\d{4})\.(\d{2})\.(\d{2})", name)
    if m:
        yy = m.group(1)[-2:]
        return f"{yy}.{m.group(2)}.{m.group(3)}"

    # 3) 구버전: MM.DD만 있으면 올해 YY를 붙임
    m = re.search(r"(\d{1,2}\.\d{1,2})", name)
    if m:
        mmdd = m.group(1)
        yy = datetime.now().strftime("%y")
        # 1자리 월/일이면 0 padding
        mm, dd = mmdd.split(".")
        mm = mm.zfill(2)
        dd = dd.zfill(2)
        return f"{yy}.{mm}.{dd}"

    # 4) fallback: 오늘 날짜
    return datetime.now().strftime("%y.%m.%d")


# ---------------------------
# 1) Result 파일 파싱
# - Summary 시트 제외
# - Raw 영역 (7행~, 2열~) 수집
# - Material: 같은 prefix에 -2가 존재하면 AL 그룹, suffix로 AL1/AL2
# ---------------------------
def parse_result_xlsx(file_path: str) -> pd.DataFrame:
    import openpyxl

    wb = openpyxl.load_workbook(file_path, data_only=True)

    # 1) prefix 중 "-2"로 끝나는 position이 있는지(=AL 그룹) 수집
    prefixes_with_2 = set()
    for ws in wb.worksheets:
        if ws.title.lower() == "summary":
            continue
        name = str(ws.title).strip()
        m = re.match(r"^(.*)-(\d+)$", name)
        if m and m.group(2) == "2":
            prefixes_with_2.add(m.group(1).strip())
        else:
            # 호환: 어딘가에 "-2"가 포함되어 있으면(예전 규칙) prefix 후보로 취급
            if "-2" in name and "-" in name:
                prefixes_with_2.add(name.rsplit("-", 1)[0].strip())

    rows = []
    for ws in wb.worksheets:
        if ws.title.lower() == "summary":
            continue

        position = str(ws.title).strip()

        # Raw 영역 값 수집
        vals = []
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

        # Material 결정
        material = "CU"
        m = re.match(r"^(.*)-(\d+)$", position)
        if m:
            prefix = m.group(1).strip()
            suffix = m.group(2).strip()
            if prefix in prefixes_with_2:
                if suffix == "1":
                    material = "AL1"
                elif suffix == "2":
                    material = "AL2"
                else:
                    material = "AL"  # 예외 suffix
            else:
                material = "CU"
        else:
            # 호환: "-2" 포함이면 AL2로 표기(기존 규칙 유지)
            if "-2" in position:
                material = "AL2"

        rows.append(
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

    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["Position", "Material", "Mean", "Std", "Min", "Max", "Range"])

    # Material 정렬순서: AL1 -> AL2 -> CU -> 기타
    order = {"AL1": 0, "AL2": 1, "CU": 2, "AL": 3}
    df["_mord"] = df["Material"].map(lambda x: order.get(str(x), 99))
    df = df.sort_values(["_mord", "Position"]).drop(columns=["_mord"]).reset_index(drop=True)
    return df


# ---------------------------
# 2) Lane 결합: WPH/KHD outer merge + Lane 컬럼 안정 처리
# ---------------------------
def combine_lane_df(wph_df: pd.DataFrame, khd_df: pd.DataFrame, lane_label: str) -> pd.DataFrame:
    merged = pd.merge(
        wph_df, khd_df, on=["Position", "Material"], how="outer", suffixes=("_WPH", "_KHD")
    )

    cols = ["Position", "Material"]
    for k in ["Mean", "Std", "Min", "Max", "Range"]:
        cols += [f"{k}_WPH", f"{k}_KHD"]

    for c in cols:
        if c not in merged.columns:
            merged[c] = np.nan

    merged = merged[cols]

    # ✅ FIX: Lane 컬럼이 이미 있으면 insert 대신 덮어쓰기 + A열로 이동
    if "Lane" in merged.columns:
        merged["Lane"] = lane_label
        merged = merged[["Lane"] + [c for c in merged.columns if c != "Lane"]]
    else:
        merged.insert(0, "Lane", lane_label)

    return merged


# ---------------------------
# 3) 엑셀 저장(양식 반영)
# - 1행: Deviation Summary (YY.MM.DD)
# - 2행: A2=Summary, B2=YY.MM.DD
# - Lane 타이틀 행(1Lane/2Lane)은 만들지 않음
# - Lane2 표는 한 칸 위로 당김(별도 타이틀 행 없으므로)
# ---------------------------
def write_summary_excel(output_path: str, lane1_df: pd.DataFrame, lane2_df: pd.DataFrame, date_str: str, sheet_name="Summary"):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    import openpyxl

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Lane1 표: startrow=2 -> 헤더는 3행, 데이터는 4행부터
        lane1_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=2)

        wb = writer.book
        ws = writer.sheets[sheet_name]

        # 스타일 정의
        title_font = Font(bold=True, size=14, color="FFFFFF")
        meta_font = Font(bold=True, size=11)
        header_font = Font(bold=True, size=11, color="FFFFFF")
        center = Alignment(horizontal="center", vertical="center")
        left = Alignment(horizontal="left", vertical="center")

        fill_title = PatternFill("solid", fgColor="1F4E79")
        fill_header = PatternFill("solid", fgColor="2F5597")
        fill_meta = PatternFill("solid", fgColor="D9E1F2")
        fill_alt = PatternFill("solid", fgColor="F7F7F7")

        thin = Side(style="thin", color="BFBFBF")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        max_col = lane1_df.shape[1]

        # 1행 타이틀
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        tcell = ws.cell(row=1, column=1, value=f"Deviation Summary ({date_str})")
        tcell.font = title_font
        tcell.fill = fill_title
        tcell.alignment = center

        # 2행 메타: Summary | YY.MM.DD (나머지는 병합)
        ws.cell(row=2, column=1, value="Summary").font = meta_font
        ws.cell(row=2, column=2, value=date_str).font = meta_font
        for c in range(1, max_col + 1):
            cell = ws.cell(row=2, column=c)
            cell.fill = fill_meta
            cell.border = border
            cell.alignment = center
        if max_col >= 3:
            ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=max_col)

        # Lane1 헤더(3행)
        header_row = 3
        for c in range(1, max_col + 1):
            cell = ws.cell(row=header_row, column=c)
            cell.font = header_font
            cell.fill = fill_header
            cell.alignment = center
            cell.border = border

        # Lane1 데이터 영역
        start_data_row_1 = header_row + 1
        end_data_row_1 = start_data_row_1 + len(lane1_df) - 1
        for r in range(start_data_row_1, end_data_row_1 + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                cell.alignment = left if c in (1, 2, 3) else center
                if (r - start_data_row_1) % 2 == 1:
                    cell.fill = fill_alt

        # Lane2 표: 타이틀 행 없이 한 줄 위로 당김
        gap = 3
        startrow2 = end_data_row_1 + gap  # ✅ +1 제거 (타이틀 행이 없으므로)
        lane2_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=startrow2)

        header_row2 = startrow2 + 1  # pandas header가 startrow2+1
        for c in range(1, max_col + 1):
            cell = ws.cell(row=header_row2, column=c)
            cell.font = header_font
            cell.fill = fill_header
            cell.alignment = center
            cell.border = border

        start_data_row_2 = header_row2 + 1
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
            from openpyxl.utils import get_column_letter
            ws.column_dimensions[get_column_letter(c)].width = w

        # 숫자 포맷
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
            # fallback: 숫자 포함
            lane_num = "1" if lane == "1lane" else "2"
            for f in files:
                name = f.name.lower().replace(" ", "")
                if dtype in name and lane_num in name:
                    return str(f)
            return None

        wph_1 = pick("wph", "1lane")
        wph_2 = pick("wph", "2lane")
        khd_1 = pick("khd", "1lane")
        khd_2 = pick("khd", "2lane")

        missing = [k for k, v in {"wph_1": wph_1, "wph_2": wph_2, "khd_1": khd_1, "khd_2": khd_2}.items() if v is None]
        if missing:
            raise Exception(
                f"Result 파일이 부족합니다: {missing}\n"
                f"ZIP 안 파일명에 WPH/KHD, 1Lane/2Lane 키워드가 포함되어 있는지 확인해줘."
            )

        df_wph_1 = parse_result_xlsx(wph_1)
        df_wph_2 = parse_result_xlsx(wph_2)
        df_khd_1 = parse_result_xlsx(khd_1)
        df_khd_2 = parse_result_xlsx(khd_2)

        lane1_df = combine_lane_df(df_wph_1, df_khd_1, "Lane1")
        lane2_df = combine_lane_df(df_wph_2, df_khd_2, "Lane2")

        out_name = f"SLB_MES_Deviation_Summary_{date_str}.xlsx"
        out_path = Path(tmpdir) / out_name

        write_summary_excel(str(out_path), lane1_df, lane2_df, date_str=date_str, sheet_name="Summary")

        out_bytes = out_path.read_bytes()
        return out_name, out_bytes

    finally:
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass
        gc.collect()
