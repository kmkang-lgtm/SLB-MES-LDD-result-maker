import pandas as pd
import numpy as np
import zipfile
import tempfile
import re
import time
import shutil
import gc
from pathlib import Path


# ---------------------------
# 0) zip 파일명에서 날짜 추출
# ---------------------------
def extract_date_from_zipname(zip_filename: str):
    name = Path(zip_filename).stem
    m = re.search(r"SLB_MES_Result_Package_(\d{1,2}\.\d{1,2})", name)
    if not m:
        return name.split("_")[-1]
    return m.group(1)


# ---------------------------
# 1) 원본 파일 파싱 규칙
# ---------------------------
def parse_one_sheet(file_path, sheet_name):
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    positions = raw.iloc[0].iloc[1:].tolist()
    data = raw.iloc[4:, 1:].copy()
    data.columns = positions
    data = data.apply(pd.to_numeric, errors="coerce")
    return data


# ---------------------------
# 2) 시트=포지션 flatten 통계 + Material 자동룰
# ---------------------------
def sheet_as_position_stats(file_path):
    with pd.ExcelFile(file_path) as xls:
        sheet_names = [s for s in xls.sheet_names if s.lower() != "summary"]

    prefix_has_2 = {}
    for s in sheet_names:
        if not isinstance(s, str) or "-" not in s:
            continue
        prefix, suffix = s.rsplit("-", 1)
        prefix = prefix.strip()
        suffix = suffix.strip()
        prefix_has_2.setdefault(prefix, False)
        if suffix == "2":
            prefix_has_2[prefix] = True

    rows = []
    for s in sheet_names:
        d = parse_one_sheet(file_path, s)
        vals = d.to_numpy().ravel()
        vals = vals[~np.isnan(vals)]
        if len(vals) == 0:
            continue

        mean = float(vals.mean())
        std  = float(vals.std(ddof=1))
        group = s.split()[0] if isinstance(s, str) else ""

        material = ""
        if isinstance(s, str) and "-" in s:
            prefix, _ = s.rsplit("-", 1)
            prefix = prefix.strip()
            material = "AL" if prefix_has_2.get(prefix, False) else "CU"

        rows.append({
            "position": s,
            "group": group,
            "Material": material,
            "count": int(len(vals)),
            "mean": mean,
            "std": std,
            "min": float(vals.min()),
            "max": float(vals.max()),
            "range": float(vals.max() - vals.min()),
        })

    return pd.DataFrame(rows)


# ---------------------------
# 3) Lane 블록(= WPH | KHD)
# ---------------------------
def make_lane_block(wph_df, khd_df):
    cols = ["position","group","Material","count","mean","std","min","max","range"]

    wph_block = wph_df[cols].copy()
    khd_block = khd_df[cols].copy()

    max_len = max(len(wph_block), len(khd_block))
    wph_block = wph_block.reindex(range(max_len))
    khd_block = khd_block.reindex(range(max_len))

    block = pd.concat([wph_block, khd_block], axis=1)
    return block


# ---------------------------
# 4) 최종 Summary 생성
# ---------------------------
def build_final_summary_single_sheet(
    wph_1lane_path, wph_2lane_path,
    khd_1lane_path, khd_2lane_path,
    output_path,
    sheet_name
):
    wph1 = sheet_as_position_stats(wph_1lane_path)
    wph2 = sheet_as_position_stats(wph_2lane_path)
    khd1 = sheet_as_position_stats(khd_1lane_path)
    khd2 = sheet_as_position_stats(khd_2lane_path)

    block_1lane = make_lane_block(wph1, khd1)
    block_2lane = make_lane_block(wph2, khd2)

    max_len = max(len(block_1lane), len(block_2lane))
    block_1lane = block_1lane.reindex(range(max_len))
    block_2lane = block_2lane.reindex(range(max_len))

    lane1_col = pd.DataFrame({"Lane": [1] * max_len})
    lane2_col = pd.DataFrame({"Lane": [2] * max_len})
    gap_cols  = pd.DataFrame({" ": [""] * max_len})

    final_df = pd.concat([lane1_col, block_1lane, gap_cols, lane2_col, block_2lane], axis=1)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=2)

        wb = writer.book
        ws = writer.sheets[sheet_name]

        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter

        title_font  = Font(bold=True, size=12)
        header_font = Font(bold=True)
        center = Alignment(horizontal="center", vertical="center")

        fill_wph = PatternFill("solid", fgColor="EAF2FF")
        fill_khd = PatternFill("solid", fgColor="F3F3F3")

        thick = Side(style="medium", color="000000")
        boundary_border = Border(right=thick)

        max_row = ws.max_row
        max_col = ws.max_column

        lane_label_w = 1
        wph_w = 9
        khd_w = 9
        lane_block_w = wph_w + khd_w
        gap_w = 1

        col_lane1 = 1
        col_wph1  = col_lane1 + lane_label_w
        col_khd1  = col_wph1 + wph_w

        col_lane2 = col_lane1 + lane_label_w + lane_block_w + gap_w
        col_wph2  = col_lane2 + lane_label_w
        col_khd2  = col_wph2 + wph_w

        for col, txt in [(col_wph1,"WPH"),(col_khd1,"KHD"),(col_wph2,"WPH"),(col_khd2,"KHD")]:
            ws.cell(row=2, column=col).value = txt
            ws.cell(row=2, column=col).font = title_font
            ws.cell(row=2, column=col).alignment = center

        for c in range(1, max_col + 1):
            cell = ws.cell(row=3, column=c)
            cell.font = header_font
            cell.alignment = center
            cell.fill = PatternFill("solid", fgColor="FFFFFF")

        for r in range(4, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.alignment = center

                if col_wph1 <= c <= col_wph1 + wph_w - 1:
                    cell.fill = fill_wph
                if col_khd1 <= c <= col_khd1 + khd_w - 1:
                    cell.fill = fill_khd

                if col_wph2 <= c <= col_wph2 + wph_w - 1:
                    cell.fill = fill_wph
                if col_khd2 <= c <= col_khd2 + khd_w - 1:
                    cell.fill = fill_khd

                if c in [col_lane1, col_lane2]:
                    cell.number_format = "0"

                int_cols = [
                    col_wph1+3, col_wph1+6, col_wph1+7, col_wph1+8,
                    col_khd1+3, col_khd1+6, col_khd1+7, col_khd1+8,
                    col_wph2+3, col_wph2+6, col_wph2+7, col_wph2+8,
                    col_khd2+3, col_khd2+6, col_khd2+7, col_khd2+8
                ]
                if c in int_cols:
                    cell.number_format = "0"

                float_cols = [
                    col_wph1+4, col_wph1+5,
                    col_khd1+4, col_khd1+5,
                    col_wph2+4, col_wph2+5,
                    col_khd2+4, col_khd2+5
                ]
                if c in float_cols:
                    cell.number_format = "0.000"

                if c == col_wph1 + wph_w - 1 or c == col_wph2 + wph_w - 1:
                    cell.border = boundary_border

        thick_side = Side(style="medium", color="000000")

        def apply_outer_border(ws, r1, r2, c1, c2):
            from openpyxl.styles import Border
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    cell = ws.cell(row=r, column=c)
                    b = cell.border
                    cell.border = Border(
                        left   = thick_side if c == c1 else b.left,
                        right  = thick_side if c == c2 else b.right,
                        top    = thick_side if r == r1 else b.top,
                        bottom = thick_side if r == r2 else b.bottom
                    )

        apply_outer_border(ws, 4, max_row, col_lane1, col_khd1 + khd_w - 1)
        apply_outer_border(ws, 4, max_row, col_lane2, col_khd2 + khd_w - 1)

        from openpyxl.utils import get_column_letter
        for c in range(1, max_col + 1):
            ws.column_dimensions[get_column_letter(c)].width = 11
        ws.column_dimensions[get_column_letter(col_lane1)].width = 6
        ws.column_dimensions[get_column_letter(col_lane2)].width = 6

        ws.freeze_panes = "A4"

    gc.collect()


# ---------------------------
# 5) ✅ Streamlit용 엔트리: zip bytes로 Summary 생성
# ---------------------------
def build_from_zip_bytes(zip_bytes: bytes, zip_filename: str):
    date_str = extract_date_from_zipname(zip_filename)
    sheet_name = date_str

    output_name = f"SLB_SV_LDD_Deviation_Summary_{date_str}.xlsx"

    tmpdir = tempfile.mkdtemp()
    try:
        # zip을 임시파일로 저장 후 extract
        zip_path = Path(tmpdir) / zip_filename
        with open(zip_path, "wb") as f:
            f.write(zip_bytes)

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(tmpdir)

        files = list(Path(tmpdir).glob("*.xlsx"))

        def pick(k1, k2):
            for f in files:
                name = f.name.lower()
                if k1.lower() in name and k2.lower() in name:
                    return str(f)
            return None

        wph_1 = pick("wph", "1lane")
        wph_2 = pick("wph", "2lane")
        khd_1 = pick("khd", "1lane")
        khd_2 = pick("khd", "2lane")

        missing = [k for k, v in {
            "WPH 1Lane": wph_1, "WPH 2Lane": wph_2,
            "KHD 1Lane": khd_1, "KHD 2Lane": khd_2
        }.items() if v is None]

        if missing:
            raise FileNotFoundError(
                f"zip 안에서 파일을 못 찾음: {missing}\n"
                f"파일명에 WPH/KHD, 1Lane/2Lane 키워드가 포함되어 있는지 확인해줘."
            )

        out_path = str(Path(tmpdir) / output_name)
        build_final_summary_single_sheet(
            wph_1, wph_2, khd_1, khd_2,
            output_path=out_path,
            sheet_name=sheet_name
        )

        with open(out_path, "rb") as f:
            summary_bytes = f.read()

        return output_name, summary_bytes

    finally:
        gc.collect()
        for _ in range(5):
            try:
                shutil.rmtree(tmpdir)
                break
            except PermissionError:
                time.sleep(0.5)
