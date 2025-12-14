# summary_engine.py
# Result ZIP (KHD/WPH x Lane1/2) → Deviation Summary 엑셀 생성
# - errors.py(UserFacingError) 기반 사용자 친화 에러 적용
# - 입력: zip bytes
# - 출력: (summary_filename, summary_bytes)

from __future__ import annotations

import io
import os
import re
import zipfile
from datetime import datetime
from typing import Dict, Tuple

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

from errors import UserFacingError, invalid_zip_error


# ---------------------------------
# Helpers
# ---------------------------------
REQUIRED_KEYS = {
    ("WPH", "1Lane"),
    ("WPH", "2Lane"),
    ("KHD", "1Lane"),
    ("KHD", "2Lane"),
}


def _extract_mmdd_from_zipname(zip_name: str) -> str:
    m = re.search(r"_(\d{2}\.\d{2})", zip_name)
    if m:
        return m.group(1)
    return datetime.now().strftime("%m.%d")


def _parse_dtype_lane(filename: str) -> Tuple[str, str]:
    """
    SLB_MES_{dtype}_Result_{lane}.xlsx → (dtype, lane)
    """
    name = os.path.basename(filename)
    m = re.search(r"SLB_MES_(KHD|WPH)_Result_(1Lane|2Lane)\.xlsx", name)
    if not m:
        raise UserFacingError(
            title="Result 파일 이름 형식이 올바르지 않습니다.",
            detail=f"파일명: {name}",
            hint="Result ZIP은 app.py에서 생성된 패키지를 사용하세요.",
            code="INVALID_RESULT_FILENAME",
        )
    return m.group(1), m.group(2)


def _collect_result_files(zip_bytes: bytes, zip_name: str) -> Dict[Tuple[str, str], bytes]:
    try:
        zf = zipfile.ZipFile(io.BytesIO(zip_bytes), "r")
    except Exception as e:
        raise invalid_zip_error(
            detail=f"파일: {zip_name}\n오류: {e}",
            context={"zip": zip_name},
        )

    files: Dict[Tuple[str, str], bytes] = {}
    for n in zf.namelist():
        if not n.lower().endswith(".xlsx"):
            continue
        if n.startswith("__MACOSX/"):
            continue

        dtype, lane = _parse_dtype_lane(n)
        files[(dtype, lane)] = zf.read(n)

    missing = REQUIRED_KEYS - set(files.keys())
    if missing:
        raise UserFacingError(
            title="Deviation Summary 생성에 필요한 Result 파일이 부족합니다.",
            detail=f"필요: {sorted(REQUIRED_KEYS)}\n누락: {sorted(missing)}",
            hint="KHD/WPH, Lane1/Lane2 Result가 모두 생성되었는지 확인하세요.",
            code="MISSING_RESULT_FILES",
        )

    return files


def _extract_position_values(wb) -> Dict[str, list]:
    """
    Result 엑셀에서 summary 시트를 제외한 모든 시트를 읽어
    position별 값 리스트를 수집
    """
    pos_vals: Dict[str, list] = {}

    for ws in wb.worksheets:
        if ws.title.lower() == "summary":
            continue

        # Raw 영역은 7행부터, 2열(B)부터 시작 (engine.py 기준)
        for c in range(2, ws.max_column + 1):
            values = []
            for r in range(7, ws.max_row + 1):
                v = ws.cell(row=r, column=c).value
                if v is None:
                    continue
                try:
                    values.append(float(v))
                except Exception:
                    continue

            if values:
                key = ws.title.strip()
                pos_vals.setdefault(key, []).extend(values)

    return pos_vals


def _material_from_position(position: str) -> str:
    """
    summary_engine 기준 Material 규칙:
    - 같은 prefix에서 '-2'가 존재하면 AL
    - 아니면 CU
    """
    if "-2" in position:
        return "AL"
    return "CU"


def _build_block(
    files: Dict[Tuple[str, str], bytes],
    dtype: str,
    lane: str,
) -> pd.DataFrame:
    """
    하나의 블록(WPH/KHD x Lane)에 대해
    position별 통계(mean/std/min/max/range) 계산
    """
    wb = openpyxl.load_workbook(io.BytesIO(files[(dtype, lane)]), data_only=True)
    pos_vals = _extract_position_values(wb)
    wb.close()

    records = []
    for pos, vals in pos_vals.items():
        arr = np.array(vals, dtype=float)
        if arr.size == 0:
            continue

        records.append(
            {
                "Position": pos,
                "Material": _material_from_position(pos),
                "Mean": float(np.mean(arr)),
                "Std": float(np.std(arr, ddof=1)) if arr.size > 1 else 0.0,
                "Min": float(np.min(arr)),
                "Max": float(np.max(arr)),
                "Range": float(np.max(arr) - np.min(arr)),
            }
        )

    if not records:
        raise UserFacingError(
            title="Result 파일에서 유효한 데이터를 찾지 못했습니다.",
            detail=f"dtype={dtype}, lane={lane}",
            hint="Result 파일이 비어있거나 템플릿 Raw 영역이 변경됐는지 확인하세요.",
            code="NO_VALID_DATA",
        )

    df = pd.DataFrame(records).sort_values("Position").reset_index(drop=True)
    return df


# ---------------------------------
# Public API
# ---------------------------------
def build_from_zip_bytes(zip_bytes: bytes, zip_name: str) -> Tuple[str, bytes]:
    """
    zip_bytes: Result ZIP bytes
    zip_name: ZIP 파일명(날짜 추출용)
    return: (summary_filename, summary_bytes)
    """
    files = _collect_result_files(zip_bytes, zip_name)
    mmdd = _extract_mmdd_from_zipname(zip_name)

    # Lane1 / Lane2 블록 생성
    blocks = {}
    for dtype, lane in REQUIRED_KEYS:
        blocks[(dtype, lane)] = _build_block(files, dtype, lane)

    # -----------------------------
    # Summary 엑셀 작성
    # -----------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Summary_{mmdd}"

    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    col = 1
    for lane in ["1Lane", "2Lane"]:
        for dtype in ["WPH", "KHD"]:
            df = blocks[(dtype, lane)]

            # 타이틀
            ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 6)
            t = ws.cell(row=1, column=col)
            t.value = f"{dtype} {lane}"
            t.font = bold
            t.alignment = center

            # 헤더
            headers = list(df.columns)
            for j, h in enumerate(headers):
                c = ws.cell(row=2, column=col + j)
                c.value = h
                c.font = bold
                c.alignment = center
                c.border = border

            # 데이터
            for i, row in df.iterrows():
                for j, h in enumerate(headers):
                    c = ws.cell(row=3 + i, column=col + j)
                    c.value = row[h]
                    c.border = border

            col += len(headers) + 1  # 블록 간 공백

    out_name = f"SLB_MES_Deviation_Summary_{mmdd}.xlsx"
    buf = io.BytesIO()
    wb.save(buf)

    return out_name, buf.getvalue()
