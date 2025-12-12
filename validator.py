# validator.py
# Pre-Validation + UI Preview용 QC 생성
# - 결과 4개(KHD/WPH x 1Lane/2Lane) 각각 대표 item 1개를 자동 선택(처음 등장한 item)
# - 시간 공백(1시간 버킷 missing) / 샘플 수 부족(low count) 검출
# - 템플릿 매핑 누락 가능성(항목명 -> 시트명) 검출

from __future__ import annotations

import io
from typing import Dict, List, Tuple, Any, Optional

import pandas as pd

from errors import UserFacingError
from engine import LANE_SHEETS, detect_dtype, item_to_sheetname, safe_to_datetime


def _hours_to_labels(hours: List[int]) -> List[int]:
    # engine hour: 0..23 -> UI label: 1..24 (0시는 24로 표시)
    return [24 if h == 0 else h for h in hours]


def _labels_to_hours(labels: List[int]) -> List[int]:
    # UI label: 1..24 -> engine hour: 0..23
    out = []
    for v in labels:
        out.append(0 if v == 24 else int(v))
    return sorted(set(out))


def _compress_hours_as_ranges_ui(hours_engine: List[int]) -> str:
    """
    engine hours(0..23) -> "10~12시, 24시" 형태로 압축(연속 구간)
    0시는 24시로 표기
    """
    if not hours_engine:
        return "-"
    labels = sorted(_hours_to_labels(sorted(set(hours_engine))))
    # labels는 1..24 (24는 자정)
    ranges = []
    start = prev = labels[0]
    for h in labels[1:]:
        if h == prev + 1:
            prev = h
        else:
            ranges.append((start, prev))
            start = prev = h
    ranges.append((start, prev))

    def _fmt(a, b):
        if a == b:
            return f"{a:02d}시"
        return f"{a:02d}~{b:02d}시"

    return ", ".join(_fmt(a, b) for a, b in ranges)


def pre_validate(
    input_files: List[Tuple[str, bytes]],
    template_sheetnames: Dict[str, List[str]],  # {"KHD":[...], "WPH":[...]}
    selected_hours: Optional[List[int]] = None,  # engine hour 0..23
    low_count_threshold: int = 3,
) -> Dict[str, Any]:
    """
    return:
    {
      "ok": bool,
      "errors": [UserFacingError],
      "warnings": [str],
      "summary_rows": [dict],  # 표 출력용(파일/결과 4개 기준)
      "by_file": {
         fname: {
            "previews": { (dtype,lane): {item, hourly_avg_series(dict label->avg), hour_counts(dict hour->count), missing_hours(list), low_hours(list)} },
            "recommend": { "exclude_hours": [...engine hours...], "exclude_labels": [...UI labels...] }
         }
      },
      "recommend_global": { "exclude_hours": [...], "exclude_labels": [...] }
    }
    """
    errors: List[UserFacingError] = []
    warnings: List[str] = []
    summary_rows: List[Dict[str, Any]] = []

    by_file: Dict[str, Any] = {}

    # 선택 시간대가 없으면 "전체(0..23)" 기준으로 QC
    target_hours = selected_hours if selected_hours else list(range(24))
    target_hours_set = set(target_hours)

    global_exclude_hours: set[int] = set()

    for fname, fbytes in input_files:
        file_pack: Dict[str, Any] = {"previews": {}, "recommend": {"exclude_hours": [], "exclude_labels": []}}
        by_file[fname] = file_pack

        try:
            xl = pd.ExcelFile(io.BytesIO(fbytes))
        except Exception as e:
            errors.append(
                UserFacingError(
                    title="엑셀 파일을 열 수 없습니다.",
                    detail=f"파일: {fname}\n오류: {e}",
                    hint="파일 손상 또는 엑셀 형식 문제인지 확인하세요.",
                    code="EXCEL_OPEN_FAILED",
                    context={"file": fname},
                )
            )
            continue

        # Lane 단위로 raw 구성(엔진과 동일하게 4개 시트 concat)
        for lane, sheets in LANE_SHEETS.items():
            # 시트 존재 체크
            missing_sheets = [sh for sh in sheets if sh not in xl.sheet_names]
            if missing_sheets:
                errors.append(
                    UserFacingError(
                        title="원본 엑셀에 필요한 시트가 없습니다.",
                        detail=f"파일: {fname}\nLane: {lane}\n누락 시트: {missing_sheets}\n현재 시트: {xl.sheet_names}",
                        hint="원본 파일 포맷(시트명)이 변경되었는지 확인하세요.",
                        code="MISSING_REQUIRED_SHEET",
                        context={"file": fname, "lane": lane},
                    )
                )
                continue

            # 4개 시트 읽기(필수 컬럼만)
            dfs = []
            for sh in sheets:
                df = xl.parse(sh)
                required_cols = ["항목명", "측정일시", "측정값"]
                miss_cols = [c for c in required_cols if c not in df.columns]
                if miss_cols:
                    errors.append(
                        UserFacingError(
                            title="원본 시트에 필요한 컬럼이 없습니다.",
                            detail=f"파일: {fname}\n시트: {sh}\n누락 컬럼: {miss_cols}\n현재 컬럼: {df.columns.tolist()}",
                            hint="원본 엑셀 포맷(컬럼명)이 변경되었는지 확인하세요.",
                            code="MISSING_REQUIRED_COLUMN",
                            context={"file": fname, "sheet": sh},
                        )
                    )
                    continue

                df = df[required_cols].copy()
                dfs.append(df)

            if not dfs:
                continue

            lane_raw = pd.concat(dfs, ignore_index=True)
            lane_raw["항목명"] = lane_raw["항목명"].astype(str)

            dt = safe_to_datetime(lane_raw["측정일시"])
            parse_ok = dt.notna().mean()
            if parse_ok < 0.8:
                warnings.append(f"[{fname} / {lane}] 측정일시 파싱 성공률 낮음: {parse_ok:.0%}")

            lane_raw["_dt"] = dt
            lane_raw = lane_raw.dropna(subset=["_dt"])
            if lane_raw.empty:
                warnings.append(f"[{fname} / {lane}] 유효한 측정일시 데이터가 없습니다.")
                continue

            lane_raw["hour"] = lane_raw["_dt"].dt.hour
            lane_raw["val"] = pd.to_numeric(lane_raw["측정값"], errors="coerce")
            lane_raw = lane_raw.dropna(subset=["val"])
            if lane_raw.empty:
                warnings.append(f"[{fname} / {lane}] 유효한 측정값 데이터가 없습니다.")
                continue

            dt_min = lane_raw["_dt"].min()
            dt_max = lane_raw["_dt"].max()
            if (dt_max - dt_min).days >= 1:
                warnings.append(
                    f"[{fname} / {lane}] 데이터가 여러 날짜에 걸쳐 있음: {dt_min.date()} ~ {dt_max.date()}"
                )

            # Lane 전체 기준 hour count (그래프 축/시간 공백 판단에 사용)
            hour_counts_all = lane_raw["hour"].value_counts().to_dict()
            present_hours = set(int(h) for h in hour_counts_all.keys())

            missing_hours = sorted(list(target_hours_set - present_hours))
            low_hours = sorted(
                [h for h in target_hours if int(hour_counts_all.get(h, 0)) < low_count_threshold]
            )

            # 템플릿 매핑 누락 가능성(항목명 -> 템플릿 시트)
            # - 모든 item을 다 체크하면 느릴 수 있는데, 실무에서 템플릿 불일치는 치명적이라 유니크 item만 체크
            unique_items = lane_raw["항목명"].dropna().unique().tolist()
            for item in unique_items:
                dtype = detect_dtype(item)
                if dtype not in template_sheetnames:
                    continue
                sheet_name = item_to_sheetname(item)
                if sheet_name not in template_sheetnames[dtype]:
                    warnings.append(
                        f"[{fname}] 템플릿 매핑 누락 가능성: dtype={dtype}, lane={lane}, "
                        f"item='{item}' → sheet='{sheet_name}'"
                    )

            # dtype별 대표 item(처음 등장)
            # 엔진은 lane_raw에서 dtype별로 groupby하므로 dtype별 대표 item을 하나씩 뽑아 preview 생성
            lane_raw["dtype"] = lane_raw["항목명"].apply(detect_dtype)

            for dtype in ["KHD", "WPH"]:
                sub = lane_raw[lane_raw["dtype"] == dtype]
                if sub.empty:
                    continue

                # 대표 item: "처음 등장한 항목명"
                rep_item = sub["항목명"].iloc[0]

                rep_df = sub[sub["항목명"] == rep_item]
                # 시간별 AVG(대표 item 기준)
                hourly_avg = (
                    rep_df.groupby("hour")["val"].mean().to_dict()
                )  # {hour:int -> avg:float}

                # UI용 시계열 dict (1..24 라벨로 변환)
                hourly_avg_ui = { (24 if h == 0 else int(h)): float(v) for h, v in hourly_avg.items() }
                # 1..24 전체 라벨에 대해 None 채워서 chart에서 빈 구간이 보이게
                series_ui = {lbl: (hourly_avg_ui.get(lbl, None)) for lbl in range(1, 25)}

                key = (dtype, lane)
                file_pack["previews"][key] = {
                    "item": rep_item,
                    "date_range": f"{dt_min.strftime('%Y-%m-%d')} ~ {dt_max.strftime('%Y-%m-%d')}",
                    "parse_ok": float(parse_ok),
                    "hour_counts": {int(k): int(v) for k, v in hour_counts_all.items()},
                    "missing_hours": missing_hours,
                    "low_hours": low_hours,
                    "hourly_avg_series": series_ui,  # {1..24 -> avg or None}
                }

                # summary row(파일/결과4개 기준)
                summary_rows.append(
                    {
                        "File": fname,
                        "Output": f"{dtype} {lane}",
                        "Preview Item": rep_item,
                        "Date Range": f"{dt_min.strftime('%Y-%m-%d')} ~ {dt_max.strftime('%Y-%m-%d')}",
                        "Parse OK %": f"{parse_ok:.0%}",
                        "Missing Hours": _compress_hours_as_ranges_ui(missing_hours),
                        "LowCount Hours": _compress_hours_as_ranges_ui(low_hours),
                    }
                )

            # 추천 제외 시간(일단 기본은 missing + low를 합치되, 실무에선 missing만 추천으로 둬도 됨)
            # 여기서는 "missing은 강추천", "low는 약추천"이므로 exclude에는 missing만 넣고,
            # low는 UI에서 참고 경고로만 쓰는 게 더 유저친화적일 때가 많음.
            exclude_hours_lane = set(missing_hours)
            global_exclude_hours |= exclude_hours_lane

        # 파일 단위 recommend 채우기
        file_pack["recommend"]["exclude_hours"] = sorted(global_exclude_hours)  # 파일별이라기보다 글로벌과 동일하게 사용
        file_pack["recommend"]["exclude_labels"] = sorted(_hours_to_labels(sorted(global_exclude_hours)))

    ok = len(errors) == 0
    recommend_global = {
        "exclude_hours": sorted(global_exclude_hours),
        "exclude_labels": sorted(_hours_to_labels(sorted(global_exclude_hours))),
    }

    return {
        "ok": ok,
        "errors": errors,
        "warnings": warnings,
        "summary_rows": summary_rows,
        "by_file": by_file,
        "recommend_global": recommend_global,
    }
