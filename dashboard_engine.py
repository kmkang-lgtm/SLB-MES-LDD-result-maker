import os, re, glob, zipfile, tempfile, shutil, datetime
from pathlib import Path
import pandas as pd
import numpy as np


# -----------------------------
# Input handling
# -----------------------------
def extract_input_from_zip_bytes(zip_bytes: bytes, zip_name: str):
    """
    Streamlit 업로드 zip(bytes)을 임시폴더에 풀고,
    내부 xlsx/xlsm 목록 반환.
    """
    tmpdir = tempfile.mkdtemp(prefix="daily_xlsx_")
    zip_path = Path(tmpdir) / (zip_name or "daily.zip")

    with open(zip_path, "wb") as f:
        f.write(zip_bytes)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(tmpdir)

    # ✅ .xlsx + .xlsm 둘 다 인식
    files = sorted(
        glob.glob(os.path.join(tmpdir, "**", "*.xlsx"), recursive=True) +
        glob.glob(os.path.join(tmpdir, "**", "*.xlsm"), recursive=True)
    )
    return files, tmpdir


def extract_input_from_file_bytes(file_bytes_list):
    """
    Summary 엑셀을 여러 개(bytes) 직접 받은 경우,
    임시폴더에 저장하고 경로 리스트 반환.
    file_bytes_list: [(filename, bytes), ...]
    """
    tmpdir = tempfile.mkdtemp(prefix="daily_xlsx_")
    files = []
    for name, b in file_bytes_list:
        p = Path(tmpdir) / Path(name).name
        with open(p, "wb") as f:
            f.write(b)
        files.append(str(p))
    return files, tmpdir


def date_from_filename(fn):
    """
    파일명 끝에 _MM.DD.xlsx 패턴이 있으면 날짜 추출.
    없으면 파일 수정일로 날짜 사용.
    """
    base = os.path.basename(fn)
    m = re.search(r"_(\d{1,2})\.(\d{2})\.(xlsx|xlsm)$", base, re.I)  # ✅ xlsm도 허용
    if m:
        month = int(m.group(1))
        day = int(m.group(2))
        year = datetime.date.today().year
        return datetime.date(year, month, day)

    ts = os.path.getmtime(fn)
    return datetime.date.fromtimestamp(ts)


# -----------------------------
# Parsing daily file
# -----------------------------
def parse_block_label_row(label_row):
    """
    상단 라벨 행에서 WPH/KHD 시작 위치 인덱스 찾아
    WPH 1Lane / KHD 1Lane / WPH 2Lane / KHD 2Lane 라벨 생성.
    """
    label_indices = [i for i, v in enumerate(label_row) if pd.notna(v)]
    labels = []
    wph_count, khd_count = 0, 0

    for i in label_indices:
        v = str(label_row[i]).strip().upper()

        # ✅ "WPH", "WPH AVG", " WPH " 등 포함으로도 인식
        if "WPH" in v:
            wph_count += 1
            labels.append((i, f"WPH {wph_count}Lane"))
        elif "KHD" in v:
            khd_count += 1
            labels.append((i, f"KHD {khd_count}Lane"))

    return labels, label_indices


def read_one_tidy(path):
    """
    일일 Summary 파일을
    (WPH1/KHD1/WPH2/KHD2) 4블록으로 쪼개 tidy 형태로 반환.
    """
    raw = pd.read_excel(path, header=None)

    label_row = list(raw.iloc[1])   # row 1 => 라벨(WPH/KHD)
    fields = list(raw.iloc[2])      # row 2 => 컬럼명

    labels, label_indices = parse_block_label_row(label_row)
    boundaries = label_indices + [len(fields)]

    data_rows = raw.iloc[3:].copy()
    tidy_parts = []

    for (start_idx, block_label), end_idx in zip(labels, boundaries[1:]):
        start = (
            start_idx - 1
            if start_idx > 0 and str(fields[start_idx - 1]).strip().lower() == "lane"
            else start_idx
        )
        end = end_idx - 1

        sub = data_rows.iloc[:, start : end + 1].copy()
        sub.columns = [str(f).strip() for f in fields[start : end + 1]]
        sub["Block"] = block_label
        tidy_parts.append(sub)

    tidy = pd.concat(tidy_parts, ignore_index=True)
    tidy = tidy[tidy["position"].notna()]

    for c in ["count", "mean", "std", "min", "max", "range", "Lane"]:
        if c in tidy.columns:
            tidy[c] = pd.to_numeric(tidy[c], errors="coerce")

    return tidy


def parse_function_line(block):
    """
    Block("WPH 1Lane") -> Function=WPH, Line=1
    """
    m = re.match(r"(WPH|KHD)\s*(\d)Lane", str(block), re.I)
    if m:
        return m.group(1).upper(), int(m.group(2))
    return None, None


# -----------------------------
# Material rule (AL1/AL2/CU)
# -----------------------------
def base_key(pos):
    s = str(pos).strip()
    s = re.sub(r"\s+", " ", s)

    parts = [p.strip() for p in s.split("-")]
    if len(parts) >= 2 and parts[-1] in ["1", "2"]:
        return "-".join(parts[:-1]).strip(), parts[-1]

    m = re.match(r"^(.*\d)\s*-\s*(\d)\s*$", s)
    if m:
        return m.group(1).strip(), m.group(2)

    return s, None


def add_material_al12(df):
    bases, lasts = [], []
    for p in df["position"].astype(str):
        b, t = base_key(p)
        bases.append(b)
        lasts.append(t)

    tmp = pd.DataFrame({"base": bases, "last": lasts})
    has2 = (
        tmp.groupby("base")["last"]
        .apply(lambda x: any(v == "2" for v in x if v is not None))
        .to_dict()
    )

    mats = []
    for b, t in zip(tmp["base"], tmp["last"]):
        if has2.get(b, False):
            if t == "1":
                mats.append("AL1")
            elif t == "2":
                mats.append("AL2")
            else:
                mats.append("AL")
        else:
            mats.append("CU")

    df["Material"] = mats
    return df


# -----------------------------
# Pivot + formulas
# -----------------------------
def make_pivot(df, metric):
    pv = df.pivot_table(
        index=["Function", "Line", "position", "Material"],
        columns="Date",
        values=metric,
        aggfunc="mean",
    ).reset_index()

    pv["Function"] = pv["Function"].ffill()
    pv["Line"] = pv["Line"].ffill()

    new_cols = []
    for c in pv.columns:
        try:
            d = pd.to_datetime(c).date()
            new_cols.append(d.isoformat())
        except Exception:
            new_cols.append(c)
    pv.columns = new_cols

    return pv


def write_pivot_with_formulas(writer, sheet_name, df):
    date_cols = [c for c in df.columns if re.match(r"\d{4}-\d{2}-\d{2}", str(c))]
    id_cols = [c for c in df.columns if c not in date_cols]
    base = df[id_cols + date_cols]

    base.to_excel(writer, index=False, sheet_name=sheet_name)
    wb = writer.book
    ws = writer.sheets[sheet_name]

    def xl_col(n):
        s = ""
        while True:
            n, r = divmod(n, 26)
            s = chr(65 + r) + s
            if n == 0:
                break
            n -= 1
        return s

    date_idx = [base.columns.get_loc(c) for c in date_cols]
    start_extra = len(base.columns)

    header_fmt = wb.add_format({"bold": True})
    for j, h in enumerate(["AVG", "MIN", "MAX"]):
        ws.write(0, start_extra + j, h, header_fmt)

    avg_fmt = wb.add_format({"bold": True, "left": 2, "num_format": "0.000"})
    min_fmt = wb.add_format({"bold": True, "num_format": "0.000"})
    max_fmt = wb.add_format({"bold": True, "num_format": "0.000"})

    for i in range(len(base)):
        excel_row = i + 2
        refs = [f"{xl_col(ci)}{excel_row}" for ci in date_idx]

        ws.write_formula(excel_row - 1, start_extra + 0,
                         f"=AVERAGE({','.join(refs)})", avg_fmt)
        ws.write_formula(excel_row - 1, start_extra + 1,
                         f"=MIN({','.join(refs)})", min_fmt)
        ws.write_formula(excel_row - 1, start_extra + 2,
                         f"=MAX({','.join(refs)})", max_fmt)

    ws.set_column(start_extra + 0, start_extra + 0, 12, avg_fmt)
    ws.set_column(start_extra + 1, start_extra + 2, 12)


# -----------------------------
# Core builder (공통)
# -----------------------------
def _build_dashboard_from_files(files, tmpdir):
    if not files:
        raise ValueError("입력 파일이 없습니다.")

    all_parts = []
    for f in files:
        try:
            t = read_one_tidy(f)
            t["Date"] = date_from_filename(f)
            all_parts.append(t)
        except Exception as e:
            # ✅ 어떤 파일이 스킵됐는지 로그 남김
            print(f"[dashboard] skip file: {f} / {e}")
            continue

    if not all_parts:
        raise ValueError("읽을 수 있는 일일 Summary 파일이 없습니다.")

    all_df = pd.concat(all_parts, ignore_index=True)

    func_line = all_df["Block"].apply(parse_function_line)
    all_df["Function"] = [x[0] for x in func_line]
    all_df["Line"] = [x[1] for x in func_line]

    all_df["Date"] = pd.to_datetime(all_df["Date"], errors="coerce").dt.date
    all_df = all_df.drop(columns=["Block", "Lane"], errors="ignore")

    all_df = add_material_al12(all_df)
    all_df = all_df.dropna(axis=1, how="all")

    desired_cols = [
        "Date", "Function", "Line", "position", "group", "Material",
        "count", "mean", "std", "min", "max", "range"
    ]
    all_df = all_df[[c for c in desired_cols if c in all_df.columns]]

    all_position = all_df.sort_values(
        ["position", "Function", "Line", "Date"]
    ).reset_index(drop=True)

    pivot_mean  = make_pivot(all_df, "mean")
    pivot_std   = make_pivot(all_df, "std")
    pivot_range = make_pivot(all_df, "range")
    heatmap_mean = pivot_mean.copy()

    out_tmp = Path(tmpdir) / "DASHBOARD.xlsx"
    with pd.ExcelWriter(out_tmp, engine="xlsxwriter", date_format="yyyy-mm-dd") as writer:
        all_df.to_excel(writer, index=False, sheet_name="All_Data")
        all_position.to_excel(writer, index=False, sheet_name="ALL POSITION")
        write_pivot_with_formulas(writer, "Pivot_Mean", pivot_mean)
        write_pivot_with_formulas(writer, "Pivot_Std", pivot_std)
        write_pivot_with_formulas(writer, "Pivot_Range", pivot_range)
        write_pivot_with_formulas(writer, "Heatmap_Mean", heatmap_mean)

    dash_bytes = out_tmp.read_bytes()

    # ✅ zip 이름 무시하고 내부 날짜 범위로 파일명 생성
    dates = sorted(all_df["Date"].dropna().unique().tolist())
    if dates:
        start_mmdd = pd.to_datetime(dates[0]).strftime("%m.%d")
        end_mmdd   = pd.to_datetime(dates[-1]).strftime("%m.%d")
        tag = start_mmdd if start_mmdd == end_mmdd else f"{start_mmdd}~{end_mmdd}"
        dash_name = f"SLB_MES_Dashboard_{tag}.xlsx"
    else:
        dash_name = "SLB_MES_Dashboard.xlsx"

    return dash_name, dash_bytes


# -----------------------------
# Public APIs for Streamlit
# -----------------------------
def build_dashboard_from_zip_bytes(zip_bytes: bytes, zip_name: str):
    files, tmpdir = extract_input_from_zip_bytes(zip_bytes, zip_name)
    try:
        return _build_dashboard_from_files(files, tmpdir)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def build_dashboard_from_file_bytes(file_bytes_list):
    """
    file_bytes_list: [(filename, bytes), ...]
    """
    files, tmpdir = extract_input_from_file_bytes(file_bytes_list)
    try:
        return _build_dashboard_from_files(files, tmpdir)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
