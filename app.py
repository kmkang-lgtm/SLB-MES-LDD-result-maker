import streamlit as st
import tempfile
import os
import zipfile
import gc
import re
from pathlib import Path

from engine import make_results_for_input
from summary_engine import build_from_zip_bytes
from dashboard_engine import (
    build_dashboard_from_zip_bytes,
    build_dashboard_from_file_bytes
)

st.set_page_config(page_title="SLB MES Result Maker", layout="wide")


# =========================================================
# 0) 비밀번호 게이트 (Secrets 기반)
# =========================================================
DEFAULT_PASSWORD = st.secrets.get("APP_PASSWORD", "")
if not DEFAULT_PASSWORD:
    st.error("관리자에게 비밀번호 설정(Secrets)을 요청하세요.")
    st.stop()

if "authed" not in st.session_state:
    st.session_state["authed"] = False

if not st.session_state["authed"]:
    st.title("SLB MES Result Maker 🔒")
    st.caption("접근하려면 비밀번호를 입력하세요.")
    pw = st.text_input("Password", type="password")
    if pw == DEFAULT_PASSWORD:
        st.session_state["authed"] = True
        st.rerun()
    else:
        st.stop()


# =========================
# 경로 설정
# =========================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_KHD_TPL = os.path.join(APP_DIR, "templates", "TEMPLATE_KHD.xlsx")
DEFAULT_WPH_TPL = os.path.join(APP_DIR, "templates", "TEMPLATE_WPH.xlsx")


# =========================
# 로고 찾기
# =========================
def find_logo_path():
    exts = ["png", "jpg", "jpeg"]
    search_dirs = [
        Path(APP_DIR) / "assets",
        Path(os.getcwd()) / "assets",
    ]
    for d in search_dirs:
        for ext in exts:
            p = d / f"logo.{ext}"
            if p.exists():
                return str(p)
    return None

logo_path_found = find_logo_path()


# =========================
# 날짜 추출 유틸 (YY.MM.DD 또는 MM.DD 둘 다 지원)
# =========================
_DATE_RE_YYMMDD = re.compile(r"(\d{2})\.(\d{2})\.(\d{2})")            # 25.12.01
_DATE_RE_MMDD   = re.compile(r"(?<!\d)(\d{1,2})\.(\d{2})(?!\d)")     # 12.01 / 8.01

def extract_mmdd(text: str):
    """
    text에서 날짜를 찾아 MM.DD 형태로 리턴.
    - 1순위: YY.MM.DD (25.12.01 -> 12.01)
    - 2순위: MM.DD     (12.01 -> 12.01)
    """
    text = text or ""

    m = _DATE_RE_YYMMDD.search(text)
    if m:
        _, mm, dd = m.groups()
        return f"{mm}.{dd}"

    m = _DATE_RE_MMDD.search(text)
    if m:
        mm, dd = m.groups()
        mm = mm.zfill(2)
        return f"{mm}.{dd}"

    return None


def extract_mmdd_from_sources(raw_files=None, raw_zip_name=None, extracted_names=None):
    """
    날짜 우선순위:
    1) raw zip 파일명에서
    2) 업로드 raw xlsx 파일명에서
    3) zip 내부 xlsx 파일명에서
    """
    if raw_zip_name:
        mmdd = extract_mmdd(raw_zip_name)
        if mmdd:
            return mmdd

    if raw_files:
        for rf in raw_files:
            mmdd = extract_mmdd(rf.name)
            if mmdd:
                return mmdd

    if extracted_names:
        for name in extracted_names:
            mmdd = extract_mmdd(name)
            if mmdd:
                return mmdd

    return None


# =========================
# 세션 상태
# =========================
if "results" not in st.session_state:
    st.session_state["results"] = []
if "zip_bytes" not in st.session_state:
    st.session_state["zip_bytes"] = None
if "zip_filename" not in st.session_state:
    st.session_state["zip_filename"] = None


def safe_gc_collect():
    """Streamlit Cloud에서 UploadedFile 버퍼 충돌(BufferError) 방지."""
    try:
        gc.collect()
    except BufferError:
        pass


def safe_read_bytes(path: Path, retries: int = 2):
    last_err = None
    for _ in range(retries + 1):
        try:
            with open(path, "rb") as f:
                return f.read()
        except PermissionError as e:
            last_err = e
            safe_gc_collect()
    raise last_err


def save_uploaded_to_temp(uploaded_file, tmp_dir: Path):
    """
    UploadedFile을 임시폴더에 저장.
    - getbuffer 대신 getvalue 사용: Streamlit Cloud 안정성↑
    """
    fname = Path(uploaded_file.name).name
    out_path = tmp_dir / fname
    data = uploaded_file.getvalue()  # ✅ bytes 복사본
    with open(out_path, "wb") as f:
        f.write(data)
    return str(out_path)


def extract_raw_zip_to_paths(raw_zip_file, tmp_dir: Path):
    """
    raw zip(폴더 압축)을 풀어서 안에 있는 xlsx 전부 찾아 경로 리스트로 반환
    """
    zip_path = Path(save_uploaded_to_temp(raw_zip_file, tmp_dir))
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(tmp_dir)

    xlsx_paths = [str(p) for p in tmp_dir.rglob("*.xlsx")]
    return xlsx_paths


# =========================
# 헤더(타이틀 + 로고)
# =========================
col_title, col_logo = st.columns([5, 1], vertical_alignment="center")
with col_title:
    st.title("SLB MES 결과 생성기")
    st.caption("KHD/WPH 원본을 파싱해 Lane1/2 Result를 템플릿 기반으로 자동 생성합니다.")
with col_logo:
    if logo_path_found:
        st.image(logo_path_found, width="stretch")
    else:
        st.caption("⚠️ logo.png 없음")


# =========================
# 사이드바 UI
# =========================
with st.sidebar:
    st.header("STEP 1) 원본 파일 선택")

    st.caption("✅ 방법 A) KHD/WPH 원본 xlsx 여러 개 업로드")
    raw_files = st.file_uploader(
        "KHD/WPH 원본 (.xlsx) - 복수 선택 가능",
        type=["xlsx"],
        accept_multiple_files=True,
        key="raw_xlsx_uploader"
    )

    st.caption("✅ 방법 B) KHD+WPH가 들어있는 폴더를 zip으로 압축해 1개 업로드")
    raw_zip = st.file_uploader(
        "원본 폴더 ZIP(선택)",
        type=["zip"],
        accept_multiple_files=False,
        key="raw_zip_uploader"
    )

    st.divider()
    st.header("STEP 2) 템플릿 (기본 자동 사용)")
    st.caption("기본 템플릿은 관리자(강경민) 관리 버전이 자동 적용됩니다.")
    st.write("기본 KHD 템플릿:", os.path.basename(DEFAULT_KHD_TPL))
    st.write("기본 WPH 템플릿:", os.path.basename(DEFAULT_WPH_TPL))

    with st.expander("템플릿을 직접 바꾸고 싶다면(옵션)", expanded=False):
        tpl_khd = st.file_uploader("KHD 템플릿 업로드(선택)", type=["xlsx"], key="tpl_khd")
        tpl_wph = st.file_uploader("WPH 템플릿 업로드(선택)", type=["xlsx"], key="tpl_wph")
        st.caption("업로드하면 해당 템플릿이 기본 템플릿보다 우선 적용됩니다.")

    st.divider()
    st.header("STEP 3) 옵션")
    raw_end_row = st.number_input(
        "Raw 끝행(차트 참조 범위 끝)",
        min_value=50, max_value=500, value=100, step=10,
        help="템플릿 차트가 참조하는 Raw 데이터의 마지막 행"
    )

    st.subheader("시간 필터(선택)")
    st.caption("선택한 시간만 결과/그래프에 포함됩니다. 비워두면 전체 자동 포함.")

    hour_options = list(range(0, 24))
    hour_labels_ui = [24 if h == 0 else h for h in hour_options]

    selected_ui = st.multiselect(
        "포함할 시간 선택",
        options=hour_labels_ui,
        default=[],
        help="예: 8,9,10만 선택하면 그 시간만 결과에 표시"
    )
    selected_hours = [0 if h == 24 else h for h in selected_ui]

    col1, col2 = st.columns(2)
    run_btn = col1.button("🚀 실행", width="stretch", key="btn-run")
    clear_btn = col2.button("🧹 결과 초기화", width="stretch", key="btn-clear")

    st.divider()
    st.markdown(
        "<div style='font-size:12px;color:gray;text-align:right;'>BYKKM</div>",
        unsafe_allow_html=True
    )


# =========================
# 결과 초기화
# =========================
if clear_btn:
    st.session_state["results"] = []
    st.session_state["zip_bytes"] = None
    st.session_state["zip_filename"] = None
    st.success("결과를 초기화했습니다. 다시 실행하세요.")


# =========================
# 메인 화면: 현재 선택 표시
# =========================
left, right = st.columns([1.2, 1])

with left:
    st.subheader("현재 선택된 원본")
    if raw_zip:
        st.write(f"- ZIP: {raw_zip.name} ({raw_zip.size/1024/1024:.1f} MB)")
    if raw_files:
        for rf in raw_files:
            st.write(f"- {rf.name} ({rf.size/1024/1024:.1f} MB)")
    if not raw_zip and not raw_files:
        st.info("왼쪽에서 원본 xlsx 또는 원본 폴더 ZIP을 선택하세요.")

with right:
    st.subheader("템플릿 적용 상태")
    st.write("✅ KHD 템플릿:",
             "기본 사용" if st.session_state.get("tpl_khd") is None else "사용자 업로드")
    st.write("✅ WPH 템플릿:",
             "기본 사용" if st.session_state.get("tpl_wph") is None else "사용자 업로드")
    st.write("Raw 끝행:", raw_end_row)

st.divider()


# =========================
# 실행
# =========================
if run_btn:
    if (not raw_files) and (raw_zip is None):
        st.error("원본 xlsx 또는 원본 폴더 ZIP을 하나 이상 선택해줘.")
        st.stop()

    if not os.path.exists(DEFAULT_KHD_TPL) or not os.path.exists(DEFAULT_WPH_TPL):
        st.error("기본 템플릿을 찾을 수 없습니다. templates 폴더 구성을 확인하세요.")
        st.stop()

    with st.spinner("파싱 및 결과 생성 중..."):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            tmp_dir = Path(tmp)

            # 템플릿 우선순위: 기본 -> 업로드
            final_khd_tpl = DEFAULT_KHD_TPL
            final_wph_tpl = DEFAULT_WPH_TPL

            if st.session_state.get("tpl_khd") is not None:
                final_khd_tpl = save_uploaded_to_temp(st.session_state["tpl_khd"], tmp_dir)
            if st.session_state.get("tpl_wph") is not None:
                final_wph_tpl = save_uploaded_to_temp(st.session_state["tpl_wph"], tmp_dir)

            templates = {"KHD": final_khd_tpl, "WPH": final_wph_tpl}

            # raw 입력을 실제 파일 경로 리스트로 통일
            raw_paths = []
            extracted_names = []

            if raw_zip is not None:
                raw_paths = extract_raw_zip_to_paths(raw_zip, tmp_dir)
                extracted_names = [Path(p).name for p in raw_paths]
            else:
                for rf in raw_files:
                    raw_paths.append(save_uploaded_to_temp(rf, tmp_dir))

            if not raw_paths:
                st.error("ZIP 안에 xlsx가 없습니다. 압축 구조를 확인해줘.")
                st.stop()

            # 날짜 기반 ZIP 네이밍
            mmdd = extract_mmdd_from_sources(
                raw_files=raw_files,
                raw_zip_name=(raw_zip.name if raw_zip else None),
                extracted_names=extracted_names
            )
            zip_base = f"SLB_MES_Result_Package_{mmdd}" if mmdd else "SLB_MES_Result_Package"
            zip_filename = f"{zip_base}.zip"

            created_paths = []
            for raw_path in raw_paths:
                created = make_results_for_input(
                    raw_path,
                    templates=templates,
                    output_dir=str(tmp_dir),
                    raw_end_row=raw_end_row,
                    selected_hours=selected_hours
                )
                created_paths.extend(created)
                safe_gc_collect()

            all_created_bytes = []
            for p in created_paths:
                p_path = Path(p)
                data = safe_read_bytes(p_path)
                all_created_bytes.append((p_path.name, data))

            zip_path = tmp_dir / zip_filename
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in created_paths:
                    zf.write(p, arcname=Path(p).name)

            zip_bytes = safe_read_bytes(zip_path)

            st.session_state["results"] = all_created_bytes
            st.session_state["zip_bytes"] = zip_bytes
            st.session_state["zip_filename"] = zip_filename

    st.success("완료! 아래에서 결과 파일을 다운로드하세요.")


# =========================
# 결과 표시
# =========================
if st.session_state["results"]:
    st.subheader("개별 결과 파일")
    for i, (filename, data) in enumerate(st.session_state["results"]):
        st.download_button(
            label=f"⬇️ {filename}",
            data=data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl-{i}-{filename}"
        )

    st.subheader("전체 ZIP")
    zip_name_for_dl = st.session_state.get("zip_filename", "SLB_MES_Result_Package.zip")
    st.download_button(
        label="⬇️ 전체 결과 ZIP 다운로드",
        data=st.session_state["zip_bytes"],
        file_name=zip_name_for_dl,
        mime="application/zip",
        key="dl-zip"
    )
else:
    st.info("원본을 선택하고 실행을 누르면 결과가 표시됩니다.")


# =========================
# Deviation Summary 생성(Zip 기반)
# =========================
st.divider()
st.subheader("Deviation Summary 생성")

zip_upload = st.file_uploader(
    "기존 SLB_MES_Result_Package_XX.XX.zip을 업로드하면 Summary를 생성합니다.",
    type=["zip"],
    key="zip_uploader_for_summary"
)

use_latest_zip = st.checkbox("방금 생성된 ZIP으로 Summary 만들기", value=False, key="chk-use-latest-zip")

if st.button("📌 Summary 생성하기", width="stretch", key="btn-build-summary"):
    try:
        if use_latest_zip:
            if st.session_state.get("zip_bytes") is None:
                st.error("먼저 결과 ZIP을 생성한 뒤 체크하세요.")
                st.stop()
            zip_bytes = st.session_state["zip_bytes"]
            zip_name = st.session_state.get("zip_filename", "SLB_MES_Result_Package.zip")
        else:
            if zip_upload is None:
                st.error("ZIP 파일을 업로드하거나, '방금 생성된 ZIP'을 선택하세요.")
                st.stop()
            # ✅ getvalue() 사용
            zip_bytes = zip_upload.getvalue()
            zip_name = zip_upload.name

        with st.spinner("Summary 생성 중..."):
            summary_name, summary_bytes = build_from_zip_bytes(zip_bytes, zip_name)

        st.success("Summary 생성 완료!")
        st.download_button(
            "⬇️ Summary 다운로드",
            data=summary_bytes,
            file_name=summary_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl-summary"
        )

    except Exception as e:
        st.error(f"Summary 생성 실패: {e}")


# =========================
# ✅ Dashboard 생성 (ZIP or 개별 Summary 업로드)
# =========================
st.divider()
st.subheader("Dashboard 생성 (여러 일자 Summary 묶음)")

st.caption(
    "✅ 방법 A) 여러 날짜 Summary 파일들을 폴더에 모아 zip으로 압축해 업로드\n"
    "✅ 방법 B) Summary 엑셀들을 개별로 여러 개 직접 업로드"
)

dash_zip = st.file_uploader(
    "방법 A) Summary 폴더 ZIP 업로드(선택)",
    type=["zip"],
    key="zip_uploader_for_dashboard"
)

dash_files = st.file_uploader(
    "방법 B) Summary 엑셀 여러 개 업로드(선택)",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    key="xlsx_uploader_for_dashboard"
)

use_latest_zip_for_dash = st.checkbox(
    "방금 생성된 ZIP으로 Dashboard 만들기",
    value=False,
    key="chk-use-latest-zip-for-dash"
)

if st.button("📊 Dashboard 생성하기", width="stretch", key="btn-build-dashboard"):
    try:
        if use_latest_zip_for_dash:
            if st.session_state.get("zip_bytes") is None:
                st.error("먼저 결과 ZIP을 생성한 뒤 체크하세요.")
                st.stop()
            zip_bytes = st.session_state["zip_bytes"]
            zip_name = st.session_state.get("zip_filename", "SLB_MES_Result_Package.zip")

            with st.spinner("Dashboard 생성 중...(최신 ZIP)"):
                dash_name, dash_bytes = build_dashboard_from_zip_bytes(zip_bytes, zip_name)

        elif dash_zip is not None:
            # ✅ getvalue() 사용
            zip_bytes = dash_zip.getvalue()
            zip_name = dash_zip.name

            with st.spinner("Dashboard 생성 중...(ZIP)"):
                dash_name, dash_bytes = build_dashboard_from_zip_bytes(zip_bytes, zip_name)

        elif dash_files:
            # ✅ getvalue()로 bytes 복사본 생성
            file_bytes_list = [(f.name, f.getvalue()) for f in dash_files]

            with st.spinner("Dashboard 생성 중...(엑셀 개별)"):
                dash_name, dash_bytes = build_dashboard_from_file_bytes(file_bytes_list)

        else:
            st.error("ZIP 또는 Summary 엑셀 파일들을 업로드하세요.")
            st.stop()

        st.success("Dashboard 생성 완료!")
        st.download_button(
            "⬇️ Dashboard 다운로드",
            data=dash_bytes,
            file_name=dash_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl-dashboard"
        )

    except Exception as e:
        st.error(f"Dashboard 생성 실패: {e}")
