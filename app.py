import streamlit as st
import tempfile
import os
import zipfile
import base64
import gc
from pathlib import Path
from engine import make_results_for_input

st.set_page_config(page_title="SLB MES Result Maker", layout="wide")

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
LOGO_PATH = os.path.join(APP_DIR, "assets", "logo.png")

# =========================
# 로고/크레딧 (상단 우측 / 하단 우측 고정)
# =========================
def find_logo_b64():
    """
    로고를 여러 기준 경로에서 찾는다.
    1) app.py 기준 APP_DIR/assets/logo.*
    2) 현재 실행 폴더(cwd)/assets/logo.*
    logo.png/jpg/jpeg 모두 허용
    """
    exts = ["png", "jpg", "jpeg"]
    search_dirs = [
        Path(APP_DIR) / "assets",
        Path(os.getcwd()) / "assets",
    ]

    for d in search_dirs:
        for ext in exts:
            p = d / f"logo.{ext}"
            if p.exists():
                with open(p, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode()
                mime = "image/png" if ext == "png" else "image/jpeg"
                return b64, mime, str(p)

    return None, None, None


logo_b64, logo_mime, logo_path_found = find_logo_b64()

# 🔎 못 찾았을 때는 경로를 화면에 보여줘서 바로 원인 확인 가능
if logo_b64 is None:
    st.warning(
        "로고 파일을 찾지 못했습니다.\n"
        f"- APP_DIR 기준: {Path(APP_DIR)/'assets'}\n"
        f"- CWD 기준: {Path(os.getcwd())/'assets'}\n"
        "assets 폴더에 logo.png(또는 jpg/jpeg)를 넣어주세요."
    )

st.markdown(
    f"""
    <style>
    .top-right-logo {{
        position: fixed;
        top: 12px;
        right: 16px;
        z-index: 9999;
        display: flex;
        align-items: center;
        opacity: 0.95;
        pointer-events: none;
        background: rgba(255,255,255,0.95);
        padding: 6px 8px;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    }}
    .top-right-logo img {{
        height: 56px;
        width: auto;
        filter: drop-shadow(0 1px 2px rgba(0,0,0,0.2));
    }}
    .bottom-right-credit {{
        position: fixed;
        bottom: 10px;
        right: 14px;
        z-index: 9999;
        font-size: 12px;
        color: rgba(120,120,120,0.9);
        font-family: Arial, sans-serif;
        letter-spacing: 0.2px;
        pointer-events: none;
    }}
    </style>

    {f"<div class='top-right-logo'><img src='data:{logo_mime};base64,{logo_b64}' /></div>" if logo_b64 else ""}
    <div class="bottom-right-credit">by.kkm</div>
    """,
    unsafe_allow_html=True
)

# =========================
# 타이틀
# =========================
st.title("SLB MES 결과 생성기")
st.caption("KHD/WPH 원본을 파싱해 Lane1/2 Result를 템플릿 기반으로 자동 생성합니다.")

# =========================
# 세션 상태(다운로드 눌러도 결과 유지)
# =========================
if "results" not in st.session_state:
    st.session_state["results"] = []     # [(filename, bytes), ...]
if "zip_bytes" not in st.session_state:
    st.session_state["zip_bytes"] = None


def safe_read_bytes(path: Path, retries: int = 2):
    """
    Windows에서 간헐적으로 파일 잠금이 남는 경우가 있어
    bytes 읽기만 가볍게 재시도.
    """
    last_err = None
    for _ in range(retries + 1):
        try:
            with open(path, "rb") as f:
                return f.read()
        except PermissionError as e:
            last_err = e
            # 재시도 전 GC로 핸들 해제 시도
            gc.collect()
    raise last_err


def save_uploaded_to_temp(uploaded_file, tmp_dir: Path):
    """
    업로드 파일을 tmp_dir에 저장.
    - Path 기반
    - 파일명 안전화
    """
    fname = Path(uploaded_file.name).name  # 혹시 경로 포함돼도 basename만
    out_path = tmp_dir / fname
    with open(out_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return str(out_path)


# =========================
# 사이드바 UI
# =========================
with st.sidebar:
    st.header("STEP 1) 원본 파일 선택")
    raw_files = st.file_uploader(
        "KHD/WPH 원본 (.xlsx) - 복수 선택 가능",
        type=["xlsx"],
        accept_multiple_files=True
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

    col1, col2 = st.columns(2)
    run_btn = col1.button("🚀 실행", use_container_width=True)
    clear_btn = col2.button("🧹 결과 초기화", use_container_width=True)

# =========================
# 결과 초기화
# =========================
if clear_btn:
    st.session_state["results"] = []
    st.session_state["zip_bytes"] = None
    st.success("결과를 초기화했습니다. 다시 실행하세요.")

# =========================
# 메인 화면: 현재 선택 표시
# =========================
left, right = st.columns([1.2, 1])

with left:
    st.subheader("현재 선택된 원본")
    if raw_files:
        for rf in raw_files:
            st.write(f"- {rf.name} ({rf.size/1024/1024:.1f} MB)")
    else:
        st.info("왼쪽에서 KHD/WPH 원본 파일을 선택하세요.")

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
    if not raw_files:
        st.error("원본 파일을 하나 이상 선택해줘.")
        st.stop()

    if not os.path.exists(DEFAULT_KHD_TPL) or not os.path.exists(DEFAULT_WPH_TPL):
        st.error("기본 템플릿을 찾을 수 없습니다. templates 폴더 구성을 확인하세요.")
        st.stop()

    with st.spinner("파싱 및 결과 생성 중..."):
        # ✅ ignore_cleanup_errors=True : 임시폴더 삭제 시 잠금 남아도 에러 방지
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            tmp_dir = Path(tmp)

            # 템플릿 경로 결정(기본 -> 업로드 있으면 덮어쓰기)
            final_khd_tpl = DEFAULT_KHD_TPL
            final_wph_tpl = DEFAULT_WPH_TPL

            if st.session_state.get("tpl_khd") is not None:
                final_khd_tpl = save_uploaded_to_temp(st.session_state["tpl_khd"], tmp_dir)
            if st.session_state.get("tpl_wph") is not None:
                final_wph_tpl = save_uploaded_to_temp(st.session_state["tpl_wph"], tmp_dir)

            templates = {"KHD": final_khd_tpl, "WPH": final_wph_tpl}

            created_paths = []
            for rf in raw_files:
                raw_path = save_uploaded_to_temp(rf, tmp_dir)

                created = make_results_for_input(
                    raw_path,
                    templates=templates,
                    output_dir=str(tmp_dir),
                    raw_end_row=raw_end_row
                )
                created_paths.extend(created)

                # ✅ engine 내부에서 openpyxl/pandas 핸들이 남을 수 있어 GC로 한번 정리
                gc.collect()

            # 결과 파일을 bytes로 보관
            all_created_bytes = []
            for p in created_paths:
                p_path = Path(p)
                data = safe_read_bytes(p_path)
                all_created_bytes.append((p_path.name, data))

            # ZIP bytes 생성
            zip_path = tmp_dir / "SLB_MES_Result_Package.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in created_paths:
                    zf.write(p, arcname=Path(p).name)

            zip_bytes = safe_read_bytes(zip_path)

            # ✅ 마지막으로 한번 더 정리
            gc.collect()

            st.session_state["results"] = all_created_bytes
            st.session_state["zip_bytes"] = zip_bytes

    st.success("완료! 아래에서 결과 파일을 다운로드하세요.")

# =========================
# 결과 표시(세션 상태 기반)
# =========================
if st.session_state["results"]:
    st.subheader("개별 결과 파일")
    for filename, data in st.session_state["results"]:
        st.download_button(
            label=f"⬇️ {filename}",
            data=data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl-{filename}"
        )

    st.subheader("전체 ZIP")
    st.download_button(
        label="⬇️ 전체 결과 ZIP 다운로드",
        data=st.session_state["zip_bytes"],
        file_name="SLB_MES_Result_Package.zip",
        mime="application/zip",
        key="dl-zip"
    )
else:
    st.info("원본을 선택하고 실행을 누르면 결과가 표시됩니다.")
