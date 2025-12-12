# ui_error.py
# Streamlit에서 에러를 “사람이 읽기 쉬운 형태”로 통일해서 보여주기 위한 공통 UI 함수

from __future__ import annotations

import traceback
from typing import Optional

import streamlit as st

from errors import UserFacingError


def show_error(
    e: Exception,
    *,
    title_prefix: str = "❗ ",
    show_dev_traceback: bool = True,
    dev_traceback_label: str = "개발자용 상세 로그",
    detail_label: str = "자세한 내용",
    context_label: str = "추가 정보",
) -> None:
    """
    - UserFacingError: title/detail/hint/context를 정돈된 UI로 표시
    - 그 외 Exception: 일반 에러 + (옵션) traceback expander 표시
    """

    if isinstance(e, UserFacingError):
        st.error(f"{title_prefix}{e.title}")

        if e.detail:
            with st.expander(detail_label):
                st.code(e.detail)

        # context는 detail과 분리해서 보고 싶을 때 유용 (로그성 정보)
        if e.context:
            with st.expander(context_label):
                st.json(e.context)

        if e.hint:
            st.info(f"💡 힌트: {e.hint}")

        return

    # 일반 예외 처리
    st.error(f"{title_prefix}처리 중 오류가 발생했습니다: {e}")

    if show_dev_traceback:
        with st.expander(dev_traceback_label):
            st.code(traceback.format_exc())


def run_with_ui_error(
    fn,
    *args,
    spinner_text: Optional[str] = None,
    **kwargs,
):
    """
    Streamlit 버튼 콜백에서 자주 쓰는 패턴:
      - 스피너(선택)
      - 예외를 show_error로 통일 출력
    사용 예:
      result = run_with_ui_error(engine.make_results_for_input, ..., spinner_text="생성 중...")
      if result is None: st.stop()
    """
    try:
        if spinner_text:
            with st.spinner(spinner_text):
                return fn(*args, **kwargs)
        return fn(*args, **kwargs)
    except Exception as e:
        show_error(e)
        return None
