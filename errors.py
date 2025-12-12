# errors.py
# 공통 사용자 친화 에러 정의 (engine.py / summary_engine.py / dashboard_engine.py / app.py에서 공유)

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional


@dataclass
class UserFacingError(Exception):
    """
    사용자에게 그대로 보여도 되는 형태로 구조화된 예외.

    title : 한 줄 요약 (Streamlit st.error에 바로 표시)
    detail: 자세한 내용 (파일명/시트명/키 값/현재 목록 등)
    hint  : 사용자가 다음에 뭘 확인하면 되는지(행동 가이드)
    code  : 내부적으로 분류용(선택)
    context: 디버깅/로그용 추가 컨텍스트(선택) - UI에는 detail로 합쳐서 보여주거나,
             필요 시 app.py에서 expander로 별도 표시 가능
    """

    title: str
    detail: str = ""
    hint: str = ""
    code: str = ""
    context: Optional[Dict[str, Any]] = None

    def __str__(self) -> str:
        # 기본 str(e)가 title로 깔끔히 나오게
        return self.title

    def to_dict(self) -> Dict[str, Any]:
        return {
            "title": self.title,
            "detail": self.detail,
            "hint": self.hint,
            "code": self.code,
            "context": self.context or {},
        }


# 아래는 자주 쓰는 “팩토리” 헬퍼들(선택). 필요 없으면 지워도 됩니다.

def missing_sheet_error(
    *,
    title: str = "필요한 시트를 찾을 수 없습니다.",
    needed: str,
    available: Optional[list[str]] = None,
    hint: str = "시트명이 변경되었는지 확인하세요.",
    code: str = "MISSING_SHEET",
    context: Optional[Dict[str, Any]] = None,
) -> UserFacingError:
    detail_lines = [f"필요 시트: {needed}"]
    if available is not None:
        detail_lines.append(f"현재 시트 목록: {available}")
    return UserFacingError(
        title=title,
        detail="\n".join(detail_lines),
        hint=hint,
        code=code,
        context=context,
    )


def missing_column_error(
    *,
    needed: str | list[str],
    available: Optional[list[str]] = None,
    title: str = "필요한 컬럼을 찾을 수 없습니다.",
    hint: str = "원본 파일 포맷(컬럼명)이 맞는지 확인하세요.",
    code: str = "MISSING_COLUMN",
    context: Optional[Dict[str, Any]] = None,
) -> UserFacingError:
    needed_list = [needed] if isinstance(needed, str) else needed
    detail_lines = [f"필요 컬럼: {needed_list}"]
    if available is not None:
        detail_lines.append(f"현재 컬럼 목록: {available}")
    return UserFacingError(
        title=title,
        detail="\n".join(detail_lines),
        hint=hint,
        code=code,
        context=context,
    )


def invalid_zip_error(
    *,
    title: str = "ZIP 파일 형식이 올바르지 않습니다.",
    detail: str = "",
    hint: str = "Result 패키지 ZIP을 업로드했는지 확인하세요.",
    code: str = "INVALID_ZIP",
    context: Optional[Dict[str, Any]] = None,
) -> UserFacingError:
    return UserFacingError(
        title=title,
        detail=detail,
        hint=hint,
        code=code,
        context=context,
    )
