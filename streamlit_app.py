"""부서 공통 문제 해결 제안서 작성기 (Streamlit Community Cloud 배포용 화면).

실제 작성 로직은 proposal_agent.py 를 그대로 사용한다.
API 키와 접속 암호는 코드가 아니라 Streamlit Secrets 에서만 읽는다.
"""

from __future__ import annotations

import hmac
from pathlib import Path

import streamlit as st

import proposal_agent as pa

MAX_FILES = 10
MAX_FILE_BYTES = 200_000
SAMPLE_DIR = Path(__file__).resolve().parent / "한국표준협회_주간보고"

st.set_page_config(page_title="공통 문제 해결 제안서", page_icon="📝", layout="centered")


def secret(name: str) -> str | None:
    try:
        return st.secrets.get(name)
    except Exception:  # secrets 미설정
        return None


# ---------------------------------------------------------------------------
# 접속 암호 (공개 주소로 배포되므로, 아무나 API 비용을 쓰지 못하게 막는다)
# ---------------------------------------------------------------------------

api_key = secret("ANTHROPIC_API_KEY")
app_password = secret("APP_PASSWORD")

st.title("📝 부서 공통 문제 해결 제안서")
st.caption("여러 부서의 주간보고에서 공통 문제를 찾아 해결 방안을 제안합니다. 원문에 없는 수치·예산·담당자·확정 일정은 만들지 않고, 근거가 부족하면 '확인 필요'로 표시합니다.")

if not api_key or not app_password:
    st.error("앱 설정(Secrets)에 ANTHROPIC_API_KEY 와 APP_PASSWORD 가 필요합니다. 배포 안내를 확인해 주세요.")
    st.stop()

if not st.session_state.get("authed"):
    with st.form("login"):
        typed = st.text_input("접속 암호", type="password")
        if st.form_submit_button("입장"):
            if hmac.compare_digest(typed.encode(), app_password.encode()):
                st.session_state["authed"] = True
                st.rerun()
            else:
                st.error("암호가 맞지 않습니다.")
    st.stop()

# ---------------------------------------------------------------------------
# 입력
# ---------------------------------------------------------------------------

st.subheader("1. 주간보고 불러오기")
uploaded = st.file_uploader(
    "부서별 주간보고 파일 (.md / .txt, 여러 개 선택 가능)",
    type=["md", "txt"],
    accept_multiple_files=True,
)
use_sample = st.checkbox("업로드 대신 연습용 예시 보고서 4건 사용", value=not uploaded, disabled=bool(uploaded))


def load_reports() -> list[pa.Report]:
    reports = []
    if uploaded:
        if len(uploaded) > MAX_FILES:
            raise pa.AgentError(f"파일은 한 번에 {MAX_FILES}개까지 처리할 수 있습니다.")
        for f in uploaded:
            data = f.getvalue()
            if len(data) > MAX_FILE_BYTES:
                raise pa.AgentError(f"'{f.name}' 파일이 너무 큽니다 (최대 {MAX_FILE_BYTES // 1000}KB).")
            try:
                text = data.decode("utf-8")
            except UnicodeDecodeError:
                raise pa.AgentError(f"'{f.name}' 파일을 UTF-8 텍스트로 읽을 수 없습니다.")
            reports.append(pa.report_from_text(f.name, text))
    elif use_sample:
        for path in sorted(SAMPLE_DIR.glob(pa.DEFAULT_PATTERN)):
            reports.append(pa.read_report(path))
        if not reports:
            raise pa.AgentError("예시 보고서를 찾지 못했습니다. 파일을 직접 업로드해 주세요.")
    else:
        raise pa.AgentError("주간보고 파일을 업로드하거나 예시 보고서를 선택해 주세요.")
    return reports


# ---------------------------------------------------------------------------
# 생성
# ---------------------------------------------------------------------------

st.subheader("2. 제안서 만들기")
if st.button("제안서 생성", type="primary"):
    st.session_state.pop("result", None)
    try:
        reports = load_reports()
        with st.status(f"보고서 {len(reports)}건으로 제안서를 만드는 중… (1~2분 걸릴 수 있습니다)", expanded=True) as status:
            document, tag = pa.generate_proposal(
                reports, api_key, pa.DEFAULT_MODEL, 16000, log=lambda m: st.write(m)
            )
            status.update(label="완료", state="complete", expanded=False)
        st.session_state["result"] = (document, tag)
    except pa.AgentError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"예기치 못한 문제로 처리를 중단했습니다: {pa.redact(str(e))}")

if "result" in st.session_state:
    document, tag = st.session_state["result"]
    st.download_button(
        "제안서 다운로드 (.md)",
        data=document.encode("utf-8"),
        file_name=f"공통문제_해결제안서_{tag}.md",
        mime="text/markdown",
    )
    st.divider()
    st.markdown(document)
