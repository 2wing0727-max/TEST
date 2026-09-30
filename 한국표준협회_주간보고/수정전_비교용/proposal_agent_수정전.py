"""
여러 부서의 주간보고(md) 를 읽어 부서 간 공통 문제를 찾고,
해결 방안을 담은 "공통 문제 해결 제안서"를 Claude API로 생성한다.

실행 예:
    python proposal_agent.py
    python proposal_agent.py --input-dir 한국표준협회_주간보고 --output-dir 한국표준협회_주간보고

제안서 구성(순서 고정):
    1. 현재 상황 / 2. 확인된 문제 / 3. 해결 방안 제안 / 4. 기대 효과 / 5. 필요한 것과 일정

규칙:
    - 문제의 근거로 보고서 이름과 원문 내용을 적는다.
    - 새로운 해결 방안은 "[제안]"으로 분명히 표시한다.
    - 원문에 없는 효과 수치, 예산, 담당자, 확정 일정은 만들지 않는다.
    - 근거가 부족한 내용은 "확인 필요"로 표시한다.
    - 오류 시 결과 파일을 만들지 않고 한국어로 이유만 출력한다.
    - API 키는 .env 파일에서만 읽고, 화면/로그/예외 메시지에 노출하지 않는다.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = BASE_DIR / "한국표준협회_주간보고"
DEFAULT_PATTERN = "*_주간보고_*.md"
DEFAULT_MODEL = "claude-sonnet-5-5"

KEY_PATTERN = re.compile(r"sk-ant-[A-Za-z0-9_\-]+")

REQUIRED_HEADINGS = [
    "## 1. 현재 상황",
    "## 2. 확인된 문제",
    "## 3. 해결 방안 제안",
    "## 4. 기대 효과",
    "## 5. 필요한 것과 일정",
]


class AgentError(Exception):
    """사용자에게 보여줄, 사유가 명확한 오류."""


def redact(text: str) -> str:
    return KEY_PATTERN.sub("[REDACTED]", text)


def load_api_key(env_path: Path) -> str:
    try:
        from dotenv import load_dotenv
    except ImportError as e:
        raise AgentError(
            "python-dotenv 패키지가 설치되어 있지 않습니다. "
            "'pip install python-dotenv anthropic' 을 먼저 실행해 주세요."
        ) from e

    load_dotenv(dotenv_path=env_path)
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise AgentError(
            f"{env_path} 파일에 ANTHROPIC_API_KEY가 설정되어 있지 않습니다. "
            "API 키를 .env 파일에 저장한 뒤 다시 실행해 주세요."
        )
    return api_key


# ---------------------------------------------------------------------------
# 1. 보고서 읽기
# ---------------------------------------------------------------------------

@dataclass
class Report:
    file_name: str
    team: str
    period: str
    text: str


def read_report(path: Path) -> Report:
    try:
        text = path.read_text(encoding="utf-8")
    except Exception as e:
        raise AgentError(f"'{path.name}' 파일을 읽는 중 오류가 발생했습니다: {redact(str(e))}")

    m = re.search(r"작성팀:\s*(.+)", text)
    team = m.group(1).strip() if m else "확인 필요"
    m = re.search(r"보고기간:\s*(.+)", text)
    period = m.group(1).strip() if m else "확인 필요"
    return Report(file_name=path.name, team=team, period=period, text=text)


# ---------------------------------------------------------------------------
# 2. Claude 호출
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """당신은 한국표준협회 내부에서 여러 부서의 주간보고를 읽고 부서 간 공통 문제를 찾아
해결 방안을 제안하는 제안서 작성 보조자입니다. 아래 보고서 원문만을 근거로 작성하세요.

반드시 지켜야 할 규칙:
1. 제안서는 아래 5개 항목을 이 순서, 이 제목 그대로 사용하세요.

## 1. 현재 상황
## 2. 확인된 문제
## 3. 해결 방안 제안
## 4. 기대 효과
## 5. 필요한 것과 일정

2. '확인된 문제'의 각 문제마다 근거가 되는 보고서 이름(파일명)과 그 보고서의 원문 내용을 적으세요.
   원문 내용은 원문 표현을 살려 적고, 여러 부서 보고서에 걸쳐 나타나는 공통 문제를 우선 다루세요.
   한 부서에서만 나온 문제는 '단일 부서 이슈'로 구분해 표시하세요.
3. 원문에 없는 새로운 해결 방안은 반드시 문장 앞에 "[제안]"을 붙여 제안임을 분명히 표시하세요.
   원문에 이미 적힌 대응(예: 협의 중, 재테스트 예정)은 "[원문 기재]"로 구분하세요.
4. 원문에 없는 효과 수치(%, 건수 감소량, 시간 단축 등), 예산·비용, 담당자 이름, 확정 일정을 만들지 마세요.
   일정은 원문에 있는 날짜만 인용하고 출처 보고서를 밝히세요. 나머지는 "확인 필요"로 적으세요.
   기대 효과는 수치 없이 정성적으로만 서술하세요.
5. 근거가 부족하거나 원문에서 알 수 없는 내용(담당자, 예산, 세부 일정, 원인 등)은 "확인 필요"라고 표시하세요.
6. 모든 보고서의 '이슈'가 하나도 빠짐없이 '확인된 문제'에 다뤄지도록 하세요.
7. 마크다운으로 작성하고, 제목(# ...)은 쓰지 말고 위 5개 '## ' 항목으로만 시작하세요."""


def call_claude(client, model: str, system: str, user_content: str, max_tokens: int) -> str:
    try:
        message = client.messages.create(
            model=model,
            max_tokens=max_tokens,
            system=system,
            messages=[{"role": "user", "content": user_content}],
        )
    except Exception as e:
        raise AgentError(f"Claude API 호출 중 오류가 발생했습니다: {redact(str(e))}")

    if message.stop_reason == "max_tokens":
        raise AgentError(
            "Claude 응답이 max_tokens 한도에 걸려 중간에 잘렸습니다. "
            "결과가 불완전해 파일을 저장하지 않았습니다. --max-tokens 값을 늘려 다시 실행해 주세요."
        )

    texts = [b.text for b in message.content if getattr(b, "type", None) == "text"]
    if not texts:
        raise AgentError("Claude 응답에서 텍스트 내용을 찾을 수 없습니다.")
    return "\n".join(texts)


def build_user_content(reports: list[Report]) -> str:
    blocks = [f"### 보고서 이름: {r.file_name}\n{r.text}" for r in reports]
    return f"다음은 부서별 주간보고 원문입니다 (총 {len(reports)}건).\n\n" + "\n\n---\n\n".join(blocks)


# ---------------------------------------------------------------------------
# 3. 자동 점검 (구조 / 근거 / 지어낸 숫자)
# ---------------------------------------------------------------------------

NUM_PATTERN = re.compile(r"\d[\d,\.]*")


def check_draft(draft: str, reports: list[Report]) -> list[str]:
    problems = []

    # 5개 항목이 정해진 순서로 있는지
    pos = -1
    for h in REQUIRED_HEADINGS:
        idx = draft.find(h)
        if idx == -1:
            problems.append(f"필수 항목 '{h}'이(가) 없습니다.")
        elif idx < pos:
            problems.append(f"항목 '{h}'의 순서가 올바르지 않습니다.")
        else:
            pos = idx

    # 모든 보고서 이름이 근거로 인용되었는지
    for r in reports:
        stem = r.file_name.rsplit(".", 1)[0]
        if stem not in draft and r.file_name not in draft:
            problems.append(f"보고서 '{r.file_name}'이(가) 근거로 인용되지 않았습니다.")

    # 원문에 없는 숫자(효과 수치, 예산 등) 감지
    source = "\n".join(r.text for r in reports)
    source_nums = {n.strip(",.") for n in NUM_PATTERN.findall(source)}
    body = draft
    # 항목 번호(1.~5.)와 [제안] 번호 매김은 제외
    body = re.sub(r"(?m)^\s*(#+\s*)?\d+\.\s", "", body)
    # 파일명 속 날짜 등은 원문 파일명에서도 허용
    source_nums |= {n.strip(",.") for n in NUM_PATTERN.findall(" ".join(r.file_name for r in reports))}
    unknown = sorted({n.strip(",.") for n in NUM_PATTERN.findall(body)} - source_nums - {""})
    if unknown:
        problems.append(
            "원문에 없는 숫자가 제안서에 나타났습니다 (지어낸 수치인지 확인 필요): " + ", ".join(unknown)
        )
    return problems


def repair_draft(client, model, draft, problems, user_content, max_tokens) -> str:
    prompt = (
        f"아래는 방금 작성한 제안서 초안입니다.\n\n{draft}\n\n---\n\n"
        "자동 점검에서 다음 문제가 발견되었습니다:\n"
        + "\n".join(f"- {p}" for p in problems)
        + "\n\n원문(아래)을 다시 참고하여 문제를 고친 전체 제안서를 처음부터 다시 작성하세요. "
        "원문에 없는 숫자는 삭제하거나 '확인 필요'로 바꾸고, 기존 규칙을 모두 지키세요.\n\n"
        f"[원문]\n{user_content}"
    )
    return call_claude(client, model, SYSTEM_PROMPT, prompt, max_tokens)


# ---------------------------------------------------------------------------
# 4. 메인 파이프라인
# ---------------------------------------------------------------------------

def run(input_dir: Path, pattern: str, output_dir: Path, model: str, env_path: Path, max_tokens: int):
    api_key = load_api_key(env_path)

    if not input_dir.exists():
        raise AgentError(f"입력 폴더를 찾을 수 없습니다: {input_dir}")
    files = sorted(input_dir.glob(pattern))
    if not files:
        raise AgentError(f"'{input_dir}' 폴더에서 '{pattern}' 패턴과 일치하는 보고서를 찾지 못했습니다.")

    print(f"[안내] 대상 폴더: {input_dir}")
    print(f"[안내] 보고서 {len(files)}건을 찾았습니다.")
    for f in files:
        print(f"  - {f.name}")

    reports = [read_report(f) for f in files]

    from anthropic import Anthropic
    client = Anthropic(api_key=api_key)
    user_content = build_user_content(reports)

    print("[안내] Claude API로 제안서를 생성합니다...")
    draft = call_claude(client, model, SYSTEM_PROMPT, user_content, max_tokens)

    print("[안내] 구조·근거·수치를 자동 점검합니다...")
    problems = check_draft(draft, reports)
    if problems:
        print(f"[안내] 점검에서 {len(problems)}건 발견하여 1회 보완 요청을 합니다.")
        for p in problems:
            print(f"  - {p}")
        draft = repair_draft(client, model, draft, problems, user_content, max_tokens)
        problems = check_draft(draft, reports)
        if problems:
            print("[경고] 보완 후에도 다음 사항이 남아 있습니다 (본문을 직접 확인해 주세요):")
            for p in problems:
                print(f"  - {p}")
    else:
        print("[안내] 점검을 모두 통과했습니다.")

    periods = {r.period for r in reports if r.period != "확인 필요"}
    period = periods.pop() if len(periods) == 1 else "확인 필요"
    if len(periods) > 1:
        print(f"[경고] 보고서들의 보고기간이 서로 다릅니다: {periods}")
    tag = re.sub(r"[^0-9A-Za-z가-힣~\-]", "", re.sub(r"\([^)]*\)", "", period)) or "기간확인필요"

    header = (
        "# 부서 공통 문제 해결 제안서 (초안)\n\n"
        f"- 대상 보고 기간: {period}\n"
        f"- 생성 일시: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"- 사용 모델: {model}\n"
        f"- 근거 보고서 ({len(reports)}건): " + ", ".join(r.file_name for r in reports) + "\n\n"
        "※ AI가 보고서 원문만을 근거로 작성한 초안입니다. '[제안]' 표시는 원문에 없는 새 제안이며, "
        "'확인 필요'는 원문만으로 알 수 없는 내용입니다. 검토 후 확정해 주세요.\n\n---\n\n"
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"공통문제_해결제안서_{tag}.md"
    output_path.write_text(header + draft, encoding="utf-8")  # 완성 후에만 기록
    return len(reports), output_path


def parse_args():
    p = argparse.ArgumentParser(description="부서 주간보고에서 공통 문제를 찾아 해결 제안서를 만든다.")
    p.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT_DIR, help="주간보고(md)가 있는 폴더")
    p.add_argument("--pattern", default=DEFAULT_PATTERN, help="보고서 파일명 패턴 (glob)")
    p.add_argument("--output-dir", type=Path, default=None, help="결과 저장 폴더 (기본: input-dir)")
    p.add_argument("--model", default=DEFAULT_MODEL, help="사용할 Claude 모델 ID")
    p.add_argument("--env-file", type=Path, default=BASE_DIR / ".env", help="API 키가 저장된 .env 경로")
    p.add_argument("--max-tokens", type=int, default=8000, help="Claude 응답 max_tokens 값")
    return p.parse_args()


def main():
    args = parse_args()
    output_dir = args.output_dir or args.input_dir
    try:
        count, output_path = run(
            args.input_dir, args.pattern, output_dir, args.model, args.env_file, args.max_tokens
        )
        print(f"\n[완료] 읽은 보고서 수: {count}건")
        print(f"[완료] 결과 저장 위치: {output_path}")
    except AgentError as e:
        print(f"[오류] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[오류] 예기치 못한 문제로 처리를 중단했습니다: {redact(str(e))}")
        sys.exit(1)


if __name__ == "__main__":
    main()
