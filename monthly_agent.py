"""
표준보급팀 담당자별 월간 실적보고(xlsx) 여러 개를 읽어
Claude API로 하나의 "팀 월간 실적 보고서 초안"을 생성한다.

실행 예:
    python monthly_agent.py
    python monthly_agent.py --input-dir 표준보급팀_영업실적보고 --output-dir 표준보급팀_영업실적보고

규칙:
    - 원문(엑셀 파일)에 없는 사실/수치를 만들지 않는다. (Claude에게 시스템 프롬프트로 강제)
    - 담당자/기한이 불명확하면 "확인 필요"로 표기한다.
    - 여러 보고서에 등장하는 동일 사안은 하나로 묶고 관련 담당자·출처 파일을 명시한다.
    - 모든 보고서의 주요 내용이 누락되지 않았는지 프로그램이 자동으로 점검하고,
      누락이 발견되면 1회 보완 요청을 한다.
    - 처리 중 오류가 발생하면 결과 파일을 만들지 않고 한국어로 이유만 출력한다.
    - API 키는 .env 파일에서만 읽고, 코드/화면(로그, 예외 메시지)에 노출하지 않는다.
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = BASE_DIR / "표준보급팀_영업실적보고"
DEFAULT_PATTERN = "*_영업실적보고_*.xlsx"
DEFAULT_MODEL = "claude-sonnet-5-5"

KEY_PATTERN = re.compile(r"sk-ant-[A-Za-z0-9_\-]+")


class AgentError(Exception):
    """이 프로그램 내에서 사용자에게 보여줄, 사유가 명확한 오류."""


def redact(text: str) -> str:
    """혹시라도 예외 메시지에 API 키 형태 문자열이 섞여 있으면 가린다."""
    return KEY_PATTERN.sub("[REDACTED]", text)


# ---------------------------------------------------------------------------
# 0. API 키 로딩 (파일에서만 읽고 화면/코드에 노출하지 않음)
# ---------------------------------------------------------------------------

def load_api_key(env_path: Path) -> str:
    try:
        from dotenv import load_dotenv
    except ImportError as e:
        raise AgentError(
            "python-dotenv 패키지가 설치되어 있지 않습니다. "
            "'pip install python-dotenv anthropic openpyxl' 을 먼저 실행해 주세요."
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
# 1. 엑셀 수식 해석기 (SUM / IFERROR / 시트간 참조 만 지원, 그 외는 원문 보존)
# ---------------------------------------------------------------------------

class FormulaResolver:
    """openpyxl로 읽은 워크북에서 이 프로젝트가 사용하는 단순한 수식
    (SUM, IFERROR, 시트간 참조, 사칙연산)만 결정론적으로 계산한다.
    해석 불가능한 수식은 값을 지어내지 않고 원문 수식 문자열을 그대로 남긴다."""

    CELL_REF = re.compile(r"\$?[A-Za-z]{1,3}\$?\d+")

    def __init__(self, wb):
        self.wb = wb
        self.cache: dict[tuple[str, str], object] = {}

    def raw(self, sheet: str, cell: str):
        return self.wb[sheet][cell].value

    def resolve(self, sheet: str, cell: str, _stack: frozenset = frozenset()):
        key = (sheet, cell)
        if key in self.cache:
            return self.cache[key]
        if key in _stack:
            return None  # 순환 참조 방지
        val = self.raw(sheet, cell)
        if isinstance(val, str) and val.startswith("="):
            resolved = self._eval_formula(sheet, val[1:].strip(), _stack | {key})
        else:
            resolved = val
        self.cache[key] = resolved
        return resolved

    def _enum_range(self, rng: str):
        from openpyxl.utils import range_boundaries, get_column_letter
        min_col, min_row, max_col, max_row = range_boundaries(rng)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                yield f"{get_column_letter(c)}{r}"

    def _eval_formula(self, sheet: str, expr: str, _stack: frozenset):
        m = re.match(r"^'?([^'!]+)'?!(\$?[A-Za-z]+\$?\d+)$", expr)
        if m:
            other_sheet, cellref = m.group(1), m.group(2).replace("$", "")
            if other_sheet in self.wb.sheetnames:
                return self.resolve(other_sheet, cellref, _stack)
            return f"[수식 미해석: ={expr}]"

        m = re.match(r"^SUM\(([^)]+)\)$", expr, re.IGNORECASE)
        if m:
            total, has_val = 0, False
            for c in self._enum_range(m.group(1).replace("$", "")):
                v = self.resolve(sheet, c, _stack)
                if isinstance(v, (int, float)):
                    total += v
                    has_val = True
            return total if has_val else None

        m = re.match(r"^IFERROR\((.+),\s*([^,]+)\)$", expr, re.IGNORECASE)
        if m:
            inner_expr, default = m.group(1), m.group(2).strip()
            try:
                val = self._eval_arith(sheet, inner_expr, _stack)
                if val is None:
                    raise ValueError
                return val
            except Exception:
                try:
                    return float(default) if "." in default else int(default)
                except Exception:
                    return default

        try:
            val = self._eval_arith(sheet, expr, _stack)
            if val is None:
                raise ValueError
            return val
        except Exception:
            return f"[수식 미해석: ={expr}]"

    def _eval_arith(self, sheet: str, expr: str, _stack: frozenset):
        def repl(mo):
            ref = mo.group(0).replace("$", "")
            v = self.resolve(sheet, ref, _stack)
            if not isinstance(v, (int, float)):
                raise ValueError("셀 값이 숫자가 아닙니다.")
            return repr(v)

        substituted = self.CELL_REF.sub(repl, expr)
        if not re.match(r"^[\d\.\+\-\*/\(\)\s]+$", substituted):
            raise ValueError("허용되지 않은 수식 형태입니다.")
        return eval(substituted, {"__builtins__": {}}, {})  # noqa: S307 (검증된 숫자식만 허용)


# ---------------------------------------------------------------------------
# 2. 개별 보고서(xlsx) 추출
# ---------------------------------------------------------------------------

EXPECTED_SHEETS = ["보고개요", "이번달_누계실적", "예상실적_연간", "이번달_영업계획", "이슈"]


@dataclass
class ReportProfile:
    file_name: str
    rep_name: str
    period: str
    text_block: str
    issue_titles: list = field(default_factory=list)


def _row_texts(ws, resolver, start_row: int, max_col: int):
    """헤더 이후의 데이터 행들을, 비어 있지 않은 셀만 모아 텍스트 줄로 변환한다."""
    lines = []
    for row in ws.iter_rows(min_row=start_row, max_col=max_col):
        values = []
        for c in row:
            v = resolver.resolve(ws.title, c.coordinate)
            if v not in (None, ""):
                values.append(str(v))
        if values:
            lines.append(" | ".join(values))
    return lines


def extract_report(path: Path) -> ReportProfile:
    import openpyxl

    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except Exception as e:
        raise AgentError(f"'{path.name}' 파일을 여는 중 오류가 발생했습니다: {redact(str(e))}")

    resolver = FormulaResolver(wb)
    has_expected_sheets = all(s in wb.sheetnames for s in EXPECTED_SHEETS)

    rep_name = "확인 필요"
    period = "확인 필요"
    issue_titles: list[str] = []
    sections: list[str] = []

    if has_expected_sheets:
        # 보고개요
        ws = wb["보고개요"]
        overview_lines = []
        for row in ws.iter_rows(min_row=3, max_row=8, max_col=3):
            label = row[0].value
            val = resolver.resolve(ws.title, row[1].coordinate) if len(row) > 1 else None
            if label and val not in (None, ""):
                overview_lines.append(f"{label}: {val}")
                if label == "담당자":
                    rep_name = str(val)
                if label == "보고 기간":
                    period = str(val)
        sections.append("[보고개요]\n" + "\n".join(overview_lines))

        # 이번달 누계실적
        ws = wb["이번달_누계실적"]
        lines = ["주차 | 목표 | 실적 | 달성률"] + _row_texts(ws, resolver, start_row=4, max_col=4)
        sections.append("[이번달 누계 실적]\n" + "\n".join(lines))

        # 예상실적(연간)
        ws = wb["예상실적_연간"]
        lines = ["월 | 구분 | 금액"] + _row_texts(ws, resolver, start_row=4, max_col=3)
        sections.append("[다음달부터 예상실적 / 연간 예상실적]\n" + "\n".join(lines))

        # 영업계획
        ws = wb["이번달_영업계획"]
        lines = ["No | 추진과제 | 목표 | 일정 | 진척현황"] + _row_texts(ws, resolver, start_row=4, max_col=5)
        sections.append("[이번달 영업계획]\n" + "\n".join(lines))

        # 이슈
        ws = wb["이슈"]
        issue_rows = _row_texts(ws, resolver, start_row=4, max_col=6)
        lines = ["No | 이슈 | 내용 | 영향도 | 대응방안 | 상태"] + issue_rows
        sections.append("[이슈]\n" + "\n".join(lines))
        for r in issue_rows:
            parts = [p.strip() for p in r.split("|")]
            if len(parts) >= 2 and parts[1]:
                issue_titles.append(parts[1])
    else:
        # 알 수 없는 구조 -> 모든 시트를 그대로 텍스트로 덤프 (사실을 만들어내지 않기 위한 대체 경로)
        for ws in wb.worksheets:
            lines = _row_texts(ws, resolver, start_row=1, max_col=ws.max_column or 6)
            sections.append(f"[{ws.title}]\n" + "\n".join(lines))

    text_block = f"### 출처 파일: {path.name} (담당자: {rep_name})\n" + "\n\n".join(sections)
    return ReportProfile(
        file_name=path.name,
        rep_name=rep_name,
        period=period,
        text_block=text_block,
        issue_titles=issue_titles,
    )


# ---------------------------------------------------------------------------
# 3. Claude 호출
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """당신은 한국표준협회 표준보급팀의 월별 팀 실적 보고서 초안을 작성하는 보조자입니다.
아래 각 담당자 보고서 원문(엑셀에서 추출한 텍스트)만을 근거로 팀 보고서를 작성하세요.

반드시 지켜야 할 규칙:
1. 원문에 없는 사실이나 수치를 만들어내지 마세요. 원문에 있는 내용만 사용하세요.
2. 담당자나 기한이 원문에서 분명하지 않으면 "확인 필요"라고 표시하세요.
3. 같은 사안(이슈, 계획 등)이 여러 보고서에 등장하면 하나로 통합하고, 관련 담당자 이름과
   출처 파일명을 함께 적으세요. 예: (관련 담당자: 김민준, 이서연 / 출처: 김민준_영업실적보고_2026년9월.xlsx, 이서연_영업실적보고_2026년9월.xlsx)
4. 제공된 모든 보고서의 주요 내용(실적, 계획, 이슈)이 하나도 빠짐없이 반영되도록 하세요.
5. 아래 5개 항목을 반드시 이 순서, 이 제목 그대로 사용하여 작성하세요.

## 1. 핵심 요약 3줄
## 2. 이번달 팀 주요 실적
## 3. 팀에서 함께 확인할 이슈와 담당·기한
## 4. 팀장이 결정해야 할 사항
## 5. 다음 달 주요 일정

마크다운 형식으로 작성하고, 각 항목 아래에는 근거가 된 담당자/출처 파일을 괄호로 표기하세요."""


def call_claude(client, model: str, system: str, user_content: str, max_tokens: int = 8000) -> str:
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
            "결과가 불완전할 수 있어 파일을 저장하지 않았습니다. "
            "--max-tokens 값을 늘리거나 입력 보고서 수를 줄여 다시 실행해 주세요."
        )

    texts = [block.text for block in message.content if getattr(block, "type", None) == "text"]
    if not texts:
        raise AgentError("Claude 응답에서 텍스트 내용을 찾을 수 없습니다.")
    return "\n".join(texts)


def build_user_content(profiles: list[ReportProfile]) -> str:
    joined = "\n\n---\n\n".join(p.text_block for p in profiles)
    return f"다음은 이번 달 담당자별 보고서 원문입니다 (총 {len(profiles)}건).\n\n{joined}"


def _title_covered(title: str, draft: str) -> bool:
    """제목이 초안에 그대로 없어도, 구성 단어의 절반 이상이 등장하면 다룬 것으로 간주한다
    (Claude가 표현을 바꿔 쓰는 경우 오탐을 줄이기 위함)."""
    if title in draft:
        return True
    tokens = [t for t in re.split(r"[^0-9A-Za-z가-힣]+", title) if len(t) >= 2]
    if not tokens:
        return title in draft
    found = sum(1 for t in tokens if t in draft)
    return found / len(tokens) >= 0.5


def check_completeness(draft: str, profiles: list[ReportProfile]) -> list[str]:
    """담당자 이름과 이슈 제목이 초안에 언급되었는지 기계적으로 점검한다."""
    missing = []
    for p in profiles:
        if p.rep_name != "확인 필요" and p.rep_name not in draft:
            missing.append(f"{p.file_name} (담당자: {p.rep_name})의 내용이 초안에서 확인되지 않습니다.")
        for title in p.issue_titles:
            if title and not _title_covered(title, draft):
                missing.append(f"{p.file_name}의 이슈 '{title}'가 초안에서 확인되지 않습니다.")
    return missing


def repair_draft(client, model: str, draft: str, missing: list[str], user_content: str, max_tokens: int) -> str:
    repair_prompt = (
        "아래는 방금 작성한 팀 월간 실적 보고서 초안입니다.\n\n"
        f"{draft}\n\n---\n\n"
        "그런데 아래 항목들이 원문 보고서에는 있지만 초안에서 누락된 것으로 보입니다:\n"
        + "\n".join(f"- {m}" for m in missing)
        + "\n\n원문 데이터(아래)를 다시 참고하여, 위 누락 항목을 알맞은 섹션에 보완한 "
        "전체 보고서를 처음부터 끝까지 다시 작성하세요. 기존 규칙(원문에 없는 내용 금지, "
        "담당자/기한 불명확 시 '확인 필요', 5개 항목 제목 그대로 유지)을 계속 지키세요.\n\n"
        f"[원문 데이터]\n{user_content}"
    )
    return call_claude(client, model, SYSTEM_PROMPT, repair_prompt, max_tokens=max_tokens)


# ---------------------------------------------------------------------------
# 4. 메인 파이프라인
# ---------------------------------------------------------------------------

def run(input_dir: Path, pattern: str, output_dir: Path, model: str, env_path: Path, max_tokens: int) -> tuple[int, Path]:
    api_key = load_api_key(env_path)  # 화면/로그에 출력하지 않음

    if not input_dir.exists():
        raise AgentError(f"입력 폴더를 찾을 수 없습니다: {input_dir}")

    all_files = sorted(input_dir.glob(pattern))
    if not all_files:
        raise AgentError(
            f"'{input_dir}' 폴더에서 '{pattern}' 패턴과 일치하는 보고서 파일을 찾지 못했습니다."
        )

    print(f"[안내] 대상 폴더: {input_dir}")
    print(f"[안내] 패턴 '{pattern}'으로 {len(all_files)}개 파일을 찾았습니다.")
    for f in all_files:
        print(f"  - {f.name}")

    profiles = [extract_report(f) for f in all_files]

    from anthropic import Anthropic
    client = Anthropic(api_key=api_key)

    user_content = build_user_content(profiles)

    print("[안내] Claude API로 초안을 생성합니다...")
    draft = call_claude(client, model, SYSTEM_PROMPT, user_content, max_tokens=max_tokens)

    print("[안내] 모든 보고서 내용이 반영되었는지 점검합니다...")
    missing = check_completeness(draft, profiles)
    if missing:
        print(f"[안내] 누락 가능성이 있는 항목 {len(missing)}건을 발견하여 보완 요청을 진행합니다.")
        for m in missing:
            print(f"  - {m}")
        draft = repair_draft(client, model, draft, missing, user_content, max_tokens=max_tokens)
        remaining = check_completeness(draft, profiles)
        if remaining:
            print("[경고] 보완 후에도 다음 항목이 초안에서 명시적으로 확인되지 않았습니다 (본문을 직접 확인해 주세요):")
            for m in remaining:
                print(f"  - {m}")
    else:
        print("[안내] 모든 담당자/이슈가 초안에서 확인되었습니다.")

    periods = {p.period for p in profiles if p.period != "확인 필요"}
    period_tag = periods.pop() if len(periods) == 1 else "기간확인필요"
    if len(periods) > 1:
        print(f"[경고] 보고서들의 '보고 기간'이 서로 다릅니다: {periods | {period_tag}}")

    safe_tag = re.sub(r"[^0-9A-Za-z가-힣]", "", period_tag)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"팀_월간실적보고_초안_{safe_tag}.md"

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    header = (
        f"# 표준보급팀 월간 실적 보고서 (초안)\n\n"
        f"- 대상 보고 기간: {period_tag}\n"
        f"- 초안 생성 일시: {generated_at}\n"
        f"- 사용 모델: {model}\n"
        f"- 참고한 담당자 보고서 ({len(profiles)}건): "
        + ", ".join(p.file_name for p in profiles)
        + "\n\n"
        "※ 이 문서는 담당자별 보고서를 근거로 AI가 자동 작성한 초안입니다. "
        "팀장 검토 후 확정해 주세요.\n\n---\n\n"
    )

    # 완전히 만들어진 뒤에만 파일을 기록한다 (오류 시 파일을 만들지 않기 위함).
    output_path.write_text(header + draft, encoding="utf-8")

    return len(profiles), output_path


def parse_args():
    parser = argparse.ArgumentParser(description="담당자 월간 보고서를 팀 월간 보고서 초안으로 통합한다.")
    parser.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT_DIR, help="담당자 보고서(xlsx)가 있는 폴더")
    parser.add_argument("--pattern", default=DEFAULT_PATTERN, help="담당자 보고서 파일명 패턴 (glob)")
    parser.add_argument("--output-dir", type=Path, default=None, help="결과 파일을 저장할 폴더 (기본: input-dir)")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="사용할 Claude 모델 ID")
    parser.add_argument("--env-file", type=Path, default=BASE_DIR / ".env", help="API 키가 저장된 .env 파일 경로")
    parser.add_argument("--max-tokens", type=int, default=8000, help="Claude 응답 max_tokens 값")
    parser.add_argument(
        "--no-pause", action="store_true",
        help="종료 시 'Enter 키를 눌러 종료' 대기를 건너뛴다 (자동화/스크립트에서 실행할 때 사용)",
    )
    return parser.parse_args()


def _pause():
    """탐색기에서 더블클릭으로 실행했을 때 콘솔 창이 바로 닫혀버리지 않도록 대기한다."""
    try:
        input("\n창을 닫으려면 Enter 키를 누르세요...")
    except EOFError:
        pass


def main():
    args = parse_args()
    output_dir = args.output_dir or args.input_dir
    exit_code = 0
    try:
        count, output_path = run(args.input_dir, args.pattern, output_dir, args.model, args.env_file, args.max_tokens)
        print(f"\n[완료] 읽은 담당자 보고서 파일 수: {count}건")
        print(f"[완료] 결과 저장 위치: {output_path}")
    except AgentError as e:
        print(f"[오류] {e}")
        exit_code = 1
    except Exception as e:  # 예상치 못한 오류도 원인만 안내하고 종료
        print(f"[오류] 예기치 못한 문제로 처리를 중단했습니다: {redact(str(e))}")
        exit_code = 1

    if not args.no_pause:
        _pause()
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
