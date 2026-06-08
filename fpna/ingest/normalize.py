"""
fpna.ingest.normalize — 센티넬/단위/스케일/타입 정규화.

핵심 보강(리서치 반영):
- number_format trailing comma = 스케일 신호(`#,##0,`=÷1000 표기, `#,##0,,`=÷1e6).
  한국 재무 엑셀에서 헤더 텍스트보다 신뢰도 높은 스케일 출처.
- 센티넬을 타입 vote 에서 분리(ptype 정수) → '…','-','N/A' 가 숫자열을 텍스트로 오판 안 함.
- 한국 회계 음수표기: 괄호 (1,234), 선행 △/▲, 후행 (-).
- 한글 단위: 천/백만/억/조.
"""
from __future__ import annotations

import datetime as _dt
import re
from collections import Counter

# 센티넬(결측 표시). 정규화하면 None.
SENTINELS = {"", "…", "...", "..", "-", "–", "—", "n/a", "na", "n.a.", ".",
             "x", "x)", "해당없음", "없음", "*", "**", "ns"}

# 한글/기호 스케일 단위 → 곱수
UNIT_SCALE = {
    "천원": 1_000, "천": 1_000,
    "백만원": 1_000_000, "백만": 1_000_000, "백만달러": 1_000_000,
    "억원": 100_000_000, "억": 100_000_000,
    "조원": 1_000_000_000_000, "조": 1_000_000_000_000,
    "원": 1, "krw": 1, "usd": 1, "%": 1,
}

_NUM_RE = re.compile(r"^[\s ]*[(△▲\-+]?[\s]*[\d,]*\.?\d+[\s]*[)%]?[\s]*$")
_PCT_RE = re.compile(r"%\s*$")
_DATE_PATTERNS = [
    re.compile(r"^\d{4}[-/.]\d{1,2}([-/.]\d{1,2})?$"),
    re.compile(r"^\d{4}년\s*\d{1,2}월(\s*\d{1,2}일)?$"),
    re.compile(r"^\d{4}\s*[Q분기]\s*\d?$", re.IGNORECASE),
]


def scale_from_number_format(fmt: str | None) -> int:
    """number_format 의 trailing comma 로 스케일 추정.

    `#,##0,`   → 1000 (천 단위 표기)
    `#,##0,,`  → 1_000_000 (백만)
    포맷 문자열 내 한글 단위("천원","백만")도 보조로 본다.
    """
    if not fmt or fmt == "General":
        return 1
    body = fmt.split(";")[0]
    # 따옴표 안 텍스트 단위 우선
    for unit, mult in UNIT_SCALE.items():
        if unit in body and mult > 1:
            return mult
    # trailing comma 카운트: 마지막 숫자 placeholder 뒤 연속 콤마
    m = re.search(r"[#0]([, ]*)\s*\"?[^\"#0]*\"?\s*$", body)
    if m:
        commas = m.group(1).count(",")
        if commas:
            return 1000 ** commas
    return 1


def parse_unit_label(text: str | None) -> tuple[str | None, int]:
    """'(단위: 천원)' 류 텍스트 → (단위명, 스케일곱수)."""
    if not text:
        return None, 1
    s = str(text)
    for unit, mult in sorted(UNIT_SCALE.items(), key=lambda kv: -len(kv[0])):
        if unit in s:
            return unit, mult
    return None, 1


def normalize_value(raw, *, fmt: str | None = None, unit_scale: int = 1):
    """단일 셀 값 정규화.

    반환: (number|str|None, sentinel:str|None, is_negative_paren:bool)
      - 숫자로 해석되면 number (스케일 적용), sentinel=None
      - 센티넬이면 (None, 매칭센티넬, False)
      - 그 외 텍스트는 (정리된 str, None, False)
    """
    if raw is None:
        return None, "", False
    if isinstance(raw, bool):
        return raw, None, False
    if isinstance(raw, (int, float)):
        # ⚠ 저장된 숫자값은 이미 실수치다. number_format 의 trailing comma 는
        # '표시 축약'일 뿐이므로 값에 곱하면 이중계상 → 곱하지 않는다.
        # 스케일/단위는 별도 unit 컬럼으로만 보존(consumer 가 해석).
        return raw, None, False
    if isinstance(raw, (_dt.datetime, _dt.date, _dt.time)):
        return raw, None, False

    s = str(raw).replace(" ", " ").strip()
    low = s.lower()
    if low in SENTINELS:
        return None, s, False

    neg = False
    body = s
    # 괄호 음수
    if body.startswith("(") and body.endswith(")"):
        neg = True
        body = body[1:-1].strip()
    # 선행 세모/역삼각 = 음수(한국 회계)
    if body[:1] in ("△", "▲", "−", "-"):
        if body[:1] in ("△", "▲", "−"):
            neg = True
        body = body[1:].strip()
    # 후행 (-)
    if body.endswith("(-)"):
        neg = True
        body = body[:-3].strip()

    is_pct = bool(_PCT_RE.search(body))
    body_num = body.rstrip("%").replace(",", "").strip()
    try:
        num = float(body_num)
        if num.is_integer():
            num = int(num)
        if neg:
            num = -num
        if is_pct:
            return num / 100.0, None, neg
        # 텍스트로 저장된 숫자: 원값 그대로(단위 접미는 unit 컬럼으로 별도 보존)
        return num, None, neg
    except (ValueError, AttributeError):
        return s, None, False


def scale_factor(fmt: str | None, unit_scale: int) -> int:
    """number_format 스케일과 단위 라벨 스케일 중 큰 쪽(둘 다 1이면 1)."""
    fmt_scale = scale_from_number_format(fmt)
    return max(fmt_scale, unit_scale, 1)


# --------------------------------------------------------------------------
# 타입 추론 (센티넬 분리 + 다수결)
# --------------------------------------------------------------------------
def regex_type(s: str) -> str:
    s = str(s).strip()
    for pat in _DATE_PATTERNS:
        if pat.match(s):
            return "DATE"
    if _PCT_RE.search(s):
        return "PCT"
    if _NUM_RE.match(s):
        return "NUM"
    return "TEXT"


def infer_column_type(values: list, *, threshold: float = 0.7) -> str:
    """열 값들의 타입 추론. 센티넬은 vote 에서 제외.

    반환: 'NUM'|'PCT'|'DATE'|'TEXT'|'MIXED'
    """
    votes = Counter()
    nonmiss = 0
    for v in values:
        if v is None:
            continue
        if isinstance(v, (int, float)):
            votes["NUM"] += 1
            nonmiss += 1
            continue
        if isinstance(v, (_dt.datetime, _dt.date)):
            votes["DATE"] += 1
            nonmiss += 1
            continue
        s = str(v).replace(" ", " ").strip()
        if s.lower() in SENTINELS:
            continue  # 센티넬 제외 — 타입 오염 방지
        votes[regex_type(s)] += 1
        nonmiss += 1
    if nonmiss == 0:
        return "TEXT"
    top, cnt = votes.most_common(1)[0]
    return top if cnt / nonmiss >= threshold else "MIXED"


__all__ = [
    "SENTINELS", "UNIT_SCALE", "scale_from_number_format", "parse_unit_label",
    "normalize_value", "scale_factor", "regex_type", "infer_column_type",
]
