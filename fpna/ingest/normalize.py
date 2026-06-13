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

# base(원) 환산용 스케일 사전. 무음 손상 방어 ①의 SSOT.
# 셀내 접미("1,234천원", "3,400억") 분해 + 단위행("(단위: 백만원)") 환산 공용.
SCALE = {
    "천원": 1_000, "천": 1_000,
    "백만원": 1_000_000, "백만": 1_000_000,
    "억원": 100_000_000, "억": 100_000_000,
    "조원": 1_000_000_000_000, "조": 1_000_000_000_000,
    "원": 1,
}
# 셀내 접미 분해용: 길이 긴 단위 우선(억원 > 억). 부동소수 본체 + 단위 접미.
# 본체는 (50)·△50·-50·50 등 음수표기 허용, 단위 뒤 닫는 괄호도 허용((50백만)).
_SCALE_SUFFIX_RE = re.compile(
    r"^\s*(\(?\s*[(△▲\-+]?\s*[\d,]*\.?\d+\s*\)?)\s*(%s)\s*\)?\s*$"
    % "|".join(sorted(SCALE, key=len, reverse=True))
)

_NUM_RE = re.compile(r"^[\s ]*[(△▲\-+]?[\s]*[\d,]*\.?\d+[\s]*[)%]?[\s]*$")
_PCT_RE = re.compile(r"%\s*$")
_DATE_PATTERNS = [
    re.compile(r"^\d{4}[-/.]\d{1,2}([-/.]\d{1,2})?$"),
    re.compile(r"^\d{4}년\s*\d{1,2}월(\s*\d{1,2}일)?$"),
    re.compile(r"^\d{4}\s*[Q분기]\s*\d?$", re.IGNORECASE),
]

# 전각(full-width) 숫자 → ASCII. 한국/일본 엑셀에 종종 섞임 → 숫자 파싱 실패 방지.
_FULLWIDTH_MAP = {ord("０") + i: ord("0") + i for i in range(10)}
_FULLWIDTH_MAP[ord("．")] = ord(".")   # 전각 마침표
_FULLWIDTH_MAP[ord("，")] = ord(",")   # 전각 콤마
_FULLWIDTH_MAP[ord("－")] = ord("-")   # 전각 하이픈
_FULLWIDTH_MAP[ord("％")] = ord("%")   # 전각 퍼센트
_ZEROWIDTH = ("​", "﻿", "‌", "‍")


def _normalize_digits(s: str) -> str:
    """전각숫자·전각구두점 → ASCII, 제로폭 문자 제거(텍스트 숫자 정규화)."""
    s = s.translate(_FULLWIDTH_MAP)
    for z in _ZEROWIDTH:
        if z in s:
            s = s.replace(z, "")
    return s


# --------------------------------------------------------------------------
# G7 각주마커 제거 — 헤더의 ¹ * 주1) 등 → 동일 논리열 키 통일.
# 위첨자 숫자(¹²³…), 별표(*†‡§), 끝부분 (주N)/주N) 패턴을 제거.
# --------------------------------------------------------------------------
_SUPERSCRIPT = "¹²³⁴⁵⁶⁷⁸⁹⁰"
# 끝의 각주 마커: 위첨자/별표 군집, 또는 (주1)·주1)·(1)·*1 류 꼬리표.
_FOOTNOTE_TAIL_RE = re.compile(
    r"(?:"
    r"[¹²³⁰-₟\*†‡§]+"   # 위첨자/별표/단검표
    r"|\s*\(?\s*주\s*\d+\s*\)?"                                  # (주1)/주1)
    r"|\s*\(\s*\d{1,2}\s*\)"                                     # (1)
    r"|\s*\*\d{1,2}"                                             # *1
    r")\s*$"
)


def strip_footnote_marker(text):
    """헤더 텍스트 끝의 각주 마커를 제거해 논리 키를 통일.

    G7: '매출¹' / '매출*' / '매출(주1)' → '매출'.
    반환: (정리키, 제거여부). 문자열 아니면 (원본, False).
    숫자만 남거나 빈 문자열이 되면(예: 마커가 본문이던 경우) 원본 유지.
    """
    if not isinstance(text, str):
        return text, False
    s = text.strip()
    cleaned = _FOOTNOTE_TAIL_RE.sub("", s).strip()
    if cleaned and cleaned != s:
        return cleaned, True
    return s, False


# --------------------------------------------------------------------------
# G3 들여쓰기 계층 — 선행공백 → 레벨(2칸=1레벨).
# openpyxl alignment.indent 와 합산해 라벨 계층 깊이 산출.
# --------------------------------------------------------------------------
def leading_space_level(text, *, spaces_per_level: int = 2) -> int:
    """라벨 앞 선행공백(NBSP 포함)을 레벨로 환산. 2칸=1레벨(내림)."""
    if not isinstance(text, str):
        return 0
    n = 0
    for ch in text:
        if ch in (" ", "\t", " ", "　"):
            n += 1 + (3 if ch in ("\t", "　") else 0)  # tab/전각공백=4칸 가중
        else:
            break
    return n // spaces_per_level



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


def scale_for_unit(unit: str | None) -> int:
    """단위 라벨(또는 '(단위: 백만원)' 잔여 텍스트) → base(원) 환산 곱수.

    SCALE 사전 기준. 매칭 없으면 1. '%'·통화기호는 환산 대상 아님 → 1.
    """
    if not unit:
        return 1
    s = str(unit)
    for u in sorted(SCALE, key=len, reverse=True):
        if u in s:
            return SCALE[u]
    return 1


def split_cell_scale(text):
    """셀내 접미 스케일 분해. '1,234천원' → ('1,234', 1000), '3,400억' → ('3,400', 1e8).

    반환: (본체문자열, 스케일곱수). 접미 없으면 (원본, 1).
    부호/괄호 음수 prefix 는 본체에 그대로 남겨 normalize_value 가 처리.
    """
    if not isinstance(text, str):
        return text, 1
    m = _SCALE_SUFFIX_RE.match(text)
    if not m:
        return text, 1
    return m.group(1).strip(), SCALE[m.group(2)]


def normalize_value(raw, *, fmt: str | None = None, unit_scale: int = 1):
    """단일 셀 값 정규화(하위호환 3-튜플 래퍼)."""
    val, sentinel, neg, _scale = normalize_value_ex(raw, fmt=fmt, unit_scale=unit_scale)
    return val, sentinel, neg


def normalize_value_ex(raw, *, fmt: str | None = None, unit_scale: int = 1):
    """단일 셀 값 정규화 + 셀내 접미 스케일 노출.

    반환: (number|str|None, sentinel:str|None, is_negative:bool, cell_scale:int)
      - 숫자로 해석되면 number(원값 그대로 — 스케일 미적용), sentinel=None
      - 센티넬이면 (None, 매칭센티넬, False, 1)
      - 그 외 텍스트는 (정리된 str, None, False, 1)
      - cell_scale = 셀 텍스트 접미('1,234천원')에서 분해한 곱수(없으면 1).
        ⚠ 값에는 곱하지 않는다(이중계상 방지). pipeline 이 우선순위(셀>블록>1)로 1회만 적용.
    """
    if raw is None:
        return None, "", False, 1
    if isinstance(raw, bool):
        return raw, None, False, 1
    if isinstance(raw, (int, float)):
        # ⚠ 저장된 숫자값은 이미 실수치다. number_format 의 trailing comma 는
        # '표시 축약'일 뿐이므로 값에 곱하면 이중계상 → 곱하지 않는다.
        # 스케일/단위는 별도 unit 컬럼으로만 보존(consumer 가 해석).
        return raw, None, False, 1
    if isinstance(raw, (_dt.datetime, _dt.date, _dt.time)):
        return raw, None, False, 1

    # 전각숫자·NBSP·꼬리공백 정규화(텍스트 숫자 보강)
    s = _normalize_digits(str(raw).replace(" ", " ")).strip()
    low = s.lower()
    if low in SENTINELS:
        return None, s, False, 1

    # 셀내 접미 스케일 분해('1,234천원' → 본체 '1,234' + scale 1000)
    body, cell_scale = split_cell_scale(s)
    neg = False
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
            return num / 100.0, None, neg, 1  # %는 스케일 환산 대상 아님
        return num, None, neg, cell_scale
    except (ValueError, AttributeError):
        return s, None, False, 1


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
    "SENTINELS", "UNIT_SCALE", "SCALE", "scale_from_number_format",
    "parse_unit_label", "scale_for_unit", "split_cell_scale",
    "normalize_value", "normalize_value_ex", "scale_factor",
    "regex_type", "infer_column_type",
    "strip_footnote_marker", "leading_space_level",
]
