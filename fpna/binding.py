"""
fpna.binding — T2 실데이터 바인딩 중앙 기계 (tidy rows → 템플릿 INPUT).

설계(자문 3R 수렴, out/consult-spine-entrypoint-3r.md):
  "형태 조립 = 템플릿(from_tidy), 검증/강제 = 중앙(bind_and_check)."
  ⛔ 순수 중앙 매퍼(god-object) 금지 — 단위/grain 의미는 템플릿마다 다르다.
  ⛔ 순수 per-template 금지 — grain 유일성·필수컬럼·단위 검증이 표류한다.
  → 형태(트리 1:N 조립)는 템플릿 어댑터(from_tidy)가 소유하고,
    **검증/강제(grain_unique·REQUIRED·UNIT_POLICY)는 이 중앙 기계가 소유**한다.

이 모듈은 런타임 경로(회사 PC) — openpyxl 도 불필요, **stdlib 만** 쓴다.
  (csv / datetime / itertools / dataclasses). pandas/numpy 금지(§1 절대 제약).

provenance(자문 C8): grain_unique 검사는 **from_tidy 가 행을 접기(groupby) 전의
pre-shape rows** 에서 한다. groupby 가 1:N 을 말아 올리면 중복 키가 사라져
침묵형 silent-merge 가 되기 때문이다. 어댑터(from_tidy)는 불신한다 — 반환된
INPUT 도 사후에 REQUIRED/UNIT_POLICY 로 무조건 재검한다.

⚠ 한계(자문 C8-2): UNIT_POLICY 는 "선언된 키가 존재/타입 일치"하는 가벼운 검증만
한다. 의미적 오선택(gross↔net, 천원↔원 같은 단위 혼동)은 **잡지 못한다** —
값이 올바른 컬럼에서 왔는지는 사람이 매핑할 때만 보증된다. 이 한계는 의도된
것으로, 단위의 의미 보증은 호출자(매핑하는 사람)의 책임으로 남긴다.
"""
from __future__ import annotations

import csv as _csv
import datetime as _dt
import math
import re
from itertools import groupby

import fpna._bootstrap  # noqa: F401

# 천단위 콤마 패턴(정수부 1-3자리 + 3자리 그룹 반복). "1,200"/"12,000,000" 통과,
# "1,2,3"/"1,23" 거부 → 다값병합·오타를 silent 흡수하지 않음.
_THOUSANDS_RE = re.compile(r"\d{1,3}(,\d{3})+(\.\d+)?")
# 흔한 날짜 포맷(ISO 외 한국 실데이터) — strptime 시도 순서.
_DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m", "%Y/%m", "%Y.%m", "%Y%m%d")


# --------------------------------------------------------------------------- #
# _coerce — tidy 셀 문자열 → 파이썬 타입 (stdlib only)                         #
# --------------------------------------------------------------------------- #
def _coerce(val, typ):
    """tidy 셀 값(문자열일 수 있음)을 typ 로 강제 변환. stdlib only.

    int/float: 콤마·통화기호(₩$€£,원)·공백 strip 후 변환. "(1,200)"=세모음수 → -1200.
    str:       strip.
    date:      ISO("YYYY-MM-DD") 파싱(이미 date 면 그대로).
    bool:      "true"/"1"/"y"/"yes" → True (대소문자 무시).
    None/빈문자: None 반환(REQUIRED 가 별도로 잡는다).
    typ=None:  변환 없이 원값.
    """
    if typ is None:
        return val
    # 이미 목표 타입(문자열 제외 — 문자열은 정제 필요)
    if typ is _dt.date and isinstance(val, _dt.date) and not isinstance(val, _dt.datetime):
        return val
    if typ is not str and not isinstance(val, str) and val is not None:
        # int/float/bool 등 이미 숫자형이면 typ 로 캐스팅만
        if typ in (int, float):
            return typ(val)
        if typ is bool:
            return bool(val)
        if typ is _dt.date and isinstance(val, _dt.datetime):
            return val.date()
        return val

    if val is None:
        return None
    s = val.strip() if isinstance(val, str) else val
    if isinstance(s, str) and s == "":
        return None

    if typ is str:
        return s
    if typ is bool:
        return str(s).strip().lower() in ("true", "1", "y", "yes", "t")
    if typ is _dt.date:
        return _parse_date(s)
    if typ in (int, float):
        return _parse_number(s, typ)
    # 미지원 타입 — 원값(어댑터가 처리)
    return val


def _parse_number(s, typ):
    """콤마·통화기호·괄호 세모음수 제거 후 int/float 변환."""
    if not isinstance(s, str):
        return typ(s)
    t = s.strip()
    neg = False
    if t.startswith("(") and t.endswith(")"):     # 회계 괄호 음수 (1,200) = -1200
        neg = True
        t = t[1:-1]
    # 통화기호·공백·'원' 제거(콤마는 천단위 검증 후 별도 처리)
    for ch in ("₩", "$", "€", "£", " ", "원"):
        t = t.replace(ch, "")
    if t in ("", "-", "—"):
        return None
    # 콤마는 천단위 구분자일 때만 허용 — "1,2,3"(다값병합·오타) silent 흡수 금지(엣지 리뷰).
    if "," in t:
        core = t[1:] if t[:1] == "-" else t
        if not _THOUSANDS_RE.fullmatch(core):
            raise ValueError("콤마가 천단위 구분자 아님(다값·오타 의심): %r" % s)
        t = t.replace(",", "")
    f = float(t)
    if not math.isfinite(f):                       # nan/inf 문자열 silent 흡수 금지
        raise ValueError("유한수 아님(nan/inf): %r" % s)
    if typ is int and f != int(f):
        # 소수를 int 로 요구하면 반올림 대신 명확 실패(결정성)
        raise ValueError("int 컬럼에 소수 값: %r" % s)
    v = int(f) if typ is int else f
    return -v if neg else v


def _parse_date(s):
    """date 파싱. ISO + 한국 실데이터 흔한 포맷(YYYY.MM.DD/YYYY-MM 등). 실패 시 ValueError."""
    if isinstance(s, _dt.datetime):
        return s.date()
    if isinstance(s, _dt.date):
        return s
    t = str(s).strip()
    try:
        return _dt.date.fromisoformat(t.replace("/", "-"))
    except ValueError:
        pass
    for fmt in _DATE_FORMATS:
        try:
            return _dt.datetime.strptime(t, fmt).date()
        except ValueError:
            continue
    raise ValueError("지원 않는 날짜 포맷: %r" % s)


# --------------------------------------------------------------------------- #
# assemble — 트리 1:N 공통 조립 (itertools.groupby 다단)                       #
# --------------------------------------------------------------------------- #
def assemble(rows, spec):
    """tidy rows(list[dict]) → 헤더 객체 1개 (트리 1:N 다단 조립).

    spec = {
      "header_cls":    Callable      # 헤더 dataclass 생성자
      "header_fields": {kw: (col, typ)}  # 헤더 dataclass kwarg ← 컬럼 매핑(+타입)
      "levels": [ <레벨 spec> ]      # 첫 레벨 1개(중첩은 그 레벨의 child 로 표현)
    }
    레벨 spec = {
      "key_cols":   [컬럼명...],     # 이 레벨 라인 식별 키(groupby 키)
      "line_cls":   Callable,        # 라인 dataclass 생성자
      "fields":     {kw: (col, typ)},  # 라인 dataclass kwarg ← 컬럼 매핑
      "child_attr": str,             # 부모가 이 레벨 list 를 받는 속성명
      "child":      <레벨 spec>       # (옵션) 자식 레벨 — 동일 구조 중첩
    }

    각 레벨은 부모 그룹 내부에서 key_cols 로 groupby → 라인 객체 list 생성.
    child 가 있으면 그 라인의 child.child_attr 에 자식 레벨 list 를 중첩한다.
    반환: header_cls(**header_kwargs, <top.child_attr> = [라인...]).
    """
    top = spec["levels"][0]
    if not rows:
        # 빈 입력 — 헤더만(라인 0). header_fields 는 None 으로.
        hkw = {kw: None for kw in spec.get("header_fields", {})}
        return spec["header_cls"](**hkw, **{top["child_attr"]: []})

    # 헤더 kwargs — 모든 행에서 동일하다고 가정(첫 행에서 취득; pre-shape 단계에서
    # grain 검사가 헤더 일관성 위반을 사실상 잡는다. 헤더는 보통 보고서 메타).
    hkw = {}
    for kw, (col, typ) in spec.get("header_fields", {}).items():
        hkw[kw] = _coerce(rows[0].get(col), typ)

    lines = _assemble_level(rows, top)
    return spec["header_cls"](**hkw, **{top["child_attr"]: lines})


def _assemble_level(rows, lv):
    """레벨 spec(lv)을 rows 위에서 key_cols 로 groupby 조립 → 라인 객체 list.

    lv 에 child(자식 레벨 spec)가 있으면 그 그룹 rows 로 재귀해
    라인의 child.child_attr 에 중첩한다(트리 1:N 다단).
    """
    key_cols = lv["key_cols"]
    keyf = lambda r: tuple(r.get(k) for k in key_cols)
    srt = sorted(rows, key=keyf)
    out = []
    for _key, grp in groupby(srt, key=keyf):
        grp = list(grp)
        lkw = {}
        for kw, (col, typ) in lv.get("fields", {}).items():
            lkw[kw] = _coerce(grp[0].get(col), typ)
        child = lv.get("child")
        if child is not None:
            lkw[child["child_attr"]] = _assemble_level(grp, child)
        out.append(lv["line_cls"](**lkw))
    return out


# --------------------------------------------------------------------------- #
# 중앙 검증/강제 — 어댑터 불신 (반환 INPUT 무조건 재검)                        #
# --------------------------------------------------------------------------- #
def _pre_shape_grain_check(rows, grain_cols):
    """from_tidy 前 pre-shape rows 에서 grain 중복 탐지(자문 C8-3).

    groupby 가 1:N 을 접으면 중복 키가 사라져 silent-merge 가 된다. 그래서
    조립 *전* 의 raw rows 에서 (grain_cols) 조합 중복을 직접 본다.
    중복/누락 발견 시 ValueError(저장 막기 전에 바인딩 단계에서 거부).
    """
    seen = set()
    dups = []
    missing = 0
    for r in rows:
        kt = tuple(r.get(c) for c in grain_cols)
        if any(v is None or (isinstance(v, str) and v.strip() == "") for v in kt):
            missing += 1
            continue
        if kt in seen:
            dups.append(kt)
        seen.add(kt)
    problems = []
    if dups:
        problems.append("중복 grain %d건: %s" % (len(dups), dups[:5]))
    if missing:
        problems.append("grain key 누락(None/빈값) 행 %d건" % missing)
    if problems:
        raise ValueError("grain_unique 위반(pre-shape) — " + "; ".join(problems))


def _assert_required(inp, required):
    """반환 INPUT 의 REQUIRED 속성이 비지 않았는지(None/빈 list/빈 문자열 금지).

    required = list[str] 속성명. 중첩 라인 리스트 속성이면 "비어있지 않은가"만 본다.
    """
    bad = []
    for attr in required or []:
        if not hasattr(inp, attr):
            bad.append("%s(속성 부재)" % attr)
            continue
        v = getattr(inp, attr)
        if v is None or (isinstance(v, (list, tuple, str, dict)) and len(v) == 0):
            bad.append(attr)
    if bad:
        raise ValueError("REQUIRED 미충족: " + ", ".join(bad))


def _assert_units(inp, unit_policy):
    """UNIT_POLICY 선언 키의 존재/타입을 가볍게 검증(자문 C8-2 한계 명시).

    unit_policy = {attr: typ}. attr 가 INPUT 에 존재하고, 값이 typ 의 인스턴스인지
    (또는 typ 가 (int,float) 면 수치인지)만 본다. ⛔ 의미적 단위 오선택은 못 잡는다.
    중첩 라인 속성 'lines.field' 표기 지원 — 각 라인의 field 가 typ 인지.
    """
    bad = []
    for attr, typ in (unit_policy or {}).items():
        if "." in attr:
            list_attr, field = attr.split(".", 1)
            seq = getattr(inp, list_attr, None) or []
            for i, line in enumerate(seq):
                if not hasattr(line, field):
                    bad.append("%s[%d](필드 부재)" % (attr, i))
                    continue
                if not _type_ok(getattr(line, field), typ):
                    bad.append("%s[%d]=%r(타입≠%s)" % (attr, i, getattr(line, field),
                                                       getattr(typ, "__name__", typ)))
        else:
            if not hasattr(inp, attr):
                bad.append("%s(속성 부재)" % attr)
                continue
            if not _type_ok(getattr(inp, attr), typ):
                bad.append("%s=%r(타입≠%s)" % (attr, getattr(inp, attr),
                                               getattr(typ, "__name__", typ)))
    if bad:
        raise ValueError("UNIT_POLICY 타입 위반: " + ", ".join(bad))


def _type_ok(value, typ):
    """value 가 typ 에 맞는지. (int,float) 는 bool 제외 수치 허용. None 은 통과
    (REQUIRED 가 None 을 전담; UNIT_POLICY 는 '값이 있다면 타입 맞나'만 본다)."""
    if value is None:
        return True
    if typ in (int, float):
        return isinstance(value, (int, float)) and not isinstance(value, bool)
    return isinstance(value, typ)


# --------------------------------------------------------------------------- #
# bind_and_check — 중앙 진입 (어댑터 호출 + 사전·사후 검증)                    #
# --------------------------------------------------------------------------- #
def bind_and_check(mod, rows):
    """tidy rows → mod.from_tidy → 검증된 INPUT. 중앙 기계의 핵심.

    절차(자문: 검증/강제는 중앙 소유):
      ① pre-shape grain 검사 — mod.GRAIN 컬럼 조합이 rows(조립 前)에서 유일한지.
         (groupby 가 행을 접기 전에 중복 탐지 — silent-merge 차단)
      ② inp = mod.from_tidy(rows) — 형태 조립은 템플릿 어댑터가 소유.
      ③ 사후 재검(어댑터 불신): _assert_required(mod.REQUIRED) +
         _assert_units(mod.UNIT_POLICY). 반환 후 무조건 다시 본다.

    mod 가 from_tidy 를 노출하지 않으면 명확한 예외.
    """
    if not hasattr(mod, "from_tidy"):
        raise NotImplementedError(
            "%s 는 아직 --csv(from_tidy) 미지원 — golden_sample() 을 쓰세요."
            % getattr(mod, "TYPE", getattr(mod, "__name__", "이 템플릿")))

    grain = getattr(mod, "GRAIN", None)
    if grain:
        _pre_shape_grain_check(rows, grain)        # ① pre-shape (조립 前)

    inp = mod.from_tidy(rows)                       # ② 형태 조립(어댑터)

    _assert_required(inp, getattr(mod, "REQUIRED", []))   # ③ 사후 재검
    _assert_units(inp, getattr(mod, "UNIT_POLICY", {}))
    return inp


def bind_from_csv(mod, csv_path):
    """csv 파일 → tidy rows(DictReader) → bind_and_check. stdlib csv only.

    utf-8-sig(엑셀 BOM) 우선 디코드. mod 에 from_tidy 없으면 bind_and_check 가
    명확히 거부(NotImplementedError) 한다.
    """
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as fh:
        rows = list(_csv.DictReader(fh))
    return bind_and_check(mod, rows)


__all__ = [
    "_coerce", "assemble",
    "bind_and_check", "bind_from_csv",
]
