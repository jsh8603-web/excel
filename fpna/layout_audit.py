"""
fpna.layout_audit — 서식/레이아웃 계약 (View Contract 의 '표현층' 보강).

배경(독립 리뷰 2026-06): 기존 그물(conserve/view_contract)은 전부 *데이터 의미*를
검증한다. **포맷·레이아웃 구조 계약이 없었다.** 그리고 골든 회귀(test_parity._cells)는
값+수식만 스냅샷해서, 서식이 엉뚱한 range 에 묻어도 초록으로 통과했다 → "양식 불균형"
버그가 사람 눈에 닿기 전까지 안 잡혔다. 이 모듈이 그 구멍을 메운다.

두 갈래로 쓴다:
  (A) fingerprint(wb)          : 셀별 정규화 서식 지문 → 골든 스냅샷(드리프트 게이트).
  (B) assert_allowed_formats() : house_style 캐논 밖의 number_format 을 hard-fail
                                 (전 템플릿 공통, 오탐 거의 0 — pipeline 게이트에 연결).
  (C) column_format_outliers() : 같은 컬럼 안 number_format 불균형 휴리스틱
                                 (외부 dead 파일용 — excel_doctor 가 소비, --fix 가능).

런타임: openpyxl 필요(qc/render 경로와 동일). 회사 PC 는 vendor/ 동봉으로 해소.
"""
from __future__ import annotations

import fpna._bootstrap  # noqa: F401  (vendor openpyxl on path)

from collections import Counter, defaultdict

from fpna import house_style as hs
from fpna.templates.base import QCReport


# --------------------------------------------------------------------------- #
# 캐논 number_format 집합 — house_style 의 FMT_* 상수에서 자동 수집            #
# --------------------------------------------------------------------------- #
def _canon_formats() -> set:
    """house_style 의 모든 FMT_* 상수 ∪ {General, @}. 캐논 = 이 집합."""
    fmts = {getattr(hs, n) for n in dir(hs)
            if n.startswith("FMT_") and isinstance(getattr(hs, n), str)}
    fmts |= {"General", "@"}
    return fmts


CANON_FORMATS = _canon_formats()


# --------------------------------------------------------------------------- #
# (A) fingerprint — 골든 스냅샷용 정규화 서식 지문                            #
# --------------------------------------------------------------------------- #
def _argb(color) -> str | None:
    """openpyxl Color → 안정적 ARGB 문자열(테마/인덱스 색은 repr 로)."""
    if color is None:
        return None
    rgb = getattr(color, "rgb", None)
    if isinstance(rgb, str):
        return rgb
    t = getattr(color, "theme", None)
    if t is not None:
        return "theme:%s/%s" % (t, getattr(color, "tint", 0))
    idx = getattr(color, "indexed", None)
    if idx is not None:
        return "indexed:%s" % idx
    return None


def _cell_fp(c) -> dict:
    """단일 셀의 서식 지문(값 제외 — 값은 test_parity 가 본다)."""
    fp: dict = {"nf": c.number_format}
    f = c.font
    if f is not None:
        fp["bold"] = bool(f.bold)
        fp["color"] = _argb(f.color)
        fp["size"] = f.size
    a = c.alignment
    if a is not None:
        fp["halign"] = a.horizontal
        fp["indent"] = a.indent or 0
    fill = c.fill
    if fill is not None and getattr(fill, "patternType", None):
        fp["fill"] = _argb(fill.fgColor)
    return fp


def fingerprint(wb) -> dict:
    """워크북 → JSON 직렬화 가능한 정규화 서식 지문.

    구조:
      {
        "<sheet>": {
          "cells":  {"A1": {nf,bold,...}, ...},   # 값이 None 아닌 셀만
          "merged": ["A1:C1", ...],                 # 정렬된 병합 범위
          "widths": {"1": 28.0, ...},               # 열폭(소수 1자리 반올림)
        }, ...
      }
    """
    out: dict = {}
    for ws in wb.worksheets:
        cells: dict = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    cells[c.coordinate] = _cell_fp(c)
        merged = sorted(str(r) for r in ws.merged_cells.ranges)
        widths = {}
        for k, dim in ws.column_dimensions.items():
            if dim.width is not None:
                widths[str(dim.min)] = round(float(dim.width), 1)
        out[ws.title] = {"cells": cells, "merged": merged, "widths": widths}
    return out


def diff_fingerprints(base: dict, cur: dict) -> list:
    """두 지문의 차이를 (경로, base값, cur값) 리스트로. 빈 리스트 = 동일."""
    diffs: list = []
    for sheet in sorted(set(base) | set(cur)):
        b = base.get(sheet, {})
        c = cur.get(sheet, {})
        bc, cc = b.get("cells", {}), c.get("cells", {})
        for coord in sorted(set(bc) | set(cc)):
            if bc.get(coord) != cc.get(coord):
                diffs.append(("%s!%s" % (sheet, coord), bc.get(coord), cc.get(coord)))
        if b.get("merged") != c.get("merged"):
            diffs.append(("%s!merged" % sheet, b.get("merged"), c.get("merged")))
        if b.get("widths") != c.get("widths"):
            diffs.append(("%s!widths" % sheet, b.get("widths"), c.get("widths")))
    return diffs


# --------------------------------------------------------------------------- #
# (B) assert_allowed_formats — 전 템플릿 공통 hard 게이트                      #
# --------------------------------------------------------------------------- #
def assert_allowed_formats(rep: QCReport, wb, *,
                           name: str = "서식 캐논(허용 number_format)") -> bool:
    """캐논(house_style FMT_*) 밖의 number_format 을 쓴 셀을 fail.

    오탐 위험이 가장 낮은 보편 규칙 — 외래/임시 서식(직접 타이핑한 '0.000' 등)이
    섞이면 즉시 적색. 표현 일관성의 1차 방어선.
    """
    bad: list = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                if c.number_format not in CANON_FORMATS:
                    bad.append("%s!%s=%r" % (ws.title, c.coordinate, c.number_format))
    ok = not bad
    rep.add(name, ok, "" if ok else "캐논 밖 서식 %d건: %s" % (len(bad), ", ".join(bad[:8])))
    return ok


# --------------------------------------------------------------------------- #
# (C) column_format_outliers — dead 파일 휴리스틱(doctor 용, --fix 대상)        #
# --------------------------------------------------------------------------- #
def column_format_outliers(wb, *, min_majority: int = 3) -> list:
    """같은 (시트, 열) 안 숫자 셀의 number_format 불균형을 탐지.

    다수 서식이 min_majority 이상이고 소수 outlier 가 (다수의 1/3 이하)면
    outlier 를 "양식 불균형 후보"로 반환. 헤더/라벨(텍스트)·수식문자열은 제외하고
    *숫자값* 셀만 본다. 반환: [{sheet, col, coord, found, expected}, ...].

    ⚠ 휴리스틱이다 — 의도된 혼합서식(예: 같은 열에 정수/비율 혼재)을 오탐할 수 있다.
       그래서 게이트가 아니라 doctor 의 *제안*으로만 쓰고, --fix 는 사용자 확인 후.
    """
    findings: list = []
    for ws in wb.worksheets:
        bycol: dict = defaultdict(list)  # col -> [(coord, nf)]
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    bycol[c.column].append((c.coordinate, c.number_format))
        for col, items in bycol.items():
            if len(items) < min_majority + 1:
                continue
            cnt = Counter(nf for _, nf in items)
            (top_nf, top_n), = cnt.most_common(1)
            if top_n < min_majority:
                continue
            for coord, nf in items:
                if nf != top_nf and cnt[nf] <= max(1, top_n // 3):
                    findings.append({"sheet": ws.title, "col": col, "coord": coord,
                                     "found": nf, "expected": top_nf})
    return findings


# --------------------------------------------------------------------------- #
# (D) 콘텐츠 타입 계약 — 텍스트가 숫자/헤더 영역에 침투하는 것을 차단            #
#     (W26 스크린샷 진단 2026-06: "헤더에 데이터", "주석이 값 셀에", #VALUE!     #
#      셋 다 동일 뿌리 = 숫자여야 할 셀에 문자열이 들어감)                       #
# --------------------------------------------------------------------------- #
# 기간/헤더 토큰 — 본문(데이터 행)에 *값*으로 나타나면 누수.
HEADER_TOKENS = {
    "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
    "JAN", "FEB", "MAR", "Q1", "Q2", "Q3", "Q4", "FY",
}


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def text_in_numeric_columns(wb, *, min_majority: int = 3) -> list:
    """다수가 숫자인 열에 끼어든 텍스트(헤더 누수/주석 누수)를 탐지(dead 파일 휴리스틱).

    두 신호를 분리한다(헤더 라벨 자체를 오탐하지 않도록):
      · 헤더 토큰(APR..FY)이 *데이터 시작행보다 아래*에 값으로 → 헤더 누수("AUG" in MAY열).
      · 그 외 텍스트가 숫자-다수 열의 데이터 구간(시작행 이상)에 → 주석 누수.
    헤더 행(데이터 위)의 정상 토큰은 건너뛴다. 반환 [{sheet,coord,value,kind}].

    ⚠ 값 문제라 --fix 자동수리 불가. 다중 헤더밴드가 있으면 오탐 가능(advisory).
    """
    findings: list = []
    for ws in wb.worksheets:
        bycol: dict = defaultdict(lambda: {"num_rows": [], "txt": []})
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None or _is_formula(v):
                    continue
                if _is_number(v):
                    bycol[c.column]["num_rows"].append(c.row)
                elif isinstance(v, str) and v.strip():
                    bycol[c.column]["txt"].append((c.row, c.coordinate, v))
        for col, d in bycol.items():
            if len(d["num_rows"]) < min_majority:
                continue
            first_num = min(d["num_rows"])
            for r, coord, v in d["txt"]:
                tok = v.strip().upper() in HEADER_TOKENS
                if tok and r > first_num:
                    findings.append({"sheet": ws.title, "col": col, "coord": coord,
                                     "value": v, "kind": "헤더토큰 누수"})
                elif not tok and r >= first_num:
                    findings.append({"sheet": ws.title, "col": col, "coord": coord,
                                     "value": v, "kind": "주석/텍스트 누수"})
    return findings


def assert_cell_content_types(rep: QCReport, wb, meta: dict, *,
                              name: str = "콘텐츠 타입(숫자영역 텍스트 차단)") -> bool:
    """선언형 게이트(생성단계). meta 가 선언하면 강제, 없으면 no-op.

      meta["numeric_regions"] = [(sheet, r0, r1, c0, c1), ...]
          → 그 사각영역의 비어있지 않은 셀은 숫자 또는 수식("=")만 허용.
            맨문자열(주석/헤더 누수)이 있으면 FAIL → #VALUE! 의 상류 차단.
      meta["header_rows"]     = {sheet: [행번호, ...]}
          → 헤더 토큰(APR..FY)이 그 행 *밖*에 값으로 나타나면 FAIL(헤더 누수).
    """
    regions = (meta or {}).get("numeric_regions") or []
    header_rows = (meta or {}).get("header_rows") or {}
    by_sheet = {ws.title: ws for ws in wb.worksheets}
    bad: list = []

    for sheet, r0, r1, c0, c1 in regions:
        ws = by_sheet.get(sheet)
        if ws is None:
            bad.append("시트없음:%s" % sheet)
            continue
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                v = ws.cell(row=r, column=c).value
                if v is None or _is_number(v) or _is_formula(v):
                    continue
                bad.append("%s!%s=%r(숫자영역 텍스트)" %
                           (sheet, ws.cell(row=r, column=c).coordinate, v))

    for sheet, rows in header_rows.items():
        ws = by_sheet.get(sheet)
        if ws is None:
            continue
        allowed = set(rows)
        for row in ws.iter_rows():
            for c in row:
                if c.row in allowed:
                    continue
                if isinstance(c.value, str) and c.value.strip().upper() in HEADER_TOKENS:
                    bad.append("%s!%s=%r(헤더토큰 본문누수)" % (sheet, c.coordinate, c.value))

    ok = not bad
    rep.add(name, ok, "" if ok else "%d건: %s" % (len(bad), "; ".join(bad[:8])))
    return ok


__all__ = ["fingerprint", "diff_fingerprints", "assert_allowed_formats",
           "column_format_outliers", "text_in_numeric_columns",
           "assert_cell_content_types", "HEADER_TOKENS", "CANON_FORMATS"]
