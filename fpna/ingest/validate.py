"""
fpna.ingest.validate — tidy 스키마 검증 + 수식 스멜 스캔.

- 스키마: dataclass + 수동 제약(타입/필수/범위). pydantic 미사용.
- 스멜: openpyxl 수식 문자열 정규식 — 하드코딩 상수, 과도 중첩, 시트간 참조 과다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .cells import Cell, ERROR_LITERALS


@dataclass
class TidyRow:
    """언피벗 결과 1행. (entity, period, metric, value) + 메타."""
    entity: str | None
    period: str | None
    metric: str | None
    value: object
    unit: str | None = None
    row_role: str = "data"        # data|subtotal|total|header
    level: int = 0
    src_row: int = 0
    src_col: int = 0
    # 무음 손상 방어 메타(결정적). value 는 base(원) 환산 후 값.
    scale_applied: int = 1        # value 에 적용된 환산 곱수(셀>블록>1)
    scale_source: str = "none"    # cell|block|none — 환산 곱수 출처
    raw_value: object = None      # 원본 셀값(오류값/환산전 보존). value 와 다를 때만 채움
    flags: str = ""               # ;구분 플래그(DITTO_FILLED 등)


@dataclass
class ValidationReport:
    n_rows: int = 0
    n_rejected: int = 0
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.n_rejected == 0 and not self.errors


def validate_rows(rows: list[TidyRow], *, require_period: bool = False,
                  numeric_metric: bool = True) -> tuple[list[TidyRow], ValidationReport]:
    """수동 제약체크. 위반 행은 reject 하고 리포트에 누적."""
    rep = ValidationReport(n_rows=len(rows))
    kept: list[TidyRow] = []
    for i, r in enumerate(rows):
        problems = []
        if r.entity is None and r.metric is None:
            problems.append("entity/metric 모두 결측")
        if require_period and r.period is None:
            problems.append("period 필수인데 결측")
        if numeric_metric and r.row_role == "data" and r.value is not None:
            if not isinstance(r.value, (int, float)):
                problems.append("value 가 비숫자: %r" % (r.value,))
        if isinstance(r.value, str) and r.value in ERROR_LITERALS:
            problems.append("value 에 수식에러: %s" % r.value)
        if problems:
            rep.n_rejected += 1
            rep.errors.append("row#%d (src r%d,c%d): %s"
                              % (i, r.src_row, r.src_col, "; ".join(problems)))
        else:
            kept.append(r)
    if rep.n_rejected:
        rep.warnings.append("총 %d행 중 %d행 reject" % (rep.n_rows, rep.n_rejected))
    return kept, rep


# --------------------------------------------------------------------------
# 수식 스멜 스캔
# --------------------------------------------------------------------------
_FUNC_RE = re.compile(r"([A-Z][A-Z0-9_.]+)\s*\(")
_SHEET_REF_RE = re.compile(r"(?:'[^']+'|[A-Za-z0-9_가-힣]+)!")
_EXT_REF_RE = re.compile(r"\[[^\]]+\.xls[xmb]?\]")
_NUM_CONST_RE = re.compile(r"(?<![A-Za-z0-9_!.$:])\d+\.?\d*")


@dataclass
class Smell:
    cell: str
    kind: str
    detail: str


def scan_formula_smells(cells: list[Cell], *, max_nesting: int = 4,
                        max_sheet_refs: int = 3) -> list[Smell]:
    """수식 셀에서 스멜 탐지.

    - 하드코딩 상수: 수식 안에 셀참조 아닌 매직 넘버(0,1,100 등 흔한 건 제외).
    - 과도한 중첩: 괄호 깊이 > max_nesting.
    - 시트간 참조 과다: 한 수식에서 다른 시트 참조 > max_sheet_refs.
    - 외부파일 참조: [book.xlsx] → 경로 유출 위험 경고.
    """
    smells: list[Smell] = []
    benign = {"0", "1", "2", "12", "100", "1000", "0.5", "1.0"}
    for c in cells:
        if not c.formula:
            continue
        coord = "R%dC%d" % (c.row, c.col)
        f = c.formula

        # 외부 참조
        if _EXT_REF_RE.search(f):
            smells.append(Smell(coord, "external_ref",
                                "외부 파일 참조(경로 마스킹 권장): %s" % f[:60]))
        # 중첩 깊이
        depth, mx = 0, 0
        for ch in f:
            if ch == "(":
                depth += 1
                mx = max(mx, depth)
            elif ch == ")":
                depth -= 1
        if mx > max_nesting:
            smells.append(Smell(coord, "deep_nesting", "괄호 깊이 %d" % mx))
        # 시트간 참조
        n_sheet = len(_SHEET_REF_RE.findall(f))
        if n_sheet > max_sheet_refs:
            smells.append(Smell(coord, "many_sheet_refs", "시트참조 %d회" % n_sheet))
        # 하드코딩 상수
        consts = [m for m in _NUM_CONST_RE.findall(f) if m not in benign]
        if consts:
            smells.append(Smell(coord, "hardcoded_const",
                                "수식 내 상수: %s" % ", ".join(consts[:5])))
    return smells


__all__ = ["TidyRow", "ValidationReport", "validate_rows",
           "Smell", "scan_formula_smells"]
