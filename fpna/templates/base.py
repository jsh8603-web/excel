"""
fpna.templates.base — 템플릿 공통 베이스(QC 게이트 + 빌더 헬퍼).

각 템플릿 유형 모듈은 다음을 제공한다(덕 타이핑):
  - INPUT: dataclass 입력 스키마
  - build(data, *, mode="create", base_path=None) -> openpyxl.Workbook
  - golden_sample() -> INPUT (재무 의미 없는 구조 골든 데이터)
  - qc(wb, data) -> QCReport

QC 게이트(작업6): 수식에러 0, 합계 교차검증, 부호규약, 단위/포맷 일관.
미통과 시 render 가 산출을 보류한다.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401

from openpyxl.workbook.workbook import Workbook

from fpna.ingest.cells import ERROR_LITERALS


@dataclass
class QCReport:
    template: str
    passed: bool = True
    checks: list = field(default_factory=list)   # (name, ok, detail)

    def add(self, name: str, ok: bool, detail: str = "") -> None:
        self.checks.append((name, bool(ok), detail))
        if not ok:
            self.passed = False

    def summary(self) -> str:
        lines = ["QC[%s]: %s" % (self.template, "PASS" if self.passed else "FAIL")]
        for name, ok, detail in self.checks:
            lines.append("  [%s] %s%s" % ("OK" if ok else "XX", name,
                                          (" — " + detail) if detail else ""))
        return "\n".join(lines)


def qc_no_formula_errors(wb: Workbook, rep: QCReport) -> None:
    """모든 시트의 셀에서 Excel 에러 리터럴(#REF! 등) 탐지."""
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v in ERROR_LITERALS:
                    bad.append("%s!%s=%s" % (ws.title, c.coordinate, v))
    rep.add("수식에러 0건", not bad,
            "" if not bad else "발견: " + ", ".join(bad[:8]))


def qc_totals(label: str, computed: float, expected: float, rep: QCReport,
              *, tol: float = 1e-6) -> None:
    """파이썬 재계산값(computed) vs 기대/셀값(expected) 교차검증."""
    if expected is None:
        rep.add("합계검증:%s" % label, False, "기대값 None")
        return
    ok = abs(computed - expected) <= tol * max(1.0, abs(expected))
    rep.add("합계검증:%s" % label, ok,
            "" if ok else "계산=%.6g 기대=%.6g" % (computed, expected))


def qc_sign(label: str, value, expected_sign: str, rep: QCReport) -> None:
    """부호규약: expected_sign ∈ {'+','-','any'}."""
    if value is None or expected_sign == "any":
        rep.add("부호:%s" % label, True)
        return
    ok = (value >= 0) if expected_sign == "+" else (value <= 0)
    rep.add("부호:%s" % label, ok, "" if ok else "값=%s 기대부호=%s" % (value, expected_sign))


__all__ = ["QCReport", "qc_no_formula_errors", "qc_totals", "qc_sign"]
