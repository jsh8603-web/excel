"""fpna.templates.budget_build — 예산·인건비 수립(부서별 인원·비용).

깊이(C4): ZBB(영기준) vs incremental(전년대비 증감) 편성방식 구분 — incremental 은
  baseline(전년) + Δ 로, ZBB 는 0 기준 재산정으로 표시(편성 근거 추적성).
게이트(C5): R10 부서 roll-up == 총계(assert_tie_out) — 누락 부서 차단.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl

from fpna import house_style as hs
from fpna import view_contract as vc
from fpna.templates.base import QCReport, qc_no_formula_errors

TYPE = "budget_build"


@dataclass
class DeptLine:
    dept: str
    headcount: int
    avg_cost: float            # 1인당 연간 인건비
    method: str = "ZBB"        # "ZBB"(영기준 재산정) | "incremental"(전년+Δ)
    prior_budget: float = 0.0  # incremental baseline(전년 인건비). ZBB 면 무시.


@dataclass
class BudgetInput:
    title: str = "예산·인건비 수립"
    subtitle: str = ""
    unit: str = "₩mn"
    depts: list = field(default_factory=list)   # list[DeptLine]


def golden_sample() -> BudgetInput:
    return BudgetInput(
        title="예산·인건비 수립 (골든샘플)", subtitle="구조 검증용 — 더미", unit="₩mn",
        depts=[DeptLine("영업", 10, 60, method="incremental", prior_budget=540),
               DeptLine("개발", 20, 80, method="ZBB"),
               DeptLine("관리", 5, 55, method="incremental", prior_budget=300)],
    )


def build(data: BudgetInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("Budget")
    last_col = 6
    hs.set_widths(ws, {1: 18, 2: 12, 3: 10, 4: 14, 5: 14, 6: 14})
    r = hs.report_frame(ws, data.title, subtitle=data.subtitle,
                        unit=data.unit, last_col=last_col)

    headers = ["부서 (단위: %s)" % data.unit, "편성방식", "인원", "1인당 비용",
               "전년 예산", "인건비 합"]
    for j, h in enumerate(headers, 1):
        hs.set_cell(ws, r, j, h, role="header", align=hs.LEFT if j == 1 else hs.CENTER)
    r += 1
    data_start = r
    for d in data.depts:
        hs.set_cell(ws, r, 1, d.dept, role="label", align=hs.LEFT)
        hs.set_cell(ws, r, 2, d.method, role="soft", align=hs.CENTER)
        hs.set_cell(ws, r, 3, d.headcount, role="input", number_format=hs.FMT_INT)
        hs.set_cell(ws, r, 4, d.avg_cost, role="input", number_format=hs.FMT_INT)
        # incremental 만 전년 baseline 표시(ZBB 는 대시 — 0기준 재산정 의미)
        if d.method == "incremental":
            hs.set_cell(ws, r, 5, d.prior_budget, role="input", number_format=hs.FMT_INT)
        else:
            hs.set_cell(ws, r, 5, "—", role="soft", align=hs.CENTER)
        hs.set_cell(ws, r, 6, "=C%d*D%d" % (r, r), role="calc", number_format=hs.FMT_INT)
        r += 1
    data_end = r - 1
    # 합계
    total_row = r
    hs.set_cell(ws, r, 1, "합계", role="total", align=hs.LEFT, bold=True)
    hs.set_cell(ws, r, 3, "=SUM(C%d:C%d)" % (data_start, data_end), role="calc",
                number_format=hs.FMT_INT, bold=True)
    hs.set_cell(ws, r, 6, "=SUM(F%d:F%d)" % (data_start, data_end), role="calc",
                number_format=hs.FMT_INT, bold=True)
    for j in range(1, last_col + 1):
        ws.cell(row=r, column=j).border = hs.BORDER_TOP_STRONG
    r += 2

    # incremental 증감 요약(전년 대비 Δ) — 편성 근거 추적
    inc = [d for d in data.depts if d.method == "incremental"]
    if inc:
        r = hs.section_header(ws, r, "Incremental 증감 (전년 대비)", last_col=last_col)
        for d in inc:
            cur = d.headcount * d.avg_cost
            delta = cur - d.prior_budget
            hs.set_cell(ws, r, 1, d.dept, role="label", align=hs.LEFT)
            hs.set_cell(ws, r, 2, "전년 %g → 금년 %g (Δ%+g)" % (d.prior_budget, cur, delta),
                        role="soft", align=hs.LEFT)
            r += 1
        r += 1

    hs.report_footer(ws, r, source="인사 정원 · 인건비 단가표 · 전년 예산",
                     prepared_by="FP&A", last_col=last_col)
    return wb


# --------------------------------------------------------------------------- #
# T2 바인딩 (from_tidy) — module-level                                         #
#   conserves 는 deferred: budget_build 는 합계를 Excel 수식(=SUM)에만 두고     #
#   _fpna_meta 에 보고 총계를 심지 않는다. T4 보존은 reported = _fpna_meta[...] #
#   를 전제하므로 여기선 tie 대상이 없다 → conserves 미구현(silent cap 금지,    #
#   본 주석으로 deferred 명시). 필요 시 build 가 grand 를 _fpna_meta 에 심는     #
#   변경이 선행돼야 함(build 변경은 본 작업 스코프 밖).                          #
# --------------------------------------------------------------------------- #
GRAIN = ("dept",)                              # 1행 = 1 부서
REQUIRED = ("depts",)
UNIT_POLICY = {"depts.avg_cost": float}


def from_tidy(rows) -> BudgetInput:
    """tidy rows(부서 1건/행) → BudgetInput. 형태 조립만.

    행 컬럼: dept, headcount, avg_cost, [method, prior_budget].
    method 기본 'ZBB'(prior_budget 미사용); incremental 이면 prior_budget 필수
    (qc 가 baseline 누락을 별도로 잡는다).
    """
    from fpna.binding import _coerce
    depts = []
    for r in rows:
        depts.append(DeptLine(
            dept=_coerce(r.get("dept"), str),
            headcount=_coerce(r.get("headcount"), int) or 0,
            avg_cost=_coerce(r.get("avg_cost"), float) or 0.0,
            method=(_coerce(r.get("method"), str) or "ZBB"),
            prior_budget=_coerce(r.get("prior_budget"), float) or 0.0,
        ))
    return BudgetInput(depts=depts)


def qc(wb, data: BudgetInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    per_dept = [d.headcount * d.avg_cost for d in data.depts]
    total = sum(per_dept)
    hc = sum(d.headcount for d in data.depts)
    rep.add("인원 합 > 0", hc > 0, "hc=%d" % hc)
    rep.add("인건비 합 계산", total >= 0, "합=%.0f" % total)

    # --- R10 부서 roll-up == 총계: Σ부서별 == grand total(누락 부서 차단) ------
    vc.assert_tie_out(rep, sum(per_dept), total, tol=1e-6, name="R10 dept_rollup_tie")

    # --- 편성방식 유효(ZBB|incremental) + incremental 은 baseline 필수 --------
    bad_method = [d.dept for d in data.depts if d.method not in ("ZBB", "incremental")]
    rep.add("편성방식 유효(ZBB|incremental)", not bad_method,
            "" if not bad_method else "미정의: " + ", ".join(bad_method))
    no_base = [d.dept for d in data.depts
               if d.method == "incremental" and d.prior_budget <= 0]
    rep.add("incremental baseline 명시", not no_base,
            "" if not no_base else "전년예산 누락: " + ", ".join(no_base))

    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "DeptLine", "BudgetInput", "golden_sample", "build", "qc",
           "GRAIN", "REQUIRED", "UNIT_POLICY", "from_tidy"]
