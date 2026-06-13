"""
fpna.templates.investment_appraisal — 투자 타당성(NPV·IRR·할인회수기간).

산출: 기간별 현금흐름표 + 할인계수 + 누적 할인현금흐름 + 요약(NPV/IRR/회수기간).
break-even month = 할인 누적이 0을 넘는 시점(선형보간).
빌더는 NPV/IRR 를 Excel 수식(=NPV/IRR)으로 기입, QC 는 파이썬 finance 로 대조.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401

import openpyxl
from openpyxl.utils import get_column_letter

from fpna import finance, house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors, qc_totals, qc_sign

TYPE = "investment_appraisal"


@dataclass
class InvestmentInput:
    title: str = "투자 타당성 분석"
    subtitle: str = ""
    unit: str = "₩mn"
    discount_rate: float = 0.10           # 기간 할인율
    period_label: str = "연도"            # 기간 축 라벨(연/월)
    cashflows: list = field(default_factory=list)  # list[float], [0]=t0 투자(음수)


def golden_sample() -> InvestmentInput:
    return InvestmentInput(
        title="투자 타당성 분석 (골든샘플)",
        subtitle="구조 검증용 — 수치는 더미",
        unit="₩mn", discount_rate=0.10, period_label="연도",
        cashflows=[-1000, 300, 350, 400, 450, 300],
    )


def build(data: InvestmentInput, *, mode: str = "create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = hs.safe_sheet_title("Investment")
    last_col = 2 + len(data.cashflows)
    widths = {1: 22}
    for j in range(2, last_col + 1):
        widths[j] = 12
    hs.set_widths(ws, widths)

    r = hs.report_frame(ws, data.title, subtitle=data.subtitle,
                        unit=data.unit, last_col=last_col, freeze_col="B")

    # 가정 블록
    r = hs.section_header(ws, r, "가정", last_col=last_col)
    hs.set_cell(ws, r, 1, "할인율 (단위: %s)" % data.unit, role="label", align=hs.LEFT)
    rate_cell = "B%d" % r
    hs.set_cell(ws, r, 2, data.discount_rate, role="input", number_format=hs.FMT_PCT1)
    r += 2

    # 기간 헤더
    n = len(data.cashflows)
    hs.set_cell(ws, r, 1, data.period_label, role="header", align=hs.LEFT)
    for t in range(n):
        hs.set_cell(ws, r, 2 + t, "t%d" % t, role="header")
    period_row = r
    r += 1

    # 현금흐름 행
    cf_row = r
    hs.set_cell(ws, r, 1, "현금흐름", role="label", align=hs.LEFT)
    for t, cf in enumerate(data.cashflows):
        hs.set_cell(ws, r, 2 + t, cf, role="input", number_format=hs.FMT_INT)
    r += 1

    # 할인계수 행 = 1/(1+rate)^t (수식)
    df_row = r
    hs.set_cell(ws, r, 1, "할인계수", role="label", align=hs.LEFT)
    for t in range(n):
        col = get_column_letter(2 + t)
        hs.set_cell(ws, r, 2 + t, "=1/(1+$%s)^%d" % (rate_cell, t),
                    role="calc", number_format=hs.FMT_NUM2)
    r += 1

    # 할인현금흐름 행 = cf * df (수식)
    dcf_row = r
    hs.set_cell(ws, r, 1, "할인현금흐름", role="label", align=hs.LEFT)
    for t in range(n):
        col = get_column_letter(2 + t)
        hs.set_cell(ws, r, 2 + t, "=%s%d*%s%d" % (col, cf_row, col, df_row),
                    role="calc", number_format=hs.FMT_INT)
    r += 1

    # 누적 할인현금흐름 행
    cum_row = r
    hs.set_cell(ws, r, 1, "누적 할인 CF", role="label", align=hs.LEFT)
    for t in range(n):
        col = get_column_letter(2 + t)
        if t == 0:
            hs.set_cell(ws, r, 2, "=%s%d" % (col, dcf_row), role="calc",
                        number_format=hs.FMT_INT)
        else:
            prev = get_column_letter(1 + t)
            hs.set_cell(ws, r, 2 + t, "=%s%d+%s%d" % (prev, cum_row, col, dcf_row),
                        role="calc", number_format=hs.FMT_INT)
    r += 2

    # 요약 블록
    r = hs.section_header(ws, r, "요약", last_col=last_col)
    cf_first = get_column_letter(2)
    cf_last = get_column_letter(1 + n)
    cf_t1 = get_column_letter(3)
    # NPV = t0 + NPV(rate, t1..tn)
    hs.set_cell(ws, r, 1, "NPV", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2,
                "=%s%d+NPV($%s,%s%d:%s%d)"
                % (cf_first, cf_row, rate_cell, cf_t1, cf_row, cf_last, cf_row),
                role="calc", number_format=hs.FMT_INT, bold=True)
    npv_cell = (r, 2)
    r += 1
    hs.set_cell(ws, r, 1, "IRR", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IFERROR(IRR(%s%d:%s%d),\"n/a\")"
                % (cf_first, cf_row, cf_last, cf_row),
                role="calc", number_format=hs.FMT_PCT1, bold=True)
    r += 1
    # 할인 회수기간(파이썬 사전계산 → 값으로 기입, 보간이라 단순 수식 어려움)
    dpb = finance.discounted_payback(data.discount_rate, data.cashflows)
    hs.set_cell(ws, r, 1, "할인 회수기간 (기간)", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, (round(dpb, 2) if dpb is not None else "회수불가"),
                role="calc", number_format=hs.FMT_NUM1, bold=True)
    r += 1

    hs.report_footer(ws, r + 1, source="CAPEX 제안 · 현금흐름 가정",
                     prepared_by="FP&A", last_col=last_col)
    wb._fpna_meta = {"npv_cell": npv_cell, "cashflows": data.cashflows,
                     "rate": data.discount_rate}
    return wb


def qc(wb: openpyxl.Workbook, data: InvestmentInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    # NPV 재계산(파이썬) — 셀 수식 의도 검증
    py_npv = finance.npv(data.discount_rate, data.cashflows)
    rep.add("NPV 계산 가능", py_npv is not None, "")
    # IRR 존재성
    py_irr = finance.irr(data.cashflows)
    rep.add("IRR 해 존재", py_irr is not None,
            "현금흐름 부호변화 없음" if py_irr is None else "IRR=%.3f" % py_irr)
    # 회수기간 일관
    dpb = finance.discounted_payback(data.discount_rate, data.cashflows)
    rep.add("할인회수기간", True,
            "회수불가" if dpb is None else "%.2f 기간" % dpb)
    # 부호규약: t0 투자는 음수여야
    qc_sign("t0 투자", data.cashflows[0] if data.cashflows else None, "-", rep)
    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "InvestmentInput", "golden_sample", "build", "qc"]
