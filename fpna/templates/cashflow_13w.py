"""fpna.templates.cashflow_13w — 13주 단기 현금흐름/유동성."""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl
from openpyxl.utils import get_column_letter

from fpna import house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors

TYPE = "cashflow_13w"


@dataclass
class CashInput:
    title: str = "13주 현금흐름"
    subtitle: str = ""
    unit: str = "₩mn"
    opening: float = 0.0
    weeks: int = 13
    inflows: list = field(default_factory=list)    # 주별 유입
    outflows: list = field(default_factory=list)   # 주별 유출


def golden_sample() -> CashInput:
    return CashInput(
        title="13주 현금흐름 (골든샘플)", subtitle="구조 검증용 — 더미", unit="₩mn",
        opening=500, weeks=13,
        inflows=[120] * 13, outflows=[100, 110, 130, 90, 100, 140, 95, 100, 120, 110, 100, 130, 90],
    )


def build(data: CashInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("Cash13W")
    n = data.weeks; last_col = 1 + n
    hs.style_sheet(ws, freeze="B6")
    hs.set_widths(ws, {1: 18, **{j: 9 for j in range(2, last_col + 1)}})
    r = hs.title_block(ws, data.title, data.subtitle, last_col=last_col)

    hs.set_cell(ws, r, 1, "항목 (단위: %s)" % data.unit, role="header", align=hs.LEFT)
    for w in range(n):
        hs.set_cell(ws, r, 2 + w, "W%d" % (w + 1), role="header")
    r += 1
    in_row = r
    hs.set_cell(ws, r, 1, "유입", role="label", align=hs.LEFT)
    for w in range(n):
        hs.set_cell(ws, r, 2 + w, data.inflows[w], role="input", number_format=hs.FMT_INT)
    r += 1
    out_row = r
    hs.set_cell(ws, r, 1, "유출", role="label", align=hs.LEFT)
    for w in range(n):
        hs.set_cell(ws, r, 2 + w, data.outflows[w], role="input", number_format=hs.FMT_INT)
    r += 1
    net_row = r
    hs.set_cell(ws, r, 1, "순현금", role="label", align=hs.LEFT)
    for w in range(n):
        col = get_column_letter(2 + w)
        hs.set_cell(ws, r, 2 + w, "=%s%d-%s%d" % (col, in_row, col, out_row),
                    role="calc", number_format=hs.FMT_INT)
    r += 1
    bal_row = r
    hs.set_cell(ws, r, 1, "기말 잔액", role="label", align=hs.LEFT, bold=True)
    for w in range(n):
        col = get_column_letter(2 + w)
        if w == 0:
            hs.set_cell(ws, r, 2, "=%g+%s%d" % (data.opening, col, net_row),
                        role="calc", number_format=hs.FMT_INT, bold=True)
        else:
            prev = get_column_letter(1 + w)
            hs.set_cell(ws, r, 2 + w, "=%s%d+%s%d" % (prev, bal_row, col, net_row),
                        role="calc", number_format=hs.FMT_INT, bold=True)
    r += 2
    hs.add_line_chart(ws, anchor="A%d" % r, data_min_col=2, data_max_col=last_col,
                      data_min_row=bal_row, data_max_row=bal_row, cat_col=1,
                      title="주간 잔액(유동성)")
    return wb


def qc(wb, data: CashInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    ok_len = len(data.inflows) == data.weeks == len(data.outflows)
    rep.add("주 수 일치", ok_len, "" if ok_len else "inflow/outflow 길이≠weeks")
    # 파이썬 잔액 시뮬 → 최소잔액 음수 경고
    bal = data.opening; min_bal = bal
    for w in range(data.weeks):
        bal += data.inflows[w] - data.outflows[w]; min_bal = min(min_bal, bal)
    rep.add("유동성(최소잔액≥0)", min_bal >= 0, "최소잔액=%.0f" % min_bal)
    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "CashInput", "golden_sample", "build", "qc"]
