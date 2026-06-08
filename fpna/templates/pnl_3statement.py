"""fpna.templates.pnl_3statement — 손익계산서(단순 3-statement 연결 골격)."""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl

from fpna import house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors, qc_totals

TYPE = "pnl_3statement"


@dataclass
class PnLInput:
    title: str = "손익계산서"
    subtitle: str = ""
    unit: str = "₩mn"
    revenue: float = 1000.0
    cogs: float = 600.0
    sga: float = 200.0
    da: float = 50.0           # 감가상각
    interest: float = 20.0
    tax_rate: float = 0.22


def golden_sample() -> PnLInput:
    return PnLInput(title="손익계산서 (골든샘플)", subtitle="구조 검증용 — 더미")


def build(data: PnLInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("PnL")
    last_col = 2
    hs.style_sheet(ws, freeze="A5")
    hs.set_widths(ws, {1: 28, 2: 16})
    r = hs.title_block(ws, data.title, data.subtitle, last_col=last_col)
    hs.set_cell(ws, r, 1, "항목 (단위: %s)" % data.unit, role="header", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "금액", role="header"); r += 1

    def line(name, value=None, formula=None, *, total=False, indent=0, sign="+"):
        nonlocal r
        hs.set_cell(ws, r, 1, name, role="total" if total else "label",
                    align=hs.indent(indent), bold=total)
        if formula:
            hs.set_cell(ws, r, 2, formula, role="calc", number_format=hs.FMT_INT, bold=total)
        else:
            hs.set_cell(ws, r, 2, value, role="input", number_format=hs.FMT_INT)
        if total:
            for j in range(1, last_col + 1):
                ws.cell(row=r, column=j).border = hs.BORDER_TOP
        cur = r; r += 1; return cur

    rev = line("매출", data.revenue)
    cogs = line("(-) 매출원가", data.cogs, indent=1)
    gp = line("매출총이익", formula="=B%d-B%d" % (rev, cogs), total=True)
    sga = line("(-) 판관비", data.sga, indent=1)
    da = line("(-) 감가상각", data.da, indent=1)
    ebit = line("영업이익(EBIT)", formula="=B%d-B%d-B%d" % (gp, sga, da), total=True)
    inte = line("(-) 이자비용", data.interest, indent=1)
    ebt = line("세전이익(EBT)", formula="=B%d-B%d" % (ebit, inte), total=True)
    hs.set_cell(ws, r, 1, "법인세율", role="label", align=hs.indent(1))
    hs.set_cell(ws, r, 2, data.tax_rate, role="input", number_format=hs.FMT_PCT1)
    tax_rate_r = r; r += 1
    tax = line("(-) 법인세", formula="=MAX(0,B%d)*B%d" % (ebt, tax_rate_r), indent=1)
    ni = line("당기순이익", formula="=B%d-B%d" % (ebt, tax), total=True)

    # 마진 요약
    r += 1
    r = hs.section_header(ws, r, "마진", last_col=last_col)
    hs.set_cell(ws, r, 1, "영업이익률", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(B%d=0,\"\",B%d/B%d)" % (rev, ebit, rev), role="calc",
                number_format=hs.FMT_PCT1); r += 1
    hs.set_cell(ws, r, 1, "순이익률", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(B%d=0,\"\",B%d/B%d)" % (rev, ni, rev), role="calc",
                number_format=hs.FMT_PCT1); r += 1
    return wb


def qc(wb, data: PnLInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    gp = data.revenue - data.cogs
    ebit = gp - data.sga - data.da
    ebt = ebit - data.interest
    ni = ebt - max(0, ebt) * data.tax_rate
    qc_totals("매출총이익", gp, data.revenue - data.cogs, rep)
    rep.add("EBIT 계산", True, "EBIT=%.0f" % ebit)
    rep.add("당기순이익 계산", True, "NI=%.0f" % ni)
    rep.add("세율 [0,1)", 0 <= data.tax_rate < 1, "tax=%.2f" % data.tax_rate)
    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "PnLInput", "golden_sample", "build", "qc"]
