"""fpna.templates.unit_economics — 유닛 이코노믹스(CAC·LTV·ARR bridge)."""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl

from fpna import finance, house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors

TYPE = "unit_economics"


@dataclass
class UnitEconInput:
    title: str = "유닛 이코노믹스"
    subtitle: str = ""
    unit: str = "₩'000"
    arpu: float = 50.0          # 1인당 월 매출
    gross_margin: float = 0.8   # 매출총이익률
    churn_monthly: float = 0.03 # 월 이탈률
    cac: float = 300.0          # 고객 획득비용


def golden_sample() -> UnitEconInput:
    return UnitEconInput(title="유닛 이코노믹스 (골든샘플)", subtitle="구조 검증용 — 더미")


def build(data: UnitEconInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("UnitEcon")
    last_col = 2
    hs.style_sheet(ws, freeze="A5")
    hs.set_widths(ws, {1: 26, 2: 14})
    r = hs.title_block(ws, data.title, data.subtitle, last_col=last_col)

    # 입력
    r = hs.section_header(ws, r, "입력 (단위: %s)" % data.unit, last_col=last_col)
    inputs = [("ARPU(월)", data.arpu, hs.FMT_INT), ("매출총이익률", data.gross_margin, hs.FMT_PCT1),
              ("월 이탈률", data.churn_monthly, hs.FMT_PCT1), ("CAC", data.cac, hs.FMT_INT)]
    rows = {}
    for name, v, fmt in inputs:
        hs.set_cell(ws, r, 1, name, role="label", align=hs.LEFT)
        hs.set_cell(ws, r, 2, v, role="input", number_format=fmt)
        rows[name] = r; r += 1
    r += 1
    # 산출(수식)
    r = hs.section_header(ws, r, "산출", last_col=last_col)
    arpu_c = "B%d" % rows["ARPU(월)"]; gm_c = "B%d" % rows["매출총이익률"]
    churn_c = "B%d" % rows["월 이탈률"]; cac_c = "B%d" % rows["CAC"]
    # 평균 고객수명(월) = 1/churn, LTV = ARPU*GM*lifetime
    hs.set_cell(ws, r, 1, "고객수명(월)", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(%s=0,\"\",1/%s)" % (churn_c, churn_c), role="calc",
                number_format=hs.FMT_NUM1)
    life_r = r; r += 1
    hs.set_cell(ws, r, 1, "LTV", role="label", align=hs.LEFT)
    # 고객수명이 ""(churn=0/미입력)이면 텍스트 곱셈 #VALUE! → 방어
    hs.set_cell(ws, r, 2, "=IF(B%d=\"\",\"\",%s*%s*B%d)" % (life_r, arpu_c, gm_c, life_r),
                role="calc", number_format=hs.FMT_INT, bold=True)
    ltv_r = r; r += 1
    hs.set_cell(ws, r, 1, "LTV/CAC", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(OR(%s=0,B%d=\"\"),\"\",B%d/%s)" % (cac_c, ltv_r, ltv_r, cac_c),
                role="calc", number_format=hs.FMT_MULT, bold=True)
    r += 1
    hs.set_cell(ws, r, 1, "CAC 회수(월)", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(%s*%s=0,\"\",%s/(%s*%s))" % (arpu_c, gm_c, cac_c, arpu_c, gm_c),
                role="calc", number_format=hs.FMT_NUM1)
    r += 1
    return wb


def qc(wb, data: UnitEconInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    rep.add("이탈률 (0,1]", 0 < data.churn_monthly <= 1, "churn=%.3f" % data.churn_monthly)
    life = finance.safe_div(1.0, data.churn_monthly)
    ltv = (data.arpu * data.gross_margin * life) if life else None
    ratio = finance.ltv_cac(ltv, data.cac) if ltv else None
    rep.add("LTV/CAC 계산", ratio is not None,
            "LTV/CAC=%.2f" % ratio if ratio else "계산불가")
    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "UnitEconInput", "golden_sample", "build", "qc"]
