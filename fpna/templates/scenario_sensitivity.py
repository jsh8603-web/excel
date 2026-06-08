"""fpna.templates.scenario_sensitivity — 시나리오/민감도(토네이도)."""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl

from fpna import house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors

TYPE = "scenario_sensitivity"


@dataclass
class Driver:
    name: str
    base: float
    low: float
    high: float
    impact_per_unit: float    # 결과지표에 대한 단위당 민감도


@dataclass
class ScenarioInput:
    title: str = "시나리오·민감도"
    subtitle: str = ""
    unit: str = "₩mn"
    base_outcome: float = 1000.0
    drivers: list = field(default_factory=list)   # list[Driver]


def golden_sample() -> ScenarioInput:
    return ScenarioInput(
        title="시나리오·민감도 (골든샘플)", subtitle="구조 검증용 — 더미", unit="₩mn",
        base_outcome=1000,
        drivers=[Driver("판매량", 100, 90, 115, 5),
                 Driver("단가", 50, 45, 55, 8),
                 Driver("원가율", 0.6, 0.65, 0.55, -600)],
    )


def build(data: ScenarioInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("Scenario")
    last_col = 4
    hs.style_sheet(ws, freeze="A6")
    hs.set_widths(ws, {1: 18, 2: 14, 3: 14, 4: 14})
    r = hs.title_block(ws, data.title, data.subtitle, last_col=last_col)
    hs.set_cell(ws, r, 1, "Base 결과 (단위: %s)" % data.unit, role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, data.base_outcome, role="input", number_format=hs.FMT_INT, bold=True)
    r += 2

    for j, h in enumerate(["드라이버", "Low 영향", "High 영향", "스윙(|폭|)"], 1):
        hs.set_cell(ws, r, j, h, role="header", align=hs.LEFT if j == 1 else hs.CENTER)
    r += 1
    data_start = r
    rows = []
    for d in data.drivers:
        low_impact = (d.low - d.base) * d.impact_per_unit
        high_impact = (d.high - d.base) * d.impact_per_unit
        swing = abs(high_impact - low_impact)
        rows.append((d.name, low_impact, high_impact, swing))
    # 스윙 큰 순(토네이도)
    rows.sort(key=lambda x: -x[3])
    for name, lo, hi, sw in rows:
        hs.set_cell(ws, r, 1, name, role="label", align=hs.LEFT)
        hs.set_cell(ws, r, 2, lo, role="calc", number_format=hs.FMT_INT)
        hs.set_cell(ws, r, 3, hi, role="calc", number_format=hs.FMT_INT)
        hs.set_cell(ws, r, 4, sw, role="calc", number_format=hs.FMT_INT)
        r += 1
    data_end = r - 1
    r += 1
    hs.add_bar_chart(ws, anchor="A%d" % r, data_min_col=4, data_max_col=4,
                     data_min_row=data_start, data_max_row=data_end, cat_col=1,
                     title="토네이도(스윙 크기)")
    return wb


def qc(wb, data: ScenarioInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    rep.add("드라이버 존재", len(data.drivers) > 0, "n=%d" % len(data.drivers))
    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "Driver", "ScenarioInput", "golden_sample", "build", "qc"]
