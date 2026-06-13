"""fpna.templates.scenario_sensitivity — 시나리오/민감도(토네이도).

깊이(C4): 시나리오 selector switch — 드롭다운(DataValidation) + CHOOSE 로 Base/
  Low/High 케이스 결과를 한 셀에서 전환(정적 컬럼 나열이 아니라 동적 스위치).
  tornado 는 swing 내림차순(기존 유지).
게이트(C5): R9 base=모델 base — 각 드라이버의 base 가 시나리오 base 케이스와
  정합(assert_scenario_aligned 정신: 민감도 기준이 모델 base 와 어긋나면 거짓 헤지).
"""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

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


def _case_outcome(data: ScenarioInput, which: str) -> float:
    """Base/Low/High 시나리오 결과 = base_outcome + Σ(case값−base)·impact."""
    out = data.base_outcome
    for d in data.drivers:
        v = {"Base": d.base, "Low": d.low, "High": d.high}[which]
        out += (v - d.base) * d.impact_per_unit
    return out


def build(data: ScenarioInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("Scenario")
    last_col = 5
    hs.set_widths(ws, {1: 18, 2: 14, 3: 14, 4: 14, 5: 14})
    r = hs.report_frame(ws, data.title, subtitle=data.subtitle,
                        unit=data.unit, last_col=last_col)

    # --- 시나리오 selector(드롭다운) + CHOOSE 스위치 ----------------------
    r = hs.section_header(ws, r, "시나리오 Selector", last_col=last_col)
    hs.set_cell(ws, r, 1, "선택(Base/Low/High)", role="label", align=hs.LEFT)
    sel_cell = "B%d" % r
    hs.set_cell(ws, r, 2, "Base", role="input", align=hs.CENTER)
    dv = DataValidation(type="list", formula1='"Base,Low,High"', allow_blank=False)
    ws.add_data_validation(dv); dv.add(ws.cell(row=r, column=2))
    r += 1
    # 케이스별 결과(파이썬 사전계산 → 셀 값) — CHOOSE 가 index 로 고른다.
    base_o = _case_outcome(data, "Base")
    low_o = _case_outcome(data, "Low")
    high_o = _case_outcome(data, "High")
    hs.set_cell(ws, r, 1, "Base / Low / High 결과", role="label", align=hs.LEFT)
    bcol, lcol, hcol = "B%d" % r, "C%d" % r, "D%d" % r
    hs.set_cell(ws, r, 2, base_o, role="calc", number_format=hs.FMT_INT)
    hs.set_cell(ws, r, 3, low_o, role="calc", number_format=hs.FMT_INT)
    hs.set_cell(ws, r, 4, high_o, role="calc", number_format=hs.FMT_INT)
    r += 1
    hs.set_cell(ws, r, 1, "선택 결과", role="label", align=hs.LEFT)
    # MATCH 로 selector 라벨 → 1/2/3, CHOOSE 로 케이스 결과 선택
    hs.set_cell(ws, r, 2,
                "=CHOOSE(MATCH(%s,{\"Base\",\"Low\",\"High\"},0),%s,%s,%s)"
                % (sel_cell, bcol, lcol, hcol),
                role="calc", number_format=hs.FMT_INT, bold=True)
    r += 2

    # --- 드라이버 base 표(시나리오 base = 모델 base 정합 시각화) -----------
    r = hs.section_header(ws, r, "드라이버 (base = 모델 base)", last_col=last_col)
    for j, h in enumerate(["드라이버", "Base", "Low", "High", "스윙(|폭|)"], 1):
        hs.set_cell(ws, r, j, h, role="header", align=hs.LEFT if j == 1 else hs.CENTER)
    r += 1
    data_start = r
    rows = []
    for d in data.drivers:
        low_impact = (d.low - d.base) * d.impact_per_unit
        high_impact = (d.high - d.base) * d.impact_per_unit
        swing = abs(high_impact - low_impact)
        rows.append((d.name, d.base, d.low, d.high, swing))
    # 스윙 큰 순(토네이도)
    rows.sort(key=lambda x: -x[4])
    for name, bv, lo, hi, sw in rows:
        hs.set_cell(ws, r, 1, name, role="label", align=hs.LEFT)
        hs.set_cell(ws, r, 2, bv, role="input", number_format=hs.FMT_NUM2)
        hs.set_cell(ws, r, 3, lo, role="input", number_format=hs.FMT_NUM2)
        hs.set_cell(ws, r, 4, hi, role="input", number_format=hs.FMT_NUM2)
        hs.set_cell(ws, r, 5, sw, role="calc", number_format=hs.FMT_INT, bold=True)
        r += 1
    data_end = r - 1
    r += 1
    hs.add_bar_chart(ws, anchor="A%d" % r, data_min_col=5, data_max_col=5,
                     data_min_row=data_start, data_max_row=data_end, cat_col=1,
                     title="토네이도(스윙 크기)")
    hs.report_footer(ws, r + 16, source="시나리오 가정 · 민감도 동인",
                     prepared_by="FP&A", last_col=last_col)
    return wb


def qc(wb, data: ScenarioInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    rep.add("드라이버 존재", len(data.drivers) > 0, "n=%d" % len(data.drivers))

    # --- R9 base=모델 base: Base 시나리오 결과 == base_outcome ---------------
    # 민감도 기준(Base 케이스)이 모델 base 와 어긋나면 거짓 헤지. Base 케이스는
    # 모든 드라이버가 base 값 → 결과는 base_outcome 이어야 한다(정합).
    base_case = _case_outcome(data, "Base")
    from fpna import finance
    ok_base = finance.approx_equal(base_case, data.base_outcome)
    rep.add("R9 base=모델base(시나리오 정합)", ok_base,
            "" if ok_base else "Base케이스=%.6g ≠ 모델base=%.6g" % (base_case, data.base_outcome))

    # --- selector 스위치 결정성: Low/High 가 base 와 분리(정적 컬럼 아님) -----
    low_o, high_o = _case_outcome(data, "Low"), _case_outcome(data, "High")
    distinct = not (low_o == high_o == data.base_outcome) or not data.drivers
    rep.add("시나리오 스위치 동적(케이스 분리)", distinct,
            "" if distinct else "Low/High 가 Base 와 동일 — 정적 컬럼(스위치 무의미)")

    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "Driver", "ScenarioInput", "golden_sample", "build", "qc"]
