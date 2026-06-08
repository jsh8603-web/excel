"""fpna.templates.rolling_forecast — 롤링 포캐스트(실적 + 전망 갱신)."""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl
from openpyxl.utils import get_column_letter

from fpna import house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors

TYPE = "rolling_forecast"


@dataclass
class ForecastInput:
    title: str = "롤링 포캐스트"
    subtitle: str = ""
    unit: str = "₩mn"
    periods: list = field(default_factory=list)
    actual_until: int = 0           # 인덱스 < actual_until = 실적, 이후 = 전망
    series: dict = field(default_factory=dict)   # {metric: [값,...]}


def golden_sample() -> ForecastInput:
    return ForecastInput(
        title="롤링 포캐스트 (골든샘플)", subtitle="구조 검증용 — 더미", unit="₩mn",
        periods=["Q1", "Q2", "Q3", "Q4"], actual_until=2,
        series={"매출": [100, 110, 120, 130]},
    )


def build(data: ForecastInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("Forecast")
    n = len(data.periods); last_col = 1 + n
    hs.style_sheet(ws, freeze="B5")
    hs.set_widths(ws, {1: 18, **{j: 11 for j in range(2, last_col + 1)}})
    r = hs.title_block(ws, data.title, data.subtitle, last_col=last_col)

    hs.set_cell(ws, r, 1, "지표 (단위: %s)" % data.unit, role="header", align=hs.LEFT)
    for j, p in enumerate(data.periods, start=2):
        tag = " (A)" if (j - 2) < data.actual_until else " (F)"
        hs.set_cell(ws, r, j, str(p) + tag, role="header")
    r += 1
    data_start = r
    for metric, vals in data.series.items():
        hs.set_cell(ws, r, 1, metric, role="label", align=hs.LEFT)
        for j, v in enumerate(vals, start=2):
            # 실적=입력(파랑), 전망=입력(파랑)이되 음영으로 구분
            role = "input"
            fill = None if (j - 2) < data.actual_until else hs.FILL_BAND
            hs.set_cell(ws, r, j, v, role=role, number_format=hs.FMT_INT, fill=fill)
        r += 1
    data_end = r - 1

    # FY 합계
    r += 1
    hs.set_cell(ws, r, 1, "FY 합계", role="total", align=hs.LEFT, bold=True)
    fc = get_column_letter(2); lc = get_column_letter(last_col)
    hs.set_cell(ws, r, 2, "=SUM(%s%d:%s%d)" % (fc, data_start, lc, data_start),
                role="calc", number_format=hs.FMT_INT, bold=True)
    r += 2
    hs.add_line_chart(ws, anchor="A%d" % r, data_min_col=2, data_max_col=last_col,
                      data_min_row=data_start, data_max_row=data_end, cat_col=1,
                      title="실적+전망")
    return wb


def qc(wb, data: ForecastInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    rep.add("actual_until 범위", 0 <= data.actual_until <= len(data.periods),
            "actual_until=%d" % data.actual_until)
    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "ForecastInput", "golden_sample", "build", "qc"]
