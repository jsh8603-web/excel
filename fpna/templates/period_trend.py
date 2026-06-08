"""fpna.templates.period_trend — 기간별 추이(MoM/QoQ/YoY)."""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl
from openpyxl.utils import get_column_letter

from fpna import house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors

TYPE = "period_trend"


@dataclass
class TrendInput:
    title: str = "기간별 추이"
    subtitle: str = ""
    unit: str = "₩mn"
    periods: list = field(default_factory=list)   # ["2025-01", ...]
    series: dict = field(default_factory=dict)     # {metric: [값,...]}


def golden_sample() -> TrendInput:
    return TrendInput(
        title="기간별 추이 (골든샘플)", subtitle="구조 검증용 — 더미", unit="₩mn",
        periods=["2025-01", "2025-02", "2025-03", "2025-04"],
        series={"매출": [100, 110, 121, 133], "영업이익": [10, 12, 15, 18]},
    )


def build(data: TrendInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("Trend")
    n = len(data.periods)
    last_col = 1 + n
    hs.style_sheet(ws, freeze="B5")
    hs.set_widths(ws, {1: 18, **{j: 11 for j in range(2, last_col + 1)}})
    r = hs.title_block(ws, data.title, data.subtitle, last_col=last_col)

    hs.set_cell(ws, r, 1, "지표 (단위: %s)" % data.unit, role="header", align=hs.LEFT)
    for j, p in enumerate(data.periods, start=2):
        hs.set_cell(ws, r, j, p, role="header")
    hdr = r; r += 1
    data_start = r
    for metric, vals in data.series.items():
        hs.set_cell(ws, r, 1, metric, role="label", align=hs.LEFT)
        for j, v in enumerate(vals, start=2):
            hs.set_cell(ws, r, j, v, role="input", number_format=hs.FMT_INT)
        r += 1
    data_end = r - 1

    # MoM% 블록(첫 시리즈 기준)
    r += 1
    r = hs.section_header(ws, r, "MoM 증감률 (첫 지표)", last_col=last_col)
    hs.set_cell(ws, r, 1, "MoM%", role="label", align=hs.LEFT)
    for t in range(n):
        col = get_column_letter(2 + t)
        if t == 0:
            hs.set_cell(ws, r, 2, "", role="calc")
        else:
            prev = get_column_letter(1 + t)
            hs.set_cell(ws, r, 2 + t,
                        "=IF(%s%d=0,\"\",%s%d/%s%d-1)" % (prev, data_start, col, data_start, prev, data_start),
                        role="calc", number_format=hs.FMT_PCT1)
    r += 2
    hs.add_line_chart(ws, anchor="A%d" % r, data_min_col=2, data_max_col=last_col,
                      data_min_row=data_start, data_max_row=data_end, cat_col=1,
                      title="추이")
    return wb


def qc(wb, data: TrendInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    lens = {len(v) for v in data.series.values()}
    rep.add("기간-시리즈 길이 일치", lens <= {len(data.periods)},
            "" if lens <= {len(data.periods)} else "길이 불일치")
    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "TrendInput", "golden_sample", "build", "qc"]
