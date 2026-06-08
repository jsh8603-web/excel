"""fpna.templates.board_kpi_pack — 이사회 KPI 팩/대시보드."""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl

from fpna import house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors

TYPE = "board_kpi_pack"


@dataclass
class KPI:
    name: str
    actual: float
    target: float
    fmt: str = "int"        # int|pct|mult


@dataclass
class BoardInput:
    title: str = "이사회 KPI 팩"
    subtitle: str = ""
    period: str = ""
    kpis: list = field(default_factory=list)   # list[KPI]


def golden_sample() -> BoardInput:
    return BoardInput(
        title="이사회 KPI 팩 (골든샘플)", subtitle="구조 검증용 — 더미", period="2025-06",
        kpis=[KPI("매출", 1120, 1000), KPI("영업이익률", 0.18, 0.15, "pct"),
              KPI("신규고객", 320, 300), KPI("LTV/CAC", 3.2, 3.0, "mult")],
    )


def _fmt(kind: str) -> str:
    return {"pct": hs.FMT_PCT1, "mult": hs.FMT_MULT}.get(kind, hs.FMT_INT)


def build(data: BoardInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("BoardKPI")
    last_col = 5
    hs.style_sheet(ws, freeze="A6")
    hs.set_widths(ws, {1: 22, 2: 14, 3: 14, 4: 12, 5: 10})
    r = hs.title_block(ws, data.title,
                       (data.subtitle + ("  ·  " + data.period if data.period else "")).strip(" ·"),
                       last_col=last_col)
    for j, h in enumerate(["KPI", "실적", "목표", "달성", "상태"], 1):
        hs.set_cell(ws, r, j, h, role="header", align=hs.LEFT if j == 1 else hs.CENTER)
    r += 1
    for k in data.kpis:
        fmt = _fmt(k.fmt)
        hs.set_cell(ws, r, 1, k.name, role="label", align=hs.LEFT)
        hs.set_cell(ws, r, 2, k.actual, role="input", number_format=fmt)
        hs.set_cell(ws, r, 3, k.target, role="input", number_format=fmt)
        hs.set_cell(ws, r, 4, "=IF(C%d=0,\"\",B%d/C%d)" % (r, r, r), role="calc",
                    number_format=hs.FMT_PCT0)
        # 상태: 달성≥100% = ●(양) 아니면 ▼
        ach = (k.actual / k.target) if k.target else None
        status = "달성" if (ach is not None and ach >= 1.0) else "미달"
        col = hs.POS_FG if status == "달성" else hs.NEG_FG
        c = hs.set_cell(ws, r, 5, status, role="calc", align=hs.CENTER)
        c.font = hs.font(col, bold=True)
        r += 1
    return wb


def qc(wb, data: BoardInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    rep.add("KPI 존재", len(data.kpis) > 0, "n=%d" % len(data.kpis))
    rep.add("목표 0 없음(달성 분모)", all(k.target != 0 for k in data.kpis),
            "목표 0 존재" if any(k.target == 0 for k in data.kpis) else "")
    return rep


__all__ = ["TYPE", "KPI", "BoardInput", "golden_sample", "build", "qc"]
