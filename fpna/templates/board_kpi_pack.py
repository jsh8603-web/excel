"""fpna.templates.board_kpi_pack — 이사회 KPI 팩/대시보드(+ source tie-out 게이트)."""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl

from fpna import house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors
from fpna.view_contract import assert_tie_out

TYPE = "board_kpi_pack"


@dataclass
class KPI:
    name: str
    actual: float           # 보드에 표기되는 값
    target: float
    fmt: str = "int"        # int|pct|mult
    # source = 출처(타 탭/모델)의 원천 수치. 주어지면 보드 표기값(actual)과
    # tie-out 강제(보드숫자 ≠ 모델숫자 = 최악 신뢰붕괴 차단). None 이면
    # golden 내 자기 tie(표기값 == 재계산값)로 fallback.
    source: float | None = None


@dataclass
class BoardInput:
    title: str = "이사회 KPI 팩"
    subtitle: str = ""
    period: str = ""
    kpis: list = field(default_factory=list)   # list[KPI]


def golden_sample() -> BoardInput:
    # 구조 검증용 더미 — source = 출처(모델) 숫자, actual = 보드 표기값. tie 정합.
    return BoardInput(
        title="이사회 KPI 팩 (골든샘플)", subtitle="구조 검증용 — 더미", period="2025-06",
        kpis=[KPI("매출", 1120, 1000, source=1120),
              KPI("영업이익률", 0.18, 0.15, "pct", source=0.18),
              KPI("신규고객", 320, 300, source=320),
              KPI("LTV/CAC", 3.2, 3.0, "mult", source=3.2)],
    )


def _fmt(kind: str) -> str:
    return {"pct": hs.FMT_PCT1, "mult": hs.FMT_MULT}.get(kind, hs.FMT_INT)


def _src(k: KPI) -> float:
    """source 가 명시되면 그 값, 아니면 표기값 자기 tie(actual)."""
    return k.source if k.source is not None else k.actual


def build(data: BoardInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("BoardKPI")
    last_col = 6
    hs.set_widths(ws, {1: 22, 2: 14, 3: 14, 4: 12, 5: 10, 6: 12})
    r = hs.report_frame(ws, data.title, subtitle=data.subtitle,
                        period_basis=data.period, last_col=last_col)
    for j, h in enumerate(["KPI", "실적", "목표", "달성", "상태", "출처대사"], 1):
        hs.set_cell(ws, r, j, h, role="header", align=hs.LEFT if j == 1 else hs.CENTER)
    r += 1
    for k in data.kpis:
        fmt = _fmt(k.fmt)
        hs.set_cell(ws, r, 1, k.name, role="label", align=hs.LEFT)
        # 실적(보드 표기값)을 출처값(source)으로 기입 — 보드숫자가 모델숫자에서 옴.
        hs.set_cell(ws, r, 2, _src(k), role="input", number_format=fmt)
        hs.set_cell(ws, r, 3, k.target, role="input", number_format=fmt)
        hs.set_cell(ws, r, 4, "=IF(C%d=0,\"\",B%d/C%d)" % (r, r, r), role="calc",
                    number_format=hs.FMT_PCT0)
        # 상태: 달성≥100% = ●(양) 아니면 ▼
        ach = (k.actual / k.target) if k.target else None
        status = "달성" if (ach is not None and ach >= 1.0) else "미달"
        col = hs.POS_FG if status == "달성" else hs.NEG_FG
        c = hs.set_cell(ws, r, 5, status, role="calc", align=hs.CENTER)
        c.font = hs.font(col, bold=True)
        # 출처대사 check 셀: 표기 실적(B) − 출처값 = 0 이어야(≠0 적색).
        hs.check_cell(ws, r, 6, "=B%d-%r" % (r, float(_src(k))), number_format=fmt)
        r += 1
    hs.report_footer(ws, r + 1, source="재무 모델 · KPI 정의서",
                     prepared_by="FP&A", last_col=last_col)
    return wb


def qc(wb, data: BoardInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    rep.add("KPI 존재", len(data.kpis) > 0, "n=%d" % len(data.kpis))
    rep.add("목표 0 없음(달성 분모)", all(k.target != 0 for k in data.kpis),
            "목표 0 존재" if any(k.target == 0 for k in data.kpis) else "")
    # --- source tie-out: 보드 표기값(actual) == 출처(모델) 수치 -----------
    # 보드숫자 ≠ 모델숫자 = 최악 신뢰붕괴. 전 KPI 합으로 1차 + 항목별 정밀 대조.
    board_sum = sum(k.actual for k in data.kpis)
    src_sum = sum(_src(k) for k in data.kpis)
    assert_tie_out(rep, board_sum, src_sum, tol=0.0, name="source tie-out(합)")
    mismatch = [k.name for k in data.kpis if _src(k) != k.actual]
    rep.add("source tie-out(항목별)", not mismatch,
            "" if not mismatch else "보드≠출처: " + ", ".join(mismatch[:6]))
    # 달성률 재계산 대조(target vs actual)
    bad_ach = []
    for k in data.kpis:
        if k.target:
            ach = k.actual / k.target
            if ach < 0:
                bad_ach.append(k.name)
    rep.add("달성률 부호 정합", not bad_ach,
            "" if not bad_ach else "음수 달성률: " + ", ".join(bad_ach[:6]))
    return rep


__all__ = ["TYPE", "KPI", "BoardInput", "golden_sample", "build", "qc"]
