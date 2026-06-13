"""fpna.templates.unit_economics — 유닛 이코노믹스(CAC·LTV·ARR bridge).

깊이(C4): LTV 를 *할인 기여이익* 기반으로 정의(현 단순 무한등비 위험 보정).
  - 현 단순화: LTV = ARPU·GM·(1/churn) = 무한 수명·무할인(과대평가 위험).
  - 보강: LTV = Σ_{t} m·(retention^t)/(1+r)^t = m·R/(1+r−R)  (R=1−churn, r=월 할인율)
    유한 호라이즌 닫힌형(geometric). 할인율 0·churn>0 이면 m/churn 으로 수렴(현 식 = 특수해).
임계: LTV/CAC ≥ 3.0(SaaS 관례) · CAC 회수 ≤ 12개월 — QC 가 신호(경고는 detail).
"""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl

from fpna import finance, house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors

TYPE = "unit_economics"

LTV_CAC_FLOOR = 3.0        # SaaS 관례 임계(이하 = 경고)
PAYBACK_CAP_MONTHS = 12.0  # CAC 회수 상한(초과 = 경고)


@dataclass
class UnitEconInput:
    title: str = "유닛 이코노믹스"
    subtitle: str = ""
    unit: str = "₩'000"
    arpu: float = 50.0          # 1인당 월 매출
    gross_margin: float = 0.8   # 매출총이익률
    churn_monthly: float = 0.03 # 월 이탈률
    cac: float = 300.0          # 고객 획득비용
    discount_monthly: float = 0.0  # 월 할인율(0=무할인=현 단순식과 동일)


def golden_sample() -> UnitEconInput:
    return UnitEconInput(title="유닛 이코노믹스 (골든샘플)", subtitle="구조 검증용 — 더미",
                         discount_monthly=0.01)


def _contribution(data: UnitEconInput) -> float:
    """월 기여이익 m = ARPU × GM."""
    return data.arpu * data.gross_margin


def discounted_ltv(data: UnitEconInput):
    """할인 기여이익 LTV 닫힌형. R=1−churn, r=할인율.

    LTV = Σ_{t≥0} m·R^t/(1+r)^t = m·(1+r)/(1+r−R)  (t=0 부터 — 첫 기 기여 포함).
    분모 = churn + r. churn=0 & r=0 이면 0 → None(무한). r=0 이면 정확히 m/churn
    (현 단순식의 특수해 — 무할인이면 기존 동작과 일치, 할인 있으면 시간가치 차감).
    반환 (ltv, lifetime_months).
    """
    m = _contribution(data)
    R = 1.0 - data.churn_monthly
    r = data.discount_monthly
    denom = 1.0 + r - R           # = churn + r
    if denom <= 0:
        return None, None
    ltv = m * (1.0 + r) / denom
    life = (1.0 / data.churn_monthly) if data.churn_monthly > 0 else None
    return ltv, life


def build(data: UnitEconInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("UnitEcon")
    last_col = 2
    hs.set_widths(ws, {1: 28, 2: 14})
    r = hs.report_frame(ws, data.title, subtitle=data.subtitle,
                        unit=data.unit, last_col=last_col)

    # 입력
    r = hs.section_header(ws, r, "입력 (단위: %s)" % data.unit, last_col=last_col)
    inputs = [("ARPU(월)", data.arpu, hs.FMT_INT), ("매출총이익률", data.gross_margin, hs.FMT_PCT1),
              ("월 이탈률", data.churn_monthly, hs.FMT_PCT1), ("월 할인율", data.discount_monthly, hs.FMT_PCT1),
              ("CAC", data.cac, hs.FMT_INT)]
    rows = {}
    for name, v, fmt in inputs:
        hs.set_cell(ws, r, 1, name, role="label", align=hs.LEFT)
        hs.set_cell(ws, r, 2, v, role="input", number_format=fmt)
        rows[name] = r; r += 1
    r += 1
    # 산출(수식)
    r = hs.section_header(ws, r, "산출", last_col=last_col)
    arpu_c = "B%d" % rows["ARPU(월)"]; gm_c = "B%d" % rows["매출총이익률"]
    churn_c = "B%d" % rows["월 이탈률"]; disc_c = "B%d" % rows["월 할인율"]
    cac_c = "B%d" % rows["CAC"]
    # 월 기여이익 m = ARPU·GM
    hs.set_cell(ws, r, 1, "월 기여이익 m=ARPU·GM", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=%s*%s" % (arpu_c, gm_c), role="calc", number_format=hs.FMT_INT)
    m_r = r; r += 1
    # 평균 고객수명(월) = 1/churn
    hs.set_cell(ws, r, 1, "고객수명(월)", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(%s=0,\"\",1/%s)" % (churn_c, churn_c), role="calc",
                number_format=hs.FMT_NUM1)
    life_r = r; r += 1
    # 할인 기여이익 LTV = m·R/(1+r−R), R=1−churn. 분모=churn+r ≤0 이면 NA(무한 방어).
    hs.set_cell(ws, r, 1, "LTV (할인 기여이익)", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2,
                "=IF((%s+%s)<=0,\"\",B%d*(1+%s)/(%s+%s))"
                % (churn_c, disc_c, m_r, disc_c, churn_c, disc_c),
                role="calc", number_format=hs.FMT_INT, bold=True)
    ltv_r = r; r += 1
    hs.set_cell(ws, r, 1, "LTV/CAC (≥%.1f 권장)" % LTV_CAC_FLOOR, role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(OR(%s=0,B%d=\"\"),\"\",B%d/%s)" % (cac_c, ltv_r, ltv_r, cac_c),
                role="calc", number_format=hs.FMT_MULT, bold=True)
    r += 1
    hs.set_cell(ws, r, 1, "CAC 회수(월, ≤%.0f 권장)" % PAYBACK_CAP_MONTHS,
                role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(B%d=0,\"\",%s/B%d)" % (m_r, cac_c, m_r),
                role="calc", number_format=hs.FMT_NUM1)
    r += 2
    # 임계 판정 라인(파이썬 사전계산 — 신호)
    ltv, _ = discounted_ltv(data)
    ratio = finance.ltv_cac(ltv, data.cac) if ltv is not None else None
    m = _contribution(data)
    pb = (data.cac / m) if m > 0 else None
    flags = []
    if ratio is not None and ratio < LTV_CAC_FLOOR:
        flags.append("LTV/CAC=%.1f < %.1f" % (ratio, LTV_CAC_FLOOR))
    if pb is not None and pb > PAYBACK_CAP_MONTHS:
        flags.append("회수=%.1f월 > %.0f" % (pb, PAYBACK_CAP_MONTHS))
    hs.set_cell(ws, r, 1, "건전성", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "양호" if not flags else "주의", role="soft",
                align=hs.LEFT)
    r += 1

    hs.report_footer(ws, r + 1, source="과금 · 코호트 지표",
                     note="; ".join(flags) if flags else "",
                     prepared_by="FP&A", last_col=last_col)
    return wb


def qc(wb, data: UnitEconInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    rep.add("이탈률 (0,1]", 0 < data.churn_monthly <= 1, "churn=%.3f" % data.churn_monthly)
    rep.add("할인율 ≥ 0", data.discount_monthly >= 0, "r=%.4f" % data.discount_monthly)

    ltv, life = discounted_ltv(data)
    rep.add("LTV(할인 기여이익) 계산", ltv is not None,
            "LTV=%.1f" % ltv if ltv is not None else "분모 churn+r ≤ 0(무한 — NA)")

    # 무할인 한정: 닫힌형 LTV == 현 단순식 m/churn 과 일치(특수해 회귀 — 정합성)
    if data.discount_monthly == 0 and data.churn_monthly > 0 and ltv is not None:
        simple = _contribution(data) / data.churn_monthly
        rep.add("무할인 LTV == m/churn(특수해 정합)",
                finance.approx_equal(ltv, simple),
                "" if finance.approx_equal(ltv, simple)
                else "닫힌형=%.6g 단순=%.6g" % (ltv, simple))
    # 할인이 있으면 할인 LTV < 무할인 LTV(시간가치 — 과대평가 보정 확인)
    if data.discount_monthly > 0 and data.churn_monthly > 0 and ltv is not None:
        simple = _contribution(data) / data.churn_monthly
        rep.add("할인 LTV < 무할인 LTV(시간가치 보정)", ltv < simple,
                "할인=%.6g 무할인=%.6g" % (ltv, simple))

    ratio = finance.ltv_cac(ltv, data.cac) if ltv is not None else None
    rep.add("LTV/CAC 계산", ratio is not None,
            "LTV/CAC=%.2f" % ratio if ratio else "계산불가")
    # 임계는 *경고*(detail) — 구조 더미라 게이트로 막지 않음(실데이터서 판단).
    if ratio is not None:
        rep.add("LTV/CAC ≥ %.1f(경고만)" % LTV_CAC_FLOOR, True,
                "OK(%.1f)" % ratio if ratio >= LTV_CAC_FLOOR
                else "주의: %.1f < %.1f" % (ratio, LTV_CAC_FLOOR))

    rep.add("단위 표기", bool(data.unit))
    return rep


__all__ = ["TYPE", "UnitEconInput", "golden_sample", "build", "qc",
           "discounted_ltv", "LTV_CAC_FLOOR", "PAYBACK_CAP_MONTHS"]
