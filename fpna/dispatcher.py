"""
fpna.dispatcher — 요청 텍스트 + tidy 데이터 컬럼으로 템플릿 유형 라우팅.

순차 cascade(빠른 판정). 판정 신호 = (a) 요청 텍스트 키워드 (b) tidy 컬럼 단서
(예: budget&actual 동시 존재 → variance, 기간축만 → period_trend).
"""
from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass
class DispatchResult:
    template: str
    reason: str
    score: int = 0


# (template, [키워드정규식]) — 위에서부터 우선
_RULES = [
    # 고정비 FP&A — 구체 키워드라 일반 variance/손익보다 앞에 둔다.
    # forward_da 는 감가 단어를 쓰므로 depreciation_schedule 보다 앞(미래/forward 한정).
    ("fc_forward_da", [r"미래\s*감가", r"forward\s*d&?a", r"감가.*투영", r"투영.*감가",
                       r"미래\s*상각", r"capex.*감가", r"향후\s*감가"]),
    ("fc_depreciation_schedule", [r"감가\s*상각", r"depreciation", r"상각\s*스케줄",
                                  r"내용\s*연수", r"자산\s*대장.*상각"]),
    ("fc_runrate_normalized", [r"런\s*레이트", r"run[\s-]*rate", r"정규화.*비용",
                               r"연환산", r"normaliz", r"1회성\s*제외", r"베이스라인"]),
    ("fc_cuttability_ladder", [r"절감\s*가능", r"cuttab", r"해지\s*가능", r"비용\s*절감",
                               r"고정비\s*절감", r"time[\s-]*to[\s-]*exit", r"감축\s*여력"]),
    ("fc_driver_unitcost", [r"동인\s*단가", r"단위\s*원가", r"unit\s*cost", r"대당",
                            r"㎡당", r"면적당", r"kwh당", r"활동\s*동인"]),
    ("fc_prepaid_rollforward", [r"선급", r"prepaid", r"롤\s*포워드", r"roll[\s-]*forward",
                                r"선급비용\s*상각"]),
    ("fc_variance_bridge", [r"고정비.*브리지", r"고정비.*변동", r"고정비.*요인",
                            r"고정비.*walk", r"고정비.*워크", r"고정비\s*bridge"]),
    ("fc_maturity_wall", [r"만기\s*도래", r"maturity\s*wall", r"약정\s*만기",
                          r"갱신\s*도래", r"만기\s*벽"]),
    ("investment_appraisal", [r"npv", r"irr", r"투자\s*타당성", r"회수\s*기간",
                              r"payback", r"할인\s*현금", r"capex", r"투자\s*검토"]),
    ("variance", [r"예실", r"plan\s*vs\s*actual", r"예산\s*대비\s*실적",
                  r"변동\s*분석", r"variance", r"차이\s*분석", r"bridge", r"브리지"]),
    ("period_trend", [r"mom", r"qoq", r"yoy", r"전월\s*대비", r"전분기",
                      r"전년\s*대비", r"추이", r"trend", r"기간별"]),
    ("rolling_forecast", [r"롤링", r"rolling", r"포캐스트\s*갱신", r"forecast\s*update",
                          r"re-?forecast", r"전망\s*갱신"]),
    ("budget_build", [r"예산\s*수립", r"budget\s*build", r"인건비", r"headcount",
                      r"인원\s*계획", r"예산\s*편성"]),
    ("cashflow_13w", [r"13\s*주", r"13[\s-]*week", r"단기\s*현금", r"주간\s*현금",
                      r"유동성", r"liquidity", r"자금\s*수지"]),
    ("unit_economics", [r"unit\s*economics", r"유닛\s*이코노믹스", r"cac", r"ltv",
                        r"arr", r"코호트", r"cohort", r"구독", r"churn", r"리텐션"]),
    ("scenario_sensitivity", [r"시나리오", r"민감도", r"scenario", r"sensitivity",
                              r"토네이도", r"tornado", r"what[\s-]*if", r"데이터\s*테이블"]),
    ("pnl_3statement", [r"손익", r"p&l", r"pnl", r"3\s*statement", r"재무제표",
                        r"손익계산서", r"income\s*statement"]),
    ("board_kpi_pack", [r"이사회", r"board", r"kpi", r"대시보드", r"dashboard",
                        r"경영\s*보고", r"월간\s*보고"]),
]


def _columns_signal(columns: list[str] | None) -> tuple[str | None, str]:
    """tidy 컬럼/메트릭 단서로 보조 판정."""
    if not columns:
        return None, ""
    low = [str(c).lower() for c in columns]
    joined = " ".join(low)
    has_budget = any("budget" in c or "계획" in c or "plan" in c or "예산" in c for c in low)
    has_actual = any("actual" in c or "실적" in c for c in low)
    if has_budget and has_actual:
        return "variance", "컬럼에 계획+실적 동시 존재"
    if re.search(r"cac|ltv|arr|churn|mrr", joined):
        return "unit_economics", "컬럼에 구독지표"
    if re.search(r"cash|현금|자금", joined):
        return "cashflow_13w", "컬럼에 현금 단서"
    return None, ""


def dispatch(request_text: str = "", *, columns: list[str] | None = None,
             metrics: list[str] | None = None) -> DispatchResult:
    """요청 텍스트 + 컬럼 단서 → DispatchResult.

    텍스트 키워드를 우선 cascade 로 평가, 동점/무매칭 시 컬럼 신호로 보강,
    그래도 없으면 기본값 pnl_3statement / board_kpi_pack.
    """
    text = (request_text or "").lower()
    pool = " ".join(filter(None, [text] + [str(m).lower() for m in (metrics or [])]))

    best = None
    for tmpl, pats in _RULES:
        hits = sum(1 for p in pats if re.search(p, pool))
        if hits and (best is None or hits > best.score):
            best = DispatchResult(tmpl, "텍스트 키워드 %d개 매칭" % hits, hits)

    col_tmpl, col_reason = _columns_signal(columns)
    if best is None and col_tmpl:
        return DispatchResult(col_tmpl, col_reason, 1)
    if best is None:
        return DispatchResult("pnl_3statement", "무매칭 → 기본값(손익)", 0)
    # 컬럼이 variance 를 강하게 시사하면 보정(계획+실적 동시존재는 강한 구조 신호).
    # 텍스트 매칭이 약하거나(score≤1) variance 가 아닌 일반 보고류면 컬럼 신호 우선.
    if col_tmpl == "variance" and best.template != "variance" and (
            best.score <= 1
            or best.template in ("period_trend", "pnl_3statement", "board_kpi_pack")):
        return DispatchResult("variance", "컬럼(계획+실적)이 텍스트보다 강함", best.score + 1)
    return best


__all__ = ["dispatch", "DispatchResult"]
