"""
fpna.finance — 순수 파이썬 재무계산(numpy 미사용).

NPV/IRR/payback/CAGR/variance/비율. QC 단계가 셀 수식값을 재계산해 대조하는 데도 쓴다.
모든 함수는 결정적이며 부동소수 안정성에 유의(이분법/뉴턴 혼합).
"""
from __future__ import annotations

import math


def npv(rate: float, cashflows: list[float]) -> float:
    """NPV. cashflows[0] = t0(보통 음의 투자). rate = 기간 할인율."""
    return sum(cf / (1.0 + rate) ** t for t, cf in enumerate(cashflows))


def _npv_deriv(rate: float, cashflows: list[float]) -> float:
    return sum(-t * cf / (1.0 + rate) ** (t + 1) for t, cf in enumerate(cashflows))


def irr(cashflows: list[float], *, guess: float = 0.1,
        tol: float = 1e-7, max_iter: int = 200) -> float | None:
    """IRR. 뉴턴법 시도 후 실패하면 이분법 fallback. 해 없으면 None.

    부호 변화가 없으면(전부 +또는 -) IRR 미정 → None.
    """
    signs = {1 if cf > 0 else (-1 if cf < 0 else 0) for cf in cashflows}
    if not ({1, -1} <= signs):
        return None

    # 뉴턴법
    rate = guess
    for _ in range(max_iter):
        f = npv(rate, cashflows)
        if abs(f) < tol:
            return rate
        d = _npv_deriv(rate, cashflows)
        if d == 0:
            break
        new = rate - f / d
        if new <= -0.999999:    # 발산 방지
            break
        if abs(new - rate) < tol:
            return new
        rate = new

    # 이분법 fallback: [-0.9999, 10] 구간 부호변화 탐색
    lo, hi = -0.9999, 10.0
    f_lo, f_hi = npv(lo, cashflows), npv(hi, cashflows)
    if f_lo * f_hi > 0:
        return None
    for _ in range(max_iter):
        mid = (lo + hi) / 2
        f_mid = npv(mid, cashflows)
        if abs(f_mid) < tol:
            return mid
        if f_lo * f_mid < 0:
            hi, f_hi = mid, f_mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2


def discounted_payback(rate: float, cashflows: list[float]) -> float | None:
    """할인 회수기간(기간 단위, 선형보간). 회수 못 하면 None."""
    cum = 0.0
    prev_cum = 0.0
    for t, cf in enumerate(cashflows):
        disc = cf / (1.0 + rate) ** t
        prev_cum = cum
        cum += disc
        if cum >= 0 and t > 0:
            # prev_cum < 0 <= cum 사이 보간
            if disc == 0:
                return float(t)
            frac = -prev_cum / disc
            return (t - 1) + frac
    return None


def payback(cashflows: list[float]) -> float | None:
    """단순 회수기간(비할인)."""
    return discounted_payback(0.0, cashflows)


def cagr(begin: float, end: float, periods: float) -> float | None:
    if begin <= 0 or periods <= 0:
        return None
    return (end / begin) ** (1.0 / periods) - 1.0


def variance(actual: float, plan: float) -> float:
    """절대 차이(실적 - 계획)."""
    return actual - plan


def variance_pct(actual: float, plan: float) -> float | None:
    if plan == 0:
        return None
    return (actual - plan) / abs(plan)


def safe_div(num: float, den: float) -> float | None:
    return None if den == 0 else num / den


# --- 흔한 재무비율(분모 0 방어) ---
def gross_margin(revenue: float, cogs: float) -> float | None:
    return safe_div(revenue - cogs, revenue)


def operating_margin(operating_income: float, revenue: float) -> float | None:
    return safe_div(operating_income, revenue)


def current_ratio(current_assets: float, current_liab: float) -> float | None:
    return safe_div(current_assets, current_liab)


def ltv_cac(ltv: float, cac: float) -> float | None:
    return safe_div(ltv, cac)


def approx_equal(a, b, *, rel: float = 1e-6, abs_: float = 1e-6) -> bool:
    if a is None or b is None:
        return a is b
    return math.isclose(a, b, rel_tol=rel, abs_tol=abs_)


__all__ = [
    "npv", "irr", "discounted_payback", "payback", "cagr",
    "variance", "variance_pct", "safe_div",
    "gross_margin", "operating_margin", "current_ratio", "ltv_cac",
    "approx_equal",
]
