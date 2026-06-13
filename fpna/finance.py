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


# --- 감가상각(정액법) — dims 비의존 순수 원시값 인터페이스 ---
def straight_line_depreciation(acq_cost: float, salvage: float,
                               life_months: int) -> float:
    """정액법 월 상각액 = (취득가 - 잔존가) / 내용연수(월)."""
    if life_months <= 0:
        return 0.0
    return (acq_cost - salvage) / life_months


def depreciation_schedule(acq_cost: float, salvage: float, life_months: int, *,
                          n_periods: int, start_index: int = 0
                          ) -> list[tuple[float, float, float]]:
    """기간별 (opening, dep, closing) 장부가 스케줄.

    start_index = 가동(in-service) 시작 period 의 0-base 인덱스 (그 전 기간은 dep=0).
    내용연수 마지막 달은 잔존가에 정확히 닿도록 잔액을 상각(부동소수 오차 흡수).
    QC(R11)가 이 스케줄 합을 GL 상각비와 대사한다.
    """
    monthly = straight_line_depreciation(acq_cost, salvage, life_months)
    rows: list[tuple[float, float, float]] = []
    accum = 0.0
    done = 0
    for i in range(n_periods):
        opening = acq_cost - accum
        if i >= start_index and done < life_months:
            dep = monthly
            if done == life_months - 1:        # 마지막 달: 잔존가에 정확히 도달
                dep = opening - salvage
            accum += dep
            done += 1
        else:
            dep = 0.0
        rows.append((opening, dep, acq_cost - accum))
    return rows


def depreciation_schedule_ext(acq_cost: float, salvage: float, life_months: int, *,
                              n_periods: int, start_index: int = 0,
                              first_period_factor: float = 1.0,
                              disposal_index: int | None = None,
                              impair_index: int | None = None,
                              impair_to: float | None = None
                              ) -> list[tuple[float, float, float]]:
    """확장 정액법 스케줄 — C11 보완(부분월/처분/손상). 기존 함수는 그대로 유지.

    기존 depreciation_schedule 의 상위호환 래퍼(모든 옵션 기본값 = 원래 동작).

    first_period_factor
        가동 1차월 일할/반월 비율(0<f≤1). 1차월 상각 = monthly*f, 부족분은 정상적으로
        내용연수 마지막 달이 잔액을 흡수(closing=salvage 보존). 이중연환산 방지.
    disposal_index
        처분 period 의 0-base 인덱스. 그 period 부터(포함) 상각 중단·closing 고정.
        처분손익(carrying-처분가)은 범위 밖(메모만) — 여기선 D&A 중단만.
    impair_index / impair_to
        손상 이벤트 period(0-base) 와 손상 후 carrying(=새 depreciable base 상단).
        그 period 의 opening 을 impair_to 로 base-reset, 잔여 내용연수에 prospective 재배분.

    반환 = [(opening, dep, closing), ...] (n_periods 행, 전수 유지).
    closing 누적 정합: Σdep = (취득가-잔존) 에서 처분/손상으로 조정된 양만큼 차감.
    """
    if life_months <= 0:
        return [(acq_cost, 0.0, acq_cost) for _ in range(n_periods)]

    monthly = straight_line_depreciation(acq_cost, salvage, life_months)
    rows: list[tuple[float, float, float]] = []
    book = acq_cost            # 현재 장부가(opening)
    done = 0                   # 소진한 내용연수(월) — 마지막 달 흡수 판정
    base_salvage = salvage     # base-reset(손상) 후 잔존가는 유지
    disposed = False

    for i in range(n_periods):
        opening = book

        # 손상: 이 period 의 opening 을 impair_to 로 base-reset(prospective 재배분)
        if impair_index is not None and i == impair_index and impair_to is not None:
            opening = impair_to
            book = impair_to
            remain = max(life_months - done, 1)
            monthly = max(book - base_salvage, 0.0) / remain   # 잔여기간 재배분

        if disposed or (disposal_index is not None and i >= disposal_index):
            disposed = True
            rows.append((opening, 0.0, opening))   # 처분 후 D&A 중단·장부가 고정
            continue

        if i >= start_index and done < life_months:
            f = first_period_factor if i == start_index else 1.0
            dep = monthly * f
            if done == life_months - 1:            # 마지막 달: 잔존가에 정확히 도달
                dep = opening - base_salvage
            dep = min(dep, max(opening - base_salvage, 0.0))   # 음수/과상각 방지
            book = opening - dep
            done += 1
        else:
            dep = 0.0
        rows.append((opening, dep, book))
    return rows


__all__ = [
    "npv", "irr", "discounted_payback", "payback", "cagr",
    "variance", "variance_pct", "safe_div",
    "gross_margin", "operating_margin", "current_ratio", "ltv_cac",
    "approx_equal",
    "straight_line_depreciation", "depreciation_schedule",
    "depreciation_schedule_ext",
]
