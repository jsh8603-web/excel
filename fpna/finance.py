"""
fpna.finance — 순수 파이썬 재무계산(numpy 미사용).

NPV/IRR/payback/CAGR/variance/비율. QC 단계가 셀 수식값을 재계산해 대조하는 데도 쓴다.
모든 함수는 결정적이며 부동소수 안정성에 유의(이분법/뉴턴 혼합).
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field


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


def mirr(cashflows: list[float], finance_rate: float, reinvest_rate: float
         ) -> float | None:
    """MIRR(수정 IRR). 음 현금흐름은 finance_rate 로 t0 까지 할인(PV),
    양 현금흐름은 reinvest_rate 로 마지막 기간까지 복리(FV).

    MIRR = (FV_positive / -PV_negative) ** (1/n) - 1.
    음 흐름 또는 양 흐름이 없으면(부호 한쪽뿐) 미정 → None.
    Excel MIRR() 정의와 일치한다(QC 가 셀 수식과 대조).
    """
    n = len(cashflows) - 1
    if n <= 0:
        return None
    pv_neg = 0.0
    fv_pos = 0.0
    for t, cf in enumerate(cashflows):
        if cf < 0:
            pv_neg += cf / (1.0 + finance_rate) ** t
        elif cf > 0:
            fv_pos += cf * (1.0 + reinvest_rate) ** (n - t)
    if pv_neg == 0 or fv_pos == 0:
        return None                       # 한쪽 부호만 → MIRR 미정
    return (fv_pos / -pv_neg) ** (1.0 / n) - 1.0


def wacc(equity_value: float, debt_value: float, cost_equity: float,
         cost_debt: float, tax_rate: float) -> float | None:
    """가중평균자본비용. WACC = E/V·Re + D/V·Rd·(1−tax).

    V = E + D. V ≤ 0 이면 미정 → None. tax_rate 는 한계세율(0..1).
    """
    v = equity_value + debt_value
    if v <= 0:
        return None
    we = equity_value / v
    wd = debt_value / v
    return we * cost_equity + wd * cost_debt * (1.0 - tax_rate)


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


# --------------------------------------------------------------------------- #
# 정규화 run-rate (A2) — robust 위치추정 × 계절지수 deseasonalize             #
#   레퍼런스(차용, 코드 미반입): X-13ARIMA-SEATS 계절지수 개념(확인) +        #
#   Tukey IQR / Hampel MAD robust 위치추정. 산식만 차용.                       #
#   ★Hampel(1974) "The Influence Curve and Its Role in Robust Estimation",    #
#   JASA 69:383-393 기반 MAD 규칙(확인). 단 "Hampel identifier"라는 명명의    #
#   원전은 미확정 — 명명 출처 확정 인용 금지(MAD robust σ 산식만 차용).        #
#   핵심: one-off(단발 대형계상)를 robust mask 로 제외 후, 활성월(active)      #
#   기준 factor 로 연환산 — 12 하드코딩 금지(부분기간 이중연환산 방지).        #
# --------------------------------------------------------------------------- #
def _median(xs: list[float]) -> float:
    s = sorted(xs)
    n = len(s)
    if n == 0:
        return 0.0
    m = n // 2
    return s[m] if n % 2 else (s[m - 1] + s[m]) / 2.0


def median_abs_deviation(xs: list[float], *, scale: float = 1.4826) -> float:
    """MAD = median(|x - median(x)|) × scale. scale=1.4826 → 정규분포 σ 일치.

    Hampel(1974) JASA 69:383-393 기반 MAD 규칙. "Hampel identifier" 명명의 원전은
    미확정 — robust σ 추정 산식만 차용(명명 출처 확정 인용 금지).
    """
    if not xs:
        return 0.0
    med = _median(xs)
    return _median([abs(x - med) for x in xs]) * scale


def tukey_fences(xs: list[float], *, k: float = 1.5) -> tuple[float, float]:
    """Tukey IQR fence (Q1 - k·IQR, Q3 + k·IQR). one-off 상·하한 경계."""
    if not xs:
        return (0.0, 0.0)
    s = sorted(xs)
    n = len(s)

    def _q(p: float) -> float:
        # 선형보간 분위수(stdlib, numpy 불필요)
        idx = p * (n - 1)
        lo = int(math.floor(idx))
        hi = int(math.ceil(idx))
        if lo == hi:
            return s[lo]
        return s[lo] + (s[hi] - s[lo]) * (idx - lo)

    q1, q3 = _q(0.25), _q(0.75)
    iqr = q3 - q1
    return (q1 - k * iqr, q3 + k * iqr)


@dataclass
class RunRateResult:
    """정규화 run-rate 결과. 모든 중간값을 박제(QC 재계산·정직성)."""
    actual_total: float            # Σ원계열(전체)
    one_off_total: float           # Σ마스킹된 one-off
    normalized_total: float        # actual - one_off (마스킹 후 합)
    active_months: int             # 정상(비마스킹) 활성월 수 → factor 기준
    monthly_run_rate: float        # normalized_total / active_months
    annualized: float              # monthly_run_rate × 12 (월런레이트의 연환산)
    masked_index: list = field(default_factory=list)        # 마스킹된 원계열 인덱스
    seasonal_factors: list = field(default_factory=list)    # 적용 계절지수(period→factor)
    over_masking: bool = False     # 마스킹 비율 과다 경고(>50%)


def seasonal_indices(series: list[float], season_len: int = 12) -> list[float]:
    """계절지수(평균=1 정규화). period p 의 지수 = mean(series[p::season_len]) / grand_mean.

    X-13 의 multiplicative 계절성분 개념만 차용(ARIMA 전체 X). 길이 < season_len 또는
    grand_mean=0 이면 전부 1.0(계절성 없음). 반환 길이 = len(series)(각 원소의 지수).
    """
    n = len(series)
    if n == 0:
        return []
    grand = sum(series) / n
    if grand == 0 or n < season_len:
        return [1.0] * n
    bucket_mean: dict[int, float] = {}
    for p in range(season_len):
        vals = series[p::season_len]
        if vals:
            bucket_mean[p] = (sum(vals) / len(vals)) / grand
    return [bucket_mean.get(i % season_len, 1.0) for i in range(n)]


def normalized_run_rate(series: list[float], *, season_len: int = 12,
                        annualize_factor: int = 12,
                        deseasonalize: bool = True,
                        tukey_k: float = 1.5,
                        hampel_sigmas: float = 3.0,
                        over_mask_threshold: float = 0.5) -> RunRateResult:
    """one-off 마스킹(robust) × 계절 조정 후 월 run-rate·연환산.

    절차:
      1) (옵션) 계절지수로 deseasonalize → 비교 가능한 계열.
      2) robust mask: Tukey IQR fence **또는** Hampel(median±k·MAD) 밖이면 one-off.
         두 신호의 합집합(보수적 — 둘 중 하나라도 outlier 면 마스킹).
      3) 정상(비마스킹) 항목만으로 월 run-rate = Σnormalized / active_months.
         ★active_months = 비마스킹 항목 수(12 하드코딩 금지 — 부분기간 이중연환산 방지).
      4) annualized = monthly_run_rate × annualize_factor.

    one_off_total 은 *원계열*(season 조정 전)에서 마스킹분을 합산(tie-out: actual =
    normalized + one_off 가 원계열 단위로 성립). normalized_total 도 원계열 단위.
    """
    n = len(series)
    actual_total = sum(series)
    if n == 0:
        return RunRateResult(0.0, 0.0, 0.0, 0, 0.0, 0.0, [], [], False)

    factors = seasonal_indices(series, season_len) if deseasonalize else [1.0] * n
    # deseasonalized 계열(마스킹 판정용). factor 0 방어.
    des = [(series[i] / factors[i]) if factors[i] else series[i] for i in range(n)]

    lo, hi = tukey_fences(des, k=tukey_k)
    med = _median(des)
    mad = median_abs_deviation(des)
    masked_idx: list[int] = []
    for i, x in enumerate(des):
        tukey_out = (x < lo) or (x > hi)
        hampel_out = (mad > 0) and (abs(x - med) > hampel_sigmas * mad)
        if tukey_out or hampel_out:
            masked_idx.append(i)

    masked_set = set(masked_idx)
    one_off_total = sum(series[i] for i in masked_idx)            # 원계열 단위
    normalized_total = actual_total - one_off_total              # tie-out 보존
    active = n - len(masked_idx)
    monthly = (normalized_total / active) if active > 0 else 0.0
    annualized = monthly * annualize_factor
    over = (n > 0) and (len(masked_idx) / n > over_mask_threshold)
    return RunRateResult(
        actual_total=actual_total, one_off_total=one_off_total,
        normalized_total=normalized_total, active_months=active,
        monthly_run_rate=monthly, annualized=annualized,
        masked_index=masked_idx, seasonal_factors=factors, over_masking=over)


# --------------------------------------------------------------------------- #
# LMDI variance decomposition (A4/⑤) — rate + volume, 잔차=0                  #
#   레퍼런스(차용): Horngren rate/volume variance(확인) + LMDI                #
#   (Ang 2005 Energy Policy 확인; 음수/0 처리 Ang&Liu 2007 확인). 로그평균    #
#   가중으로 완전분해(residual≡0).                                            #
#   ★0/음수 분기(자문 재확인): 순수 0 은 1e-20 로 치환하면 로그평균이 극한    #
#   수렴해 분해 가능(0 통째 flag 는 과보수). 음수만 로그 미정의 → 산술        #
#   fallback(decomp_undefined). C0,C1>0 보장 후 LMDI, 0 포함시 small-ε 치환.   #
# --------------------------------------------------------------------------- #
def _log_mean(a: float, b: float) -> float:
    """로그평균 L(a,b) = (a-b)/(ln a - ln b). a==b 면 a. a,b>0 전제."""
    if a == b:
        return a
    return (a - b) / (math.log(a) - math.log(b))


# LMDI 0 치환용 small-ε (Ang&Liu 2007: 순수 0 은 작은 양수로 치환 시 로그평균 극한 수렴).
_LMDI_EPS = 1e-20


@dataclass
class LmdiResult:
    rate_effect: float       # 단가(P) 효과
    volume_effect: float     # 수량(Q) 효과
    total: float             # P1Q1 - P0Q0 (검증 대상)
    residual: float          # total - (rate+volume) — LMDI 면 ≈0
    undefined: bool = False  # 음수로 로그평균 불가 → 산술 fallback 사용
    zero_substituted: bool = False  # 순수 0 을 ε 치환해 분해(Ang&Liu 2007)


def variance_decomp_lmdi(p0: float, q0: float, p1: float, q1: float) -> LmdiResult:
    """비용 ΔC = P1Q1 - P0Q0 를 rate(단가) + volume(수량) 으로 완전분해(잔차 0).

    LMDI-I 가법분해: 가중치 = 로그평균 L(C1,C0).
      rate_effect   = L(C1,C0) · ln(P1/P0)
      volume_effect = L(C1,C0) · ln(Q1/Q0)
      rate + volume == C1 - C0 (residual ≡ 0, 항등).

    0/음수 분기(자문 재확인 — Ang&Liu 2007):
      - 순수 0 (P/Q/C 중 0 이 있고 음수는 없음): 0 → 1e-20 치환 시 로그평균이
        극한 수렴해 분해 가능. zero_substituted=True flag. (0 통째 undefined 는 과보수.)
      - 음수 (P/Q/C 중 하나라도 < 0): 로그 미정의 → Horngren 산술 fallback:
          rate=(P1-P0)·Q1,  volume=P0·(Q1-Q0)  (잔차는 mix 로 흡수되며 0 아닐 수 있음).
        undefined=True flag (decomp_undefined).
    """
    c0, c1 = p0 * q0, p1 * q1
    total = c1 - c0
    vals = (c0, c1, p0, p1, q0, q1)
    # 음수가 하나라도 있으면 로그 미정의 → Horngren 산술 fallback (decomp_undefined)
    if min(vals) < 0:
        rate = (p1 - p0) * q1
        vol = p0 * (q1 - q0)
        resid = total - (rate + vol)
        return LmdiResult(rate, vol, total, resid, undefined=True)
    # 음수 없음. 순수 0 은 ε 치환해 로그평균 극한 수렴 → 분해 가능.
    has_zero = min(vals) == 0
    pe0 = p0 if p0 > 0 else _LMDI_EPS
    pe1 = p1 if p1 > 0 else _LMDI_EPS
    qe0 = q0 if q0 > 0 else _LMDI_EPS
    qe1 = q1 if q1 > 0 else _LMDI_EPS
    ce0 = pe0 * qe0
    ce1 = pe1 * qe1
    L = _log_mean(ce1, ce0)
    rate = L * math.log(pe1 / pe0)
    vol = L * math.log(qe1 / qe0)
    # 잔차는 *실제* total(0 치환 전 ΔC) 기준으로 보고 (tie-out 정직성).
    resid = total - (rate + vol)
    return LmdiResult(rate, vol, total, resid, undefined=False,
                      zero_substituted=has_zero)


# --------------------------------------------------------------------------- #
# stickiness proxy (A3 보조) — sticky costs(ABJ 2003) 탄력                    #
#   레퍼런스(확인): Anderson-Banker-Janakiraman 2003 JAR. 비용이 매출 상승    #
#   시보다 하락 시 *덜* 줄어드는 비대칭성(원가 점착성). ⛔ 단일신호로 cuttable  #
#   판정 금지 — cuttability_rung(Contract 속성)의 *보조지표*로만 사용.         #
# --------------------------------------------------------------------------- #
@dataclass
class StickinessResult:
    up_elasticity: float | None     # 활동 ↑ 구간 Δcost%/Δactivity% 평균
    down_elasticity: float | None   # 활동 ↓ 구간 Δcost%/Δactivity% 평균
    asymmetry: float | None         # up - down (>0 = sticky: 내릴 때 덜 준다)
    sticky: bool                    # asymmetry > threshold (보조 신호일 뿐)
    n_up: int = 0
    n_down: int = 0


def stickiness_proxy(costs: list[float], activity: list[float], *,
                     asymmetry_threshold: float = 0.0) -> StickinessResult:
    """원가 점착성(ABJ 2003) 근사 — 활동 증가/감소 구간 비용 탄력성 비대칭.

    각 인접 기간 (Δlog cost / Δlog activity) 를 활동 증감 방향으로 분리 평균.
    sticky = up_elasticity > down_elasticity (활동 하락 시 비용이 덜 줄어듦).

    ⛔ 단일신호 금지: 반환의 sticky 는 cuttability *보조* 신호. 등급 판정은
    dims.cuttability_rung(계약 속성)이 주(主)이고 이 값은 modulate 보조로만.
    """
    ups: list[float] = []
    downs: list[float] = []
    for i in range(1, len(costs)):
        c0, c1 = costs[i - 1], costs[i]
        a0, a1 = activity[i - 1], activity[i]
        if None in (c0, c1, a0, a1) or c0 <= 0 or c1 <= 0 or a0 <= 0 or a1 <= 0:
            continue
        d_a = math.log(a1 / a0)
        if d_a == 0:
            continue
        elas = math.log(c1 / c0) / d_a
        (ups if a1 > a0 else downs).append(elas)
    up = (sum(ups) / len(ups)) if ups else None
    down = (sum(downs) / len(downs)) if downs else None
    asym = (up - down) if (up is not None and down is not None) else None
    sticky = (asym is not None) and (asym > asymmetry_threshold)
    return StickinessResult(up, down, asym, sticky, len(ups), len(downs))


# --------------------------------------------------------------------------- #
# K-IFRS 1116 (IFRS 16) 리스 — 사용권자산·리스부채 상각 (C6 fc_lease_ifrs16)    #
#   리스부채 = Σ 미래 리스료의 현재가치(증분차입이자율 할인).                   #
#   매기: 이자 = 기초부채 × 기간할인율 / 부채상각 = 지급액 − 이자 /             #
#         기말부채 = 기초부채 + 이자 − 지급액.                                  #
#   사용권자산 = 리스부채(초기) + 선급리스료 − 리스인센티브 + 초기직접원가.     #
#   사용권자산 상각 = 정액(리스기간). rent-free(무상기간)는 지급액 0 이지만     #
#   비용은 정액 인식(지급≠비용; 부채는 이자만 증가).                            #
# --------------------------------------------------------------------------- #
@dataclass
class LeaseScheduleRow:
    period_index: int        # 0-base 기간 인덱스
    payment: float           # 해당 기간 리스료 지급액(rent-free 면 0)
    opening_liab: float      # 기초 리스부채
    interest: float          # 이자(기초부채 × rate)
    principal: float         # 부채상각(지급−이자)
    closing_liab: float      # 기말 리스부채
    rou_open: float          # 기초 사용권자산
    rou_amort: float         # 사용권자산 상각(정액)
    rou_close: float         # 기말 사용권자산


def lease_liability_pv(payments: list[float], rate: float) -> float:
    """리스부채 초기 측정 = Σ 지급액 / (1+rate)^(t+1) (기말 지급 가정).

    payments[t] = t 기간(0-base) 리스료. rate = 기간 증분차입이자율.
    """
    return sum(p / (1.0 + rate) ** (t + 1) for t, p in enumerate(payments))


def lease_schedule(payments: list[float], rate: float, *,
                   initial_direct_costs: float = 0.0,
                   prepaid: float = 0.0, incentives: float = 0.0
                   ) -> list[LeaseScheduleRow]:
    """K-IFRS 1116 리스 스케줄(부채 상각 + 사용권자산 정액상각).

    부채(t): interest = opening×rate / principal = payment − interest /
             closing = opening + interest − payment.
    사용권자산 = 부채초기 + prepaid − incentives + initial_direct_costs,
             정액상각(리스기간 n). rent-free(payment=0)도 자산상각은 정액(정액화).
    마지막 기간은 잔액(closing≈0 / rou_close≈0)을 정확히 흡수(부동소수 보정).
    """
    n = len(payments)
    if n == 0:
        return []
    liab0 = lease_liability_pv(payments, rate)
    rou0 = liab0 + prepaid - incentives + initial_direct_costs
    rou_monthly = rou0 / n
    rows: list[LeaseScheduleRow] = []
    liab = liab0
    rou = rou0
    for t, pay in enumerate(payments):
        interest = liab * rate
        principal = pay - interest                # 부채상각 = 지급 − 이자
        closing = liab + interest - pay
        if t == n - 1:
            closing = 0.0                         # 마지막 기간 잔액 흡수(부동소수 보정)
        ra = rou_monthly if t < n - 1 else rou    # 마지막 달 ROU 잔여 흡수
        rows.append(LeaseScheduleRow(
            period_index=t, payment=pay, opening_liab=liab, interest=interest,
            principal=principal, closing_liab=closing,
            rou_open=rou, rou_amort=ra, rou_close=rou - ra))
        liab = closing
        rou = rou - ra
    return rows


# --------------------------------------------------------------------------- #
# 공통비 다대다 배부 (C6 fc_allocation) — pool×driver 가중 배부, 보존           #
# --------------------------------------------------------------------------- #
def allocate_pool(amount: float, weights: dict) -> dict:
    """단일 풀 amount 를 weights(key→가중) 비례 배부. Σ배부 == amount(잔여 흡수).

    Σweight ≤ 0 이면 빈 dict(배부 불가 — 호출측이 UNALLOCATED 처리).
    마지막 key 가 rounding 잔여를 흡수해 보존(Σ=amount).
    """
    keys = list(weights.keys())
    total = sum(max(weights[k], 0.0) for k in keys)
    if total <= 0:
        return {}
    out: dict = {}
    assigned = 0.0
    for i, k in enumerate(keys):
        if i == len(keys) - 1:
            out[k] = amount - assigned
        else:
            a = amount * max(weights[k], 0.0) / total
            out[k] = a
            assigned += a
    return out


__all__ = [
    "npv", "irr", "mirr", "wacc", "discounted_payback", "payback", "cagr",
    "variance", "variance_pct", "safe_div",
    "gross_margin", "operating_margin", "current_ratio", "ltv_cac",
    "approx_equal",
    "straight_line_depreciation", "depreciation_schedule",
    "depreciation_schedule_ext",
    # 정규화 run-rate (A2)
    "median_abs_deviation", "tukey_fences", "seasonal_indices",
    "normalized_run_rate", "RunRateResult",
    # LMDI variance decomp (A4/⑤)
    "variance_decomp_lmdi", "LmdiResult",
    # stickiness proxy (A3 보조)
    "stickiness_proxy", "StickinessResult",
    # K-IFRS 1116 리스 (C6 fc_lease_ifrs16)
    "LeaseScheduleRow", "lease_liability_pv", "lease_schedule",
    # 공통비 다대다 배부 (C6 fc_allocation)
    "allocate_pool",
]
