"""
fpna.dims — FP&A conformed dimension + 회계 캘린더 (Kimball star schema).

이 모듈은 "테이블 더미"가 아니라 **fact + conformed dimension** 구조를 강제하는 토대다.
View Contract(R1·R8·R9·R10)가 여기서 만든 차원·grain 위에서만 성립한다.

설계 원칙
---------
- stdlib only (datetime/calendar/dataclass). openpyxl·pandas·numpy 미사용.
- 시간축 ruler 는 **오직 AccountingCalendar 에서만** 생성한다(R1 근거 — 데이터에서
  날짜를 뽑아 열을 만드는 행위 금지).
- 모든 fact 는 grain 을 선언해야 한다(Kimball: grain first). 미선언 = 빌드 실패(R8).

6 conformed dimension
---------------------
  1) Period       — AccountingCalendar 가 생성 (fiscal year/period/quarter/cutoff)
  2) Scenario     — Actual/Budget/FC1/FC2/PriorYear (variance 는 이 축의 차이로만)
  3) Account      — code→group→fs_line parent-child 계층 + cost behavior 속성
  4) CostCenter   — cost center→dept→entity + 분야(domain) 매핑
  5) Asset/Vendor — 자산번호·임대인·리스사 (고정비는 계약·자산 단위 추적)
  6) CostBehavior — fixed/variable/semi_fixed/committed + 판정 기준
"""
from __future__ import annotations

import calendar as _cal
import datetime as _dt
from dataclasses import dataclass, field


# --------------------------------------------------------------------------- #
# 2) Scenario 차원                                                            #
# --------------------------------------------------------------------------- #
# variance = 데이터 필터가 아니라 이 축의 차이로만 계산(R9). 같은 grain 에 공존.
SCENARIOS: tuple[str, ...] = ("Actual", "Budget", "FC1", "FC2", "PriorYear")


def is_scenario(s: str) -> bool:
    return s in SCENARIOS


# --------------------------------------------------------------------------- #
# 6) CostBehavior 차원 + 4) domain(분야)                                      #
# --------------------------------------------------------------------------- #
BEHAVIORS: tuple[str, ...] = ("fixed", "variable", "semi_fixed", "committed")

# 각 분야에서 끌어온 비용이 "왜 고정인가"의 판정 기준을 차원 속성으로 명문화.
BEHAVIOR_CRITERIA: dict[str, str] = {
    "fixed": "조업도와 무관하게 기간 일정 (임차료·감가상각·고정급여)",
    "variable": "조업도에 비례 (재료비·연동 수수료)",
    "semi_fixed": "구간별 계단식 (설비 증설 시 점프 — 준고정)",
    "committed": "계약상 약정으로 해지 전까지 고정 (리스·장기 유지보수)",
}

# 4) CostCenter 의 분야 속성 축 — 직무 핵심(차량/부동산/유틸리티/고정부품).
DOMAINS: tuple[str, ...] = ("vehicle", "property", "utility", "fixed_parts")
DOMAIN_LABEL: dict[str, str] = {
    "vehicle": "차량",
    "property": "부동산",
    "utility": "유틸리티",
    "fixed_parts": "고정부품",
}


def is_behavior(b: str) -> bool:
    return b in BEHAVIORS


def is_domain(d: str) -> bool:
    return d in DOMAINS


# --------------------------------------------------------------------------- #
# 1) Period 차원 + AccountingCalendar                                         #
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Period:
    """회계 기간 1개. 단순 일자가 아니라 fy/period/quarter/cutoff 를 갖는다."""
    fy: int                       # fiscal year (start-label 기준)
    period: int                   # 1..periods_per_year
    quarter: int                  # 1..4
    cutoff_date: _dt.date | None  # 기간 말일(cutoff)
    label: str                    # "FY2024-P03"

    @property
    def ordinal(self) -> int:
        """연속 ruler 정렬용 단조 정수 (12 period/year 가정)."""
        return self.fy * 12 + (self.period - 1)


# 4-4-5 주 배분(분기당 4+4+5 주 = 13주, 1년 52주, 12 period).
_W445: tuple[int, ...] = (4, 4, 5, 4, 4, 5, 4, 4, 5, 4, 4, 5)


@dataclass
class AccountingCalendar:
    """회사의 회계 캘린더. 시간축 ruler 의 유일한 발생원.

    scheme="calendar": 역월/이월. fiscal_year_start_month 부터 매 캘린더 월 = 1 period.
    scheme="445":      4-4-5 주 배분. anchor_date(anchor_fy 의 P1 시작일) 필수.
    """
    fiscal_year_start_month: int = 1     # 1 = 역월(달력연도와 일치)
    scheme: str = "calendar"             # "calendar" | "445"
    anchor_date: _dt.date | None = None  # 445 전용: anchor_fy 의 P1 시작일
    anchor_fy: int | None = None         # 445 전용
    periods_per_year: int = 12

    def __post_init__(self) -> None:
        if self.scheme not in ("calendar", "445"):
            raise ValueError("scheme 은 'calendar' 또는 '445': %r" % self.scheme)
        if not (1 <= self.fiscal_year_start_month <= 12):
            raise ValueError("fiscal_year_start_month 는 1..12")
        if self.periods_per_year != 12:
            raise ValueError("periods_per_year 는 현재 12 만 지원")
        if self.scheme == "445" and (self.anchor_date is None or self.anchor_fy is None):
            raise ValueError("445 scheme 은 anchor_date + anchor_fy 필요")

    # -- cutoff 계산 -------------------------------------------------------- #
    def _calendar_cutoff(self, fy: int, p: int) -> _dt.date:
        idx = (self.fiscal_year_start_month - 1) + (p - 1)
        year = fy + idx // 12
        month = idx % 12 + 1
        last_day = _cal.monthrange(year, month)[1]
        return _dt.date(year, month, last_day)

    def _w445_cutoff(self, fy: int, p: int) -> _dt.date:
        assert self.anchor_date is not None and self.anchor_fy is not None
        fy_start = self.anchor_date + _dt.timedelta(weeks=52 * (fy - self.anchor_fy))
        weeks_before = sum(_W445[: p - 1])
        p_start = fy_start + _dt.timedelta(weeks=weeks_before)
        p_end = p_start + _dt.timedelta(weeks=_W445[p - 1]) - _dt.timedelta(days=1)
        return p_end

    # -- 단일/연속 period --------------------------------------------------- #
    def period(self, fy: int, p: int) -> Period:
        if not (1 <= p <= self.periods_per_year):
            raise ValueError("period 는 1..%d: %r" % (self.periods_per_year, p))
        quarter = (p - 1) // 3 + 1
        cutoff = (self._calendar_cutoff(fy, p) if self.scheme == "calendar"
                  else self._w445_cutoff(fy, p))
        return Period(fy, p, quarter, cutoff, "FY%d-P%02d" % (fy, p))

    def periods(self, start: tuple[int, int], end: tuple[int, int]) -> list[Period]:
        """start..end (둘 다 (fy, period), inclusive) 의 **연속** ruler.

        결측 기간을 건너뛰지 않는다(R1). 표시 시간축은 항상 이 리스트로 생성한다.
        """
        s = start[0] * 12 + (start[1] - 1)
        e = end[0] * 12 + (end[1] - 1)
        if e < s:
            raise ValueError("end(%r) < start(%r)" % (end, start))
        out: list[Period] = []
        for o in range(s, e + 1):
            out.append(self.period(o // 12, o % 12 + 1))
        return out

    def ruler(self, start: tuple[int, int], end: tuple[int, int]) -> list[str]:
        """연속 시간축 라벨 리스트 (R1 검증용)."""
        return [p.label for p in self.periods(start, end)]


# --------------------------------------------------------------------------- #
# 3) Account 차원 (parent-child 계층)                                         #
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Account:
    code: str
    name: str
    group: str                 # 계정그룹
    fs_line: str               # 재무제표 라인 (P&L/BS)
    behavior: str              # CostBehavior (fixed/variable/...)
    parent: str | None = None  # 상위 계정 code (없으면 최상위)


def account_leaves(accounts: list[Account]) -> list[Account]:
    """자식이 없는(leaf) 계정만."""
    parents = {a.parent for a in accounts if a.parent}
    return [a for a in accounts if a.code not in parents]


def rollup(accounts: list[Account], leaf_values: dict[str, float]) -> dict[str, float]:
    """leaf_values(leaf code→금액) → 모든 노드의 (자신+자손) 합계 dict.

    R10(계층 정합성)에서 이 rollup 합과 leaf 합의 tie-out 을 검증한다.
    """
    children: dict[str, list[str]] = {}
    for a in accounts:
        if a.parent:
            children.setdefault(a.parent, []).append(a.code)
    memo: dict[str, float] = {}

    def total(code: str) -> float:
        if code in memo:
            return memo[code]
        v = float(leaf_values.get(code, 0.0))
        for c in children.get(code, []):
            v += total(c)
        memo[code] = v
        return v

    return {a.code: total(a.code) for a in accounts}


# --------------------------------------------------------------------------- #
# 4) CostCenter 차원                                                          #
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class CostCenter:
    code: str
    dept: str
    entity: str
    domain: str   # DOMAINS 중 하나 (vehicle/property/utility/fixed_parts)


# --------------------------------------------------------------------------- #
# 5) Asset/Vendor 차원                                                        #
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Asset:
    asset_no: str
    vendor: str                       # 임대인/리스사/공급처
    domain: str                       # DOMAINS 중 하나
    acq_cost: float                   # 취득가
    life_months: int                  # 내용연수(월)
    salvage: float = 0.0              # 잔존가
    in_service: tuple[int, int] | None = None  # 가동 시점 (fy, period)


# --------------------------------------------------------------------------- #
# 7) Contract 차원 (자문 §2 — 재발 인스턴스)                                  #
#   active window(start~end)는 "재발 인스턴스" 속성이고 외생 master 여야       #
#   invariant 성립(Fact 파생은 순환 → 꼬리 결측 못 잡음). Account/Asset 과     #
#   평행 instance-carrier(흡수 금지). R12·R14·cuttability 의 공통 전제.        #
# --------------------------------------------------------------------------- #
RECURRENCE: tuple[str, ...] = ("monthly", "quarterly", "annual", "one_time")
CONTRACT_STATUS: tuple[str, ...] = ("active", "terminated", "suspended")
LIFECYCLE_STATES: tuple[str, ...] = ("encumbered", "expended", "liquidated", "cancelled")

_CADENCE: dict[str, int | None] = {"monthly": 1, "quarterly": 3, "annual": 12, "one_time": None}


@dataclass(frozen=True)
class Contract:
    """재발 고정비 계약 1건. R12 엔진 = recurrence + status 쌍."""
    contract_id: str                      # PK
    account_id: str                       # FK → Account (N:1)
    counterparty: str
    start_date: _dt.date
    end_date: _dt.date | None             # None = evergreen(상시 active)
    recurrence: str                       # RECURRENCE
    amount_per_period: float
    currency: str = "KRW"
    status: str = "active"                # CONTRACT_STATUS — ended vs missing 구분의 핵심
    asset_id: str | None = None           # FK → Asset (N:1 nullable; SaaS=None)
    lifecycle_state: str = "encumbered"   # LIFECYCLE_STATES

    def __post_init__(self) -> None:
        if self.recurrence not in RECURRENCE:
            raise ValueError("recurrence 는 %s 중 하나: %r" % (RECURRENCE, self.recurrence))
        if self.status not in CONTRACT_STATUS:
            raise ValueError("status 는 %s 중 하나: %r" % (CONTRACT_STATUS, self.status))


def expected_presence(contract: Contract, periods: list) -> set:
    """R12 엔진: 활성 약정이 각 기간에 계상돼야 하는 {(contract_id, period_label)}.

    status != active → 빈 집합(terminated/suspended 는 결측 정당, flag 억제).
    evergreen(end_date=None) → 상시 active(skip 금지). recurrence cadence 로 정렬.
    """
    if contract.status != "active":
        return set()
    cad = _CADENCE[contract.recurrence]
    active = []
    for p in periods:
        cd = p.cutoff_date
        if cd is None:
            active.append(p)
            continue
        if cd < contract.start_date:
            continue
        if contract.end_date is not None and cd > contract.end_date:
            continue
        active.append(p)
    if not active:
        return set()
    if cad is None:                              # one_time: 첫 활성 기간만
        return {(contract.contract_id, active[0].label)}
    base = active[0].ordinal
    return {(contract.contract_id, p.label) for p in active
            if (p.ordinal - base) % cad == 0}


# --------------------------------------------------------------------------- #
# 내부 tidy fact 표현                                                         #
# --------------------------------------------------------------------------- #
@dataclass
class Fact:
    """내부 tidy fact. 표시(wide)는 view_contract.cross_tab 으로 전수 전개한다.

    grain 미선언 = 빌드 실패(Kimball: declare the grain first / R8).
    """
    grain: str                 # "1행 = 1 period × 1 cost_center × 1 account × 1 scenario"
    grain_keys: tuple[str, ...]
    rows: list[dict] = field(default_factory=list)

    def __post_init__(self) -> None:
        if not self.grain or not self.grain_keys:
            raise ValueError(
                "Fact 는 grain 선언 의무 (Kimball: grain first). "
                "grain(문장) + grain_keys(튜플) 둘 다 필요"
            )

    def key_tuples(self) -> list[tuple]:
        return [tuple(r.get(k) for k in self.grain_keys) for r in self.rows]

    def distinct_keys(self) -> set[tuple]:
        return set(self.key_tuples())

    def has_duplicate_grain(self) -> bool:
        kt = self.key_tuples()
        return len(kt) != len(set(kt))


__all__ = [
    # Scenario
    "SCENARIOS", "is_scenario",
    # Behavior / domain
    "BEHAVIORS", "BEHAVIOR_CRITERIA", "DOMAINS", "DOMAIN_LABEL",
    "is_behavior", "is_domain",
    # Period / calendar
    "Period", "AccountingCalendar",
    # Account
    "Account", "account_leaves", "rollup",
    # CostCenter / Asset
    "CostCenter", "Asset",
    # Fact
    "Fact",
]
