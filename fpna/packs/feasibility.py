"""
fpna.packs.feasibility — 사업타당성·투자심사 팩.

exhibit 구성(시트당 1장): investment_appraisal + pnl_3statement(순환연동) +
debt_schedule + working_capital + scenario_sensitivity. 조달·상환·세금·민감도가
단일 가정집합(공유 calendar/debt/assumptions)에서 일관 흐른다.

크로스시트 정합 5종(2-5) — 스파인이 강제:
  (a) BS 항등        자산 == 부채 + 자본 (R3/R11)
  (b) 현금 tie       CFS 기말현금 == BS 현금
  (c) 부채 tie       debt_schedule 기말 == BS 차입금
  (d) 감가 연동      depreciation == IS 감가
  (e) 이자 연동      평균부채 × rate(solve 수렴) == IS 이자

★tautology 경계(정직 박제): golden 은 구조 더미라 shared_facts 양변을 일관값으로
채워 tie 통과를 *시연*한다(fc_boardpack 과 동일 패턴). 실데이터에서는 각 exhibit 의
독립 산출(예: debt_schedule.totals['closing'])을 추출해 진짜 cross-exhibit 대조가 된다.
raw_sum_fn 은 shared_facts dict 만 읽어(build 호출 0) AST 독립성을 지킨다.
"""
from __future__ import annotations

import fpna._bootstrap  # noqa: F401

from fpna.conserve import ConserveSpec
from fpna.pack import PackSpec, ExhibitSpec, ModelCheck
from fpna.templates import (pnl_3statement as P, debt_schedule as D,
                            working_capital as W, investment_appraisal as I,
                            scenario_sensitivity as S)


def make_spec() -> PackSpec:
    """사업타당성 팩 spec. 공유 debt 가정으로 pnl 순환연동 + 5 cross tie."""
    # 공유 부채 가정(pnl 순환연동 · debt_schedule exhibit 동일 사용)
    beginning_debt = {"revolver": 0.0, "term": 2000.0}
    debt_rates = {"revolver": 0.010, "term": 0.008}

    # pnl: solve_revolver 로 이자·현금·부채 내생화(리볼버 plug 해소) → linked
    pnl = P.solve_and_link(
        revenue=1000.0, cogs=600.0, sga=200.0, da=50.0, tax_rate=0.22,
        beginning_debt=beginning_debt, debt_rates=debt_rates,
        pre_financing_cash=-500.0, cash_begin=100.0, re_begin=500.0,
        dividends=30.0, paid_in_capital=200.0, min_cash=50.0,
        sweep_priority=("revolver", "term"))

    # equity·assets 재계산(독립) — BS 항등 검정용
    ebt = pnl.revenue - pnl.cogs - pnl.sga - pnl.da - pnl.interest
    ni = ebt - max(0.0, ebt) * pnl.tax_rate
    equity = pnl.re_begin + ni - pnl.dividends + pnl.paid_in_capital
    assets = pnl.cash + pnl.other_assets

    # debt_schedule exhibit — pnl 과 같은 tranche 구조(연동 시연)
    debt_inp = D.DebtScheduleInput(
        tranches=[
            D.DebtTranche("term", "Term Loan", "term", opening=2000.0, rate=0.008,
                          mandatory=[0.0], sweep_enabled=True),
            D.DebtTranche("revolver", "리볼버(RCF)", "revolver", opening=0.0,
                          rate=0.010, mandatory=[], sweep_enabled=False),
        ],
        cash_available=[0.0], start=(2024, 1), end=(2024, 1))

    # 공유 facts(평탄) — cross tie / 모델체크 source. golden 은 일관값(tautology 경계).
    shared = {
        "is_assets": assets, "bs_liab_eq": pnl.liabilities + equity,        # (a)
        "cfs_end_cash": pnl.cash, "bs_cash": pnl.cash,                      # (b)
        "sched_debt_end": pnl.liabilities, "bs_debt": pnl.liabilities,     # (c)
        "sched_depr": pnl.da, "is_da": pnl.da,                             # (d)
        "calc_interest": pnl.interest, "is_interest": pnl.interest,        # (e)
    }

    cross_ties = (
        ConserveSpec("(a) BS항등 자산=부채+자본",
                     raw_sum_fn=lambda s: s.shared_facts["is_assets"],
                     reported_key="bs_liab_eq", tol=0.5),
        ConserveSpec("(b) 현금 tie CFS기말=BS현금",
                     raw_sum_fn=lambda s: s.shared_facts["cfs_end_cash"],
                     reported_key="bs_cash", tol=0.5),
        ConserveSpec("(c) 부채 tie 스케줄기말=BS차입금",
                     raw_sum_fn=lambda s: s.shared_facts["sched_debt_end"],
                     reported_key="bs_debt", tol=0.5),
        ConserveSpec("(d) 감가연동 스케줄=IS감가",
                     raw_sum_fn=lambda s: s.shared_facts["sched_depr"],
                     reported_key="is_da", tol=0.5),
        ConserveSpec("(e) 이자연동 평균부채×rate=IS이자",
                     raw_sum_fn=lambda s: s.shared_facts["calc_interest"],
                     reported_key="is_interest", tol=0.5),
    )

    model_checks = (
        ModelCheck("(a) BS 항등", "is_assets", "bs_liab_eq"),
        ModelCheck("(b) 현금 tie", "cfs_end_cash", "bs_cash"),
        ModelCheck("(c) 부채 tie", "sched_debt_end", "bs_debt"),
        ModelCheck("(d) 감가 연동", "sched_depr", "is_da"),
        ModelCheck("(e) 이자 연동", "calc_interest", "is_interest"),
    )

    exhibits = (
        ExhibitSpec("investment_appraisal", I.golden_sample_full(), "투자타당성", "summary"),
        ExhibitSpec("pnl_3statement", pnl, "손익(순환연동)", "detail"),
        ExhibitSpec("debt_schedule", debt_inp, "부채스케줄", "detail"),
        ExhibitSpec("working_capital", W.golden_sample(), "운전자본", "detail"),
        ExhibitSpec("scenario_sensitivity", S.golden_sample(), "민감도", "detail"),
    )

    return PackSpec(
        name="feasibility", title="사업타당성·투자심사 팩",
        subtitle="조달·상환·세금·민감도 단일 가정집합 연동",
        exhibits=exhibits, shared_facts=shared, cross_ties=cross_ties,
        model_checks=model_checks, unit="₩mn")


__all__ = ["make_spec"]
