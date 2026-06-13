"""
tests/test_fc.py — fpna-fixed-cost-tables 스킬 회귀 (stdlib unittest).

실행: py -m unittest tests.test_fc   또는   py -m unittest discover tests
회사 PC 무설치 동작. 합성 재무수치 없음 — golden_sample 의 구조 더미만 사용.

커버:
  - dims: 회계 캘린더 연속성/445 주배분, account rollup, grain 강제
  - view_contract: R1·R2·R3·R8·R9·R10·R11 핵심 불변식
  - 도메인 템플릿: golden build+qc PASS, dispatch 라우팅, render 게이트
  - R6 메타: fc 빌더 소스에 금지 휴리스틱 토큰 부재
"""
from __future__ import annotations

import datetime
import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fpna._bootstrap  # noqa: F401

from fpna import dims, finance
from fpna import view_contract as vc
from fpna.templates.base import QCReport
from fpna.dispatcher import dispatch
from fpna.templates import available, get_template
from fpna.render import render


class TestDims(unittest.TestCase):
    def test_calendar_continuous_ruler(self):
        cal = dims.AccountingCalendar(fiscal_year_start_month=1)
        ps = cal.periods((2024, 1), (2025, 3))
        self.assertEqual(len(ps), 15)
        ords = [p.ordinal for p in ps]
        self.assertEqual(ords, list(range(ords[0], ords[0] + 15)))  # 결측 없는 연속
        self.assertEqual(ps[0].cutoff_date, datetime.date(2024, 1, 31))
        self.assertEqual(ps[-1].cutoff_date, datetime.date(2025, 3, 31))

    def test_calendar_fiscal_start_month(self):
        cal = dims.AccountingCalendar(fiscal_year_start_month=4)
        p1 = cal.period(2024, 1)
        self.assertEqual(p1.cutoff_date, datetime.date(2024, 4, 30))
        self.assertEqual(p1.quarter, 1)

    def test_calendar_445_weeks(self):
        cal = dims.AccountingCalendar(scheme="445", anchor_date=datetime.date(2024, 1, 1),
                                      anchor_fy=2024)
        w = cal.periods((2024, 1), (2024, 3))
        # 4주·4주·5주 → cutoff 1/28, 2/25, 3/31
        self.assertEqual([p.cutoff_date for p in w],
                         [datetime.date(2024, 1, 28), datetime.date(2024, 2, 25),
                          datetime.date(2024, 3, 31)])

    def test_445_requires_anchor(self):
        with self.assertRaises(ValueError):
            dims.AccountingCalendar(scheme="445")

    def test_account_rollup(self):
        accs = [dims.Account("T", "합계", "g", "P&L", "fixed", None),
                dims.Account("A", "a", "g", "P&L", "fixed", "T"),
                dims.Account("B", "b", "g", "P&L", "fixed", "T")]
        self.assertEqual({a.code for a in dims.account_leaves(accs)}, {"A", "B"})
        roll = dims.rollup(accs, {"A": 100.0, "B": 50.0})
        self.assertEqual(roll["T"], 150.0)

    def test_fact_requires_grain(self):
        with self.assertRaises(ValueError):
            dims.Fact("", (), [])
        f = dims.Fact("1행=1자산", ("asset",), [{"asset": "A"}])
        self.assertFalse(f.has_duplicate_grain())


class TestViewContract(unittest.TestCase):
    def setUp(self):
        self.cal = dims.AccountingCalendar()

    def test_r8_grain_duplicate(self):
        f = dims.Fact("g", ("k",), [{"k": "A"}, {"k": "A"}])
        rep = QCReport("t")
        vc.assert_grain(rep, f)
        self.assertFalse(rep.passed)

    def test_r1_time_ruler_missing(self):
        f = dims.Fact("g", ("period",),
                      [{"period": "FY2024-P01"}, {"period": "FY2024-P03"}])
        rep = QCReport("t")
        vc.assert_time_ruler(rep, f, self.cal, (2024, 1), (2024, 3), period_key="period")
        self.assertFalse(rep.passed)   # P02 결측 → 실패

    def test_r2_full_outer_status(self):
        m = vc.full_outer([{"k": 1, "a": 10}], [{"k": 2, "b": 20}], "k")
        self.assertEqual(len(m), 2)
        self.assertEqual({r["k"]: r["match_status"] for r in m},
                         {1: "LEFT_ONLY", 2: "RIGHT_ONLY"})

    def test_r3_tie_out_zero_tol(self):
        rep = QCReport("t")
        vc.assert_tie_out(rep, 100.0, 100.0)
        self.assertTrue(rep.passed)
        rep2 = QCReport("t")
        vc.assert_tie_out(rep2, 100.0, 99.999)
        self.assertFalse(rep2.passed)   # tol=0

    def test_r5_cross_tab_no_data(self):
        f = dims.Fact("g", ("cc",), [{"cc": "X", "sc": "Actual", "v": 5}])
        ct = vc.cross_tab(f, "cc", "sc", "v", all_columns=["Actual", "Budget"])
        self.assertEqual(ct["rows"][0]["Budget"], "NO_DATA")   # 결측 = NO_DATA(생략 금지)

    def test_r9_scenario_aligned(self):
        rep = QCReport("t")
        vc.assert_scenario_aligned(rep, {("X",)}, {("X",), ("Y",)})
        self.assertFalse(rep.passed)   # Budget-only 미정렬 노출 필요

    def test_r10_hierarchy_ties(self):
        accs = [dims.Account("T", "t", "g", "P&L", "fixed", None),
                dims.Account("A", "a", "g", "P&L", "fixed", "T")]
        rep = QCReport("t")
        vc.assert_hierarchy_ties(rep, accs, {"A": 50.0})
        self.assertTrue(rep.passed)

    def test_r11_master_to_gl_reason(self):
        rep = QCReport("t")
        vc.assert_master_to_gl(rep, 100.0, 90.0, reasons=["일회성"])
        self.assertTrue(rep.passed)    # 사유 명세 시 허용
        rep2 = QCReport("t")
        vc.assert_master_to_gl(rep2, 100.0, 90.0)
        self.assertFalse(rep2.passed)  # 사유 없으면 실패


class TestDepreciation(unittest.TestCase):
    def test_schedule_sums_to_depreciable_base(self):
        sch = finance.depreciation_schedule(1200.0, 0.0, 12, n_periods=15, start_index=0)
        self.assertAlmostEqual(sum(r[1] for r in sch), 1200.0, places=6)
        self.assertAlmostEqual(sch[-1][2], 0.0, places=6)          # 종료 closing=잔존

    def test_schedule_salvage_and_delay(self):
        sch = finance.depreciation_schedule(1000.0, 100.0, 9, n_periods=12, start_index=2)
        self.assertEqual([round(sch[i][1], 6) for i in range(2)], [0.0, 0.0])  # 가동 전 0
        self.assertAlmostEqual(sum(r[1] for r in sch), 900.0, places=6)
        self.assertAlmostEqual(sch[-1][2], 100.0, places=6)        # 잔존가 도달


class TestFcTemplates(unittest.TestCase):
    def test_registered(self):
        av = available()
        self.assertIn("fc_depreciation_schedule", av)
        self.assertIn("fc_variance_bridge", av)

    def test_golden_build_qc_pass(self):
        for t in ("fc_depreciation_schedule", "fc_variance_bridge"):
            mod = get_template(t)
            data = mod.golden_sample()
            rep = mod.qc(mod.build(data), data)
            self.assertTrue(rep.passed, "%s QC FAIL:\n%s" % (t, rep.summary()))

    def test_dispatch_routing(self):
        self.assertEqual(dispatch("자산 감가상각 스케줄").template, "fc_depreciation_schedule")
        self.assertEqual(dispatch("고정비 변동요인 브리지").template, "fc_variance_bridge")
        self.assertEqual(dispatch("예실 변동 분석").template, "variance")  # 회귀: 기존 유지

    def test_render_gate_saves(self):
        import tempfile
        for t in ("fc_depreciation_schedule", "fc_variance_bridge"):
            data = get_template(t).golden_sample()
            with tempfile.TemporaryDirectory() as d:
                out = os.path.join(d, t + ".xlsx")
                res = render(t, data, out)
                self.assertTrue(res.saved and res.qc.passed)
                self.assertTrue(os.path.exists(out))

    def test_bridge_tie_out_fails_when_unbalanced(self):
        mod = get_template("fc_variance_bridge")
        data = mod.golden_sample()
        data.factors = data.factors[:-1]   # 잔차 제거 → 합 불일치
        rep = mod.qc(mod.build(data), data)
        self.assertFalse(rep.passed)


class TestC11Depreciation(unittest.TestCase):
    """C11: R12 MISSING_ACCRUAL + 부분월/처분/손상 확장."""
    def _mod(self):
        return get_template("fc_depreciation_schedule")

    def test_missing_accrual_emit_and_conserved(self):
        mod = self._mod()
        data = mod.golden_sample()
        data.missing_accruals = {"V-001": ["FY2024-P05"]}
        data.gl_dep_by_period = {}             # 총액 R11 대조 면제(결측으로 합 감소)
        wb = mod.build(data)
        meta = wb._fpna_meta
        # 1층 emit: ledger 1행 + 시트 surfaced flag 1개
        led = meta["anomaly_ledger"]
        self.assertEqual(len(led), 1)
        self.assertEqual(led.rows[0]["anomaly_type"], "MISSING_ACCRUAL")
        self.assertEqual(led.rows[0]["grain"], ("V-001",))
        self.assertEqual(meta["surfaced_flags"], 1)
        # 2층 보존: 은폐 0 → passed True (anomaly 존재는 저장 막지 않음)
        rep = mod.qc(wb, data)
        self.assertTrue(rep.passed, rep.summary())

    def test_missing_accrual_hidden_fails(self):
        """surfaced flag 를 0 으로 위조하면 anomaly_conserved 가 FAIL(은폐 차단)."""
        mod = self._mod()
        data = mod.golden_sample()
        data.missing_accruals = {"V-001": ["FY2024-P05"]}
        data.gl_dep_by_period = {}
        wb = mod.build(data)
        wb._fpna_meta["surfaced_flags"] = 0    # 은폐 시뮬레이션
        rep = mod.qc(wb, data)
        self.assertFalse(rep.passed)

    def test_no_anomaly_golden_clean(self):
        mod = self._mod()
        data = mod.golden_sample()
        wb = mod.build(data)
        self.assertEqual(len(wb._fpna_meta["anomaly_ledger"]), 0)
        self.assertEqual(wb._fpna_meta["surfaced_flags"], 0)
        self.assertTrue(mod.qc(wb, data).passed)

    def test_disposal_stops_da(self):
        """처분월부터 D&A 중단 + 이후 0. window 자동 축소로 가짜 결측 없음."""
        mod = self._mod()
        data = mod.golden_sample()
        data.disposals = {"V-001": (2024, 7)}
        data.gl_dep_by_period = {}
        wb = mod.build(data)
        fact = wb._fpna_meta["fact"]
        post = [r["dep"] for r in fact.rows
                if r["asset_no"] == "V-001" and r["period"] >= "FY2024-P07"]
        self.assertTrue(all((d or 0.0) == 0.0 for d in post))
        self.assertEqual(len(wb._fpna_meta["anomaly_ledger"]), 0)
        self.assertTrue(mod.qc(wb, data).passed)

    def test_partial_first_month_and_impairment(self):
        mod = self._mod()
        # 부분월: 1차월 일할 0.5 — 전체 상각 합은 여전히 (취득가-잔존)에 수렴
        data = mod.golden_sample()
        data.first_period_factor = {"V-001": 0.5}
        data.gl_dep_by_period = {}
        wb = mod.build(data)
        self.assertTrue(mod.qc(wb, data).passed)
        # 손상: base-reset 이벤트가 빌드/QC 를 깨지 않음
        data2 = mod.golden_sample()
        data2.impairments = {"P-001": ((2024, 6), 60_000.0)}
        data2.gl_dep_by_period = {}
        wb2 = mod.build(data2)
        self.assertTrue(mod.qc(wb2, data2).passed)


class TestC12Bridge(unittest.TestCase):
    """C12: timing/permanent 분류 + 잔차 버킷 명시."""
    def _mod(self):
        return get_template("fc_variance_bridge")

    def test_kind_subtotals_in_meta(self):
        mod = self._mod()
        data = mod.golden_sample()
        wb = mod.build(data)
        meta = wb._fpna_meta
        # golden: timing=-20, permanent=30+45+12=87, residual=3
        self.assertAlmostEqual(meta["timing_sum"], -20.0)
        self.assertAlmostEqual(meta["perm_sum"], 87.0)
        self.assertAlmostEqual(meta["resid_sum"], 3.0)
        self.assertTrue(mod.qc(wb, data).passed)

    def test_bad_kind_fails(self):
        mod = self._mod()
        data = mod.golden_sample()
        data.factors[0].kind = "bogus"
        rep = mod.qc(mod.build(data), data)
        self.assertFalse(rep.passed)

    def test_missing_residual_bucket_fails(self):
        mod = self._mod()
        data = mod.golden_sample()
        # 잔차 버킷을 permanent 로 바꿔 residual 부재화 → C12-2 FAIL
        for f in data.factors:
            if f.kind == "residual":
                f.kind = "permanent"
        rep = mod.qc(mod.build(data), data)
        self.assertFalse(rep.passed)


class TestNewMeasures(unittest.TestCase):
    """A2~A6 신규 measure (finance/dims/view_contract 추가분)."""

    def test_normalized_run_rate_tie_and_mask(self):
        # 11×100 정상 + 1×1000 one-off → 마스킹 1, tie-out 보존, factor=active_months
        s = [100.0] * 11 + [1000.0]
        r = finance.normalized_run_rate(s, deseasonalize=False)
        self.assertEqual(r.masked_index, [11])
        self.assertAlmostEqual(r.actual_total, r.normalized_total + r.one_off_total)
        self.assertEqual(r.active_months, 11)            # 12 하드코딩 금지 — 마스킹 후 11
        self.assertAlmostEqual(r.monthly_run_rate, 100.0)
        self.assertAlmostEqual(r.annualized, 1200.0)

    def test_normalized_run_rate_no_outlier_uses_all(self):
        # step(50→70) 은 one-off 아님 — 마스킹 0 기대
        s = [50.0] * 6 + [70.0] * 6
        r = finance.normalized_run_rate(s, deseasonalize=False)
        self.assertEqual(r.masked_index, [])
        self.assertEqual(r.active_months, 12)

    def test_lmdi_residual_zero(self):
        d = finance.variance_decomp_lmdi(10, 100, 12, 80)   # 단가↑ 수량↓
        self.assertGreater(d.rate_effect, 0)                # rate > 0
        self.assertLess(d.volume_effect, 0)                 # volume < 0
        self.assertAlmostEqual(d.residual, 0.0, places=9)   # 완전분해 잔차 0
        self.assertAlmostEqual(d.rate_effect + d.volume_effect, d.total, places=9)

    def test_lmdi_undefined_fallback(self):
        d = finance.variance_decomp_lmdi(10, 0, 12, 80)     # 0 수량 → 로그 미정의
        self.assertTrue(d.undefined)                        # decomp_undefined flag

    def test_stickiness_asymmetry(self):
        # 활동 하락 시 비용 덜 줄어듦 → up_elas > down_elas (sticky)
        costs = [100.0, 110.0, 108.0]
        act = [100.0, 120.0, 90.0]
        st = finance.stickiness_proxy(costs, act)
        self.assertIsNotNone(st.asymmetry)
        self.assertTrue(st.sticky)

    def test_cuttability_rung_contract_driven(self):
        c = dims.Contract("L1", "6010", "임대", datetime.date(2023, 1, 1),
                          datetime.date(2025, 6, 30), "monthly", 100.0)
        # notice 0 → 잔여 17M ≤ locked_horizon(12)? 17>12 → committed (장기)
        res = dims.cuttability_rung(c, as_of=datetime.date(2024, 1, 31), notice_months=0)
        self.assertEqual(res["rung"], "committed")
        # evergreen + notice 미정의 → committed(locked)
        ev = dims.Contract("U1", "6300", "전력", datetime.date(2023, 1, 1),
                           None, "monthly", 100.0)
        res2 = dims.cuttability_rung(ev, as_of=datetime.date(2024, 1, 31), notice_months=0)
        self.assertEqual(res2["rung"], "committed")
        self.assertIsNone(res2["earliest_exit_m"])

    def test_cuttability_stickiness_does_not_change_rung(self):
        """⛔ 단일신호 금지: stickiness 가 등급을 바꾸지 않는다(보조 신호)."""
        c = dims.Contract("L1", "6010", "임대", datetime.date(2023, 1, 1),
                          datetime.date(2025, 6, 30), "monthly", 100.0)
        a = dims.cuttability_rung(c, as_of=datetime.date(2024, 1, 31), notice_months=0,
                                  stickiness=0.9)["rung"]
        b = dims.cuttability_rung(c, as_of=datetime.date(2024, 1, 31), notice_months=0,
                                  stickiness=None)["rung"]
        self.assertEqual(a, b)

    def test_ratio_na_reasons(self):
        rep = QCReport("t")
        led = vc.AnomalyLedger()
        # 0분모 → ZERO_DENOM
        v, reason = vc.assert_ratio_na(rep, 100, 0, led)
        self.assertIsNone(v)
        self.assertEqual(reason, "ZERO_DENOM")
        # 0분자 → 0 (NA 아님)
        v2, r2 = vc.ratio_or_na(0, 5)
        self.assertEqual(v2, 0.0)
        self.assertIsNone(r2)
        # 결측분자 → MISSING_NUM
        _, r3 = vc.ratio_or_na(None, 5)
        self.assertEqual(r3, "MISSING_NUM")
        # 단위불일치 → UNIT_MISMATCH
        _, r4 = vc.ratio_or_na(10, 5, num_unit="sqft", den_unit="㎡",
                               require_same_unit=True)
        self.assertEqual(r4, "UNIT_MISMATCH")
        self.assertEqual(len(led), 1)   # ZERO_DENOM 1건만 emit


class TestNewFcTemplates(unittest.TestCase):
    """A2~A6 신규 템플릿 5종 — golden build+qc / dispatch / render 게이트 / 엣지."""
    _TYPES = ("fc_runrate_normalized", "fc_cuttability_ladder", "fc_driver_unitcost",
              "fc_forward_da", "fc_prepaid_rollforward")

    def test_registered(self):
        av = available()
        for t in self._TYPES:
            self.assertIn(t, av)

    def test_golden_build_qc_pass(self):
        for t in self._TYPES:
            mod = get_template(t)
            data = mod.golden_sample()
            rep = mod.qc(mod.build(data), data)
            self.assertTrue(rep.passed, "%s QC FAIL:\n%s" % (t, rep.summary()))

    def test_dispatch_routing(self):
        self.assertEqual(dispatch("비용 정규화 런레이트 연환산").template,
                         "fc_runrate_normalized")
        self.assertEqual(dispatch("고정비 절감 가능성 사다리").template,
                         "fc_cuttability_ladder")
        self.assertEqual(dispatch("차량 대당 단위원가").template, "fc_driver_unitcost")
        self.assertEqual(dispatch("미래 감가상각 투영").template, "fc_forward_da")
        self.assertEqual(dispatch("선급비용 롤포워드").template, "fc_prepaid_rollforward")
        # 회귀: 과거 감가 스케줄은 여전히 depreciation_schedule
        self.assertEqual(dispatch("자산 감가상각 스케줄").template,
                         "fc_depreciation_schedule")

    def test_render_gate_saves(self):
        import tempfile
        for t in self._TYPES:
            data = get_template(t).golden_sample()
            with tempfile.TemporaryDirectory() as d:
                out = os.path.join(d, t + ".xlsx")
                res = render(t, data, out)
                self.assertTrue(res.saved and res.qc.passed, "%s render FAIL" % t)
                self.assertTrue(os.path.exists(out))

    def test_runrate_tie_breaks_when_normalized_corrupted(self):
        """A2 tie: normalized 가 actual−one_off 와 안 맞으면 FAIL(메타 위조)."""
        mod = get_template("fc_runrate_normalized")
        data = mod.golden_sample()
        wb = mod.build(data)
        wb._fpna_meta["sum_norm"] += 50.0   # tie 깨기
        self.assertFalse(mod.qc(wb, data).passed)

    def test_driver_unitcost_na_conserved(self):
        """A4/R17: golden 에 ZERO_DENOM·UNIT_MISMATCH 2건 → ledger==surfaced 보존."""
        mod = get_template("fc_driver_unitcost")
        data = mod.golden_sample()
        wb = mod.build(data)
        meta = wb._fpna_meta
        self.assertEqual(len(meta["anomaly_ledger"]), 2)        # ZERO_DENOM + UNIT_MISMATCH
        self.assertEqual(meta["surfaced_flags"], 2)
        self.assertTrue(mod.qc(wb, data).passed)
        # 은폐 시뮬레이션 → FAIL
        wb._fpna_meta["surfaced_flags"] = 0
        self.assertFalse(mod.qc(wb, data).passed)

    def test_driver_blended_not_column_sum(self):
        """A4: blended = Σcost/Σqty (열 합산 아님). 정상 2라인만 가중."""
        mod = get_template("fc_driver_unitcost")
        data = mod.golden_sample()
        wb = mod.build(data)
        meta = wb._fpna_meta
        # 정상 라인: cost 12000(qty20) + 30000(qty1000) → Σ42000/Σ1020
        self.assertAlmostEqual(meta["blended"], 42_000.0 / 1_020.0, places=6)

    def test_prepaid_clamp_no_negative(self):
        """A6: 조기해지 999 요청 > 잔액 → clamp(음수 잔액 차단), chain 연속."""
        mod = get_template("fc_prepaid_rollforward")
        data = mod.golden_sample()
        wb = mod.build(data)
        rolled = wb._fpna_meta["rolled"]
        for key, rr in rolled.items():
            for beg, add, amort, end in rr:
                self.assertGreaterEqual(end, -1e-9)            # 음수 잔액 없음
                self.assertLessEqual(amort, beg + add + 1e-9)  # clamp
        self.assertTrue(mod.qc(wb, data).passed)

    def test_forward_da_sums_to_depreciable(self):
        """A5: 투영창이 내용연수 포함하면 Σ미래D&A == cost−salvage."""
        mod = get_template("fc_forward_da")
        data = mod.golden_sample()   # start 2024 end 2026(36M) — 두 자산 내용연수 24/36 포함
        wb = mod.build(data)
        # CX-01 (24M, 2024-01 가동) → 36M 창에 전부 포함 → Σ=24000
        fact = wb._fpna_meta["fact"]
        got = sum(r["dep"] for r in fact.rows if r["asset_no"] == "CX-01")
        self.assertAlmostEqual(got, 24_000.0, places=6)
        self.assertTrue(mod.qc(wb, data).passed)


class TestNoForbiddenHeuristic(unittest.TestCase):
    """R6: fc 빌더 소스에 샘플링/head/top-N 휴리스틱 토큰이 없어야 한다."""
    def test_fc_builders_clean(self):
        tdir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "fpna", "templates")
        for fn in ("fc_depreciation_schedule.py", "fc_variance_bridge.py",
                   "fc_runrate_normalized.py", "fc_cuttability_ladder.py",
                   "fc_driver_unitcost.py", "fc_forward_da.py",
                   "fc_prepaid_rollforward.py"):
            with open(os.path.join(tdir, fn), encoding="utf-8") as fh:
                src = fh.read()
            rep = QCReport("r6")
            vc.assert_no_forbidden_heuristic(rep, src)
            self.assertTrue(rep.passed, "%s 에 금지 휴리스틱 토큰" % fn)


if __name__ == "__main__":
    unittest.main()
