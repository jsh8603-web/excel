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


class TestNoForbiddenHeuristic(unittest.TestCase):
    """R6: fc 빌더 소스에 샘플링/head/top-N 휴리스틱 토큰이 없어야 한다."""
    def test_fc_builders_clean(self):
        tdir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "fpna", "templates")
        for fn in ("fc_depreciation_schedule.py", "fc_variance_bridge.py"):
            with open(os.path.join(tdir, fn), encoding="utf-8") as fh:
                src = fh.read()
            rep = QCReport("r6")
            vc.assert_no_forbidden_heuristic(rep, src)
            self.assertTrue(rep.passed, "%s 에 금지 휴리스틱 토큰" % fn)


if __name__ == "__main__":
    unittest.main()
