"""
tests/test_fpna.py — stdlib unittest 회귀(pytest 불필요).

실행: py -m unittest tests.test_fpna   또는   py -m unittest discover tests
회사 PC에서도 설치 없이 동작.
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fpna._bootstrap  # noqa: F401

from fpna import finance
from fpna.dispatcher import dispatch
from fpna.templates import available, get_template
from fpna.render import render

FIXTURE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures", "messy_sample.xlsx")


class TestFinance(unittest.TestCase):
    def test_npv_irr_consistency(self):
        cfs = [-1000, 300, 350, 400, 450, 300]
        r = finance.irr(cfs)
        self.assertIsNotNone(r)
        # IRR 에서 NPV ≈ 0
        self.assertAlmostEqual(finance.npv(r, cfs), 0.0, places=4)

    def test_irr_no_sign_change(self):
        self.assertIsNone(finance.irr([100, 200, 300]))

    def test_discounted_payback(self):
        self.assertIsNotNone(finance.discounted_payback(0.1, [-1000, 300, 350, 400, 450]))
        self.assertIsNone(finance.discounted_payback(0.1, [-1000, 10, 10]))

    def test_safe_div(self):
        self.assertIsNone(finance.safe_div(1, 0))


class TestIngest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not os.path.isfile(FIXTURE):
            from tests.make_fixtures import build_messy
            build_messy(FIXTURE)

    def test_ingest_structure(self):
        from fpna.ingest import ingest_workbook
        res = ingest_workbook(FIXTURE)
        self.assertEqual(res.n_blocks, 2)
        self.assertEqual(res.report.n_rejected, 0)
        rows = {(r.entity, r.period, r.metric): r.value for r in res.tidy_rows}
        # 괄호음수 / 세모음수 / 센티넬 / 소계
        self.assertEqual(rows[("제품B", "2024", "비용")], -50)
        self.assertEqual(rows[("기타", "2024", "매출")], -30)
        self.assertIsNone(rows[("제품B", "2025", "비용")])  # "-" 센티넬
        subtotals = [r for r in res.tidy_rows if r.row_role == "subtotal"]
        self.assertTrue(len(subtotals) >= 1)

    def test_percent_text(self):
        from fpna.ingest import ingest_workbook
        res = ingest_workbook(FIXTURE)
        rows = {r.entity: r.value for r in res.tidy_rows}
        self.assertAlmostEqual(rows["가동률"], 0.85, places=4)


class TestDispatch(unittest.TestCase):
    def test_keyword_routing(self):
        self.assertEqual(dispatch("NPV IRR 투자 타당성").template, "investment_appraisal")
        self.assertEqual(dispatch("예실 변동 브리지").template, "variance")
        self.assertEqual(dispatch("MoM 추이").template, "period_trend")

    def test_column_signal(self):
        d = dispatch("월간 보고", columns=["계획", "실적", "항목"])
        self.assertEqual(d.template, "variance")


class TestTemplatesQC(unittest.TestCase):
    def test_all_golden_pass_qc(self):
        for t in available():
            mod = get_template(t)
            data = mod.golden_sample()
            with tempfile.TemporaryDirectory() as tmp:
                out = os.path.join(tmp, "%s.xlsx" % t)
                res = render(t, data, out)
                self.assertTrue(res.qc.passed, "%s QC FAIL: %s" % (t, res.qc.summary()))
                self.assertTrue(res.saved)
                self.assertTrue(os.path.isfile(out))


if __name__ == "__main__":
    unittest.main(verbosity=2)
