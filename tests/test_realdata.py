"""
tests/test_realdata.py — 실무 데이터 적용 보완 (infer / listing / recommend / autobind).

실행: py -m unittest tests.test_realdata
검증: 문자 식별자 위주 데이터(measure 희소)도 "적용 템플릿 없음" 없이 착지 +
  자동 매핑 + 피벗 Σ 보존 + 정리표 fidelity(행보존·소계 tie). 합성 재무수치 0(구조 더미).
"""
from __future__ import annotations

import copy
import unittest

import fpna._bootstrap  # noqa: F401

from fpna.autobind import autobind, pivot_conserved
from fpna.dispatcher import recommend_from_roles
from fpna.infer import infer_columns, summarize
from fpna.pipeline import run_report
from fpna.templates import get_template, listing


# --- infer: 컬럼 의미 추론 ----------------------------------------------------
class InferTest(unittest.TestCase):
    def test_analytic_shape(self):
        rows = [{"기간": "2024-01", "계정": "매출", "금액": 1000},
                {"기간": "2024-02", "계정": "매출", "금액": 1100}]
        s = summarize(infer_columns(rows))
        self.assertIn("기간", s["time"])
        self.assertIn("금액", s["measure"])
        self.assertIn("계정", s["dimension"])

    def test_id_heavy_no_measure(self):
        """식별자 위주(measure 0) → has_measure False(정리표 신호)."""
        rows = [{"id": "P-1", "zone": "A", "st": "on"},
                {"id": "P-2", "zone": "A", "st": "off"},
                {"id": "P-3", "zone": "B", "st": "on"},
                {"id": "P-4", "zone": "B", "st": "on"}]
        s = summarize(infer_columns(rows))
        self.assertFalse(s["has_measure"])
        self.assertIn("id", s["id"])
        self.assertIn("zone", s["dimension"])

    def test_deterministic(self):
        rows = [{"a": "x", "b": 1}, {"a": "y", "b": 2}]
        self.assertEqual(summarize(infer_columns(rows)), summarize(infer_columns(rows)))

    def test_empty(self):
        self.assertEqual(infer_columns([]), [])

    def test_low_card_number_is_dimension(self):
        """저카디 숫자(코드값)는 measure 아님."""
        rows = [{"code": 1, "v": 100}, {"code": 1, "v": 200}, {"code": 2, "v": 300},
                {"code": 2, "v": 400}, {"code": 1, "v": 500}, {"code": 2, "v": 600},
                {"code": 1, "v": 700}, {"code": 2, "v": 800}, {"code": 1, "v": 900},
                {"code": 2, "v": 1000}, {"code": 1, "v": 1100}]
        s = summarize(infer_columns(rows))
        self.assertIn("code", s["dimension"])     # 2개 고유값 = 범주코드
        self.assertIn("v", s["measure"])


# --- recommend: 항상 착지점 --------------------------------------------------
class RecommendTest(unittest.TestCase):
    def test_no_measure_to_listing(self):
        s = {"time": [], "measure": [], "dimension": ["z"], "id": ["p"],
             "has_measure": False, "n_measure": 0, "n_dimension": 1}
        self.assertEqual(recommend_from_roles(s).template, "listing")

    def test_time_measure_to_trend(self):
        s = {"time": ["t"], "measure": ["m"], "dimension": [], "id": [],
             "has_measure": True, "n_measure": 1, "n_dimension": 0}
        self.assertEqual(recommend_from_roles(s).template, "period_trend")

    def test_two_measure_to_variance(self):
        s = {"time": [], "measure": ["plan", "actual"], "dimension": ["item"], "id": [],
             "has_measure": True, "n_measure": 2, "n_dimension": 1}
        self.assertEqual(recommend_from_roles(s).template, "variance")

    def test_never_empty(self):
        """어떤 조합도 템플릿을 반환(적용 없음 종착 제거)."""
        s = {"time": [], "measure": [], "dimension": [], "id": [],
             "has_measure": False, "n_measure": 0, "n_dimension": 0}
        self.assertTrue(recommend_from_roles(s).template)


# --- listing: 정리표 fidelity ------------------------------------------------
class ListingTest(unittest.TestCase):
    def test_golden_spine_passes(self):
        res = run_report(listing, listing.golden_sample())
        self.assertTrue(res.qc.passed, res.qc.summary())
        self.assertIsNotNone(res.receipt)

    def test_row_preservation(self):
        inp = listing.golden_sample()
        wb = listing.build(inp)
        self.assertEqual(wb._fpna_meta["n_emitted"], wb._fpna_meta["n_rows_in"])
        self.assertTrue(listing.qc(wb, inp).passed)

    def test_subtotal_tie(self):
        inp = listing.golden_sample()       # group_by=구역, number_cols=수량
        wb = listing.build(inp)
        rep = listing.qc(wb, inp)
        self.assertTrue(any("소계" in c[0] and c[1] for c in rep.checks))
        # 그룹 소계 합 == 총계
        self.assertAlmostEqual(
            sum(g[1]["수량"] for g in wb._fpna_meta["subtotals"]),
            wb._fpna_meta["grand_total"], places=6)

    def test_no_numbers_listing(self):
        """measure 0(전부 문자) 데이터도 정리표로 떨어짐 + 스파인 통과."""
        inp = listing.ListingInput(
            title="식별자 목록", headers=["id", "zone", "st"],
            rows=[{"id": "P-1", "zone": "A", "st": "on"},
                  {"id": "P-2", "zone": "B", "st": "off"}],
            number_cols=[], group_by="", show_total=False)
        res = run_report(listing, inp)
        self.assertTrue(res.qc.passed, res.qc.summary())

    def test_conserve_total(self):
        """CONSERVE_SPECS: Σ입력 == 보고 총계."""
        inp = listing.golden_sample()
        wb = listing.build(inp)
        raw = sum(listing._num(r.get("수량")) for r in inp.rows)
        self.assertAlmostEqual(raw, wb._fpna_meta["grand_total"], places=6)


# --- autobind: spec 없는 자동 매핑 -------------------------------------------
class AutobindTest(unittest.TestCase):
    def test_id_heavy_to_listing(self):
        rows = [{"parcel": "P-1", "zone": "A", "qty": 5},
                {"parcel": "P-2", "zone": "A", "qty": 7},
                {"parcel": "P-3", "zone": "B", "qty": 3}]
        t, inp = autobind(rows)
        self.assertEqual(t, "listing")
        self.assertEqual(inp.number_cols, ["qty"])
        self.assertEqual(inp.group_by, "zone")
        self.assertTrue(run_report(get_template(t), inp).qc.passed)

    def test_time_measure_pivot_conserved(self):
        rows = [{"기간": "2024-01", "금액": 1000}, {"기간": "2024-01", "금액": 200},
                {"기간": "2024-02", "금액": 1100}]
        t, inp = autobind(rows)
        self.assertEqual(t, "period_trend")
        self.assertEqual(inp.series["금액"][0], 1200.0)        # 같은 기간 합산
        self.assertTrue(pivot_conserved(rows, inp, t))         # Σ 보존
        self.assertTrue(run_report(get_template(t), inp).qc.passed)

    def test_pivot_conservation_catches_loss(self):
        """피벗 합이 long 합과 다르면 False(왜곡 검출)."""
        rows = [{"기간": "2024-01", "금액": 1000}, {"기간": "2024-02", "금액": 500}]
        t, inp = autobind(rows)
        broke = copy.deepcopy(inp)
        broke.series["금액"][0] = 999.0                        # 인위적 왜곡
        self.assertFalse(pivot_conserved(rows, broke, t))

    def test_empty_safe(self):
        t, inp = autobind([])
        self.assertEqual(t, "listing")


# --- analyze: 다중시트 워크북 구조 스캔 + 추천 (P3) --------------------------
class AnalyzeTest(unittest.TestCase):
    def _multi_wb(self, path):
        from openpyxl import Workbook
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "추이"
        ws1.append(["기간", "금액"])
        ws1.append(["2024-01", 1000])
        ws1.append(["2024-02", 1100])
        ws2 = wb.create_sheet("목록")
        ws2.append(["id", "zone", "status"])
        ws2.append(["P-1", "A", "on"])
        ws2.append(["P-2", "B", "off"])
        ws2.append(["P-3", "B", "on"])
        wb.save(path)

    def test_multi_sheet_tags_and_recommend(self):
        import os
        import tempfile
        from fpna.analyze import analyze_workbook
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "multi.xlsx")
            self._multi_wb(p)
            res = analyze_workbook(p)
            shapes = {s["sheet"]: s["shape"] for s in res["sheets"]}
            tmpl = {s["sheet"]: s["template"] for s in res["sheets"]}
            self.assertEqual(shapes["추이"], "시계열형")
            self.assertEqual(shapes["목록"], "식별자 목록형")
            self.assertEqual(tmpl["추이"], "period_trend")
            self.assertEqual(tmpl["목록"], "listing")          # measure 0 → 정리표
            self.assertEqual(res["recommendation"]["kind"], "multi")

    def test_cmd_analyze_rc(self):
        import os
        import tempfile
        from main import cmd_analyze
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "multi.xlsx")
            self._multi_wb(p)
            self.assertEqual(cmd_analyze([p]), 0)
        self.assertEqual(cmd_analyze(["no_such.xlsx"]), 2)


if __name__ == "__main__":
    unittest.main()
