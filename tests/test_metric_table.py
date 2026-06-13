"""
tests.test_metric_table — 일반 라벨+단일숫자 표(metric_table) 회귀.

검사: golden build+qc PASS · 디스패처 라우팅(구체형 무회귀) · 상황 프로파일 ·
필드계약 위반 차단(부호/범위/허용값/not_null) · 결측 NO_DATA 노출(저장 허용) ·
tie-out 불일치 차단 · grain(라벨 유일성).
stdlib unittest 만 사용(pytest 불필요).
"""
import unittest

import fpna._bootstrap  # noqa: F401

from fpna.dispatcher import dispatch
from fpna.templates import metric_table as M

F = M.FieldSpec
R = M.MetricRow


class TestMetricTableGolden(unittest.TestCase):
    def test_golden_build_qc_pass(self):
        data = M.golden_sample()
        wb = M.build(data, mode="create")
        rep = M.qc(wb, data)
        self.assertTrue(rep.passed, rep.summary())

    def test_golden_has_nodata_surfaced(self):
        # 골든은 결측(물량 B)을 NO_DATA 로 노출하지만 위반은 0 → PASS 유지
        data = M.golden_sample()
        wb = M.build(data, mode="create")
        self.assertGreaterEqual(wb._fpna_meta["surfaced_flags"], 1)
        self.assertEqual(wb._fpna_meta["n_violation"], 0)


class TestMetricTableDispatch(unittest.TestCase):
    def test_generic_table_routes_to_metric_table(self):
        for txt in ("투자비 명세 내역서", "부서별 원가 명세", "항목별 금액표",
                    "수량표 집계", "line item 표"):
            self.assertEqual(dispatch(txt).template, "metric_table", txt)

    def test_specific_templates_still_win(self):
        # 회귀 보호 — 구체 분석형이 metric_table 보다 먼저 이겨야 한다.
        self.assertEqual(dispatch("예실 변동 브리지").template, "variance")
        self.assertEqual(dispatch("자산 감가상각 스케줄").template,
                         "fc_depreciation_schedule")
        self.assertEqual(dispatch("인원 계획 인건비 증원").template, "headcount_plan")
        self.assertEqual(dispatch("NPV IRR 투자 타당성").template,
                         "investment_appraisal")


class TestMetricTableProfiles(unittest.TestCase):
    def test_suggest_profile(self):
        self.assertEqual(M.suggest_profile("투자비 내역"), "capex")
        self.assertEqual(M.suggest_profile("임금 원가 명세"), "labor_cost")
        self.assertEqual(M.suggest_profile("상각비 명세"), "depreciation")
        self.assertEqual(M.suggest_profile("물량 집계표"), "volume")
        self.assertEqual(M.suggest_profile("그냥 금액 표"), "mixed")

    def test_profile_fields_nonempty(self):
        for p in ("volume", "labor_cost", "depreciation", "capex", "mixed"):
            self.assertTrue(M.profile_fields(p), p)


class TestMetricTableValidationGate(unittest.TestCase):
    def _qc(self, data):
        wb = M.build(data, mode="create")
        return M.qc(wb, data)

    def test_sign_and_accepted_values_and_not_null_block(self):
        bad = M.MetricTableInput(
            unit="₩", profile="capex",
            fields=[F("투자비", kind="currency", nullable=False, min_value=0),
                    F("등급", kind="count", accepted_values=[1, 2, 3])],
            rows=[R("A", {"투자비": -50.0, "등급": 9}),   # 부호+허용값 위반
                  R("B", {"투자비": None, "등급": 1})],     # not_null 위반
        )
        rep = self._qc(bad)
        self.assertFalse(rep.passed)

    def test_range_violation_blocks(self):
        bad = M.MetricTableInput(
            fields=[F("물량", kind="quantity", nullable=False,
                      min_value=0, max_value=100)],
            rows=[R("A", {"물량": 150.0})],   # 상한 위반
        )
        self.assertFalse(self._qc(bad).passed)

    def test_tie_out_mismatch_blocks(self):
        bad = M.MetricTableInput(
            fields=[F("투자비", kind="currency", nullable=False, tie_total=100.0)],
            rows=[R("A", {"투자비": 60.0}), R("B", {"투자비": 30.0})],  # Σ=90≠100
        )
        self.assertFalse(self._qc(bad).passed)

    def test_clean_input_passes_with_nodata(self):
        ok = M.MetricTableInput(
            fields=[F("상각비", kind="currency", nullable=False, min_value=0),
                    F("비고물량", kind="quantity", nullable=True, min_value=0)],
            rows=[R("A", {"상각비": 10.0, "비고물량": None}),   # 결측 허용 → NO_DATA
                  R("B", {"상각비": 20.0, "비고물량": 5.0})],
        )
        self.assertTrue(self._qc(ok).passed)

    def test_duplicate_label_grain_blocks(self):
        bad = M.MetricTableInput(
            fields=[F("금액", kind="currency", sign="any")],
            rows=[R("A", {"금액": 1.0}), R("A", {"금액": 2.0})],  # 라벨 중복
        )
        self.assertFalse(self._qc(bad).passed)


if __name__ == "__main__":
    unittest.main()
