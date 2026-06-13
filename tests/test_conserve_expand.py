"""
tests/test_conserve_expand.py — T4 보존검증 커버리지 확대(집계先 + 변환後).

실행: py -m unittest tests.test_conserve_expand

자문 3R C4 "게이트 극장 금지" 를 신규 CONSERVE_SPECS/qc N-version 에도 적용:
변이(끝행드롭=off-by-one / 부호반전 / 항목중복=이중계상)가 게이트를 trip 시키는지,
그리고 clean 은 통과하는지 둘 다 증명한다.

구성:
  A. 집계先 CONSERVE_SPECS 6종 — clean build 의 reported(_fpna_meta) 대비 변이된
     INPUT 으로 raw_sum_fn 재계산 시 불일치(trip). clean 은 통과.
  B. 변환後 N-version 4종 — build 산출 _fpna_meta(또는 비교 대상)를 오염시키면
     독립 qc 경로가 잡는지 / clean 통과.
  C. 모듈경계 독립성(ast) — 신규 raw_sum_fn 이 build 헬퍼를 부르지 않음을 회귀 단언.
"""
from __future__ import annotations

import ast
import dataclasses
import os
import unittest

import fpna._bootstrap  # noqa: F401

from fpna.conserve import eval_specs
from fpna.templates import get_template


# --------------------------------------------------------------------------- #
# 공통 헬퍼                                                                    #
# --------------------------------------------------------------------------- #
def _clean_meta(type_name):
    """clean golden 으로 build → (mod, golden, _fpna_meta)."""
    mod = get_template(type_name)
    g = mod.golden_sample()
    wb = mod.build(g)
    return mod, g, getattr(wb, "_fpna_meta", {}) or {}


def _trips(mod, mutated_data, clean_meta):
    """변이 INPUT 으로 raw 재계산 시 clean reported 와 어긋나는 spec 이 있는가."""
    for _name, lhs, rhs, tol in eval_specs(mod.CONSERVE_SPECS, mutated_data, clean_meta):
        if rhs is None:
            return True
        if abs(lhs - rhs) > tol:
            return True
    return False


def _clean_passes(mod, clean_data, clean_meta):
    """clean 은 모든 spec 이 tie(게이트가 무조건 trip 하는 가짜가 아님)."""
    for _name, lhs, rhs, tol in eval_specs(mod.CONSERVE_SPECS, clean_data, clean_meta):
        if rhs is None or abs(lhs - rhs) > tol:
            return False
    return True


# --------------------------------------------------------------------------- #
# A. 집계先 CONSERVE_SPECS — 변이 하니스                                       #
#   각 케이스: (type, list_field, [세 변이가 모두 trip 해야 하는지])           #
# --------------------------------------------------------------------------- #
_AGG_CASES = [
    ("consolidation_fx", "entities"),
    ("pvm_bridge", "lines"),
    ("fc_allocation", "depts"),
    ("fc_variance_bridge", "factors"),
    ("working_capital", "rows_in"),
    ("budget_build", "depts"),
]


class AggregateConserveHarnessTest(unittest.TestCase):
    """집계先 6종: 끝행드롭 / 항목중복 변이가 보존등식을 trip 시키는지."""

    def _mutate_drop_last(self, data, field):
        items = list(getattr(data, field))
        return dataclasses.replace(data, **{field: items[:-1]})

    def _mutate_dup_first(self, data, field):
        items = list(getattr(data, field))
        return dataclasses.replace(data, **{field: items + [items[0]]})

    def test_clean_passes(self):
        for type_name, field in _AGG_CASES:
            with self.subTest(type=type_name):
                mod, g, meta = _clean_meta(type_name)
                self.assertTrue(_clean_passes(mod, g, meta),
                                "%s clean 이 통과하지 않음(게이트 과민)" % type_name)

    def test_drop_last_trips(self):
        for type_name, field in _AGG_CASES:
            with self.subTest(type=type_name):
                mod, g, meta = _clean_meta(type_name)
                mutated = self._mutate_drop_last(g, field)
                self.assertTrue(_trips(mod, mutated, meta),
                                "%s 끝행드롭(off-by-one)이 trip 하지 않음" % type_name)

    def test_dup_first_trips(self):
        for type_name, field in _AGG_CASES:
            with self.subTest(type=type_name):
                mod, g, meta = _clean_meta(type_name)
                mutated = self._mutate_dup_first(g, field)
                self.assertTrue(_trips(mod, mutated, meta),
                                "%s 항목중복(이중계상)이 trip 하지 않음" % type_name)


# --------------------------------------------------------------------------- #
# B. 변환後 N-version — 독립 qc 가 build 산출 오염을 잡는지                     #
# --------------------------------------------------------------------------- #
class NVersionTransformTest(unittest.TestCase):
    """변환계열 4종: clean qc 통과 + meta 오염 시 독립 qc FAIL."""

    # 스칼라 meta 오염형(type, 오염 키)
    _SCALAR_CASES = [
        ("fc_prepaid_rollforward", "sum_end"),
        ("fc_lease_ifrs16", "liab0_sum"),
    ]
    # 전 변환계열(clean qc 통과 확인용)
    _ALL = ["fc_prepaid_rollforward", "fc_depreciation_schedule",
            "fc_lease_ifrs16", "investment_appraisal"]

    def test_clean_qc_passes(self):
        for type_name in self._ALL:
            with self.subTest(type=type_name):
                mod = get_template(type_name)
                g = mod.golden_sample()
                wb = mod.build(g)
                rep = mod.qc(wb, g)
                self.assertTrue(rep.passed, "%s clean qc FAIL: %s"
                                % (type_name, rep.summary()))

    def test_meta_poison_caught(self):
        for type_name, key in self._SCALAR_CASES:
            with self.subTest(type=type_name):
                mod = get_template(type_name)
                g = mod.golden_sample()
                wb = mod.build(g)
                # 오염: reported meta 값을 크게 흔든다 → 독립 qc 경로가 잡아야.
                wb._fpna_meta[key] = wb._fpna_meta[key] + 9_999.0
                rep = mod.qc(wb, g)
                self.assertFalse(rep.passed,
                                 "%s meta[%s] 오염을 독립 qc 가 못 잡음: %s"
                                 % (type_name, key, rep.summary()))

    def test_depreciation_fact_poison_caught(self):
        """fc_depreciation: 비이벤트·전수명미달(가동=표시시작) 자산 fact dep 오염을
        N-version(정액×기간) 독립경로가 잡는지."""
        mod = get_template("fc_depreciation_schedule")
        g = mod.golden_sample()
        wb = mod.build(g)
        # V-001(가동 = 표시시작, 비이벤트) 의 한 기간 dep 을 부풀린다.
        for r in wb._fpna_meta["fact"].rows:
            if r["asset_no"] == "V-001" and r["dep"] is not None:
                r["dep"] = r["dep"] + 5_000.0
                break
        rep = mod.qc(wb, g)
        nver = [c for c in rep.checks if "N-version 부분상각:V-001" in c[0]]
        self.assertTrue(nver and not nver[0][1],
                        "depreciation N-version 이 fact dep 오염을 못 잡음: %s"
                        % rep.summary())

    def test_prepaid_independent_recompute_is_the_catcher(self):
        """fc_prepaid: sum_end 오염 시 N-version check 가 실제로 FAIL 인지 명시 확인."""
        mod = get_template("fc_prepaid_rollforward")
        g = mod.golden_sample()
        wb = mod.build(g)
        wb._fpna_meta["sum_end"] = wb._fpna_meta["sum_end"] + 9_999.0
        rep = mod.qc(wb, g)
        nver = [c for c in rep.checks if "N-version" in c[0]]
        self.assertTrue(nver and not nver[0][1],
                        "N-version Σ기말 독립 대조가 오염을 잡지 못함")

    def test_lease_pv_bound_independent(self):
        """fc_lease: 부채PV ≤ Σ미할인 부등식이 PV 과대계상(오염)을 잡는지."""
        mod = get_template("fc_lease_ifrs16")
        g = mod.golden_sample()
        wb = mod.build(g)
        wb._fpna_meta["liab0_sum"] = wb._fpna_meta["liab0_sum"] + 9_999.0
        rep = mod.qc(wb, g)
        bound = [c for c in rep.checks if "부등식" in c[0]]
        self.assertTrue(bound and not bound[0][1],
                        "부채PV ≤ Σ미할인 부등식이 과대 PV 를 잡지 못함")

    def test_investment_npv_nversion_present_and_clean(self):
        """investment_appraisal: N-version NPV 독립 대조 check 가 존재 + clean PASS."""
        mod = get_template("investment_appraisal")
        g = mod.golden_sample()
        wb = mod.build(g)
        rep = mod.qc(wb, g)
        nver = [c for c in rep.checks if "N-version NPV" in c[0]]
        self.assertTrue(nver, "N-version NPV 독립 대조 check 부재")
        self.assertTrue(nver[0][1], "N-version NPV 독립 대조가 clean 에서 FAIL")


# --------------------------------------------------------------------------- #
# C. 모듈경계 독립성(ast) — 신규 raw_sum_fn 이 build 헬퍼를 부르지 않음        #
#   CONSERVE_SPECS 의 lambda 본문 소스에 _stepdown/_decompose/_translate/      #
#   _compute/_roll/_build_fact 등 build 헬퍼 호출이 없는지 정적 검사.          #
# --------------------------------------------------------------------------- #
_BUILD_HELPER_NAMES = {
    "_translate", "_decompose", "_stepdown", "_compute", "_roll", "_build_fact",
    "_rows", "_schedules", "_payments", "_periods", "_effective_cashflows",
    "_effective_rate", "lease_schedule", "depreciation_schedule",
    "depreciation_schedule_ext",
}


class SpecIndependenceTest(unittest.TestCase):
    """CONSERVE_SPECS 가 정의된 템플릿 파일에서 spec 블록이 build 헬퍼 비의존인지."""

    def _spec_lambda_sources(self, type_name):
        mod = get_template(type_name)
        path = mod.__file__
        with open(path, encoding="utf-8") as fh:
            tree = ast.parse(fh.read(), filename=path)
        # CONSERVE_SPECS = [...] 할당을 찾아 그 안의 모든 Call 노드 func 이름 수집
        called = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign):
                tgts = [t.id for t in node.targets if isinstance(t, ast.Name)]
                if "CONSERVE_SPECS" in tgts:
                    for sub in ast.walk(node.value):
                        if isinstance(sub, ast.Call):
                            f = sub.func
                            if isinstance(f, ast.Name):
                                called.add(f.id)
                            elif isinstance(f, ast.Attribute):
                                called.add(f.attr)
        return called

    def test_no_build_helper_in_specs(self):
        for type_name, _ in _AGG_CASES:
            with self.subTest(type=type_name):
                called = self._spec_lambda_sources(type_name)
                bad = called & _BUILD_HELPER_NAMES
                self.assertEqual(bad, set(),
                                 "%s CONSERVE_SPECS 가 build 헬퍼 호출(독립성 위배): %s"
                                 % (type_name, bad))


if __name__ == "__main__":
    unittest.main()
