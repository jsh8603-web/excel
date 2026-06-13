"""
tests/test_binding.py — T2 실데이터 바인딩 + T4 보존(conserves) 회귀.

실행: py -m unittest tests.test_binding
회사 PC 무설치(stdlib + openpyxl) 동작. 합성 재무수치 없음 — 각 템플릿의
golden_sample() INPUT 을 tidy rows 로 역직렬화해 round-trip 동치를 단언한다.

커버(T2 from_tidy):
  fc_maturity_wall / headcount_plan / budget_build / cost_allocation /
  cohort_retention / debt_schedule / variance (7종)
conserves(T4) 커버: fc_maturity_wall / headcount_plan / cost_allocation /
  cohort_retention (4종). budget_build·debt_schedule·variance 는 deferred
  (각 모듈 주석에 사유 박제 — _fpna_meta 보고총계 부재 또는 복잡로직 중복).
"""
from __future__ import annotations

import csv as _csv
import datetime as _dt
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fpna._bootstrap  # noqa: F401

from fpna import binding
from fpna.binding import bind_and_check, bind_from_csv, _coerce, assemble
from fpna.pipeline import run_report
from fpna.templates import get_template


# --------------------------------------------------------------------------- #
# golden INPUT → tidy rows 역직렬화 헬퍼 (템플릿별)                            #
# --------------------------------------------------------------------------- #
def _rows_fc_maturity_wall(inp):
    out = []
    for c in inp.contracts:
        out.append({
            "contract_id": c.contract_id, "account_id": c.account_id,
            "counterparty": c.counterparty,
            "start_date": c.start_date.isoformat(),
            "end_date": (c.end_date.isoformat() if c.end_date else ""),
            "recurrence": c.recurrence,
            "amount_per_period": str(c.amount_per_period),
            "status": c.status,
        })
    return out


def _rows_headcount_plan(inp):
    # dept × period 펼치기 (period 라벨은 합성 — 순서만 의미)
    out = []
    for ln in inp.lines:
        for i, hc in enumerate(ln.headcount):
            out.append({
                "dept": ln.dept, "period": "P%02d" % (i + 1),
                "headcount": str(hc),
                "base_salary_annual": str(ln.base_salary_annual),
                "loading_rate": str(ln.loading_rate),
                "grade": ln.grade,
            })
    return out


def _rows_budget_build(inp):
    out = []
    for d in inp.depts:
        out.append({
            "dept": d.dept, "headcount": str(d.headcount),
            "avg_cost": str(d.avg_cost), "method": d.method,
            "prior_budget": str(d.prior_budget),
        })
    return out


def _rows_cost_allocation(inp):
    out = []
    for p in inp.pools:
        for cc in inp.cost_centers:
            w = inp.driver_weights.get(p.pool_id, {}).get(cc, "")
            out.append({
                "pool_id": p.pool_id, "cost_center": cc,
                "weight": ("" if w == "" else str(w)),
                "pool_label": p.label, "amount": str(p.amount),
                "driver": p.driver, "cc_label": inp.cc_labels.get(cc, cc),
            })
    return out


def _rows_cohort_retention(inp):
    out = []
    for ln in inp.cohorts:
        for s in ln.steps:
            out.append({
                "cohort": ln.cohort, "age": str(s.age),
                "start_mrr": str(s.start_mrr), "churn": str(s.churn),
                "contraction": str(s.contraction), "expansion": str(s.expansion),
            })
    return out


def _rows_debt_schedule(inp):
    # period 라벨 = P01.. (순서만 의미). cash_available 는 period 단위.
    out = []
    nP = max((len(t.mandatory) for t in inp.tranches), default=0)
    nP = max(nP, len(inp.cash_available))
    for t in inp.tranches:
        for i in range(nP):
            mand = t.mandatory[i] if i < len(t.mandatory) else 0.0
            cash = inp.cash_available[i] if i < len(inp.cash_available) else 0.0
            out.append({
                "tranche_id": t.tranche_id, "period": "P%02d" % (i + 1),
                "label": t.label, "kind": t.kind,
                "opening": str(t.opening), "rate": str(t.rate),
                "sweep_enabled": str(t.sweep_enabled),
                "mandatory": str(mand), "cash_available": str(cash),
            })
    return out


def _rows_variance(inp):
    out = []
    for it in inp.items:
        out.append({
            "name": it.name, "plan": str(it.plan), "actual": str(it.actual),
            "cost_nature": str(it.cost_nature), "level": str(it.level),
            "is_total": str(it.is_total), "key": it.key,
        })
    return out


# (type, rows_fn, [(attr, getter)] 핵심필드 동치 검사기)
_SERIALIZERS = {
    "fc_maturity_wall": _rows_fc_maturity_wall,
    "headcount_plan": _rows_headcount_plan,
    "budget_build": _rows_budget_build,
    "cost_allocation": _rows_cost_allocation,
    "cohort_retention": _rows_cohort_retention,
    "debt_schedule": _rows_debt_schedule,
    "variance": _rows_variance,
}
_CONSERVES_TYPES = ("fc_maturity_wall", "headcount_plan",
                    "cost_allocation", "cohort_retention")


class TestCoerce(unittest.TestCase):
    def test_number_strip(self):
        self.assertEqual(_coerce("1,200", int), 1200)
        self.assertEqual(_coerce("₩1,200원", float), 1200.0)
        self.assertEqual(_coerce("(1,200)", int), -1200)        # 회계 괄호 음수
        self.assertEqual(_coerce("$3.5", float), 3.5)

    def test_int_with_decimal_raises(self):
        with self.assertRaises(ValueError):
            _coerce("3.5", int)

    def test_str_and_blank(self):
        self.assertEqual(_coerce("  ab ", str), "ab")
        self.assertIsNone(_coerce("", float))
        self.assertIsNone(_coerce("   ", int))

    def test_date_iso(self):
        self.assertEqual(_coerce("2024-09-30", _dt.date), _dt.date(2024, 9, 30))
        self.assertEqual(_coerce("2024/09/30", _dt.date), _dt.date(2024, 9, 30))
        self.assertEqual(_coerce(_dt.date(2024, 1, 1), _dt.date), _dt.date(2024, 1, 1))

    def test_bool(self):
        self.assertTrue(_coerce("True", bool))
        self.assertTrue(_coerce("yes", bool))
        self.assertFalse(_coerce("0", bool))
        self.assertFalse(_coerce("False", bool))

    def test_none_typ_passthrough(self):
        self.assertEqual(_coerce("raw", None), "raw")


class TestAssembleTree(unittest.TestCase):
    def test_two_level_groupby(self):
        # cohort_retention 형태로 2단 트리 조립 검증
        from fpna.templates.cohort_retention import (CohortRetentionInput,
                                                     CohortLine, CohortStep)
        rows = [
            {"cohort": "A", "age": "0", "start_mrr": "100", "churn": "0",
             "contraction": "0", "expansion": "0"},
            {"cohort": "A", "age": "1", "start_mrr": "100", "churn": "10",
             "contraction": "0", "expansion": "0"},
            {"cohort": "B", "age": "0", "start_mrr": "50", "churn": "0",
             "contraction": "0", "expansion": "0"},
        ]
        spec = {
            "header_cls": CohortRetentionInput, "header_fields": {},
            "levels": [{"key_cols": ["cohort"], "line_cls": CohortLine,
                        "fields": {"cohort": ("cohort", str)},
                        "child_attr": "cohorts",
                        "child": {"child_attr": "steps", "key_cols": ["age"],
                                  "line_cls": CohortStep,
                                  "fields": {"age": ("age", int),
                                             "start_mrr": ("start_mrr", float),
                                             "churn": ("churn", float),
                                             "contraction": ("contraction", float),
                                             "expansion": ("expansion", float)}}}],
        }
        inp = assemble(rows, spec)
        self.assertEqual([c.cohort for c in inp.cohorts], ["A", "B"])
        self.assertEqual(len(inp.cohorts[0].steps), 2)          # A: age 0,1
        self.assertEqual(len(inp.cohorts[1].steps), 1)          # B: age 0
        self.assertEqual(inp.cohorts[0].steps[1].churn, 10.0)

    def test_empty_rows(self):
        from fpna.templates.budget_build import BudgetInput, DeptLine  # noqa: F401
        spec = {"header_cls": BudgetInput, "header_fields": {},
                "levels": [{"key_cols": ["dept"], "line_cls": DeptLine,
                            "fields": {"dept": ("dept", str)},
                            "child_attr": "depts"}]}
        inp = assemble([], spec)
        self.assertEqual(inp.depts, [])


class TestBindRoundTrip(unittest.TestCase):
    """golden INPUT → tidy rows → bind_and_check → 핵심 필드 복원 동치."""

    def test_all_covered_templates_round_trip(self):
        for t, rows_fn in _SERIALIZERS.items():
            mod = get_template(t)
            golden = mod.golden_sample()
            rows = rows_fn(golden)
            inp = bind_and_check(mod, rows)
            self._assert_equiv(t, golden, inp)

    def _assert_equiv(self, t, golden, inp):
        if t == "fc_maturity_wall":
            self.assertEqual(len(inp.contracts), len(golden.contracts))
            g = {c.contract_id: c for c in golden.contracts}
            for c in inp.contracts:
                self.assertAlmostEqual(c.amount_per_period,
                                       g[c.contract_id].amount_per_period)
                self.assertEqual(c.end_date, g[c.contract_id].end_date)
                self.assertEqual(c.recurrence, g[c.contract_id].recurrence)
        elif t == "headcount_plan":
            self.assertEqual(len(inp.lines), len(golden.lines))
            g = {l.dept: l for l in golden.lines}
            for l in inp.lines:
                self.assertEqual(l.headcount, g[l.dept].headcount)
                self.assertAlmostEqual(l.base_salary_annual,
                                       g[l.dept].base_salary_annual)
                self.assertAlmostEqual(l.loading_rate, g[l.dept].loading_rate)
        elif t == "budget_build":
            self.assertEqual(len(inp.depts), len(golden.depts))
            g = {d.dept: d for d in golden.depts}
            for d in inp.depts:
                self.assertEqual(d.headcount, g[d.dept].headcount)
                self.assertAlmostEqual(d.avg_cost, g[d.dept].avg_cost)
                self.assertEqual(d.method, g[d.dept].method)
        elif t == "cost_allocation":
            self.assertEqual(set(inp.cost_centers), set(golden.cost_centers))
            self.assertEqual({p.pool_id for p in inp.pools},
                             {p.pool_id for p in golden.pools})
            gp = {p.pool_id: p for p in golden.pools}
            for p in inp.pools:
                self.assertAlmostEqual(p.amount, gp[p.pool_id].amount)
            # driver_weights 복원
            for pid, ccw in golden.driver_weights.items():
                for cc, w in ccw.items():
                    if w:   # 0 가중은 빈셀로 직렬화되어 생략됨(의도)
                        self.assertAlmostEqual(inp.driver_weights[pid][cc], w)
        elif t == "cohort_retention":
            self.assertEqual({c.cohort for c in inp.cohorts},
                             {c.cohort for c in golden.cohorts})
            g = {c.cohort: c for c in golden.cohorts}
            for c in inp.cohorts:
                self.assertEqual(len(c.steps), len(g[c.cohort].steps))
                gs = {s.age: s for s in g[c.cohort].steps}
                for s in c.steps:
                    self.assertAlmostEqual(s.end_mrr, gs[s.age].end_mrr)
        elif t == "debt_schedule":
            self.assertEqual({x.tranche_id for x in inp.tranches},
                             {x.tranche_id for x in golden.tranches})
            g = {x.tranche_id: x for x in golden.tranches}
            for x in inp.tranches:
                self.assertAlmostEqual(x.opening, g[x.tranche_id].opening)
                self.assertAlmostEqual(x.rate, g[x.tranche_id].rate)
                self.assertEqual(x.kind, g[x.tranche_id].kind)
                # mandatory 동치 — 빈 벡터([])와 [0.0,..]는 의미상 동일(0 의무상환).
                # zero-padding 후 비교(revolver 의 mandatory=[] 직렬화 아티팩트 흡수).
                a, b = x.mandatory, g[x.tranche_id].mandatory
                n = max(len(a), len(b))
                ap = a + [0.0] * (n - len(a))
                bp = b + [0.0] * (n - len(b))
                for u, v in zip(ap, bp):
                    self.assertAlmostEqual(u, v)
            self.assertEqual(inp.cash_available, golden.cash_available)
        elif t == "variance":
            self.assertEqual(len(inp.items), len(golden.items))
            g = {it.name: it for it in golden.items}
            for it in inp.items:
                self.assertAlmostEqual(it.plan, g[it.name].plan)
                self.assertAlmostEqual(it.actual, g[it.name].actual)
                self.assertEqual(it.is_total, g[it.name].is_total)
                self.assertEqual(it.cost_nature, g[it.name].cost_nature)


class TestBindFromCsv(unittest.TestCase):
    def test_csv_round_trip(self):
        # fc_maturity_wall 1종 csv 왕복(tmpfile)
        mod = get_template("fc_maturity_wall")
        golden = mod.golden_sample()
        rows = _rows_fc_maturity_wall(golden)
        cols = list(rows[0].keys())
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "contracts.csv")
            with open(p, "w", encoding="utf-8-sig", newline="") as fh:
                w = _csv.DictWriter(fh, fieldnames=cols)
                w.writeheader()
                w.writerows(rows)
            inp = bind_from_csv(mod, p)
        self.assertEqual(len(inp.contracts), len(golden.contracts))
        ids = {c.contract_id for c in inp.contracts}
        self.assertEqual(ids, {c.contract_id for c in golden.contracts})


class TestConservesViaSpine(unittest.TestCase):
    """conserves 커버 템플릿: run_report 스파인이 T4 보존 게이트 통과."""

    def test_conserves_pass(self):
        for t in _CONSERVES_TYPES:
            mod = get_template(t)
            golden = mod.golden_sample()
            self.assertTrue(hasattr(mod, "conserves"),
                            "%s conserves 누락" % t)
            with tempfile.TemporaryDirectory() as d:
                out = os.path.join(d, t + ".xlsx")
                res = run_report(mod, golden, out_path=out)
                self.assertTrue(res.saved, "%s 미저장" % t)
                self.assertTrue(res.qc.passed,
                                "%s 스파인 QC FAIL:\n%s" % (t, res.qc.summary()))
                self.assertTrue(os.path.exists(out))
                # T4 보존 체크가 실제로 리포트에 등재됐는지(이름에 'T4 보존')
                names = [n for n, _, _ in res.qc.checks if n.startswith("T4 보존")]
                self.assertTrue(names, "%s T4 보존 체크 미등재" % t)

    def test_conserves_catches_corrupted_meta(self):
        """build 보고 총계(_fpna_meta)를 위조하면 conserves(T4)가 FAIL."""
        # fc_maturity_wall: grand 위조 → 독립 재산출과 불일치 → 저장 거부
        mod = get_template("fc_maturity_wall")
        golden = mod.golden_sample()

        class _Shim:                              # build 출력 grand 를 오염시키는 래퍼
            TYPE = mod.TYPE

            def build(self, data, *, mode="create", base_path=None):
                wb = mod.build(data, mode=mode, base_path=base_path)
                wb._fpna_meta["grand"] += 999.0   # 보고 총계 오염
                return wb

            conserves = staticmethod(mod.conserves)
            qc = staticmethod(mod.qc)

        with tempfile.TemporaryDirectory() as d:
            out = os.path.join(d, "x.xlsx")
            res = run_report(_Shim(), golden, out_path=out)
            self.assertFalse(res.saved)           # 보존 위반 → 미저장
            self.assertFalse(res.qc.passed)
            self.assertFalse(os.path.exists(out))


class TestGrainPreShape(unittest.TestCase):
    """grain_unique 를 pre-shape rows 에서 검사(자문 C8-3: groupby 접기 前)."""

    def test_duplicate_grain_rejected(self):
        mod = get_template("fc_maturity_wall")
        golden = mod.golden_sample()
        rows = _rows_fc_maturity_wall(golden)
        rows.append(dict(rows[0]))                # contract_id 중복 1건
        with self.assertRaises(ValueError):
            bind_and_check(mod, rows)

    def test_missing_grain_key_rejected(self):
        mod = get_template("cohort_retention")
        golden = mod.golden_sample()
        rows = _rows_cohort_retention(golden)
        rows[0]["age"] = ""                       # grain key(age) 누락
        with self.assertRaises(ValueError):
            bind_and_check(mod, rows)

    def test_multilevel_dup_seen_before_groupby(self):
        """1:N 트리에서 (cohort,age) 중복은 groupby 접힘 前 pre-shape 가 잡는다."""
        mod = get_template("cohort_retention")
        golden = mod.golden_sample()
        rows = _rows_cohort_retention(golden)
        rows.append(dict(rows[0]))                # (cohort,age) 중복
        with self.assertRaises(ValueError):
            bind_and_check(mod, rows)


class TestRequiredAndUnits(unittest.TestCase):
    def test_required_empty_rejected(self):
        mod = get_template("variance")
        with self.assertRaises(ValueError):
            bind_and_check(mod, [])               # items 비면 REQUIRED FAIL

    def test_unsupported_template_raises(self):
        # from_tidy 미노출 템플릿 → 명확한 NotImplementedError
        mod = get_template("pvm_bridge")
        self.assertFalse(hasattr(mod, "from_tidy"))
        with self.assertRaises(NotImplementedError):
            bind_and_check(mod, [{"x": "1"}])


if __name__ == "__main__":
    unittest.main()
