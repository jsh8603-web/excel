"""
tests/test_pack.py — 표준 COA + 리볼버 솔버 + 다중 exhibit pack.

실행: py -m unittest tests.test_pack
검증: coa taxonomy 무결성 / solve_revolver 순환 수렴·sweep·floor·한도·토글 /
  build_pack 스파인 경유·receipt·Control·모델체크·cross tie 차단. 합성 재무수치 0(구조 더미).
"""
from __future__ import annotations

import os
import tempfile
import unittest

import fpna._bootstrap  # noqa: F401

from openpyxl import load_workbook

from fpna import coa, finance
from fpna.conserve import ConserveSpec
from fpna.pack import PackSpec, ExhibitSpec, ModelCheck, build_pack, _assemble
from fpna.templates import get_template


# --- COA taxonomy -------------------------------------------------------------
class CoaTest(unittest.TestCase):
    def test_json_equals_golden(self):
        j, g = coa.load_coa(), coa.golden_coa()
        self.assertEqual(set(l.code for l in j), set(l.code for l in g))
        self.assertEqual(len(j), 37)

    def test_validate_clean(self):
        self.assertEqual(coa.validate_coa(coa.load_coa()), [])

    def test_validate_catches_bad_parent(self):
        bad = coa.golden_coa() + [coa.FsLine("X", "x", "x", "IS", "x", "+", "NOPE")]
        probs = coa.validate_coa(bad)
        self.assertTrue(any("parent" in p for p in probs))

    def test_validate_catches_bad_statement(self):
        bad = [coa.FsLine("Y", "y", "y", "ZZ", "y", "+", None)]
        self.assertTrue(any("statement" in p for p in coa.validate_coa(bad)))

    def test_account_to_fs_line_priority(self):
        idx = coa.coa_index(coa.load_coa())

        class A:
            code = "Z"
            fs_line = "revenue"
        self.assertEqual(coa.account_to_fs_line(A(), idx), "revenue")  # fs_line 우선

        class B:
            code = "IS_REV"
            fs_line = None
        self.assertEqual(coa.account_to_fs_line(B(), idx), "revenue")  # code 매핑

        class C:
            code = "UNKNOWN"
            fs_line = None
        self.assertIsNone(coa.account_to_fs_line(C(), idx))            # 미매핑 = None(은폐 금지)

    def test_core_lines_present(self):
        idx = coa.coa_index(coa.load_coa())
        for need in ("BS_ASSETS", "BS_LIAB", "BS_EQ", "IS_NI", "CF_NET", "BS_RE", "BS_CASH"):
            self.assertIn(need, idx)

    def test_statement_split(self):
        lines = coa.load_coa()
        self.assertTrue(all(l.statement == "IS" for l in coa.statement_lines(lines, "IS")))
        self.assertGreaterEqual(len(coa.statement_lines(lines, "BS")), 10)


# --- solve_revolver 순환 솔버 -------------------------------------------------
class SolveRevolverTest(unittest.TestCase):
    def test_floor_draw_converges_to_min_cash(self):
        """선파이낸싱 부족 → 리볼버 floor 인출, end_cash=min_cash 수렴(이자 순환 포함)."""
        r = finance.solve_revolver(
            [-100.0], beginning_cash=0.0,
            beginning_debt={"revolver": 0.0, "senior": 1000.0},
            rates={"revolver": 0.01, "senior": 0.01},
            min_cash=50.0, sweep_priority=("revolver", "senior"))
        p = r.periods[0]
        self.assertTrue(r.converged)
        self.assertAlmostEqual(p["end_cash"], 50.0, places=6)
        self.assertGreater(p["revolver_draw"], 100.0)   # 이자만큼 더 인출(순환)

    def test_sweep_repays_debt(self):
        r = finance.solve_revolver(
            [500.0], beginning_cash=0.0,
            beginning_debt={"revolver": 0.0, "senior": 1000.0},
            rates={"revolver": 0.0, "senior": 0.0},
            min_cash=50.0, sweep_priority=("senior",))
        p = r.periods[0]
        self.assertAlmostEqual(p["debt_balance"], 550.0, places=6)   # 450 상환
        self.assertAlmostEqual(p["end_cash"], 50.0, places=6)
        self.assertAlmostEqual(p["repay"]["senior"], 450.0, places=6)

    def test_revolver_limit_clamps_draw(self):
        r = finance.solve_revolver(
            [-300.0], beginning_cash=0.0, beginning_debt={"revolver": 0.0},
            rates={"revolver": 0.0}, min_cash=0.0, revolver_limit=100.0)
        p = r.periods[0]
        self.assertAlmostEqual(p["revolver_draw"], 100.0, places=6)   # 한도 clamp
        self.assertAlmostEqual(p["end_cash"], -200.0, places=6)       # 부족 잔존(은폐 금지)

    def test_sweep_toggle_off_accumulates_cash(self):
        r = finance.solve_revolver(
            [500.0], beginning_cash=0.0, beginning_debt={"senior": 1000.0},
            rates={"senior": 0.0}, min_cash=50.0, sweep_priority=("senior",),
            sweep_enabled=False)
        p = r.periods[0]
        self.assertAlmostEqual(p["debt_balance"], 1000.0, places=6)   # 상환 0
        self.assertAlmostEqual(p["end_cash"], 500.0, places=6)        # 현금 적재

    def test_multi_period_rolls_balances(self):
        r = finance.solve_revolver(
            [-50.0, 200.0], beginning_cash=0.0,
            beginning_debt={"revolver": 0.0}, rates={"revolver": 0.01},
            min_cash=0.0, sweep_priority=("revolver",))
        self.assertEqual(len(r.periods), 2)
        # 1기 draw → 2기 잉여로 일부 paydown(잔액 연속)
        self.assertGreater(r.periods[0]["revolver_balance"], 0.0)
        self.assertLess(r.periods[1]["revolver_balance"], r.periods[0]["revolver_balance"])


# --- pack 오케스트레이터 ------------------------------------------------------
def _demo_spec(assets=1000.0, liab_eq=1000.0):
    pnl = get_template("pnl_3statement")
    wc = get_template("working_capital")
    return PackSpec(
        name="demo", title="데모 연동 팩",
        exhibits=(ExhibitSpec("pnl_3statement", pnl.golden_sample(), "손익", "summary"),
                  ExhibitSpec("working_capital", wc.golden_sample(), "운전자본", "detail")),
        shared_facts={"assets": assets, "liab_eq": liab_eq,
                      "cfs_end_cash": 50.0, "bs_cash": 50.0},
        cross_ties=(ConserveSpec("BS항등", raw_sum_fn=lambda s: s.shared_facts["assets"],
                                 reported_key="liab_eq", tol=0.5),
                    ConserveSpec("현금 tie", raw_sum_fn=lambda s: s.shared_facts["cfs_end_cash"],
                                 reported_key="bs_cash", tol=0.5)),
        model_checks=(ModelCheck("자산=부채+자본", "assets", "liab_eq"),
                      ModelCheck("현금 tie", "cfs_end_cash", "bs_cash")))


class PackTest(unittest.TestCase):
    def test_build_pack_spine_and_receipt(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "pack.xlsx")
            res = build_pack(_demo_spec(), out_path=p)
            self.assertTrue(res.saved)
            self.assertIsNotNone(res.receipt)
            self.assertTrue(res.qc.passed, res.qc.summary())
            self.assertTrue(os.path.isfile(p))

    def test_control_first_sheet(self):
        wb = _assemble(_demo_spec())
        self.assertTrue(wb.worksheets[0].title.startswith("Control"))
        titles = [w.title for w in wb.worksheets]
        self.assertEqual(len(titles), 3)        # Control + 2 exhibit

    def test_model_check_mismatch_blocks_save(self):
        """모델체크 불일치(자산≠부채+자본) → qc FAIL → 저장 차단(스파인 우회 불가)."""
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "bad.xlsx")
            res = build_pack(_demo_spec(assets=1000.0, liab_eq=1234.0), out_path=p)
            self.assertFalse(res.qc.passed)
            self.assertFalse(res.saved)
            self.assertFalse(os.path.isfile(p))   # receipt 없으면 파일 자체가 없음

    def test_cross_tie_mismatch_fails(self):
        """cross tie(BS항등) raw≠reported → 스파인 T4 보존 FAIL."""
        spec = _demo_spec()
        bad = PackSpec(
            name=spec.name, title=spec.title, exhibits=spec.exhibits,
            shared_facts={"assets": 1000.0, "liab_eq": 999.0,
                          "cfs_end_cash": 50.0, "bs_cash": 50.0},
            cross_ties=spec.cross_ties, model_checks=())   # 모델체크 비우고 tie만 검증
        res = build_pack(bad)
        self.assertFalse(res.qc.passed)
        self.assertTrue(any("BS항등" in c[0] and not c[1] for c in res.qc.checks),
                        res.qc.summary())

    def test_graft_preserves_cells(self):
        """graft 후 exhibit 시트에 데이터 셀이 보존(빈 시트 아님)."""
        wb = _assemble(_demo_spec())
        pnl_ws = wb.worksheets[1]
        non_empty = sum(1 for row in pnl_ws.iter_rows() for c in row if c.value is not None)
        self.assertGreater(non_empty, 5)


# --- pnl solve_and_link (순환 plug 해소) --------------------------------------
class PnlSolveLinkTest(unittest.TestCase):
    def test_solve_and_link_passes_3statement_tie(self):
        """현금부족 → 리볼버 draw plug 해소 → linked PnLInput 이 3-statement tie 전수통과."""
        P = get_template("pnl_3statement")
        inp = P.solve_and_link(
            revenue=1000, cogs=600, sga=200, da=50, tax_rate=0.22,
            beginning_debt={"revolver": 0.0, "term": 2000.0},
            debt_rates={"revolver": 0.01, "term": 0.008},
            pre_financing_cash=-500.0, cash_begin=100.0, re_begin=500.0,
            dividends=30.0, paid_in_capital=200.0, min_cash=50.0,
            sweep_priority=("revolver", "term"))
        self.assertTrue(inp.linked)
        self.assertGreater(inp.liabilities, 2000.0)   # 리볼버 draw 발생(plug)
        self.assertAlmostEqual(inp.cash, 50.0, places=6)  # 현금 = min_cash 수렴
        rep = P.qc(P.build(inp), inp)
        self.assertTrue(rep.passed, rep.summary())

    def test_surplus_no_draw(self):
        """선파이낸싱 잉여 → 리볼버 draw 0, 기존 부채 유지/감소."""
        P = get_template("pnl_3statement")
        inp = P.solve_and_link(
            revenue=1000, cogs=600, sga=200, da=50, tax_rate=0.22,
            beginning_debt={"revolver": 0.0, "term": 1000.0},
            debt_rates={"revolver": 0.01, "term": 0.008},
            pre_financing_cash=300.0, cash_begin=100.0, re_begin=500.0,
            dividends=30.0, paid_in_capital=200.0, min_cash=50.0,
            sweep_priority=("revolver", "term"))
        rep = P.qc(P.build(inp), inp)
        self.assertTrue(rep.passed, rep.summary())
        self.assertLessEqual(inp.liabilities, 1000.0)   # draw 없음(잉여)


# --- feasibility 팩 (5 exhibit + 5 cross tie) ---------------------------------
class FeasibilityPackTest(unittest.TestCase):
    def test_registry_and_build(self):
        from fpna.packs import get_pack, available
        self.assertIn("feasibility", available())
        spec = get_pack("feasibility").make_spec()
        self.assertEqual(len(spec.exhibits), 5)
        self.assertEqual(len(spec.cross_ties), 5)
        res = build_pack(spec)
        self.assertTrue(res.qc.passed, res.qc.summary())
        # 5 cross tie(T4 보존) 전수 통과
        ties = [c for c in res.qc.checks if c[0].startswith("T4 보존")]
        self.assertEqual(len(ties), 5)
        self.assertTrue(all(c[1] for c in ties), [c for c in ties if not c[1]])

    def test_six_sheets(self):
        from fpna.packs import get_pack
        wb = _assemble(get_pack("feasibility").make_spec())
        self.assertEqual(len(wb.worksheets), 6)        # Control + 5 exhibit
        self.assertTrue(wb.worksheets[0].title.startswith("Control"))


# --- dispatcher pack 게이트 ---------------------------------------------------
class PackRouteTest(unittest.TestCase):
    def test_pack_stage_and_resolve(self):
        from fpna.dispatcher import classify_stage, resolve_pack, route
        for txt in ("사업타당성 검토", "투자심사 연동 모델"):
            stage, cmd = classify_stage(txt)
            self.assertEqual(stage, "pack", txt)
            self.assertIn("feasibility", cmd)
        self.assertEqual(resolve_pack("사업타당성"), "feasibility")
        self.assertEqual(route("투자심사 연동")["stage"], "pack")

    def test_single_intent_not_pack(self):
        """단일 의도(감가상각만)는 pack 아님 → analysis 위임."""
        from fpna.dispatcher import route
        r = route("감가상각 스케줄")
        self.assertEqual(r["stage"], "analysis")
        self.assertEqual(r["template"], "fc_depreciation_schedule")


if __name__ == "__main__":
    unittest.main()
