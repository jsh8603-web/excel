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

    # ------------------------------------------------------------------
    # 무음 손상 3종 방어 골든(G1 단위전파 / G2 혼재 / G4 ditto / G6 오류값).
    # ⚠ 합성 *구조* 더미 — 재무 수치는 의미 없음(스케일/플래그 로직 검증용).
    # ------------------------------------------------------------------
    @staticmethod
    def _ingest_wb(builder):
        """openpyxl.Workbook 빌더 콜백 → 임시파일 저장 → ingest 결과."""
        import openpyxl
        from fpna.ingest import ingest_workbook
        wb = openpyxl.Workbook()
        builder(wb.active)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as fh:
            path = fh.name
        try:
            wb.save(path)
            return ingest_workbook(path)
        finally:
            try:
                os.remove(path)
            except OSError:
                pass

    def test_g1_unit_scale_propagation(self):
        """G1: 블록 내부 '(단위: 백만원)' → 블록 전체 값 base(원) 환산 + raw 보존."""
        def build(ws):
            ws["A1"] = "(단위: 백만원)"
            ws["A2"], ws["B2"], ws["C2"], ws["D2"] = "계정", "2023", "2024", "2025"
            ws["A3"], ws["B3"], ws["C3"], ws["D3"] = "매출", 1000, 1200, 1500
            ws["A4"], ws["B4"], ws["C4"], ws["D4"] = "비용", 600, 800, 900
            ws["A5"], ws["B5"], ws["C5"], ws["D5"] = "이익", 400, 400, 600
        res = self._ingest_wb(build)
        rows = {(r.entity, r.period): r for r in res.tidy_rows}
        tr = rows[("매출", "2024")]
        self.assertEqual(tr.value, 1200 * 1_000_000)   # 백만원 → 원
        self.assertEqual(tr.scale_applied, 1_000_000)
        self.assertEqual(tr.scale_source, "block")
        self.assertEqual(tr.raw_value, 1200)           # 환산 전 원본 보존
        # 단위행이 데이터로 오인되지 않음(period/entity 정상 복원)
        self.assertEqual(rows[("이익", "2025")].value, 600 * 1_000_000)

    def test_g2_scale_heterogeneous(self):
        """G2: 블록단위(백만) 위에 셀 접미(억)가 우선 + 한 열 스케일 2종 → smell."""
        def build(ws):
            ws["A1"] = "(단위: 백만원)"
            ws["A2"], ws["B2"], ws["C2"] = "항목", "2024", "2025"
            ws["A3"], ws["B3"], ws["C3"] = "매출", 1000, "3,400억"
            ws["A4"], ws["B4"], ws["C4"] = "비용", 500, 600
            ws["A5"], ws["B5"], ws["C5"] = "이익", 500, 700
        res = self._ingest_wb(build)
        rows = {(r.entity, r.period): r for r in res.tidy_rows}
        # 셀 접미(억=1e8)가 블록(백만=1e6)을 덮어씀
        cell_tr = rows[("매출", "2025")]
        self.assertEqual(cell_tr.value, 3400 * 100_000_000)
        self.assertEqual(cell_tr.scale_source, "cell")
        # 같은 열 다른 셀은 블록 스케일
        self.assertEqual(rows[("비용", "2025")].scale_source, "block")
        self.assertEqual(rows[("비용", "2025")].value, 600 * 1_000_000)
        kinds = {s["kind"] for s in res.smells}
        self.assertIn("SCALE_HETEROGENEOUS", kinds)

    def test_g4_ditto_fill_down(self):
        """G4: 카테고리 열 상동 빈칸을 위→아래 단방향 충전 + DITTO_FILLED 플래그."""
        def build(ws):
            ws["A1"], ws["B1"], ws["C1"], ws["D1"] = "부문", "제품", "매출", "수량"
            ws["A2"], ws["B2"], ws["C2"], ws["D2"] = "국내", "제품A", 100, 10
            ws["A3"], ws["B3"], ws["C3"], ws["D3"] = None, "제품B", 200, 20
            ws["A4"], ws["B4"], ws["C4"], ws["D4"] = "해외", "제품C", 300, 30
            ws["A5"], ws["B5"], ws["C5"], ws["D5"] = None, "제품D", 400, 40
        res = self._ingest_wb(build)
        # 빈칸이 채워져 국내/해외가 모든 행에 전파
        ents = {r.entity for r in res.tidy_rows}
        self.assertEqual(ents, {"국내", "해외"})
        filled = [r for r in res.tidy_rows if "DITTO_FILLED" in r.flags]
        self.assertTrue(filled)                       # 일부 행은 ditto 충전 표시
        # 충전된 행은 직전 부문 값 상속(제품B → 국내, 제품D → 해외)
        b_rows = [r for r in res.tidy_rows if r.metric and "제품B" in r.metric]
        self.assertTrue(all(r.entity == "국내" for r in b_rows))
        kinds = {s["kind"] for s in res.smells}
        self.assertIn("DITTO_FILLED", kinds)

    def test_g6_error_cell_preserved(self):
        """G6: #DIV/0!/#N/A/#REF! → value=null + raw 보존 + ERROR_CELL smell.

        ⛔ 0/NaN coerce 금지(reject 도 아님 — 행은 유지, 값만 null)."""
        def build(ws):
            ws["A1"], ws["B1"], ws["C1"] = "계정", "2024", "2025"
            ws["A2"], ws["B2"], ws["C2"] = "매출", 1000, 1200
            ws["A3"], ws["B3"], ws["C3"] = "증가율", "#DIV/0!", "#N/A"
            ws["A4"], ws["B4"], ws["C4"] = "참조", "#REF!", 50
        res = self._ingest_wb(build)
        rows = {(r.entity, r.period): r for r in res.tidy_rows}
        err = rows[("증가율", "2024")]
        self.assertIsNone(err.value)                  # 0/NaN 아님 — null
        self.assertEqual(err.raw_value, "#DIV/0!")    # 원본 보존
        self.assertIn("ERROR_CELL", err.flags)
        # 정상 값은 그대로
        self.assertEqual(rows[("매출", "2024")].value, 1000)
        self.assertEqual(rows[("참조", "2025")].value, 50)
        # reject 되지 않고 smell 로만 표면화
        self.assertEqual(res.report.n_rejected, 0)
        err_smells = [s for s in res.smells if s["kind"] == "ERROR_CELL"]
        self.assertEqual(len(err_smells), 3)

    def test_fullwidth_and_suffix_normalization(self):
        """전각숫자/NBSP/접미 정규화 — normalize_value_ex 단위 검증."""
        from fpna.ingest.normalize import normalize_value_ex, split_cell_scale
        self.assertEqual(split_cell_scale("1,234천원"), ("1,234", 1000))
        self.assertEqual(split_cell_scale("3,400억"), ("3,400", 100_000_000))
        # 전각숫자 + 접미
        val, sent, neg, scale = normalize_value_ex("１，２３４천원")
        self.assertEqual(val, 1234)
        self.assertEqual(scale, 1000)
        # 퍼센트는 스케일 환산 대상 아님
        val, sent, neg, scale = normalize_value_ex("85%")
        self.assertAlmostEqual(val, 0.85, places=4)
        self.assertEqual(scale, 1)


class TestDispatch(unittest.TestCase):
    def test_keyword_routing(self):
        self.assertEqual(dispatch("NPV IRR 투자 타당성").template, "investment_appraisal")
        self.assertEqual(dispatch("예실 변동 브리지").template, "variance")
        self.assertEqual(dispatch("MoM 추이").template, "period_trend")

    def test_column_signal(self):
        d = dispatch("월간 보고", columns=["계획", "실적", "항목"])
        self.assertEqual(d.template, "variance")


class TestProfile(unittest.TestCase):
    """정제 마트 → SHAPE 추출. ⚠ 결정적 구조 픽스처(난수·실금액 아님)."""

    SEAS = [0.80, 0.77, 0.86, 0.92, 0.95, 0.99, 0.98, 0.98, 1.03, 1.11, 1.24, 1.37]

    def _make_mart(self, path):
        import csv
        rows = [["entity", "account", "period", "budget", "actual"]]
        for ent in ("E1", "E2"):
            for acc in ("Rev", "COGS"):
                for yi in range(2):
                    for mo in range(1, 13):
                        pi = yi * 12 + (mo - 1)
                        b = round(1000 * (1.02 ** pi) * self.SEAS[mo - 1])
                        a = round(b * 1.05)  # actual = budget×1.05 → 완전 상관
                        rows.append([ent, acc, "%d-%02d-01" % (2023 + yi, mo), b, a])
        with open(path, "w", encoding="utf-8-sig", newline="") as fh:
            csv.writer(fh).writerows(rows)

    def test_axes_recovered(self):
        from fpna.profile import profile_table
        with tempfile.TemporaryDirectory() as tmp:
            csvp = os.path.join(tmp, "mart.csv")
            self._make_mart(csvp)
            spec = profile_table(csvp)
            cols = spec["tables"]["mart"]["columns"]
            self.assertEqual(cols["entity"]["type"], "choice")
            self.assertEqual(cols["entity"]["n"], 2)
            self.assertEqual(cols["period"]["type"], "date")
            self.assertEqual(cols["budget"]["type"], "measure")
            # 추세 ≈ 0.02 (입력), 시즌 12개·12월>1월
            self.assertAlmostEqual(cols["budget"]["trend"], 0.02, places=2)
            seas = cols["budget"]["seasonality"]
            self.assertEqual(len(seas), 12)
            self.assertGreater(seas[11], seas[0])
            # actual = budget×1.05 → 상관 ≈ 1.0
            self.assertEqual(cols["actual"].get("corr_with"), "budget")
            self.assertGreater(cols["actual"]["corr"], 0.99)

    def test_no_value_leak(self):
        from fpna.profile import profile_table, emit_yaml
        with tempfile.TemporaryDirectory() as tmp:
            csvp = os.path.join(tmp, "mart.csv")
            self._make_mart(csvp)
            text = emit_yaml(profile_table(csvp))
            self.assertIn("EDIT_ME_scale", text)          # base = 자리표시자
            self.assertNotIn("1000", text)                # 절대 금액 미유출
            self.assertNotIn("1050", text)


class TestCrypto(unittest.TestCase):
    def test_rfc8439_vectors(self):
        from fpna._chacha import test_vectors
        self.assertTrue(test_vectors())

    def test_roundtrip_unicode_newline(self):
        from fpna.crypto import encrypt_text, decrypt_text
        msg = "다국어 ünïcode 1,234 (단위:천원)\n줄바꿈\t탭 ✓"
        arm = encrypt_text("pw-암호", msg)
        self.assertEqual(decrypt_text("pw-암호", arm), msg)

    def test_wrong_passphrase_rejected(self):
        from fpna.crypto import encrypt_text, decrypt_text
        arm = encrypt_text("right-pw", "secret")
        with self.assertRaises(ValueError):
            decrypt_text("wrong-pw", arm)

    def test_tamper_rejected(self):
        import base64
        from fpna.crypto import encrypt_text, decrypt_text
        arm = encrypt_text("pw", "secret payload")
        blob = bytearray(base64.b64decode(arm))
        blob[-1] ^= 0x01  # ciphertext 1비트 변조
        tampered = base64.b64encode(bytes(blob)).decode()
        with self.assertRaises(ValueError):
            decrypt_text("pw", tampered)

    def test_mail_split_roundtrip(self):
        from fpna.crypto import encrypt_text, to_mail_text, decrypt_text
        msg = "line %d data 1234 한글\n" * 300
        arm = encrypt_text("pw", msg)
        parts = to_mail_text(arm, max_lines=5, msg_id="T")
        self.assertGreater(len(parts), 1)                  # 여러 통으로 분할
        for p in parts:
            self.assertLessEqual(p.count("\n") + 1, 5)     # 메일당 줄수 한정 준수
        combined = "\n".join(reversed(parts))              # 역순으로 붙여도
        self.assertEqual(decrypt_text("pw", combined), msg)  # 정렬·복원


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
