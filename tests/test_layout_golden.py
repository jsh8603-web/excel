"""
tests/test_layout_golden.py — 서식/레이아웃 골든 스냅샷 회귀 (자문 2026-06).

목적: test_parity 가 값+수식만 봐서 못 잡던 "양식 불균형"(서식이 엉뚱한 range 에
묻음)을 골든 드리프트로 잡는다. 전 템플릿을 run_report 로 빌드 → layout_audit.fingerprint
→ tests/golden_layout/<type>.json 과 대조.

자가 생성(snapshot 패턴): 베이스라인이 없으면 *현재 출력으로 민팅*하고 그 케이스를
skip 한다(메시지: 검토 후 커밋). 베이스라인이 있으면 엄격 비교 — 차이 있으면 FAIL.
의도된 서식 변경은 baseline JSON 을 지우고 재실행(re-bless)하면 된다.

실행: py -m unittest tests.test_layout_golden
"""
from __future__ import annotations

import json
import os
import tempfile
import unittest

import fpna._bootstrap  # noqa: F401

from openpyxl import load_workbook

from fpna import layout_audit
from fpna.pipeline import run_report
from fpna.templates import available, get_template

_BASE_DIR = os.path.join(os.path.dirname(__file__), "golden_layout")


def _fingerprint_for(type_name: str) -> dict:
    mod = get_template(type_name)
    data = mod.golden_sample()
    with tempfile.TemporaryDirectory() as d:
        res = run_report(mod, data, out_path=os.path.join(d, "g.xlsx"))
        # QC 미통과면 저장이 안 되므로(receipt 없음) 빌드된 wb 를 직접 지문화.
        wb = res.wb if not res.saved else load_workbook(res.out_path)
    return layout_audit.fingerprint(wb)


class LayoutGoldenTest(unittest.TestCase):
    def test_all_templates_layout_stable(self):
        os.makedirs(_BASE_DIR, exist_ok=True)
        minted, drifted = [], []
        for t in available():
            cur = _fingerprint_for(t)
            path = os.path.join(_BASE_DIR, "%s.json" % t)
            if not os.path.exists(path):
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(cur, f, ensure_ascii=False, indent=1, sort_keys=True)
                minted.append(t)
                continue
            with open(path, encoding="utf-8") as f:
                base = json.load(f)
            diffs = layout_audit.diff_fingerprints(base, cur)
            if diffs:
                drifted.append((t, diffs[:6]))
        if drifted:
            msg = "\n".join(
                "  [%s] " % t + "; ".join("%s: %s→%s" % (p, b, c) for p, b, c in ds)
                for t, ds in drifted)
            self.fail("레이아웃 골든 드리프트:\n" + msg +
                      "\n(의도된 변경이면 해당 golden_layout/*.json 삭제 후 재실행)")
        if minted:
            self.skipTest("베이스라인 신규 민팅(검토 후 커밋): " + ", ".join(minted))


if __name__ == "__main__":
    unittest.main()
