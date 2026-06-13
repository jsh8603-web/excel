"""
tests/test_stage.py — classify_stage / route 단계 라우팅 회귀 (stdlib unittest).

실행: py -m unittest tests.test_stage
회사 PC 무설치 동작. 합성 재무수치 없음 — 요청 텍스트 분류만 검증.

커버:
  - classify_stage: 더러운 엑셀→ingest / 반출 스키마→profile / 암복호화→transport
                    / 분석 요청→analysis(기본값) + 파일 단서(messy) 보강
  - route: 단계별 next_command + analysis 일 때 dispatch 위임으로 template 채움
  - dispatch 하위호환: 기존 라우팅 시그니처·결과 불변
"""
from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fpna._bootstrap  # noqa: F401

from fpna.dispatcher import classify_stage, route, dispatch, DispatchResult


class TestClassifyStage(unittest.TestCase):
    def test_ingest_messy_excel(self):
        for req in ("시트가 엉망인 더러운 엑셀 정리해줘",
                    "병합 셀 누더기 파일 정형화",
                    "이 비정형 엑셀 깨끗하게 만들어"):
            stage, cmd = classify_stage(req)
            self.assertEqual(stage, "ingest", req)
            self.assertIn("ingest", cmd)

    def test_profile_shape_export(self):
        stage, cmd = classify_stage("정제 마트를 회사에서 집으로 SHAPE 스키마로 반출")
        self.assertEqual(stage, "profile")
        self.assertIn("profile", cmd)

    def test_transport_crypto(self):
        for req in ("이 텍스트 암호화해서 메일 본문으로 보낼래",
                    "받은 암호문 복호화해줘"):
            stage, cmd = classify_stage(req)
            self.assertEqual(stage, "transport", req)
            self.assertIn("encrypt", cmd)

    def test_analysis_is_default(self):
        # 단계 키워드 없는 분석 요청은 analysis 로 떨어진다.
        for req in ("고정비 만기 도래 분석", "손익 계산서 만들어줘", "예실 변동 분석"):
            stage, _ = classify_stage(req)
            self.assertEqual(stage, "analysis", req)

    def test_messy_file_hint_boosts_ingest(self):
        # 텍스트 신호가 없어도 누더기 파일 단서가 있으면 ingest.
        stage, _ = classify_stage("이거 좀 봐줘", has_messy_file=True)
        self.assertEqual(stage, "ingest")
        # 단, 이미 tidy 라고 표시되면 analysis 유지.
        stage2, _ = classify_stage("이거 좀 봐줘", has_messy_file=True,
                                   has_clean_table=True)
        self.assertEqual(stage2, "analysis")


class TestRoute(unittest.TestCase):
    def test_route_non_analysis_has_no_template(self):
        r = route("더러운 엑셀 정리해줘")
        self.assertEqual(r["stage"], "ingest")
        self.assertIsNone(r["template"])
        self.assertIn("ingest", r["next_command"])

    def test_route_analysis_fills_template_via_dispatch(self):
        r = route("고정비 만기 도래")
        self.assertEqual(r["stage"], "analysis")
        # dispatch 와 동일 템플릿이어야 한다(위임).
        self.assertEqual(r["template"], dispatch("고정비 만기 도래").template)
        self.assertEqual(r["template"], "fc_maturity_wall")
        self.assertIn("render", r["next_command"])
        self.assertIn(r["template"], r["next_command"])

    def test_route_passes_columns_to_dispatch(self):
        cols = ["월", "budget", "actual"]
        r = route("월별 보고", columns=cols)
        self.assertEqual(r["stage"], "analysis")
        self.assertEqual(r["template"], dispatch("월별 보고", columns=cols).template)


class TestDispatchBackcompat(unittest.TestCase):
    """route/classify_stage 추가가 기존 dispatch 라우팅을 바꾸지 않음을 박제."""

    def test_dispatch_signature_unchanged(self):
        res = dispatch("예실 변동 분석")
        self.assertIsInstance(res, DispatchResult)
        self.assertEqual(res.template, "variance")

    def test_dispatch_default_pnl(self):
        self.assertEqual(dispatch("그냥 표 하나").template, "pnl_3statement")

    def test_dispatch_column_signal_unchanged(self):
        res = dispatch("월별 추이", columns=["월", "예산", "실적"])
        self.assertEqual(res.template, "variance")


if __name__ == "__main__":
    unittest.main()
