"""
fpna.render — 템플릿 빌드 + QC 게이트 + 저장.

생성모드(create): house_style 로 처음부터 작성.
채우기모드(fill): templates_base/ 의 베이스 .xlsx 를 열어 데이터 셀만 채움
  (값·서식·수식 보존). ⚠ openpyxl 라운드트립 한계로 베이스의 차트/이미지/피벗은
  빠질 수 있음 → 차트 많은 베이스는 생성모드 권장.

QC 미통과 시 산출 보류(작업6 게이트). 강제 저장은 force=True.
"""
from __future__ import annotations

import os
import warnings
from dataclasses import dataclass

import fpna._bootstrap  # noqa: F401

from fpna.templates import get_template
from fpna.templates.base import QCReport


@dataclass
class RenderResult:
    template: str
    out_path: str | None
    qc: QCReport
    saved: bool


def _legacy_render(type_name: str, data, out_path: str, *, mode: str = "create",
                   base_path: str | None = None, force: bool = False) -> RenderResult:
    """구 render 본문 — 패리티 테스트 전용 보존(tests/test_parity.py).

    ⛔ 프로덕션 직접 호출 금지(스파인 우회). run_report 와의 셀 동치 비교 기준선일 뿐.
    """
    mod = get_template(type_name)
    wb = mod.build(data, mode=mode, base_path=base_path)
    rep = mod.qc(wb, data)
    saved = False
    if rep.passed or force:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
        wb.save(out_path)
        saved = True
    return RenderResult(type_name, out_path if saved else None, rep, saved)


def render(type_name: str, data, out_path: str, *, mode: str = "create",
           base_path: str | None = None, force: bool = False) -> RenderResult:
    """[deprecated] run_report 스파인 위임 셸. 심볼만 유지(하위호환).

    독립 save 경로가 사라져 스파인 우회가 물리적으로 불가능하다. 신규 코드는
    fpna.pipeline.run_report 를 직접 호출할 것.
    """
    warnings.warn("render() 는 deprecated — fpna.pipeline.run_report 로 위임됩니다.",
                  DeprecationWarning, stacklevel=2)
    from fpna.pipeline import run_report
    mod = get_template(type_name)
    res = run_report(mod, data, out_path=out_path, mode=mode, base_path=base_path, force=force)
    return RenderResult(type_name, res.out_path, res.qc, res.saved)


def render_golden(type_name: str, out_dir: str = "out/golden") -> RenderResult:
    """유형의 골든샘플을 스파인(run_report)으로 빌드·QC·저장(회귀 테스트용)."""
    from fpna.pipeline import run_report
    mod = get_template(type_name)
    data = mod.golden_sample()
    res = run_report(mod, data, out_path=os.path.join(out_dir, "%s_golden.xlsx" % type_name))
    return RenderResult(type_name, res.out_path, res.qc, res.saved)


__all__ = ["render", "render_golden", "_legacy_render", "RenderResult"]
