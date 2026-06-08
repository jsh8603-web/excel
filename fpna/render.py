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


def render(type_name: str, data, out_path: str, *, mode: str = "create",
           base_path: str | None = None, force: bool = False) -> RenderResult:
    """템플릿 빌드 → QC → (통과 시) 저장."""
    mod = get_template(type_name)
    wb = mod.build(data, mode=mode, base_path=base_path)
    rep = mod.qc(wb, data)
    saved = False
    if rep.passed or force:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
        wb.save(out_path)
        saved = True
    return RenderResult(type_name, out_path if saved else None, rep, saved)


def render_golden(type_name: str, out_dir: str = "out/golden") -> RenderResult:
    """유형의 골든샘플을 빌드·QC·저장(회귀 테스트용)."""
    mod = get_template(type_name)
    data = mod.golden_sample()
    return render(type_name, data, os.path.join(out_dir, "%s_golden.xlsx" % type_name))


__all__ = ["render", "render_golden", "RenderResult"]
