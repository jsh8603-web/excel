"""
fpna.pipeline — 모든 엑셀 산출의 필수 경로(스파인) + receipt 게이트.

핵심(자문 §2): **검증(View Contract)을 특수상황 끝이 아니라 메인으로 이동.**
run_report 를 거치면 어떤 템플릿이든 공통 검증(수식에러·grain·anomaly 보존)이
**base-owned 로 강제**된다. 템플릿이 자율 qc 에서 빠뜨려도 스파인이 잡는다.

점진 채택(회귀 0): 기존 `fpna.render.render(type,data,out)` 는 유지(하위호환).
run_report 는 신규 통로로, ReportTemplate 또는 기존 build/qc 덕타이핑 모듈을 받는다.
receipt(GatePass)는 **스파인 내부에서만 mint** 되고 `_render_with_receipt` 가 이를
요구 → 스파인을 우회하면 저장이 불가능하다.

ALWAYS(spine): grain → 차원정합 → View Contract 게이트 → QC(불변식+anomaly 보존) → render.
ONLY IF(pointer): ingest/profile/crypto 는 스파인 밖(편입 금지).
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Protocol, runtime_checkable

import fpna._bootstrap  # noqa: F401

from openpyxl.workbook.workbook import Workbook

from fpna.templates.base import QCReport, qc_no_formula_errors
from fpna import view_contract as vc


# --------------------------------------------------------------------------- #
# ReportTemplate Protocol (덕타이핑, 점진)                                     #
#   필수: TYPE + build(data)->Workbook.                                        #
#   구방식 호환: qc(wb,data)->QCReport.                                        #
#   신방식(스파인 강화): invariants()->list / detectors()->list.              #
#   스파인 강화 hook(선택): fact_of(wb) / ledger_of(wb) / surfaced_of(wb).     #
# --------------------------------------------------------------------------- #
@runtime_checkable
class ReportTemplate(Protocol):
    TYPE: str

    def build(self, data) -> Workbook: ...


# --------------------------------------------------------------------------- #
# receipt — 스파인 내부에서만 mint                                            #
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class GatePass:
    """필수 경로를 통과했다는 증명. _render_with_receipt 가 요구한다.

    스파인(run_report) 안에서만 생성된다. 외부에서 만들지 말 것(우회).
    """
    template: str
    base_gate_ok: bool
    anomaly_conserved: bool


@dataclass
class RunResult:
    wb: Workbook
    qc: QCReport
    receipt: GatePass | None
    saved: bool
    out_path: str | None


# --------------------------------------------------------------------------- #
# base-owned 게이트 — 전 템플릿 공통 ("검증 메인 이동")                        #
# --------------------------------------------------------------------------- #
def _meta(wb) -> dict:
    return getattr(wb, "_fpna_meta", {}) or {}


def _fact_of(wb, template):
    """템플릿이 노출한 내부 tidy Fact. hook 우선, 없으면 _fpna_meta['fact']."""
    if hasattr(template, "fact_of"):
        return template.fact_of(wb)
    return _meta(wb).get("fact")


def _ledger_of(wb, template):
    if hasattr(template, "ledger_of"):
        return template.ledger_of(wb)
    return _meta(wb).get("anomaly_ledger")


def _surfaced_of(wb, template):
    if hasattr(template, "surfaced_of"):
        return template.surfaced_of(wb)
    return _meta(wb).get("surfaced_flags")


def _base_owned_gate(rep: QCReport, wb, data, template) -> bool:
    """모든 산출에 공통 적용되는 필수 검증. 템플릿 재량 아님."""
    qc_no_formula_errors(wb, rep)
    # grain: 내부 tidy fact 를 노출한 템플릿은 grain 정합을 강제 검증
    fact = _fact_of(wb, template)
    if fact is not None:
        vc.assert_grain(rep, fact)
    # anomaly 2층 보존: ledger 를 노출하면 |ledger|==surfaced 강제
    led = _ledger_of(wb, template)
    if led is not None:
        vc.assert_anomaly_conserved(rep, led, _surfaced_of(wb, template) or 0)
    return rep.passed


def _template_checks(rep: QCReport, wb, data, template) -> None:
    """템플릿 고유 검증. 신방식(invariants/detectors) 우선, 없으면 구 qc merge."""
    if hasattr(template, "invariants"):
        for inv in template.invariants():
            inv(rep, wb, data)
        getters = getattr(template, "detectors", None)
        if callable(getters):
            for det in getters():
                det(rep, wb, data)
    elif hasattr(template, "qc"):
        sub = template.qc(wb, data)
        for name, ok, detail in sub.checks:
            rep.add(name, ok, detail)


# --------------------------------------------------------------------------- #
# 스파인 + receipt 강제 저장                                                   #
# --------------------------------------------------------------------------- #
def _render_with_receipt(wb, out_path: str, receipt: GatePass | None, force: bool) -> None:
    """receipt 없이는 저장 불가(스파인 우회 차단). force 는 명시적 예외."""
    if receipt is None and not force:
        raise RuntimeError("GatePass 없이 저장 불가 — 필수 경로(run_report)를 우회했습니다.")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    wb.save(out_path)


def run_report(template, data, *, out_path: str | None = None,
               force: bool = False) -> RunResult:
    """유일한 스파인. grain→차원정합→View Contract→QC→render 순서를 소유한다.

    template = ReportTemplate(또는 기존 build/qc 덕타이핑 모듈). receipt 는 내부에서만
    mint 되며, QC 통과(rep.passed) 시에만 발급된다. out_path 주면 receipt 있을 때만 저장.
    """
    type_name = getattr(template, "TYPE", None) or getattr(template, "__name__", "report")

    wb = template.build(data)                       # ③ build (content)
    rep = QCReport(type_name)
    _base_owned_gate(rep, wb, data, template)        # ②④ base 공통 게이트 + anomaly 보존
    _template_checks(rep, wb, data, template)        # ⑤ 템플릿 고유 불변식

    receipt = GatePass(type_name, True, True) if rep.passed else None

    saved = False
    if out_path is not None and (receipt is not None or force):
        _render_with_receipt(wb, out_path, receipt, force)   # ⑥ receipt 요구
        saved = True
    return RunResult(wb, rep, receipt, saved, out_path if saved else None)


__all__ = ["ReportTemplate", "GatePass", "RunResult", "run_report"]
