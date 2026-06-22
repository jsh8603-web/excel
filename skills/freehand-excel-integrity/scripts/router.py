#!/usr/bin/env python3
"""
router.py — 신호 기반 capability+backend 라우팅 (퍼센트 사전확률 제거).

설계 원칙
  · CAPABILITY(전략)와 BACKEND(툴)는 직교. 둘 다 *신호*의 결정함수(확률 아님).
  · BACKEND 는 3개로 통합: openpyxl(헤드리스 편집/쓰기) · xlsxwriter(신규 대량/차트) ·
    xlwings(라이브 Excel; 피벗/슬라이서/외부링크 등 Excel 고유기능은 .api(COM)로 강하).
    → pywin32 COM 은 독립 백엔드가 아니라 xlwings 의 탈출구.
  · PREFLIGHT: 선택 백엔드의 런타임이 실제 있나 확인. 없으면 폴백 + **DOWNGRADE 명시**(침묵 금지).
  · 정적 값 vs 수식: 재계산 엔진/게이트 단계가 갈린다.

decide(task) → 결정 dict. task 신호:
  op: create|edit|repair      input_quality: tidy|messy
  has_template: bool          features: {pivot,slicer,external_link,power_query,data_model,vba,chart,table}
  value_mode: static|formula  size: small|large
"""
from __future__ import annotations

import shutil

EXCEL_ONLY_FEATURES = {"pivot", "slicer", "external_link", "power_query",
                       "data_model", "vba", "connection"}


def detect_env() -> dict:
    env = {"openpyxl": False, "xlsxwriter": False, "pywin32": False,
           "excel": False, "libreoffice": False, "formulas": False}
    for mod in ("openpyxl", "xlsxwriter", "formulas"):
        try:
            __import__(mod); env[mod] = True
        except Exception:
            pass
    try:
        import win32com  # noqa
        env["pywin32"] = True; env["excel"] = True  # COM 가능 ≈ Excel 설치
    except Exception:
        pass
    if shutil.which("soffice") or shutil.which("libreoffice"):
        env["libreoffice"] = True
    return env


def _recalc_engine(env: dict):
    if env.get("pywin32"):
        return "pywin32"
    if env.get("libreoffice"):
        return "libreoffice"
    if env.get("formulas"):
        return "formulas"
    return None


def decide(task: dict, env: dict | None = None) -> dict:
    env = env or detect_env()
    op = task.get("op", "create")
    features = set(task.get("features", []))
    value_mode = task.get("value_mode", "static")
    size = task.get("size", "small")
    downgrades, reasons = [], []

    # ── ① CAPABILITY ─────────────────────────────────────────────
    if task.get("input_quality") == "messy":
        capability = "ingest→tidy→(template|freehand)"
        reasons.append("입력 누더기 → 정제 먼저")
    elif task.get("has_template"):
        capability = "template+qc"
        reasons.append("표준 양식 有")
    else:
        capability = "freehand+contract"
        reasons.append("양식 없음 → 프리핸드+계약")

    # ── ② BACKEND (3개, COM=xlwings 탈출구) ──────────────────────
    excel_feature = bool(features & EXCEL_ONLY_FEATURES)
    live_edit = (op in ("edit", "repair"))

    if excel_feature or (op == "edit" and task.get("live_open")):
        backend = "xlwings"
        reasons.append("Excel 고유기능/라이브 편집 → xlwings(+.api COM)")
        if not env.get("excel"):
            # 폴백 + 다운그레이드 명시
            if excel_feature:
                lost = sorted(features & EXCEL_ONLY_FEATURES)
                downgrades.append("Excel 부재 → %s 불가. openpyxl 로 폴백(해당 기능은 정적 대체/생략)."
                                  % ", ".join(lost))
            backend = "openpyxl"
            reasons.append("DOWNGRADE: Excel 없음 → openpyxl")
    elif op == "repair" or op == "edit":
        backend = "openpyxl"  # 기존 파일 편집/수리는 읽기-편집 가능한 openpyxl
        reasons.append("기존 파일 편집/수리 → openpyxl(비파괴 스타일/값)")
    elif op == "create" and size == "large" and "chart" in features:
        backend = "xlsxwriter"
        reasons.append("신규 대량+차트 → xlsxwriter")
        if op == "repair":  # 방어: xlsxwriter 는 편집 불가
            backend = "openpyxl"
    else:
        backend = "openpyxl"
        reasons.append("신규 헤드리스 → openpyxl")

    # 방어: xlsxwriter 로 편집/수리 라우팅 금지
    if backend == "xlsxwriter" and op in ("edit", "repair"):
        downgrades.append("xlsxwriter 는 기존파일 편집 불가 → openpyxl 로 교정")
        backend = "openpyxl"

    # PREFLIGHT: 선택 백엔드 런타임 존재?
    need = {"openpyxl": "openpyxl", "xlsxwriter": "xlsxwriter", "xlwings": "excel"}[backend]
    if not env.get(need):
        if backend == "xlsxwriter" and env.get("openpyxl"):
            downgrades.append("xlsxwriter 미설치 → openpyxl 폴백"); backend = "openpyxl"
        elif backend == "openpyxl" and env.get("xlsxwriter") and op == "create":
            downgrades.append("openpyxl 미설치 → xlsxwriter 폴백(신규)"); backend = "xlsxwriter"
        else:
            downgrades.append("PREFLIGHT FAIL: %s 런타임 없음 — 진행 불가" % backend)

    # ── ③ 재계산/게이트 ──────────────────────────────────────────
    recalc = _recalc_engine(env) if value_mode == "formula" else None
    if value_mode == "formula" and recalc is None:
        downgrades.append("수식인데 재계산 엔진 없음 → 정적 값+contract.expected 권장")

    gate = ["contract-coverage", "xlsx_doctor"]
    if op in ("edit", "repair"):
        gate.append("roundtrip-gate(before-snapshot)")
    if value_mode == "formula":
        gate.append("recalc(%s)" % (recalc or "none"))
    if env.get("pywin32"):
        gate.append("verify_xlsx(--recalc)")

    return {
        "capability": capability, "backend": backend, "recalc_engine": recalc,
        "gate": gate, "downgrades": downgrades, "reasons": reasons,
        "value_mode": value_mode,
    }


def explain(task: dict, env: dict | None = None) -> str:
    d = decide(task, env)
    lines = ["CAPABILITY: %s" % d["capability"],
             "BACKEND   : %s" % d["backend"],
             "RECALC    : %s" % d["recalc_engine"],
             "GATE      : %s" % " → ".join(d["gate"])]
    if d["downgrades"]:
        lines.append("DOWNGRADE : " + " ; ".join(d["downgrades"]))
    lines.append("WHY       : " + " | ".join(d["reasons"]))
    return "\n".join(lines)


if __name__ == "__main__":
    # 데모 케이스
    cases = [
        ("신규 대량 차트(헤드리스)", {"op": "create", "size": "large", "features": ["chart"], "value_mode": "static"}),
        ("피벗 요청(Excel 없음)", {"op": "create", "features": ["pivot"], "value_mode": "static"}),
        ("기존 파일 수리(수식)", {"op": "repair", "value_mode": "formula"}),
        ("외부링크 워크북 편집", {"op": "edit", "features": ["external_link"], "live_open": True}),
    ]
    env = detect_env()
    print("ENV:", {k: v for k, v in env.items() if v})
    for name, t in cases:
        print("\n■", name)
        print(explain(t, env))
