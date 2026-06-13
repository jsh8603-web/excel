"""
fpna.analyze — 다중시트 워크북 구조 스캔 + 템플릿/팩 추천.

"이 파일 어떻게 처리하지?"의 진입점. 워크북의 각 시트를 첫 N행만 읽어 infer 로
컬럼 의미를 추론하고, recommend 로 시트별 착지 템플릿을 정한다. 시트가 여러 개면
연동(pack) 후보를 안내한다. ⛔ 도메인 사전 0 — 시트 역할은 *구조*(시계열/항목·금액/
식별자 목록)로만 태그한다(손익/BS 같은 도메인 라벨 X).
"""
from __future__ import annotations

import fpna._bootstrap  # noqa: F401

import openpyxl

from fpna.dispatcher import recommend_from_roles
from fpna.infer import infer_columns, summarize


def _sheet_to_rows(ws, sample: int) -> tuple:
    """워크시트 → (headers, list[dict]). 첫 비어있지 않은 행=헤더, 이후 데이터."""
    headers = None
    out: list = []
    for row in ws.iter_rows(values_only=True):
        if headers is None:
            if any(c is not None for c in row):
                headers = [str(c) if c is not None else "col%d" % j
                           for j, c in enumerate(row)]
            continue
        if len(out) >= sample:
            break
        if all(c is None for c in row):
            continue
        out.append({headers[j]: row[j] for j in range(min(len(headers), len(row)))})
    return headers, out


def _shape_tag(summ: dict) -> str:
    """구조 태그(도메인 무관)."""
    if summ["time"] and summ["has_measure"]:
        return "시계열형"
    if summ["has_measure"] and (summ["dimension"] or summ["id"]):
        return "항목·금액형"
    if not summ["has_measure"]:
        return "식별자 목록형"
    return "수치형"


def analyze_workbook(path: str, *, sample: int = 50) -> dict:
    """워크북 경로 → {sheets: [...], recommendation: {...}}. 시트별 추론 + 워크북 추천."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheets: list = []
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            headers, rows = _sheet_to_rows(ws, sample)
            if not rows:
                sheets.append({"sheet": name, "shape": "빈/비정형",
                               "template": None, "summary": None})
                continue
            summ = summarize(infer_columns(rows))
            rec = recommend_from_roles(summ)
            sheets.append({"sheet": name, "shape": _shape_tag(summ),
                           "template": rec.template, "reason": rec.reason,
                           "summary": summ})
    finally:
        wb.close()
    return {"sheets": sheets, "recommendation": _recommend(sheets)}


def _recommend(sheets: list) -> dict:
    """시트 조합 → 워크북 수준 추천(단일/다중연동/없음)."""
    data_sheets = [s for s in sheets if s["template"]]
    if not data_sheets:
        return {"kind": "none", "detail": "데이터 시트 없음(빈/비정형) — ingest 로 정형화 먼저"}
    if len(data_sheets) == 1:
        s = data_sheets[0]
        return {"kind": "single", "template": s["template"],
                "detail": "단일 데이터 시트 '%s' → %s (py main.py render %s --csv ...)"
                % (s["sheet"], s["template"], s["template"])}
    listed = ", ".join("%s=%s" % (s["sheet"], s["template"]) for s in data_sheets)
    return {"kind": "multi",
            "detail": "%d개 데이터 시트 → 시트들이 *연동*(공유 가정·크로스시트 tie)이면 "
            "pack(packs.md 카탈로그), 독립이면 시트별 단일. 시트별 추천: %s"
            % (len(data_sheets), listed)}


__all__ = ["analyze_workbook", "_sheet_to_rows", "_shape_tag", "_recommend"]
