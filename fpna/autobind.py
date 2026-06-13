"""
fpna.autobind — tidy rows → 추천 템플릿 + 자동 INPUT 구성 (수동 spec 없이).

binding.assemble 이 사람이 짠 spec({필드:(컬럼,타입)})을 요구하는 반면, autobind 는
infer(컬럼 의미 추론) + recommend(템플릿 선택)로 **실데이터를 spec 없이 템플릿 INPUT 에
자동으로 꽂는다**. measure 가 없으면 listing(정리표)으로 떨어져 "적용 템플릿 없음"이
사라진다.

피벗 시 ★Σ 보존 검증(pivot_conserved): long 원본 measure 합 == 산출 합(왜곡 0).
⛔ 도메인 사전 0. 골든/예시는 무의미 더미.
"""
from __future__ import annotations

import fpna._bootstrap  # noqa: F401

from fpna.infer import infer_columns, summarize
from fpna.dispatcher import recommend_from_roles
from fpna.templates.listing import ListingInput, _num
from fpna.templates.period_trend import TrendInput
from fpna.templates.variance import VarianceInput, LineItem


def _to_listing(rows: list, summ: dict, title: str):
    headers = list(rows[0].keys()) if rows else []
    group_by = summ["dimension"][0] if summ["dimension"] else ""
    return ListingInput(
        title=title or "데이터 정리표", subtitle="자동 매핑(infer)",
        headers=headers, rows=rows, number_cols=list(summ["measure"]),
        group_by=group_by, show_total=bool(summ["measure"]))


def _to_trend(rows: list, summ: dict, title: str):
    """time × measure 피벗 — 같은 기간의 measure 는 합산(Σ 보존)."""
    time_col = summ["time"][0]
    periods = sorted(set(str(r.get(time_col)) for r in rows if r.get(time_col) is not None))
    series: dict = {}
    for m in summ["measure"]:
        series[m] = [sum(_num(r.get(m)) for r in rows if str(r.get(time_col)) == p)
                     for p in periods]
    return TrendInput(title=title or "기간 추이", subtitle="자동 매핑(infer)",
                      periods=periods, series=series)


def _to_variance(rows: list, summ: dict, title: str):
    measures = summ["measure"]
    dim = (summ["dimension"] or summ["id"] or ["항목"])[0]
    items = [LineItem(str(r.get(dim)), plan=_num(r.get(measures[0])),
                      actual=_num(r.get(measures[1]))) for r in rows]
    return VarianceInput(title=title or "예실 비교", subtitle="자동 매핑(infer)", items=items)


_BUILDERS = {"listing": _to_listing, "period_trend": _to_trend, "variance": _to_variance}


def autobind(rows: list, *, template: str | None = None, title: str = ""):
    """rows(list[dict]) → (template_name, INPUT). template 미지정 시 추론 추천.

    measure 0 → listing 으로 안전 착지(never "적용 템플릿 없음").
    """
    if not rows:
        return "listing", ListingInput(title=title or "빈 데이터")
    summ = summarize(infer_columns(rows))
    rec = template or recommend_from_roles(summ).template
    fn = _BUILDERS.get(rec, _to_listing)
    return rec, fn(rows, summ, title)


def build_checked(rows: list, *, template: str | None = None, title: str = "",
                  out_path: str | None = None, force: bool = False):
    """rows → autobind → **run_report 스파인 경유** 빌드+QC+(저장). RunResult 반환.

    ★ mod.build(inp) 직접 호출은 스파인(QC·receipt)을 우회한다 — 그러면 계약 위반·
    결측 은폐가 걸러지지 않고 저장된다. 라이브러리로 실데이터를 빌드할 때는 이 헬퍼를
    써서 검증을 강제한다(main.py render/pack/report 와 동일 경로). out_path 주면 QC
    통과 시에만 저장된다.
    """
    from fpna.pipeline import run_report
    from fpna.templates import get_template
    t, inp = autobind(rows, template=template, title=title)
    return run_report(get_template(t), inp, out_path=out_path, force=force)


def pivot_conserved(rows: list, inp, template: str) -> bool:
    """피벗 Σ 보존: long 원본 measure 합 == 산출 합(왜곡·누락 0). period_trend 전용."""
    if template != "period_trend" or not rows:
        return True
    summ = summarize(infer_columns(rows))
    for m in summ["measure"]:
        long_sum = sum(_num(r.get(m)) for r in rows)
        piv_sum = sum(inp.series.get(m, []))
        if abs(long_sum - piv_sum) > 0.5:
            return False
    return True


__all__ = ["autobind", "build_checked", "pivot_conserved"]
