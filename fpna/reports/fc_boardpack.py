"""
fpna.reports.fc_boardpack — 고정비 만기 보드팩 (다중시트 제본, B 실행경로).

`fc_maturity_wall`(단일시트 분석)의 데이터를 회계법인 워크페이퍼 형식으로 제본한다:
  표지(자동) · 요약(만기 총 연환산) · 상세(계약별 연환산 line) · 검증(자동 크로스tie).

설계(자문 3R §B / plan R3):
- 각 builder = (ws, ctx) -> facts dict. ws 에 house_style 로 직접 작성하고,
  크로스시트 tie 용 facts(메모리값)를 반환한다.
- 크로스시트 tie = ConserveSpec 재사용: 상세.total(독립 합) == 요약.total(보고값).
  build_report 가 facts 를 평탄화('상세.total' 등)해서 eval_specs 로 대조한다.
- 데이터 없으면 fc_maturity_wall.golden_sample() 의 contracts(구조 더미)를 재사용한다.
- 연환산 = amount_per_period × (연 발생 횟수). active + 미만료만 포함.
"""
from __future__ import annotations

import fpna._bootstrap  # noqa: F401

from fpna import house_style as hs
from fpna.conserve import ConserveSpec
from fpna.dims import AccountingCalendar
from fpna.report import ReportSpec, SheetSpec

# 연 발생 횟수(연환산 계수) — fc_maturity_wall 과 의미 동일(코드경로는 독립).
_PER_YEAR = {"monthly": 12, "quarterly": 4, "annual": 1, "one_time": 1}


def _active_lines(data):
    """data → [(contract_id, counterparty, account_id, expiry, annualized), ...].

    fc_maturity_wall 의 build 헬퍼(_rows/_annualized)를 부르지 않고 INPUT.contracts 를
    직접 순회한다(독립 경로). active + 미만료 계약만 연환산해 line 으로 만든다.
    """
    cal = AccountingCalendar(fiscal_year_start_month=data.fy_start_month)
    ref = cal.period(*data.report_period).cutoff_date
    lines = []
    for c in data.contracts:
        if c.status != "active":
            continue
        if c.end_date is not None:
            rem = (c.end_date.year - ref.year) * 12 + (c.end_date.month - ref.month)
            if rem < 0:
                continue                       # 이미 만료 제외
        ann = c.amount_per_period * _PER_YEAR.get(c.recurrence, 1)
        expiry = c.end_date.isoformat() if c.end_date else "evergreen"
        lines.append((c.contract_id, c.counterparty, c.account_id, expiry, ann))
    return ref, lines


def _summary_builder(data):
    """요약 시트 builder 팩토리: 만기 총 연환산 1줄. facts={'total': Σ}."""
    def build(ws, ctx):
        ref, lines = _active_lines(data)
        total = sum(x[4] for x in lines)
        r = hs.report_frame(ws, "요약 — 약정 만기 총액", subtitle="active 계약 연환산 합계",
                            unit=data.unit, as_of=ref.isoformat(), last_col=2)
        hs.set_widths(ws, {1: 28, 2: 18})
        hs.set_cell(ws, r, 1, "총 연환산 약정(active)", role="label", align=hs.LEFT)
        hs.set_cell(ws, r, 2, total, role="total", number_format=hs.FMT_INT)
        hs.set_cell(ws, r + 1, 1, "계약 건수", role="label", align=hs.LEFT)
        hs.set_cell(ws, r + 1, 2, len(lines), role="calc", number_format=hs.FMT_INT,
                    align=hs.CENTER)
        hs.report_footer(ws, r + 3, source="상세 시트 합계로 추적(크로스tie)",
                         prepared_by="FP&A", last_col=2)
        return {"total": total, "n": len(lines)}
    return build


def _detail_builder(data):
    """상세 시트 builder 팩토리: 계약별 연환산 line + 합계. facts={'total': Σ}."""
    def build(ws, ctx):
        ref, lines = _active_lines(data)
        r = hs.report_frame(ws, "상세 — 계약별 연환산", subtitle="active + 미만료 약정",
                            unit=data.unit, as_of=ref.isoformat(), last_col=5)
        hs.set_widths(ws, {1: 10, 2: 18, 3: 10, 4: 12, 5: 14})
        for j, h in enumerate(["계약", "거래처", "계정", "만기", "연환산"], start=1):
            hs.set_cell(ws, r, j, h, role="header",
                        align=hs.LEFT if j <= 2 else hs.CENTER)
        r += 1
        total = 0.0
        for cid, cp, acc, expiry, ann in lines:
            hs.set_cell(ws, r, 1, cid, role="label", align=hs.LEFT)
            hs.set_cell(ws, r, 2, cp, role="label", align=hs.LEFT)
            hs.set_cell(ws, r, 3, acc, role="label", align=hs.CENTER)
            hs.set_cell(ws, r, 4, expiry, role="calc", align=hs.CENTER)
            hs.set_cell(ws, r, 5, ann, role="calc", number_format=hs.FMT_INT)
            total += ann
            r += 1
        hs.set_cell(ws, r, 1, "합계(active)", role="total", align=hs.LEFT)
        hs.set_cell(ws, r, 5, total, role="total", number_format=hs.FMT_INT)
        for j in range(1, 6):
            ws.cell(row=r, column=j).border = hs.BORDER_TOP_STRONG
        hs.report_footer(ws, r + 2, source="계약 마스터(약정 등록부)",
                         prepared_by="FP&A", last_col=5)
        return {"total": total, "n": len(lines)}
    return build


def make_spec(data=None) -> ReportSpec:
    """보드팩 ReportSpec 생성. data=None 이면 fc_maturity golden 구조더미 재사용.

    4시트: 표지(자동) + 요약(summary) + 상세(detail) + 검증(자동).
    크로스tie: 요약.total == 상세.total (ConserveSpec, build 호출 없는 독립 평탄 facts 대조).
    """
    if data is None:
        from fpna.templates import fc_maturity_wall
        data = fc_maturity_wall.golden_sample()

    ref, _ = _active_lines(data)
    return ReportSpec(
        title="고정비 만기 보드팩",
        subtitle="약정 만기 도래 — 회계법인 워크페이퍼",
        as_of=ref.isoformat(),
        unit=data.unit,
        source="계약 마스터(약정 등록부)",
        sheets=[
            SheetSpec("요약", _summary_builder(data), section="summary", title="요약"),
            SheetSpec("상세", _detail_builder(data), section="detail",
                      title="상세 명세"),
        ],
        # 크로스시트 tie: 요약 보고값 == 상세 독립 합. source=평탄 facts(build 미호출).
        cross_specs=[ConserveSpec("요약 == 상세 합",
                                  raw_sum_fn=lambda f: f["상세.total"],
                                  reported_key="요약.total")],
    )


__all__ = ["make_spec"]
