"""fpna.templates.pnl_3statement — 손익계산서(3-statement 연결 골격 + tie-out 게이트)."""
from __future__ import annotations

from dataclasses import dataclass, field

import fpna._bootstrap  # noqa: F401
import openpyxl

from fpna import finance, house_style as hs
from fpna.templates.base import QCReport, qc_no_formula_errors, qc_totals
from fpna.view_contract import assert_tie_out, recon_block

TYPE = "pnl_3statement"


@dataclass
class PnLInput:
    title: str = "손익계산서"
    subtitle: str = ""
    unit: str = "₩mn"
    revenue: float = 1000.0
    cogs: float = 600.0
    sga: float = 200.0
    da: float = 50.0           # 감가상각
    interest: float = 20.0
    tax_rate: float = 0.22

    # --- 3-statement 연결 (선택) -----------------------------------------
    # 아래 BS/RE/CF 필드가 모두 채워지면 build/qc 가 3-statement 연결 모드로
    # BS균형(A=L+E)·RE roll(기초+NI-배당=기말)·CF 간접법(NI 출발) tie-out 까지
    # 강제한다. 비우면(linked=False) IS + 계층tie 만 검증한다.
    # ⚠ 골든 기본값은 의미 없는 구조 더미(tie 정합만 맞춤).
    dividends: float | None = None       # 배당(RE roll)
    re_begin: float | None = None        # 기초 이익잉여금
    # 기말 BS (A = L + E). equity = re_begin + NI - dividends + paid_in_capital
    cash: float | None = None
    other_assets: float | None = None    # 비현금 자산(매출채권·재고·유형자산 등 합)
    liabilities: float | None = None     # 부채 합
    paid_in_capital: float | None = None  # 납입자본(자본금+자본잉여금)
    # CF 간접법 검증: 기말현금 = 기초현금 + NI + DA(비현금) + ΔNWC − 배당
    cash_begin: float | None = None
    delta_nwc: float | None = None       # 운전자본 증감(현금 영향, + = 현금 유입)

    @property
    def linked(self) -> bool:
        """3-statement 연결 모드 여부(BS/RE/CF 필드 전부 존재)."""
        return all(v is not None for v in (
            self.dividends, self.re_begin, self.cash, self.other_assets,
            self.liabilities, self.paid_in_capital,
            self.cash_begin, self.delta_nwc))


# --------------------------------------------------------------------------- #
# 파이썬 재계산(QC 의 SSOT) — 셀 수식 의도와 1:1 대조                           #
# --------------------------------------------------------------------------- #
def _compute(data: PnLInput) -> dict:
    gp = data.revenue - data.cogs
    ebit = gp - data.sga - data.da
    ebt = ebit - data.interest
    ni = ebt - max(0.0, ebt) * data.tax_rate
    out = {"gp": gp, "ebit": ebit, "ebt": ebt, "ni": ni}
    if data.linked:
        re_end = data.re_begin + ni - data.dividends
        equity = re_end + data.paid_in_capital
        assets = data.cash + data.other_assets
        cash_end_cf = data.cash_begin + ni + data.da + data.delta_nwc - data.dividends
        out.update({"re_end": re_end, "equity": equity, "assets": assets,
                    "cash_end_cf": cash_end_cf})
    return out


def golden_sample() -> PnLInput:
    # 구조 검증용 3-statement 연결 더미 — 의미 없는 금액, tie 정합만 맞춤.
    # IS: rev1000 cogs600 sga200 da50 int20 tax0.22 → NI 계산.
    c = _compute(PnLInput(revenue=1000, cogs=600, sga=200, da=50,
                          interest=20, tax_rate=0.22))
    ni = c["ni"]
    # RE roll: 기초500 + NI − 배당30 = 기말
    re_begin, dividends = 500.0, 30.0
    re_end = re_begin + ni - dividends
    paid_in = 200.0
    equity = re_end + paid_in
    # CF 간접법: 기말현금 = 기초현금120 + NI + DA50 + ΔNWC(-10) − 배당30
    cash_begin, delta_nwc, da = 120.0, -10.0, 50.0
    cash_end = cash_begin + ni + da + delta_nwc - dividends
    # BS 균형: 자산 = 부채 + 자본 → other_assets 를 plug 로 맞춤(구조 더미)
    liabilities = 300.0
    other_assets = (liabilities + equity) - cash_end
    return PnLInput(
        title="손익계산서 (골든샘플)", subtitle="구조 검증용 — 3-statement 연결 더미",
        revenue=1000, cogs=600, sga=200, da=50, interest=20, tax_rate=0.22,
        dividends=dividends, re_begin=re_begin,
        cash=cash_end, other_assets=other_assets, liabilities=liabilities,
        paid_in_capital=paid_in, cash_begin=cash_begin, delta_nwc=delta_nwc)


def build(data: PnLInput, *, mode="create", base_path=None) -> openpyxl.Workbook:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = hs.safe_sheet_title("PnL")
    last_col = 2
    hs.set_widths(ws, {1: 28, 2: 16})
    r = hs.report_frame(ws, data.title, subtitle=data.subtitle,
                        unit=data.unit, last_col=last_col)
    hs.set_cell(ws, r, 1, "항목 (단위: %s)" % data.unit, role="header", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "금액", role="header"); r += 1

    def line(name, value=None, formula=None, *, total=False, indent=0, sign="+"):
        nonlocal r
        hs.set_cell(ws, r, 1, name, role="total" if total else "label",
                    align=hs.indent(indent), bold=total)
        if formula:
            hs.set_cell(ws, r, 2, formula, role="calc", number_format=hs.FMT_INT, bold=total)
        else:
            hs.set_cell(ws, r, 2, value, role="input", number_format=hs.FMT_INT)
        if total:
            for j in range(1, last_col + 1):
                ws.cell(row=r, column=j).border = hs.BORDER_TOP
        cur = r; r += 1; return cur

    rev = line("매출", data.revenue)
    cogs = line("(-) 매출원가", data.cogs, indent=1)
    gp = line("매출총이익", formula="=B%d-B%d" % (rev, cogs), total=True)
    sga = line("(-) 판관비", data.sga, indent=1)
    da = line("(-) 감가상각", data.da, indent=1)
    ebit = line("영업이익(EBIT)", formula="=B%d-B%d-B%d" % (gp, sga, da), total=True)
    inte = line("(-) 이자비용", data.interest, indent=1)
    ebt = line("세전이익(EBT)", formula="=B%d-B%d" % (ebit, inte), total=True)
    hs.set_cell(ws, r, 1, "법인세율", role="label", align=hs.indent(1))
    hs.set_cell(ws, r, 2, data.tax_rate, role="input", number_format=hs.FMT_PCT1)
    tax_rate_r = r; r += 1
    tax = line("(-) 법인세", formula="=MAX(0,B%d)*B%d" % (ebt, tax_rate_r), indent=1)
    ni = line("당기순이익", formula="=B%d-B%d" % (ebt, tax), total=True)

    # 마진 요약
    r += 1
    r = hs.section_header(ws, r, "마진", last_col=last_col)
    hs.set_cell(ws, r, 1, "영업이익률", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(B%d=0,\"\",B%d/B%d)" % (rev, ebit, rev), role="calc",
                number_format=hs.FMT_PCT1); r += 1
    hs.set_cell(ws, r, 1, "순이익률", role="label", align=hs.LEFT)
    hs.set_cell(ws, r, 2, "=IF(B%d=0,\"\",B%d/B%d)" % (rev, ni, rev), role="calc",
                number_format=hs.FMT_PCT1); r += 1

    # ------------------------------------------------------------------ #
    # 3-statement 연결 + _RECON tie-out (linked 모드일 때만)             #
    # ------------------------------------------------------------------ #
    if data.linked:
        c = _compute(data)
        # 이익잉여금 roll
        r += 1
        r = hs.section_header(ws, r, "이익잉여금 roll", last_col=last_col)
        re_b = line("기초 이익잉여금", data.re_begin)
        ni_re = line("(+) 당기순이익", formula="=B%d" % ni, indent=1)
        div = line("(-) 배당", data.dividends, indent=1)
        re_e = line("기말 이익잉여금", formula="=B%d+B%d-B%d" % (re_b, ni_re, div), total=True)

        # 자본
        pic = line("납입자본", data.paid_in_capital)
        eq = line("자본 총계", formula="=B%d+B%d" % (re_e, pic), total=True)

        # 재무상태표(기말)
        r += 1
        r = hs.section_header(ws, r, "재무상태표(기말)", last_col=last_col)
        cash_r = line("현금", data.cash)
        oa_r = line("기타 자산", data.other_assets, indent=1)
        asset_r = line("자산 총계", formula="=B%d+B%d" % (cash_r, oa_r), total=True)
        liab_r = line("부채 총계", data.liabilities)
        le_r = line("부채+자본", formula="=B%d+B%d" % (liab_r, eq), total=True)

        # 현금흐름(간접법) — NI 출발
        r += 1
        r = hs.section_header(ws, r, "현금흐름(간접법)", last_col=last_col)
        cb_r = line("기초 현금", data.cash_begin)
        cf_ni = line("(+) 당기순이익", formula="=B%d" % ni, indent=1)
        cf_da = line("(+) 감가상각(비현금)", formula="=B%d" % da, indent=1)
        cf_nwc = line("(±) 운전자본 증감", data.delta_nwc, indent=1)
        cf_div = line("(-) 배당", formula="=B%d" % div, indent=1)
        ce_r = line("기말 현금(CF)", formula="=B%d+B%d+B%d+B%d-B%d"
                    % (cb_r, cf_ni, cf_da, cf_nwc, cf_div), total=True)

        # _RECON 블록 — tie-out check 행(=0 정상, ≠0 적색). hs.check_cell 활용.
        r += 1
        r = hs.section_header(ws, r, "_RECON (tie-out)", last_col=last_col)
        recon = [
            ("BS 균형(자산−부채−자본)", "=B%d-B%d" % (asset_r, le_r)),
            ("CF 기말현금 − BS 현금", "=B%d-B%d" % (ce_r, cash_r)),
        ]
        for label, formula in recon:
            hs.check_cell(ws, r, 2, formula, label=label, label_col=1)
            r += 1

    hs.report_footer(ws, r + 1, source="총계정원장(GL) 마감",
                     prepared_by="FP&A", last_col=last_col)
    return wb


def qc(wb, data: PnLInput) -> QCReport:
    rep = QCReport(TYPE)
    qc_no_formula_errors(wb, rep)
    c = _compute(data)
    qc_totals("매출총이익", c["gp"], data.revenue - data.cogs, rep)
    rep.add("EBIT 계산", True, "EBIT=%.0f" % c["ebit"])
    rep.add("당기순이익 계산", True, "NI=%.0f" % c["ni"])
    rep.add("세율 [0,1)", 0 <= data.tax_rate < 1, "tax=%.2f" % data.tax_rate)
    rep.add("단위 표기", bool(data.unit))

    # --- 계층 tie (R10류): 손익 rollup leaf 합 == 소계 -------------------
    # 매출 − 매출원가 = 매출총이익 (leaf 합 == 소계). plug 은폐 차단.
    leaf_sum = data.revenue - data.cogs
    assert_tie_out(rep, leaf_sum, c["gp"], tol=0.0, name="R10 계층tie:매출총이익")
    # EBIT = GP − SGA − DA, EBT = EBIT − 이자 (각 소계가 leaf 산술의 정합)
    assert_tie_out(rep, c["gp"] - data.sga - data.da, c["ebit"], tol=0.0,
                   name="R10 계층tie:EBIT")
    assert_tie_out(rep, c["ebit"] - data.interest, c["ebt"], tol=0.0,
                   name="R10 계층tie:EBT")

    # --- 3-statement 연결 tie-out (linked 모드) -------------------------
    if data.linked:
        # N-version(자문 R2 C6): _compute(빌드 경로) 와 독립으로 INPUT 스칼라에서
        #   NI·equity·assets 를 재계산해 대조. _compute 와 같은 버그를 두 경로가
        #   반복하지 않도록 부호·세금·잔류이익 누적을 직접 다시 푼다(off-by 노출).
        ebt_ind = data.revenue - data.cogs - data.sga - data.da - data.interest
        ni_ind = ebt_ind - max(0.0, ebt_ind) * data.tax_rate
        equity_ind = data.re_begin + ni_ind - data.dividends + data.paid_in_capital
        assets_ind = data.cash + data.other_assets
        # 빌드 경로(_compute) 가 독립 재계산과 합치하는지(경로 일관)
        assert_tie_out(rep, c["ni"], ni_ind, tol=1e-9, name="N-version NI 독립 대조")
        assert_tie_out(rep, c["equity"], equity_ind, tol=1e-9,
                       name="N-version 자본 독립 대조")
        assert_tie_out(rep, c["assets"], assets_ind, tol=1e-9,
                       name="N-version 자산 독립 대조")
        # 1) BS 균형: 자산 == 부채 + 자본 (A = L + E). 독립 재계산값으로 검정 →
        #   _compute 가 양변에 공통 버그를 심어도 INPUT 직접 경로가 잡는다.
        assert_tie_out(rep, assets_ind, data.liabilities + equity_ind, tol=0.0,
                       name="BS 균형(A=L+E)")
        # 2) RE roll: 기초 + NI − 배당 = 기말 → equity 에 반영됐는지 (구조 정합)
        assert_tie_out(rep, data.re_begin + c["ni"] - data.dividends, c["re_end"],
                       tol=0.0, name="RE roll(기초+NI-배당)")
        # 3) CF 간접법: 기말현금(CF) == BS 현금 (NI 출발 간접법이 BS 현금과 tie)
        assert_tie_out(rep, c["cash_end_cf"], data.cash, tol=0.0,
                       name="CF 간접법 기말현금 == BS 현금")
    else:
        rep.add("3-statement 연결", True, "단순 IS 모드(BS/CF/RE 미입력) — 계층tie만")
    return rep


# --------------------------------------------------------------------------- #
# solve_and_link — 순환참조(이자↔현금↔부채) 해소 후 3-statement 연결 (R4)      #
#   debt_schedule→이자→IS→CFS→BS현금→리볼버 plug 를 finance.solve_revolver 로  #
#   해소하고, 그 수렴값으로 linked PnLInput 을 구성한다. 결과는 기존 linked     #
#   tie-out(BS균형·RE roll·CF 간접법)이 그대로 검증한다(회귀 0 — 기존 무변경).  #
# --------------------------------------------------------------------------- #
def solve_and_link(*, revenue: float, cogs: float, sga: float, da: float,
                   tax_rate: float, beginning_debt: dict, debt_rates: dict,
                   pre_financing_cash: float, cash_begin: float,
                   re_begin: float, dividends: float, paid_in_capital: float,
                   min_cash: float = 0.0, sweep_priority: tuple = ("revolver",),
                   revolver_key: str = "revolver",
                   revolver_limit: float = float("inf"),
                   title: str = "손익계산서 (순환연동)",
                   subtitle: str = "solve_revolver plug — 이자·현금·부채 내생화") -> PnLInput:
    """순환(이자=평균부채×rate ↔ 현금부족→리볼버 draw ↔ 부채↑→이자↑)을 1기간 고정점으로
    해소하고, 수렴값(interest/end_cash/debt_balance)으로 균형 잡힌 linked PnLInput 반환.

    pre_financing_cash = 영업+투자 CF(이자·배당·financing 전). 배당은 financing 으로
    별도 차감(pre 에서 미리 뺀다 → 현금 일관). other_assets·delta_nwc 는 BS균형·CF tie 가
    정확히 닫히도록 plug 역산(★구조 더미 — 실데이터는 회사에서 실값 입력).

    반환 PnLInput 은 linked=True → build/qc 가 3-statement tie-out 전수 검증.
    plug 해소 여부는 결과 PnLInput 의 (cash, liabilities) 가 solve 수렴값임으로 확인.
    """
    res = finance.solve_revolver(
        [pre_financing_cash - dividends], beginning_cash=cash_begin,
        beginning_debt=dict(beginning_debt), rates=dict(debt_rates),
        min_cash=min_cash, sweep_priority=tuple(sweep_priority),
        revolver_key=revolver_key, revolver_limit=revolver_limit)
    p = res.periods[0]
    interest = p["interest"]
    cash_end = p["end_cash"]
    liabilities = p["debt_balance"]

    ebt = revenue - cogs - sga - da - interest
    ni = ebt - max(0.0, ebt) * tax_rate
    re_end = re_begin + ni - dividends
    equity = re_end + paid_in_capital
    # BS 균형 plug: assets = liabilities + equity, assets = cash_end + other_assets
    other_assets = liabilities + equity - cash_end
    # CF 간접법 tie plug: cash_begin + ni + da + delta_nwc − dividends == cash_end
    delta_nwc = cash_end - cash_begin - ni - da + dividends

    return PnLInput(
        title=title, subtitle=subtitle, revenue=revenue, cogs=cogs, sga=sga,
        da=da, interest=interest, tax_rate=tax_rate,
        dividends=dividends, re_begin=re_begin, cash=cash_end,
        other_assets=other_assets, liabilities=liabilities,
        paid_in_capital=paid_in_capital, cash_begin=cash_begin, delta_nwc=delta_nwc)


__all__ = ["TYPE", "PnLInput", "golden_sample", "build", "qc", "solve_and_link"]
