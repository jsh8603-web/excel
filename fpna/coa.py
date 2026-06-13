"""
fpna.coa — 표준 계정과목 taxonomy (Chart of Accounts).

R5·R6 배선: SEC US-GAAP XBRL Taxonomy(공개도메인 element 명칭) + IFRS 명칭 참조로
표준 IS/BS/CFS 라인 골격을 정의한다. dims.Account 차원을 이 taxonomy 의 fs_line 으로
정규화해, pnl_3statement / pack 의 라벨·부호·재무제표 귀속을 일관화한다.

⛔ 합성 재무수치 0 — taxonomy 는 *금액*이 아니라 *분류*다(PIT/합성 규율 무관).
정적 구조는 refdata/coa_us_gaap.json([V] 내 파생물, 공개도메인 명칭). 로드 실패 시
코드 내장 골격(golden_coa)으로 폴백 — 무설치 회사 PC 에서도 동작.

용법:
  lines = load_coa()                      # JSON → list[FsLine] (실패 시 내장 폴백)
  idx = coa_index(lines)                  # code → FsLine
  fl = account_to_fs_line(acc, by_code)   # dims.Account → fs_line 문자열
  is_lines = statement_lines(lines, "IS") # 재무제표별 라인
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass

import fpna._bootstrap  # noqa: F401

STATEMENTS = ("IS", "BS", "CFS")
SIGNS = ("+", "-")

_REFDATA = os.path.join(os.path.dirname(__file__), "refdata", "coa_us_gaap.json")


@dataclass(frozen=True)
class FsLine:
    """표준 재무제표 라인 1개.

    code        내부 고유 코드(예: "IS_REV").
    label_ko/en 표시 라벨.
    statement   재무제표 귀속 — IS / BS / CFS.
    fs_line     dims.Account.fs_line 와 매칭되는 정규 키(예: "revenue").
    sign        부호 규약 — '+'(가산) / '-'(차감). 표시·집계 방향.
    parent      상위 라인 code(소계 roll-up). None = 최상위.
    us_gaap_tag SEC US-GAAP element 명칭(공개도메인 — tag 로 사용 가능).
    ifrs_tag    IFRS Taxonomy 명칭(참조만 — 벌크복사 아님).
    """
    code: str
    label_ko: str
    label_en: str
    statement: str
    fs_line: str
    sign: str
    parent: str | None = None
    us_gaap_tag: str = ""
    ifrs_tag: str = ""


# --------------------------------------------------------------------------- #
# 로드                                                                         #
# --------------------------------------------------------------------------- #
def load_coa(path: str | None = None) -> list[FsLine]:
    """refdata/coa_us_gaap.json → list[FsLine]. 파일 부재/오류 시 내장 골격 폴백."""
    p = path or _REFDATA
    try:
        with open(p, encoding="utf-8") as fh:
            raw = json.load(fh)
        out = [_from_dict(d) for d in raw.get("lines", [])]
        return out if out else golden_coa()
    except (OSError, ValueError, KeyError):
        return golden_coa()


def _from_dict(d: dict) -> FsLine:
    return FsLine(
        code=d["code"], label_ko=d["label_ko"], label_en=d["label_en"],
        statement=d["statement"], fs_line=d["fs_line"], sign=d.get("sign", "+"),
        parent=d.get("parent"), us_gaap_tag=d.get("us_gaap_tag", ""),
        ifrs_tag=d.get("ifrs_tag", ""))


def coa_index(lines: list[FsLine]) -> dict[str, FsLine]:
    """code → FsLine. 중복 code 는 마지막이 승리(결정성)."""
    return {ln.code: ln for ln in lines}


def by_fs_line(lines: list[FsLine]) -> dict[str, FsLine]:
    """fs_line → FsLine (statement 내 유일 가정)."""
    return {ln.fs_line: ln for ln in lines}


def statement_lines(lines: list[FsLine], statement: str) -> list[FsLine]:
    """재무제표(IS/BS/CFS)별 라인."""
    return [ln for ln in lines if ln.statement == statement]


# --------------------------------------------------------------------------- #
# 매핑: dims.Account → 표준 fs_line                                           #
# --------------------------------------------------------------------------- #
def account_to_fs_line(account, by_code: dict[str, FsLine] | None = None,
                       fallback_map: dict[str, str] | None = None) -> str | None:
    """dims.Account 를 표준 fs_line 으로 정규화.

    우선순위: (1) account.fs_line 이 이미 표준 fs_line 이면 그대로
              (2) account.code 가 fallback_map 에 있으면 매핑값
              (3) account.code 가 coa code 면 그 FsLine.fs_line
    못 찾으면 None(은폐 금지 — 호출자가 미매핑 표면화).
    """
    fs = getattr(account, "fs_line", None)
    if fs:
        return fs
    code = getattr(account, "code", None)
    if fallback_map and code in fallback_map:
        return fallback_map[code]
    if by_code and code in by_code:
        return by_code[code].fs_line
    return None


def validate_coa(lines: list[FsLine]) -> list[str]:
    """taxonomy 무결성 검사 → 문제 메시지 리스트(빈 = 정상).

    (a) statement 유효 (b) sign 유효 (c) parent 존재(있으면) (d) code 유일.
    """
    problems: list[str] = []
    codes = {ln.code for ln in lines}
    seen: set[str] = set()
    for ln in lines:
        if ln.statement not in STATEMENTS:
            problems.append("%s: 잘못된 statement '%s'" % (ln.code, ln.statement))
        if ln.sign not in SIGNS:
            problems.append("%s: 잘못된 sign '%s'" % (ln.code, ln.sign))
        if ln.parent is not None and ln.parent not in codes:
            problems.append("%s: parent '%s' 부재" % (ln.code, ln.parent))
        if ln.code in seen:
            problems.append("%s: code 중복" % ln.code)
        seen.add(ln.code)
    return problems


# --------------------------------------------------------------------------- #
# golden — 내장 표준 골격 1세트 (JSON 폴백 + 테스트 기준)                       #
# --------------------------------------------------------------------------- #
def golden_coa() -> list[FsLine]:
    """IS/BS/CFS 표준 골격 1세트. JSON 부재 시 폴백. 재무수치 0(분류만)."""
    return [
        # IS
        FsLine("IS_REV", "매출액", "Revenue", "IS", "revenue", "+", None, "Revenues", "Revenue"),
        FsLine("IS_COGS", "매출원가", "Cost of Revenue", "IS", "cogs", "-", None, "CostOfRevenue", "CostOfSales"),
        FsLine("IS_GP", "매출총이익", "Gross Profit", "IS", "gross_profit", "+", None, "GrossProfit", "GrossProfit"),
        FsLine("IS_SGA", "판매관리비", "SG&A Expense", "IS", "sga", "-", "IS_OPEX", "SellingGeneralAndAdministrativeExpense", "SellingGeneralAndAdministrativeExpense"),
        FsLine("IS_DA", "감가상각비", "Depreciation & Amortization", "IS", "da", "-", "IS_OPEX", "DepreciationDepletionAndAmortization", "DepreciationAndAmortisationExpense"),
        FsLine("IS_OPEX", "영업비용", "Operating Expenses", "IS", "opex", "-", None, "OperatingExpenses", "OperatingExpense"),
        FsLine("IS_OPINC", "영업이익", "Operating Income", "IS", "operating_income", "+", None, "OperatingIncomeLoss", "ProfitLossFromOperatingActivities"),
        FsLine("IS_INT", "이자비용", "Interest Expense", "IS", "interest", "-", None, "InterestExpense", "FinanceCosts"),
        FsLine("IS_EBT", "세전이익", "Income Before Tax", "IS", "ebt", "+", None, "IncomeLossFromContinuingOperationsBeforeIncomeTaxes", "ProfitLossBeforeTax"),
        FsLine("IS_TAX", "법인세비용", "Income Tax Expense", "IS", "tax", "-", None, "IncomeTaxExpenseBenefit", "IncomeTaxExpenseContinuingOperations"),
        FsLine("IS_NI", "당기순이익", "Net Income", "IS", "net_income", "+", None, "NetIncomeLoss", "ProfitLoss"),
        # BS
        FsLine("BS_CASH", "현금및현금성자산", "Cash & Equivalents", "BS", "cash", "+", "BS_CA", "CashAndCashEquivalentsAtCarryingValue", "CashAndCashEquivalents"),
        FsLine("BS_AR", "매출채권", "Accounts Receivable", "BS", "receivables", "+", "BS_CA", "AccountsReceivableNetCurrent", "TradeAndOtherCurrentReceivables"),
        FsLine("BS_INV", "재고자산", "Inventory", "BS", "inventory", "+", "BS_CA", "InventoryNet", "Inventories"),
        FsLine("BS_CA", "유동자산", "Current Assets", "BS", "current_assets", "+", "BS_ASSETS", "AssetsCurrent", "CurrentAssets"),
        FsLine("BS_PPE", "유형자산", "Property, Plant & Equipment", "BS", "ppe", "+", "BS_NCA", "PropertyPlantAndEquipmentNet", "PropertyPlantAndEquipment"),
        FsLine("BS_NCA", "비유동자산", "Non-current Assets", "BS", "noncurrent_assets", "+", "BS_ASSETS", "AssetsNoncurrent", "NoncurrentAssets"),
        FsLine("BS_ASSETS", "자산총계", "Total Assets", "BS", "total_assets", "+", None, "Assets", "Assets"),
        FsLine("BS_AP", "매입채무", "Accounts Payable", "BS", "payables", "+", "BS_CL", "AccountsPayableCurrent", "TradeAndOtherCurrentPayables"),
        FsLine("BS_STDEBT", "단기차입금", "Short-term Debt", "BS", "st_debt", "+", "BS_CL", "ShortTermBorrowings", "ShorttermBorrowings"),
        FsLine("BS_CL", "유동부채", "Current Liabilities", "BS", "current_liab", "+", "BS_LIAB", "LiabilitiesCurrent", "CurrentLiabilities"),
        FsLine("BS_LTDEBT", "장기차입금", "Long-term Debt", "BS", "lt_debt", "+", "BS_NCL", "LongTermDebtNoncurrent", "NoncurrentBorrowings"),
        FsLine("BS_NCL", "비유동부채", "Non-current Liabilities", "BS", "noncurrent_liab", "+", "BS_LIAB", "LiabilitiesNoncurrent", "NoncurrentLiabilities"),
        FsLine("BS_LIAB", "부채총계", "Total Liabilities", "BS", "total_liab", "+", None, "Liabilities", "Liabilities"),
        FsLine("BS_PIC", "납입자본", "Paid-in Capital", "BS", "paid_in_capital", "+", "BS_EQ", "CommonStockValue", "IssuedCapital"),
        FsLine("BS_RE", "이익잉여금", "Retained Earnings", "BS", "retained_earnings", "+", "BS_EQ", "RetainedEarningsAccumulatedDeficit", "RetainedEarnings"),
        FsLine("BS_EQ", "자본총계", "Total Equity", "BS", "total_equity", "+", None, "StockholdersEquity", "Equity"),
        # CFS
        FsLine("CF_NI", "당기순이익", "Net Income", "CFS", "cf_net_income", "+", "CF_OP", "NetIncomeLoss", "ProfitLoss"),
        FsLine("CF_DA", "감가상각 가산", "D&A Add-back", "CFS", "cf_da", "+", "CF_OP", "DepreciationDepletionAndAmortization", "DepreciationAndAmortisationExpense"),
        FsLine("CF_DNWC", "운전자본 변동", "Change in NWC", "CFS", "cf_delta_nwc", "-", "CF_OP", "IncreaseDecreaseInOperatingCapital", "AdjustmentsForDecreaseIncreaseInWorkingCapital"),
        FsLine("CF_OP", "영업활동현금흐름", "Operating Cash Flow", "CFS", "cfo", "+", None, "NetCashProvidedByUsedInOperatingActivities", "CashFlowsFromUsedInOperatingActivities"),
        FsLine("CF_CAPEX", "자본적지출", "Capital Expenditure", "CFS", "capex", "-", "CF_INV", "PaymentsToAcquirePropertyPlantAndEquipment", "PurchaseOfPropertyPlantAndEquipment"),
        FsLine("CF_INV", "투자활동현금흐름", "Investing Cash Flow", "CFS", "cfi", "+", None, "NetCashProvidedByUsedInInvestingActivities", "CashFlowsFromUsedInInvestingActivities"),
        FsLine("CF_DEBT", "차입금 순증감", "Net Debt Issuance", "CFS", "cf_debt", "+", "CF_FIN", "ProceedsFromRepaymentsOfDebt", "ProceedsFromBorrowings"),
        FsLine("CF_DIV", "배당금 지급", "Dividends Paid", "CFS", "cf_dividends", "-", "CF_FIN", "PaymentsOfDividends", "DividendsPaid"),
        FsLine("CF_FIN", "재무활동현금흐름", "Financing Cash Flow", "CFS", "cff", "+", None, "NetCashProvidedByUsedInFinancingActivities", "CashFlowsFromUsedInFinancingActivities"),
        FsLine("CF_NET", "현금 순증감", "Net Change in Cash", "CFS", "net_change_cash", "+", None, "CashAndCashEquivalentsPeriodIncreaseDecrease", "IncreaseDecreaseInCashAndCashEquivalents"),
    ]


__all__ = ["FsLine", "STATEMENTS", "SIGNS", "load_coa", "coa_index", "by_fs_line",
           "statement_lines", "account_to_fs_line", "validate_coa", "golden_coa"]
