"""
fpna.dispatcher — 요청 텍스트 + tidy 데이터 컬럼으로 템플릿 유형 라우팅.

순차 cascade(빠른 판정). 판정 신호 = (a) 요청 텍스트 키워드 (b) tidy 컬럼 단서
(예: budget&actual 동시 존재 → variance, 기간축만 → period_trend).
"""
from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass
class DispatchResult:
    template: str
    reason: str
    score: int = 0


# (template, [키워드정규식]) — 위에서부터 우선
_RULES = [
    # 고정비 FP&A — 구체 키워드라 일반 variance/손익보다 앞에 둔다.
    # forward_da 는 감가 단어를 쓰므로 depreciation_schedule 보다 앞(미래/forward 한정).
    ("fc_forward_da", [r"미래\s*감가", r"forward\s*d&?a", r"감가.*투영", r"투영.*감가",
                       r"미래\s*상각", r"capex.*감가", r"향후\s*감가"]),
    ("fc_depreciation_schedule", [r"감가\s*상각", r"depreciation", r"상각\s*스케줄",
                                  r"내용\s*연수", r"자산\s*대장.*상각"]),
    ("fc_runrate_normalized", [r"런\s*레이트", r"run[\s-]*rate", r"정규화.*비용",
                               r"연환산", r"normaliz", r"1회성\s*제외", r"베이스라인"]),
    ("fc_cuttability_ladder", [r"절감\s*가능", r"cuttab", r"해지\s*가능", r"비용\s*절감",
                               r"고정비\s*절감", r"time[\s-]*to[\s-]*exit", r"감축\s*여력"]),
    ("fc_driver_unitcost", [r"동인\s*단가", r"단위\s*원가", r"unit\s*cost", r"대당",
                            r"㎡당", r"면적당", r"kwh당", r"활동\s*동인"]),
    ("fc_prepaid_rollforward", [r"선급", r"prepaid", r"롤\s*포워드", r"roll[\s-]*forward",
                                r"선급비용\s*상각"]),
    # C6 신규: 리스 자본화 / 다대다 배부(step-down)
    ("fc_lease_ifrs16", [r"리스", r"lease", r"ifrs\s*16", r"1116", r"사용권\s*자산",
                         r"리스부채", r"rent[\s-]*free"]),
    ("fc_allocation", [r"step[\s-]*down", r"단계식\s*배부", r"다대다\s*배부",
                       r"서비스\s*부서\s*배부", r"부서간\s*배부", r"연쇄\s*배부"]),
    ("fc_variance_bridge", [r"고정비.*브리지", r"고정비.*변동", r"고정비.*요인",
                            r"고정비.*walk", r"고정비.*워크", r"고정비\s*bridge"]),
    ("fc_maturity_wall", [r"만기\s*도래", r"maturity\s*wall", r"약정\s*만기",
                          r"갱신\s*도래", r"만기\s*벽"]),
    # C3 빠진 템플릿 6 — 구체 키워드라 일반 variance/손익보다 앞에 둔다.
    ("pvm_bridge", [r"pvm", r"price[\s/]*volume[\s/]*mix", r"가격[\s·]*물량[\s·]*믹스",
                    r"단가[\s·]*물량[\s·]*믹스", r"믹스\s*분해", r"물량\s*효과",
                    r"단가\s*효과"]),
    ("debt_schedule", [r"부채\s*스케줄", r"debt\s*schedule", r"리볼버", r"revolver",
                       r"cash\s*sweep", r"캐시\s*스윕", r"차입금\s*상환", r"이자\s*스케줄"]),
    ("consolidation_fx", [r"연결\s*재무", r"연결\s*환산", r"consolidat", r"세그먼트\s*합산",
                          r"환산", r"멀티\s*엔티티", r"자회사\s*합산", r"fx\s*환산"]),
    ("working_capital", [r"운전\s*자본", r"working\s*capital", r"dso", r"dpo", r"dio",
                         r"ccc", r"현금\s*전환\s*주기", r"회수\s*일", r"회전\s*일"]),
    ("dupont_roic", [r"듀폰", r"dupont", r"roic", r"roe\s*분해", r"자본\s*수익률",
                     r"eva", r"자산\s*회전율", r"재무\s*레버리지"]),
    ("cost_allocation", [r"공통비\s*배부", r"간접비\s*배부", r"배부\s*기준", r"cost\s*allocat",
                         r"배부", r"allocat"]),
    ("investment_appraisal", [r"npv", r"irr", r"투자\s*타당성", r"회수\s*기간",
                              r"payback", r"할인\s*현금", r"capex", r"투자\s*검토"]),
    ("variance", [r"예실", r"plan\s*vs\s*actual", r"예산\s*대비\s*실적",
                  r"변동\s*분석", r"variance", r"차이\s*분석", r"bridge", r"브리지"]),
    ("period_trend", [r"mom", r"qoq", r"yoy", r"전월\s*대비", r"전분기",
                      r"전년\s*대비", r"추이", r"trend", r"기간별"]),
    ("rolling_forecast", [r"롤링", r"rolling", r"포캐스트\s*갱신", r"forecast\s*update",
                          r"re-?forecast", r"전망\s*갱신"]),
    # 인원/인건비 계획은 budget_build(예산 편성)보다 구체 — 앞에 둔다.
    ("headcount_plan", [r"인원\s*계획", r"인건비", r"headcount", r"roster",
                        r"증원", r"fully[\s-]*loaded", r"인력\s*계획",
                        r"채용\s*계획", r"인원\s*편성"]),
    ("budget_build", [r"예산\s*수립", r"budget\s*build", r"예산\s*편성"]),
    ("cashflow_13w", [r"13\s*주", r"13[\s-]*week", r"단기\s*현금", r"주간\s*현금",
                      r"유동성", r"liquidity", r"자금\s*수지"]),
    # 코호트 잔존/NRR·GRR 은 일반 unit_economics 보다 구체 — 앞에 둔다.
    ("cohort_retention", [r"코호트", r"cohort", r"리텐션", r"retention", r"잔존",
                          r"nrr", r"grr", r"net\s*retention", r"gross\s*retention",
                          r"이탈\s*누적"]),
    ("unit_economics", [r"unit\s*economics", r"유닛\s*이코노믹스", r"cac", r"ltv",
                        r"arr", r"구독", r"churn"]),
    ("scenario_sensitivity", [r"시나리오", r"민감도", r"scenario", r"sensitivity",
                              r"토네이도", r"tornado", r"what[\s-]*if", r"데이터\s*테이블"]),
    ("pnl_3statement", [r"손익", r"p&l", r"pnl", r"3\s*statement", r"재무제표",
                        r"손익계산서", r"income\s*statement"]),
    ("board_kpi_pack", [r"이사회", r"board", r"kpi", r"대시보드", r"dashboard",
                        r"경영\s*보고", r"월간\s*보고"]),
    # 정리표 — 계산 없이 문자+숫자 데이터를 보기 좋게 정리(measure 희소 데이터 착지점).
    ("listing", [r"정리표", r"정리해", r"명세", r"목록", r"리스트", r"listing",
                 r"대장", r"보기\s*좋게", r"표로\s*정리", r"나열"]),
]


def _columns_signal(columns: list[str] | None) -> tuple[str | None, str]:
    """tidy 컬럼/메트릭 단서로 보조 판정."""
    if not columns:
        return None, ""
    low = [str(c).lower() for c in columns]
    joined = " ".join(low)
    has_budget = any("budget" in c or "계획" in c or "plan" in c or "예산" in c for c in low)
    has_actual = any("actual" in c or "실적" in c for c in low)
    if has_budget and has_actual:
        return "variance", "컬럼에 계획+실적 동시 존재"
    if re.search(r"cac|ltv|arr|churn|mrr", joined):
        return "unit_economics", "컬럼에 구독지표"
    if re.search(r"cash|현금|자금", joined):
        return "cashflow_13w", "컬럼에 현금 단서"
    return None, ""


def dispatch(request_text: str = "", *, columns: list[str] | None = None,
             metrics: list[str] | None = None) -> DispatchResult:
    """요청 텍스트 + 컬럼 단서 → DispatchResult.

    텍스트 키워드를 우선 cascade 로 평가, 동점/무매칭 시 컬럼 신호로 보강,
    그래도 없으면 기본값 pnl_3statement / board_kpi_pack.
    """
    text = (request_text or "").lower()
    pool = " ".join(filter(None, [text] + [str(m).lower() for m in (metrics or [])]))

    best = None
    for tmpl, pats in _RULES:
        hits = sum(1 for p in pats if re.search(p, pool))
        if hits and (best is None or hits > best.score):
            best = DispatchResult(tmpl, "텍스트 키워드 %d개 매칭" % hits, hits)

    col_tmpl, col_reason = _columns_signal(columns)
    if best is None and col_tmpl:
        return DispatchResult(col_tmpl, col_reason, 1)
    if best is None:
        return DispatchResult("pnl_3statement", "무매칭 → 기본값(손익)", 0)
    # 컬럼이 variance 를 강하게 시사하면 보정(계획+실적 동시존재는 강한 구조 신호).
    # 텍스트 매칭이 약하거나(score≤1) variance 가 아닌 일반 보고류면 컬럼 신호 우선.
    if col_tmpl == "variance" and best.template != "variance" and (
            best.score <= 1
            or best.template in ("period_trend", "pnl_3statement", "board_kpi_pack")):
        return DispatchResult("variance", "컬럼(계획+실적)이 텍스트보다 강함", best.score + 1)
    return best


# --------------------------------------------------------------------------- #
# 단계(stage) 라우팅 — "지금 파이프라인 어느 단계인가" 선행 판정                #
#   dispatch() 는 분석표 템플릿 28종을 고를 뿐 "더러운 엑셀이냐/운반이냐"는      #
#   판정하지 않는다. classify_stage 가 그 앞단을 책임진다.                       #
#   stage ∈ {ingest, profile, transport, report, analysis}.                            #
# --------------------------------------------------------------------------- #
# (stage, [키워드정규식]) — 위에서부터 우선. analysis 는 fall-through 기본값.
_STAGE_RULES = [
    # pack = 여러 장표가 *연동*돼 한 묶음(공유 가정·크로스시트 tie·Control). report 보다
    # 구체(연동이 메시지)라 최상단. 단일 의도면 아래 cascade(L1)로 위임.
    ("pack", [r"타당성", r"투자\s*심사", r"사업\s*타당성", r"feasibility", r"중기\s*계획",
              r"통합\s*모델", r"통합\s*재무", r"연동\s*워크페이퍼", r"연동\s*모델",
              r"부채\s*구조조정", r"3\s*표\s*연동", r"세\s*개\s*표\s*연동",
              r"조달.*상환.*민감도", r"연간\s*예산.*연동"]),
    ("report", [r"보드\s*팩", r"보드팩", r"제본", r"다중\s*시트", r"board\s*pack",
                r"\bpack\b", r"워크페이퍼", r"workpaper", r"크로스\s*시트", r"제본\s*보고"]),
    ("transport", [r"암호화", r"복호화", r"encrypt", r"decrypt", r"메일\s*본문",
                   r"메일로\s*보", r"운반", r"전송", r"반출.*텍스트", r"본문에\s*붙여"]),
    ("profile", [r"스키마", r"shape", r"프로파일", r"차원\s*없", r"형태만",
                 r"회사에서\s*집", r"집으로\s*반출", r"구조\s*반출", r"마트.*반출"]),
    ("ingest", [r"더럽", r"누더기", r"비정형", r"정리해", r"깨끗하게", r"정형화",
                r"시트가\s*엉망", r"병합\s*셀", r"머지\s*셀", r"제목행", r"각주.*섞",
                r"엉망.*엑셀", r"raw\s*엑셀", r"원본\s*엑셀.*정리"]),
]

# stage → 사람이 바로 칠 수 있는 다음 명령(플레이스홀더 포함).
_STAGE_COMMAND = {
    "ingest": "py main.py ingest <파일.xlsx> out/ingest",
    "profile": "py main.py profile <마트.csv> out/profile_spec.yaml",
    "transport": "py main.py encrypt <평문.txt> --mail   (받는 쪽: py main.py decrypt <암호문>)",
    "pack": "py main.py pack <name> out/pack.xlsx",
    "report": "py main.py report fc_boardpack out/pack.xlsx",
    "analysis": "py main.py dispatch \"<요청>\"  →  py main.py render <type> out/<type>.xlsx",
}

# pack 트리거 → 구현 카탈로그 name. 미구현 팩은 packs.md 설계만(단일 exhibit 위임 안내).
_PACK_RULES = [
    ("feasibility", [r"타당성", r"투자\s*심사", r"사업\s*타당성", r"feasibility",
                     r"조달.*상환", r"투자\s*검토.*연동"]),
]


def resolve_pack(request_text: str = "") -> str | None:
    """pack 트리거 텍스트 → 구현된 pack name. 없으면 None(packs.md 카탈로그 참조)."""
    text = (request_text or "").lower()
    for name, pats in _PACK_RULES:
        if any(re.search(p, text) for p in pats):
            return name
    return None


def classify_stage(request_text: str = "", *, has_messy_file: bool = False,
                   has_clean_table: bool = False) -> tuple[str, str]:
    """요청을 파이프라인 단계로 선분류한다.

    반환 = (stage, next_command). stage ∈ {ingest, profile, transport, report, analysis}.
    텍스트 키워드를 우선 판정하고, 파일 단서(messy/clean)로 보강한다.
    analysis 는 기본값 — 단계 키워드가 없으면 분석표 요청으로 본다(→ dispatch 가 템플릿 판정).
    """
    text = (request_text or "").lower()
    for stage, pats in _STAGE_RULES:
        if any(re.search(p, text) for p in pats):
            if stage == "pack":
                name = resolve_pack(text)
                if name:
                    return "pack", "py main.py pack %s out/pack.xlsx" % name
                # 트리거는 연동 묶음인데 구현 카탈로그 없음 → packs.md 참조 안내.
                return "pack", ("연동 묶음 → packs.md 카탈로그 참조. 미구현 팩은 "
                                "단일 exhibit(dispatch)로 위임. 구현: "
                                "py main.py pack feasibility out/pack.xlsx")
            return stage, _STAGE_COMMAND[stage]
    # 텍스트 신호 없을 때 파일 단서로 보강: 누더기 파일이면 ingest 가 입력을 만든다.
    if has_messy_file and not has_clean_table:
        return "ingest", _STAGE_COMMAND["ingest"]
    return "analysis", _STAGE_COMMAND["analysis"]


def route(request_text: str = "", *, columns: list[str] | None = None,
          metrics: list[str] | None = None, has_messy_file: bool = False,
          has_clean_table: bool = False) -> dict:
    """단계 라우팅 + (analysis 면) 템플릿 판정까지 한 번에.

    반환 dict 키:
      stage          — ingest/profile/transport/analysis
      next_command   — 사람이 바로 칠 수 있는 명령 문자열
      reason         — 판정 근거
      template       — analysis 일 때만(dispatch 결과). 그 외 None
    dispatch() 는 그대로 호출만 하므로 기존 라우팅은 변하지 않는다(하위호환).
    """
    stage, next_command = classify_stage(
        request_text, has_messy_file=has_messy_file, has_clean_table=has_clean_table)
    if stage != "analysis":
        return {"stage": stage, "template": None,
                "next_command": next_command,
                "reason": "단계 키워드 매칭 → %s" % stage}
    disp = dispatch(request_text, columns=columns, metrics=metrics)
    render_cmd = "py main.py render %s out/%s.xlsx" % (disp.template, disp.template)
    return {"stage": "analysis", "template": disp.template,
            "next_command": render_cmd,
            "reason": "분석표 → dispatch: %s" % disp.reason}


# --------------------------------------------------------------------------- #
# 컬럼 의미 기반 추천 (infer.summarize 결과 → 템플릿/레이아웃)                  #
#   infer 를 import 하지 않고 summary dict 만 받는다(느슨한 결합). 호출자가     #
#   fpna.infer.infer_columns→summarize 로 만들어 전달. measure 유무가 1차 분기. #
# --------------------------------------------------------------------------- #
def recommend_from_roles(summary: dict) -> DispatchResult:
    """infer.summarize(dict) → DispatchResult. "적용 템플릿 없음" 종착 제거.

    summary 키: time/measure/dimension/id(list) + has_measure/n_measure/n_dimension.
    분기: measure 0 → listing(정리표) / time+measure → period_trend /
          dim+2measure(계획·실적류) → variance / dim+1measure → listing(소계) /
          그 외 → listing.
    """
    nm = summary.get("n_measure", 0)
    has_time = bool(summary.get("time"))
    has_dim = bool(summary.get("dimension") or summary.get("id"))
    if nm == 0:
        return DispatchResult("listing", "measure 0 → 정리표(식별자 위주 데이터)", 1)
    if has_time and nm >= 1:
        return DispatchResult("period_trend", "time+measure → 기간 추이", 2)
    if nm >= 2:
        return DispatchResult("variance", "dim+2measure → 예실/비교(계획·실적 가정)", 2)
    if has_dim and nm == 1:
        return DispatchResult("listing", "dimension+1measure → 정리표(그룹 소계)", 1)
    return DispatchResult("listing", "기본 → 정리표", 0)


__all__ = ["dispatch", "DispatchResult", "classify_stage", "route", "resolve_pack",
           "recommend_from_roles"]
