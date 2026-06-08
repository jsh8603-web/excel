# Dispatcher — 템플릿 라우팅 (dispatch.md)

요청 텍스트 + tidy 데이터 단서로 어떤 템플릿 유형을 쓸지 **빠른 순차 cascade**로 판정한다.
구현: `fpna/dispatcher.py`. 호출: `py main.py dispatch "<요청 텍스트>"`.

## 판정 신호

1. **요청 텍스트 키워드** (우선) — 유형별 정규식 매칭 수가 많은 쪽.
2. **tidy 컬럼/메트릭 단서** (보조·보정) — 예: 계획+실적 컬럼 동시 존재 → variance.

## Cascade 순서

| # | 조건(키워드 예) | → 유형 |
|---|---|---|
| 1 | NPV·IRR·투자 타당성·회수기간·payback·할인현금·capex | `investment_appraisal` |
| 2 | 예실·plan vs actual·예산 대비 실적·변동분석·variance·bridge | `variance` |
| 3 | MoM·QoQ·YoY·전월대비·추이·trend·기간별 | `period_trend` |
| 4 | 롤링·rolling·포캐스트 갱신·전망 갱신 | `rolling_forecast` |
| 5 | 예산 수립·budget build·인건비·headcount·인원 계획 | `budget_build` |
| 6 | 13주·13-week·단기 현금·주간 현금·유동성 | `cashflow_13w` |
| 7 | unit economics·CAC·LTV·ARR·코호트·구독·churn | `unit_economics` |
| 8 | 시나리오·민감도·tornado·what-if·데이터테이블 | `scenario_sensitivity` |
| 9 | 손익·P&L·재무제표·손익계산서·income statement | `pnl_3statement` |
| 10 | 이사회·board·KPI·대시보드·경영보고 | `board_kpi_pack` |
| — | (무매칭 기본값) | `pnl_3statement` |

## 컬럼 단서 보정

- `budget/계획/plan` AND `actual/실적` 컬럼 동시 존재 → **variance** 로 보정(텍스트가 period_trend/pnl을 가리켜도).
- `cac/ltv/arr/churn/mrr` → unit_economics.
- `cash/현금/자금` → cashflow_13w.

## 사용 예

```
py main.py dispatch "이번 분기 NPV IRR 투자 타당성 검토"   → investment_appraisal (score 3)
py main.py dispatch "예실 변동 분석 브리지"                → variance (score 3)
py main.py dispatch "13주 단기 현금 유동성"               → cashflow_13w (score 3)
```

판정 후 `py main.py render <type> out.xlsx` 로 렌더(QC 게이트). 실데이터 바인딩은
3층 래퍼(데이터 바인딩)에서 시트명·컬럼 위치를 주입해 흡수한다.
