# Packs — 다중 exhibit 연동 묶음 (packs.md)

`dispatch.md`(L1, 단일 의도→단일 템플릿)의 **형제 문서(L2)**. "결과가 여러 장표로 **연동**돼
한 묶음으로 가는가"를 먼저 판정하고, 그렇다면 어떤 팩을 쓸지 카탈로그로 안내한다.
구현: `fpna/pack.py`(오케스트레이터) + `fpna/packs/`(카탈로그 레지스트리). 설계 근거 = R7(PFRAM
다중시트 연동)·R2(FAST Control 시트)·R8(Doubletalk ledger 대안).
호출: `py main.py pack <name> out/pack.xlsx`.

## 판정 1단계 — 단일 vs 팩

> **"결과 장표가 서로 연동돼 한 묶음으로 가나?"**
> - **No** → `dispatch.md`(단일 템플릿). 장표 하나면 팩으로 부풀리지 말 것(과투자).
> - **Yes** → 아래 팩. 팩은 **"연동이 메시지일 때"만** — 공유 가정집합에서 여러 exhibit 가
>   흐르고, 크로스시트 tie 가 검증의 핵심일 때.

`dispatcher.py`의 pack 게이트가 트리거(타당성·투자심사·중기계획·통합모델·연동·부채 구조조정)를
만나면 이 카탈로그로 라우팅한다. 미구현 팩은 단일 exhibit(dispatch)로 위임.

## 팩 공통 구조

각 팩 = **[상황 트리거] / [exhibit 구성=`_MODULES` 키] / [공유=calendar·assumptions·(coa)] /
[핵심 ties] / [왜]**. 모든 팩은:
- 단일 가정집합(공유 calendar/debt/dims)에서 각 exhibit `build()` → 시트당 1 exhibit(graft 합본).
- 첫 시트 = **Control(Index)**: exhibit↔시트 맵 + 모델체크(A=L+E·현금 tie) OK/XX + cross ties.
- **`run_report` 스파인 경유** → receipt 없이는 저장 불가(우회 차단).
- 연동은 셀참조가 아니라 **공유 facts** + `ConserveSpec`(build 호출 0, AST 독립).

크로스시트 tie 5종(공통 어휘):
(a) BS 항등 자산=부채+자본 · (b) 현금 tie CFS기말=BS현금 · (c) 부채 tie debt_schedule기말=BS차입금 ·
(d) 감가 연동 depreciation=IS감가 · (e) 이자 연동 평균부채×rate(solve)=IS이자.

## 카탈로그

### 1. 사업타당성·투자심사 팩 — `feasibility` ✅구현
- **트리거**: 사업타당성 / 투자심사 / 투자검토 연동 / 조달·상환·민감도
- **exhibit**: `investment_appraisal` + `pnl_3statement`(순환연동) + `debt_schedule` + `working_capital` + `scenario_sensitivity`
- **공유**: 단일 debt 가정({revolver, term}+rates) → `pnl_3statement.solve_and_link`가 이자·현금·부채를 내생화(리볼버 plug 해소)
- **ties**: (a)(b)(c)(d)(e) 전부
- **왜**: 조달·상환·세금·민감도가 단일 가정집합에서 일관 흐른다. 리볼버 plug(이자↔현금↔부채 순환)를 `finance.solve_revolver`가 고정점 수렴.

### 2. 연간예산·중기계획 팩 — `annual_budget` ⏳설계(packs.md만)
- **트리거**: 연간예산 연동 / 중기계획 / 통합 plan
- **exhibit**: `budget_build` + `headcount_plan` + `pnl_3statement` + `rolling_forecast`(또는 `cashflow_13w`)
- **공유**: 단일 인원·예산 가정 → 인건비→IS, 예산라인→P&L
- **ties**: 인건비→IS / 예산라인→P&L / 현금 tie (b)
- **왜**: 인원·예산이 손익·현금으로 흘러 한 번에 검증.

### 3. 이사회보고 팩 — `board_report` ⏳설계
- **트리거**: 이사회 보고 묶음 / 경영보고 통합
- **exhibit**: `board_kpi_pack` + `variance` + `period_trend` + `pnl_3statement`(요약)
- **ties**: 변동=Scenario 차이(R9) / 현금 tie (b)
- **왜**: KPI·변동·추세·손익을 한 receipt로 묶어 보고. (단순 이사회 대시보드 1장이면 `board_kpi_pack` 단일 → dispatch)

### 4. 고정비 심층 팩 — `fixed_cost_deep` ⏳설계
- **트리거**: 고정비 심층 / 고정비 통합 / 고정비 전수 연동
- **exhibit**: `fc_depreciation_schedule` + `fc_lease_ifrs16` + `fc_allocation` + `fc_runrate_normalized` + `fc_cuttability_ladder` + `fc_driver_unitcost` + `fc_prepaid_rollforward` + `fc_maturity_wall` + `pnl_3statement`
- **ties**: master→GL(R11) / allocation_conserves(R11) / 감가 (d)
- **왜**: 고정비 전수가 손익·계정에 보존되며 연결. (만기 보드팩만 필요하면 `fc_boardpack` report → `py main.py report fc_boardpack`)

### 5. 부채·구조조정 팩 — `debt_restructuring` ⏳설계
- **트리거**: 부채 구조조정 / 리파이낸싱 연동 / 워크아웃
- **exhibit**: `debt_schedule` + `pnl_3statement` + `cashflow_13w` + `scenario_sensitivity`
- **ties**: (c)(e) / 현금 (b) · (옵션) `ledger_mode` 고려
- **왜**: 상환·이자·현금·시나리오가 한 모델에서 순환 해소. solve_revolver가 plug 해소.

## 신규 팩 추가 방법

1. `fpna/packs/<name>.py` 생성, `make_spec() -> pack.PackSpec` 노출(exhibits·shared_facts·cross_ties·model_checks).
2. `fpna/packs/__init__.py`의 `_PACKS`에 등록.
3. `fpna/dispatcher.py`의 `_PACK_RULES`에 트리거 키워드 추가.
4. `py main.py pack <name> out/x.xlsx` + `py tools/verify_xlsx.py out\x.xlsx`로 검증.
5. 본 카탈로그에 ✅로 갱신.

> ⚠️ **과투자 경계**: 단일 exhibit 만 필요하면 팩으로 부풀리지 말 것. 팩은 연동이 메시지일 때만.
> 차트·조건부서식은 graft 에서 누락될 수 있다(openpyxl 제약) — 값·수식·tie 는 보존, 수식 정확성은
> `fullCalcOnLoad` + `tools/verify_xlsx`(COM)가 담당.
