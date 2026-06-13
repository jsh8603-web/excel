---
name: fpna-excel
description: >
  사내 FP&A Excel 작업의 **마스터 라우터**. "excel/엑셀" 작업은 이 스킬로 진입한다.
  요청을 받으면 먼저 **단계(stage)** 를 판정한다 — 누더기 입력(ingest)·구조 반출(profile)·
  메일 운반(transport)·분석표 산출(analysis). 분석표는 ALWAYS 필수 경로(grain 선언 → 차원 정합 →
  View Contract 게이트 → QC → render 저장)를 `run_report` 스파인이 소유해 우회 없이 거치고,
  특수 변환·운반(ingest·profile·crypto)은 ONLY IF 포인터로 상황 매칭 시에만 호출한다.
  dispatcher 가 stage 를 선분류하고(classify_stage/route) 분석표는 유형으로 라우팅한다(dispatch).
  무설치·openpyxl 단일 의존(vendor 동봉).
---

# FP&A Excel Skill — 마스터 라우터

회사 PC(설치0·외부망 차단·AI 없음)에서 `git pull` 후 **`py main.py`만으로** 동작.
런타임 의존성은 `vendor/`에 동봉한 openpyxl 하나뿐. **excel 트리거 = 이 라우터로 진입.**

## 0. 발동하면 이 순서로 따라간다 (워크스루)

스킬이 켜지면 **무조건 이 4스텝**을 밟는다. 추측하지 말고 단계부터 판정한다.

```
① stage 판정    요청이 어느 단계인가? (dispatcher.classify_stage / route)
                  messy(누더기 입력)→ingest · clean→profile(반출)/analysis(분석)
                  transport(암복호화·메일운반)→crypto
② 해당 명령 실행  단계별로 칠 명령이 정해져 있다(아래 표·§3·§5). classify_stage 가 next_command 를 그대로 돌려준다.
③ 분석이면 dispatch → render
                  analysis 면 dispatch 로 템플릿을 고르고, render 가 run_report 스파인을 태운다.
                  스파인이 grain→차원정합→View Contract→QC 게이트를 base-owned 로 강제(우회 불가).
④ 회사↔집 규율   실데이터 숫자·거래처명·원본파일 반출 금지. 구조 메타·코드·수식만 텍스트로(§6).
```

**ALWAYS vs ONLY IF** — ②~③의 두 갈래:

- **ALWAYS (필수 경로)** — *분석표를 산출*하는 모든 요청(stage=analysis). §1 의 필수 경로를 **우회 없이** 거친다. 순서·강제가 본질이고, **그 강제는 `run_report` 스파인이 소유**한다(검증은 특수상황 끝이 아니라 메인경로가 가진다).
- **ONLY IF (특수 포인터)** — ingest·profile·crypto 같은 *상황 한정* 곁작업(stage=ingest/profile/transport). 조건이 맞을 때만 §3 포인터로 안내. 자동 호출 안 함.

> 단계 판정 한 줄: "분석표(보고서/모델/테이블)를 만드는가?" → YES면 analysis=ALWAYS. "입력 정형화·구조 반출·암복호화 같은 곁작업인가?" → ingest/profile/transport=ONLY IF. `classify_stage(요청)` 가 키워드로 이 판정을 코드로 수행하고, `route(요청)` 는 analysis 면 dispatch 까지 이어 template+next_command 를 돌려준다.

## 1. ALWAYS — 필수 경로 (분석표 산출 = stage analysis, 예외 없음)

```
스파인 = fpna.pipeline.run_report(template, data, out_path)  ← main.py render 가 이걸 호출
① grain 선언       1행이 무엇인가 (fpna.dims.Fact — 미선언 = 빌드 실패, R8)
② 차원 정합        conformed dimension 으로 슬라이스 가능, orphan key 0 (base-owned)
③ View Contract    전수성·tie-out 게이트 (fpna.view_contract R1~R11) → receipt(GatePass) 발급
④ QC               base 공통 게이트 + 템플릿 불변식 + anomaly 보존
⑤ render 저장      receipt 있을 때만 저장 — receipt 없이 호출하면 RuntimeError(우회 구조적 차단)
```

- **검증은 메인경로가 소유한다(확정)**: `run_report` 가 ①~⑤ 순서를 **base-owned 로 강제**한다. 템플릿이 자율 `qc()` 에서 빠뜨려도 스파인의 `_base_owned_gate` 가 수식에러·grain·anomaly 보존을 잡는다. 저장(`render`)은 스파인 안에서만 mint 되는 receipt 를 요구하므로(토큰+wb 해시 재대조) 스파인을 우회한 저장은 불가능하다. → "검증은 특수상황 끝이 아니라 메인경로가 소유."
- **render 게이트의 의미**: `QCReport.passed=False` 는 **"산출물이 부정직하다"**(정합/tie-out/보존 위반)는 뜻이지 "데이터에 문제 있음"이 아니다. anomaly(미계상·부호반전 같은 *발견*)는 저장을 막지 않고 flag 로 노출하며, anomaly 의 **은폐**만 저장을 막는다.

## 2. 템플릿 라우팅 (stage=analysis 일 때만 → dispatch → render)

§0 의 stage 판정이 **analysis** 로 떨어졌을 때만 여기로 온다. 요청 텍스트 + tidy 컬럼 단서로 `fpna.dispatcher.dispatch` 가 유형 판정. 모두 §1 필수 경로(run_report 스파인)를 거친다.

- **의도 앵커**: 각 유형이 답하는 CEO/CFO 질문은 해당 `fpna/templates/<type>.py` 모듈 docstring 첫 줄에 박혀 있다(예: `fc_maturity_wall` = "12개월 내 만료/갱신 도래 고정비 얼마, 갱신 인상 리스크는?"). 산출 전 그 한 줄로 의도를 확인한다.

| 유형 | 트리거 신호 | 산출 |
|---|---|---|
| variance | 예실·plan vs actual·변동분석·brige | 예실 변동표 + 워터폴 |
| investment_appraisal | NPV·IRR·투자타당성·회수기간 | 투자 타당성 |
| period_trend | MoM·QoQ·YoY·추이 | 기간 추이 |
| rolling_forecast | 롤링·forecast 갱신 | 롤링 포캐스트 |
| budget_build | 예산수립·인건비·headcount | 예산/인건비 |
| cashflow_13w | 13주·단기현금·유동성 | 13주 현금 |
| unit_economics | CAC·LTV·ARR·코호트·churn | 유닛 이코노믹스 |
| scenario_sensitivity | 시나리오·민감도·토네이도 | 시나리오 민감도 |
| pnl_3statement | 손익·P&L·재무제표 | 손익 |
| board_kpi_pack | 이사회·KPI·대시보드·경영보고 | 이사회 KPI |
| **fc_depreciation_schedule** | 감가상각·상각 스케줄·내용연수 | 자산별 월 상각 + GL 대사 |
| **fc_variance_bridge** | 고정비 변동요인·고정비 브리지·walk | 고정비 5요인 워터폴 |

- **고정비 FP&A 분기**: 위 fc_* + `fpna.dims`(회계 캘린더 역월/4-4-5 + conformed dimension 6축) + `fpna.view_contract`(R1~R11). 상세 = `skills/fpna-fixed-cost-tables/SKILL.md`.
- **신규 유형 추가**: `fpna/templates/<type>.py`(TYPE/INPUT/golden_sample/build/qc 덕타이핑) → `_MODULES` 등록 → dispatcher 키워드. build=house_style만, qc=finance 재계산+view_contract.

## 3. ONLY IF — 특수 포인터 (stage ≠ analysis, 자동 호출 X)

리포트 템플릿이 아니라 스파인 **밖**의 변환·운반이다. §0 stage 판정이 ingest/profile/transport 로 떨어졌을 때만 안내한다. `classify_stage` 가 이 매칭을 키워드로 수행하고 아래 명령(next_command)을 그대로 돌려준다.

| 포인터 | stage | 언제 쓰나 (트리거 조건) | 명령 |
|---|---|---|---|
| **ingest** (on-ramp) | ingest | 입력이 **누더기/비정형** 엑셀("더럽다/병합셀/시트 엉망")일 때 — 이미 tidy면 건너뜀 | `py main.py ingest "<파일.xlsx>" out\ingest [--sheet 시트명]` |
| **profile** | profile | 정제 마트를 **회사→집 반출**(차원없는 SHAPE 스키마)할 때 | `py main.py profile "<마트.csv>" out\profile_spec.yaml` |
| **crypto** | transport | 텍스트를 **메일 본문으로 운반**(암복호화)할 때 | `py main.py encrypt <평문> --mail` / `py main.py decrypt <암호문>` |

- ingest 는 필수 경로의 **입력**을 만든다(누더기 → tidy → 그 tidy 가 §1 ①로 진입). profile/crypto 는 경계 횡단(반출/운반) 시의 side-door — 선형 must-do 아님.

## 4. 스파인 구조 (구현됨 — fpna.pipeline.run_report)

> 자문 3R(`out/consult-fixed-cost-3r.md`) 결론을 구현한 **단일 스파인**. 과거 한계("render 단일 통로 + 템플릿 자율 qc" 라 공통 게이트를 우회 가능, honor 의존)를 `run_report` + receipt token 으로 닫았다. `main.py render` 가 이 함수를 호출한다.

**구조 = 단일 pipeline 함수 + Protocol + receipt token** (상속도 render-only도 아님):

```
run_report(template, data, out_path):       # 유일한 스파인 (게이트 순서 소유)
  wb       = template.build(data, ...)       # ③ content (템플릿 재량)
  _base_owned_gate(rep, wb, data, template)  # ②④ 수식에러·grain·anomaly 보존·T4 tie-out (base 소유)
  _template_checks(rep, wb, data, template)  # ⑤ 불변식(invariants)+탐지기(detectors) 또는 구 qc
  receipt  = _mint(...) if rep.passed else None   # 통과 시에만 GatePass 발급(_MINT 토큰)
  _render_with_receipt(wb, out, receipt)     # ⑥ receipt 없이 저장 불가 → 우회 RuntimeError
```

- `ReportTemplate` Protocol: 필수 = `TYPE` + `build(data)->Workbook`. 신방식 hook(선택) = `invariants()->list` / `detectors()->list` / `conserves()`(raw INPUT vs 보고총계 T4 tie-out) / `fact_of`·`ledger_of`·`surfaced_of`. 신 hook 없으면 구 `qc(wb,data)` 를 그대로 merge(하위호환).
- **하지 말 것(함정)**: 불변식 DSL/rule engine 도입, ingest/profile/crypto 를 스파인에 편입(§3 그대로 밖), 깊은 상속.
- **함정(하지 말 것)**: 불변식 DSL/rule engine, ingest/profile/crypto 를 파이프라인에 편입, 깊은 상속.
- **하위호환**: 신 hook(invariants/detectors)을 안 단 템플릿은 구 `qc()` 가 그대로 스파인에서 merge 되므로, 28종 전부 별도 마이그레이션 없이 run_report 를 통과한다.

## 5. 진입점

```powershell
py main.py selftest                                  # 회귀(설치 0) — 골든 전체 + ingest 픽스처
py main.py dispatch "<요청 텍스트>"                  # 텍스트 → 유형 판정(analysis 단계)
py main.py render <type> out\<type>.xlsx             # run_report 스파인(QC 게이트) 후 저장
py main.py ingest "<파일.xlsx>" out\ingest           # (stage=ingest) 누더기 → tidy
py main.py profile "<마트.csv>" out\spec.yaml        # (stage=profile) SHAPE 반출
py main.py encrypt <평문> --mail / decrypt <암호문>  # (stage=transport) 메일 운반
py main.py list                                      # 유형 목록
py -S main.py selftest                               # 무설치 재현(site-packages 차단)
```

- stage 선판정은 코드로도 호출 가능: `from fpna.dispatcher import classify_stage, route` — `route(요청)` 가 stage·template·next_command 를 dict 로 돌려준다.

## 6. 워크플로 (회사 ↔ 집)

1. **회사 → 집**: 구조 메타(컬럼/시트/포맷/규모)·에러 메시지만 텍스트로. 실데이터 숫자 반출 금지.
2. **집**: ingest 로 구조 검증 → 템플릿 데이터 바인딩 → 골든+QC 검증 → `tools/verify_xlsx.py`(집-전용 Excel COM).
3. **집 → 회사**: 코드/수식/런북(텍스트). 회사에서 실데이터로 render, 결과는 회사 PC 마무리.

## 7. 제약 (반드시 준수)

- openpyxl 외 런타임 라이브러리 금지(pandas/numpy/pydantic/XlsxWriter/formulas/dbt/GE). stdlib + dataclass 직접 구현.
- 중간데이터 csv/json(stdlib)만. 차트는 openpyxl(워터폴=stacked-bar+투명 base).
- 합성 재무수치 금지(구조 골든 더미만 예외 — 명시). **QC 통과 후 출력 확정.**
- COM/xlwings/MCP는 집-전용 검증(`tools/`). 회사 런타임 의존 아님.

## 8. 관련 문서

- 고정비 FP&A: `skills/fpna-fixed-cost-tables/SKILL.md` + `skills/fpna-fixed-cost-tables/claude_snippet.md`
- 정형화 규칙: `rules/normalize_rules.md` · 디스패처: `dispatch.md` · tidy 스키마: `schema/tidy_schema.md`
- 운영 매뉴얼: `CLAUDE.md` · 자문 종합: `out/consult-fixed-cost-3r.md`
