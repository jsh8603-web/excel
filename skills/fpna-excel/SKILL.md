---
name: fpna-excel
description: >
  사내 FP&A Excel 작업의 **마스터 라우터**. "excel/엑셀" 작업은 이 스킬로 진입한다.
  모든 엑셀 산출은 ALWAYS 필수 경로(grain 선언 → 차원 정합 → View Contract 게이트 →
  QC → render 저장)를 우회 없이 거치고, 특수 변환·운반(ingest 누더기정형화·profile SHAPE추출·
  crypto 메일운반)은 ONLY IF 포인터로 상황 매칭 시에만 호출한다. 템플릿 12종(예실/투자NPV/추이/
  롤링/예산/13주현금/유닛이코노믹스/시나리오/손익/이사회KPI/고정비 감가·브리지)과 고정비 FP&A
  자산(View Contract R1~R11, conformed dimension)을 dispatcher 가 유형으로 라우팅한다.
  무설치·openpyxl 단일 의존(vendor 동봉).
---

# FP&A Excel Skill — 마스터 라우터

회사 PC(설치0·외부망 차단·AI 없음)에서 `git pull` 후 **`py main.py`만으로** 동작.
런타임 의존성은 `vendor/`에 동봉한 openpyxl 하나뿐. **excel 트리거 = 이 라우터로 진입.**

## 0. 라우팅 원칙 — ALWAYS vs ONLY IF

엑셀 작업 요청을 받으면 두 갈래로 나눈다:

- **ALWAYS (필수 경로)** — 무언가를 *산출*하는 모든 요청. 아래 §1 의 5단계를 **우회 없이** 거친다. 순서·강제가 본질.
- **ONLY IF (특수 포인터)** — 변환·운반 같은 *상황 한정* 작업. 조건이 맞을 때만 §3 포인터로 안내. 자동 호출 안 함.

> 판정: "엑셀 산출물(보고서/모델/테이블)을 만드는가?" → YES면 ALWAYS. "입력 정형화·구조 반출·암복호화 같은 곁작업인가?" → ONLY IF.

## 1. ALWAYS — 필수 경로 (모든 엑셀 산출, 예외 없음)

```
① grain 선언       1행이 무엇인가 (fpna.dims.Fact — 미선언 = 빌드 실패, R8)
② 차원 정합        conformed dimension 으로 슬라이스 가능, orphan key 0 (fpna.dims)
③ View Contract    전수성·tie-out 게이트 (fpna.view_contract R1~R11)
④ QC               불변식 통과 + anomaly 보존 검증 (템플릿 qc())
⑤ render 저장      QC 통과 시에만 저장 (fpna.render — 단일 통로)
```

- **render 게이트의 의미(확정)**: `QCReport.passed=False` 는 **"산출물이 부정직하다"**(정합/tie-out/보존 위반)는 뜻이지 "데이터에 문제 있음"이 아니다. anomaly(미계상·부호반전 같은 *발견*)는 저장을 막지 않고 flag 로 노출하며, anomaly 의 **은폐**만 저장을 막는다.
- **현재 강제 수준**: `render()` 가 단일 통로이고 각 템플릿이 `qc()` 를 구현한다. 단 qc 내용은 템플릿 재량이라 공통 게이트 우회 여지가 있다 → §4 재구조화로 강제력 격상 예정.

## 2. 템플릿 라우팅 (dispatch → render)

요청 텍스트 + tidy 컬럼 단서로 `fpna.dispatcher` 가 유형 판정. 모두 §1 필수 경로를 거친다.

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

## 3. ONLY IF — 특수 포인터 (상황 매칭 시만, 자동 호출 X)

리포트 템플릿이 아니라 스파인 **밖**의 변환·운반이다. 조건이 맞을 때만 안내한다.

| 포인터 | 트리거 조건 | 명령 |
|---|---|---|
| **ingest** (on-ramp) | 입력이 **누더기/비정형** 엑셀일 때 (이미 tidy면 건너뜀) | `py main.py ingest "<파일.xlsx>" out\ingest [--sheet 시트명]` |
| **profile** | 정제 마트를 **회사→집 반출**(차원없는 SHAPE 스키마)할 때 | `py main.py profile "<마트.csv>" out\profile_spec.yaml` |
| **crypto** | 텍스트를 **메일 본문으로 운반**(암복호화)할 때 | `py main.py encrypt <평문> --mail` / `py main.py decrypt <암호문>` |

- ingest 는 필수 경로의 **입력**을 만든다(누더기 → tidy → 그 tidy 가 §1 ①로 진입). profile/crypto 는 경계 횡단(반출/운반) 시의 side-door — 선형 must-do 아님.

## 4. 재구조화 검토 (목표 구조 — plan 연결)

> 자문 3R(`out/consult-fixed-cost-3r.md`) 결론. **현 구조의 한계**: §1 필수 경로가 "render 단일 통로 + 템플릿 자율 qc" 라 공통 게이트(②③+anomaly 보존)를 우회할 수 있다. 체크리스트는 honor 의존이라 같은 실패가 재발한다.

**목표 = 단일 pipeline 함수 + Protocol + receipt token** (상속도 render-only도 아님):

```
run_report(t, src):                 # 유일한 스파인 (게이트 순서 소유)
  grain   = t.grain_spec            # ① 선언
  conform_dimensions(src, grain)    # ② 정합 (base 소유)
  artifact = t.build(src, grain)    # ③ content (템플릿 재량)
  receipt  = view_contract_gate(...)# ④ 전수성+tie-out → GatePass 토큰 발급
  qc(artifact, t.invariants(), t.detectors(), receipt)  # ⑤ 불변식+anomaly 보존
  return render(artifact, receipt)  # ⑥ receipt 없이 호출 불가 → 우회 구조적 차단
```

- `ReportTemplate` Protocol: `grain_spec` / `build()` / `invariants()->list` / `detectors()->list`. 기존 12종은 qc 를 invariants()+detectors() 로 기계적 분할.
- **함정(하지 말 것)**: 불변식 DSL/rule engine, ingest/profile/crypto 를 파이프라인에 편입, 깊은 상속.
- **마이그레이션**: fc 2종(감가·브리지)부터 → 나머지 10종 기계적 추출. **구현은 `plan.md` 로 진행.**

## 5. 진입점

```powershell
py main.py selftest                                  # 회귀(설치 0) — 골든 12종 + ingest 픽스처
py main.py dispatch "<요청 텍스트>"                  # 텍스트 → 유형 판정
py main.py render <type> out\<type>.xlsx             # 빌드 + QC 게이트 후 저장
py main.py list                                      # 유형 목록
py -S main.py selftest                               # 무설치 재현(site-packages 차단)
```

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
