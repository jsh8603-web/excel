---
name: fpna-fixed-cost-tables
description: >
  고정비 FP&A 테이블 셋 스킬. 다중 테이블을 Kimball star schema(fact + conformed
  dimension 6축)로 보고, View Contract R1~R11(전수성·grain·tie-out·시나리오 정합)을
  코드로 강제해 데이터 엔지니어링 기본값(sparse 압축·하루치만·스냅샷만·anti-join only)을
  구조적으로 차단한다. 감가상각 스케줄·고정비 variance bridge 를 QC 게이트 통과 후 렌더.
  무설치·openpyxl 단일 의존. "고정비/감가상각 스케줄, 고정비 변동요인 브리지, 자산 대장
  상각, 계약·자산 대사, 회계 캘린더(역월/4-4-5) 테이블" 요청 시 사용.
---

# FP&A 고정비 테이블 스킬

폐쇄망 Claude 가 다중 테이블을 만나도 **"항상 올바른 테이블 셋"** 을 만들도록 강제한다.
주의 문장이 아니라 **계약 + 헬퍼 + QC 게이트 + 테스트 4중**으로 박는다.

## 핵심 원칙 — 왜 이 스킬이 필요한가

도메인 지식 없이 다중 테이블을 만나면 데이터 엔지니어링 기본값(sparse 압축·하루치만·
스냅샷만·anti-join only)으로 빠진다. 이 스킬은 그걸 **불가능하게** 만든다.

- **출력은 fact + conformed dimension** (Kimball). 평면 테이블 더미가 아니다.
- **grain 을 먼저 선언**한다(`fpna.dims.Fact`). 미선언 = 빌드 실패(R8).
- **시간축은 회계 캘린더에서만** 생성한다(`fpna.dims.AccountingCalendar`). 데이터에서
  날짜를 뽑아 열을 만드는 코드 금지(R1).
- **variance 는 시나리오 축의 차이로만**(R9). 한쪽만 있는 cost center 를 0 으로 버리지 않는다.

## 트리거

- "감가상각 스케줄", "자산 대장 상각", "depreciation", "내용연수" → **fc_depreciation_schedule**
- "고정비 변동요인", "고정비 브리지/walk", "고정비 bridge" → **fc_variance_bridge**
- "예실 변동" 일반은 기존 **variance** 로 라우팅(겹침 방지 — fc 규칙이 cascade 앞).

## 진입점

repo 루트에서 실행. 모든 진입점은 `fpna._bootstrap` 을 최우선 import 해 `vendor/` 주입.

```powershell
py main.py list                                              # 유형 목록(fc_* 포함)
py main.py dispatch "<요청 텍스트>"                          # 텍스트 → 유형 라우팅
py main.py render fc_depreciation_schedule out\fc_dep.xlsx   # 빌드 + QC 게이트 후 저장
py main.py render fc_variance_bridge out\fc_bridge.xlsx
py -m unittest tests.test_fc                                 # 22 회귀 (stdlib)
py -S main.py selftest                                       # 무설치 재현(site-packages 차단)
```

## View Contract v2 — 스킬이 강제하는 불변식(R1~R11)

각 함수는 `fpna.view_contract` 에 있고 `(rep: QCReport, ...)` 로 QCReport.add 에 누적한다.
템플릿 `qc()` 가 이를 호출 → 미통과 시 `render` 가 **저장을 보류**한다.

| R | 함수 | 강제 내용 |
|---|------|---------|
| R1 | `assert_time_ruler` | 표시 시간축 = 캘린더 연속 ruler. 결측 기간도 행 유지 + NO_DATA |
| R2 | `full_outer` / `assert_full_population` | 조인 기본 FULL OUTER. match_status ∈ {MAPPED, LEFT_ONLY, RIGHT_ONLY} |
| R3 | `recon_block` / `assert_tie_out(tol=0)` | _RECON 블록 + 원천합 vs 출력합 차이=0 + completeness/accuracy/cutoff |
| R4 | `assert_filter_declared` | 필터는 명시+헤더 텍스트+제외건수 있을 때만 |
| R5 | `cross_tab` | 내부 tidy / 표시 wide 분리. sparse 는 행·열 생략 근거 아님 |
| R6 | `assert_no_forbidden_heuristic` | sample/head/top_n/preview 토큰 미요청 시 결함 |
| R7 | `gate` / `assert_no_silent_drop` | R1+R8+tie_out+no_silent_drop 묶음 게이트 |
| R8 | `assert_grain` | 행수 == distinct grain 조합 (중복·누락 0) |
| R9 | `assert_scenario_aligned` | 비교 시나리오 같은 grain·모집단. 한쪽만 = LEFT/RIGHT_ONLY 노출 |
| R10 | `assert_hierarchy_ties` | rollup 합 == leaf 합. orphan 진단 |
| R11 | `assert_master_to_gl` / `assert_allocation_conserves` | 마스터↔GL 대사, 배부 전·후 합 tie-out |

## lib 매핑 (불변식을 코드로 박은 곳)

| 사용자 산출물 "lib/" | repo 실제 위치 | 역할 |
|---|---|---|
| 차원 모델 | `fpna/dims.py` | 회계 캘린더(역월/4-4-5) + conformed dimension 6축 + `Fact`(grain) |
| View Contract | `fpna/view_contract.py` | R1~R11 assert (dbt-utils/GE/IAASB 어휘 차용, 미설치) |
| 재무 엔진 | `fpna/finance.py` | `depreciation_schedule`(정액법) 등 — QC 재계산 SSOT |
| 도메인 테이블 | `fpna/templates/fc_*.py` | `_MODULES` 레지스트리 등록 → render/QC/CLI 재사용 |

## 6 conformed dimension (`fpna.dims`)

1) Period(AccountingCalendar) · 2) Scenario(Actual/Budget/FC/PriorYear) ·
3) Account(parent-child + behavior) · 4) CostCenter(+분야 vehicle/property/utility/fixed_parts) ·
5) Asset/Vendor · 6) CostBehavior(fixed/variable/semi_fixed/committed).

## 새 도메인 테이블 추가

1. `fpna/templates/fc_<type>.py` 에 TYPE/INPUT/golden_sample/build/qc 노출(덕 타이핑).
2. `build` 는 `fpna.house_style` 만, `qc` 는 `fpna.finance` 재계산 + `fpna.view_contract` 게이트.
3. `fpna/templates/__init__.py` `_MODULES` 등록 + `fpna/dispatcher.py` 키워드 추가.
4. `tests/test_fc.py` 에 golden build+qc PASS 단언 추가. `py main.py render` + `py -S main.py selftest` 확인.

## 제약 (반드시 준수)

- openpyxl 외 런타임 라이브러리 금지(pandas/numpy/pydantic/dbt/GE/XlsxWriter). stdlib + dataclass 직접 구현.
- 레퍼런스(Kimball·dbt-utils·Great Expectations·IFRS16·FAST·Panko·Tidy Data)는 **어휘·검증 패턴만 차용**, 코드 미반입.
- 합성 재무수치 금지(golden_sample 의 구조 더미만 예외 — 의미 없는 더미임을 명시). **QC 통과 후에만 저장.**
- 회사→집: 구조 메타·에러 텍스트만. 집→회사: 코드/수식/런북 텍스트. 실데이터는 회사 PC 에서 실행.

## 2차 PR (후속, 범위 밖)

리스 스케줄(K-IFRS 1116/IFRS 16) · 약정/런레이트 · 계약·자산 마스터 테이블 · allocation/bridge 다대다.

## 관련 문서

- repo 운영 매뉴얼: `CLAUDE.md`
- 룩 SSOT: `fpna/house_style.py` · 재무계산: `fpna/finance.py`
- 폐쇄망 repo 루트 상시 적재용 요약: `skills/fpna-fixed-cost-tables/claude_snippet.md`
