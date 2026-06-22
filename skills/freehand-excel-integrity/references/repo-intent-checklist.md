# Repo 핵심 의도 체크리스트 (rewrite self-check)

하드닝된 FP&A 파이프라인(`jsh8603-web/excel`)의 핵심 의도를, 이 스킬로 프리핸드 재작성할 때
대조하는 체크리스트. 각 항목: 스킬에서의 강제 수단과 커버리지 등급.

등급: ✅ 자동강제(게이트) · ◐ 작성 규칙(규율) · △ 부분/조건부 · ✗ 스킬 범위 밖(원 파이프라인 책임)

## A. 산출 정합성 (View Contract / conserve) — 재작성에서 가장 중요

- [ ] **R3 tie-out**: 모든 합계 = Σ parts. → `ties` ✅ ([7] 커버리지가 미선언 합계 경고)
- [ ] **R8 grain**: 1행=1항목, 중복 키 없음. → `grain` ✅
- [ ] **R10 계층 정합**: parent = Σ children(소계/총계). → `ties` ✅
- [ ] **R11 master↔GL / 배부 보존**(pre=post). → `ties` ✅
- [ ] **R14 약정 보존**(Σ인식+Σ잔여+취소=계약총액). → `ties` ✅
- [ ] **R17 비율 완전성**: 결측/0분모 → "NA", #VALUE! 금지. → `ratios` + [4] ✅
- [ ] **R9 시나리오 정합**: Actual/Budget 같은 모집단, 한쪽 키 0처리 금지. → `scenario` ✅
- [ ] **R7 no_silent_drop**: 입력 행수 == 출력 행수. → `expected_n` ✅
- [ ] **R2 모집단 전수성**: 한쪽만 키도 행으로 노출(LEFT/RIGHT_ONLY). → `scenario` + 규칙5 ◐
- [ ] **R1 시간축 전수성**: 결측 기간도 NO_DATA 행/열. 캘린더에서 생성. → 규칙10 ◐
- [ ] **R12 재발성**(활성 window 결측 노출). → 규칙11 ◐ (계약 미구현 — 필요시 확장)
- [ ] **R13 부호/오분류**: 부호 flip·고정/변동 오분류. → `fields.sign` △ + [5] 자문
- [ ] **conserve N-version 독립성**: ⚠ 계약 tie 는 집계-class(같은 합 비교). 부호·연산자
      버그까지 잡으려면 *변환식 재유도*가 필요 — 스킬 미구현, 한계로 인지. △

## B. 입력/필드 계약 (binding / metric_table FieldSpec)

- [ ] **grain_unique** (pre-shape 중복). → `grain` ✅
- [ ] **REQUIRED/결측 정직 노출**(채우지 말고 NO_DATA). → 규칙12 + [2] ◐
- [ ] **FieldSpec 범위/부호/허용값**. → `fields` ✅
- [ ] **숫자 강제**(콤마/통화/괄호음수 → 숫자, 텍스트숫자 금지). → 규칙1 + [2] ✅
- [ ] **UNIT_POLICY/원값 보존**(단위는 헤더+서식, 값 변형 금지). → 규칙6 ◐

## C. 콘텐츠 타입 (W26 버그 직격)

- [ ] **숫자영역에 텍스트 금지**(주석 누수). → `numeric_regions` + [2] ✅
- [ ] **헤더 토큰 본문 누수 금지**. → `header_rows` + [2] ✅
- [ ] **주석/명세는 comment·notes 열로**(값 셀 금지). → 규칙1 + [2] ✅

## D. 수식 무결성

- [ ] **에러셀 0**(#REF!/#VALUE!/#DIV0). → [1] ✅ (단 openpyxl 미재계산 — Excel COM 재계산 권장)
- [ ] **수식 칼럼/방향 참조**(=실적-계획 등). → `formula_refs` ✅
- [ ] **가드된 비율**. → `ratios` + [4] ✅

## E. 디자인 표준 (house_style / design_standard) — "양식 불균형" 직격

- [ ] **6단 골격**(제목·메타·본문·_RECON·출처·page_setup). → 규칙13 ◐
- [ ] **색 역할**(입력=파랑/계산=검정/링크=초록, 양=초록/음=빨강). → 규칙13 ◐
- [ ] **숫자서식 SSOT**(괄호음수, INT/PCT/MULT/백만`,,`). → 규칙6·7 + [3] △
- [ ] **합계 보더**(top strong / 최종 single+double). → 규칙7·13 ◐
- [ ] **단일 액센트·확정 폰트·gridlines off·freeze header**. → 규칙13 ◐
- ⚠ 한계: [3] 은 *파일 내부* 서식 일관성만 본다(열 다수서식 기준). repo 의 house_style 정확한
  색/폰트/스킬레톤을 강제하진 않는다 — 규칙(◐)으로만 포팅. 원 룩 100% 일치가 필요하면 원
  파이프라인의 `house_style.py` 로 렌더해야 한다.

## F. 게이트 / 규율

- [ ] **fail-closed**(QC 통과 후에만 확정). → PostToolUse 훅(exit2) ✅
- [ ] **합성 재무수치 금지 / PIT**. → 규칙12 ◐
- [ ] **회사↔집 반출 규율**(실데이터·거래처명·원본 반출 금지, 구조/코드/수식만). → ✗
      스킬 범위 밖. 작업 환경 정책으로 별도 관리 필요(이 스킬은 산출 정합성만 다룸).
- [ ] **row_role data-only 합산**(소계/총계 재합산 금지, 이중계상). → 규칙8 ◐

## G. 파이프라인 단계 (스킬 범위 밖 — 의식적으로 제외)

- ✗ **stage 라우팅**(ingest/profile/transport/analysis 분류) — 원 파이프라인 책임.
- ✗ **ingest 정형화**(누더기→tidy 8단계, value-like 판정) — 원 파이프라인 책임.
- ✗ **Kimball fact+conformed dimension star schema** — 이 스킬은 *출력 정합성* 게이트지
  데이터모델 강제기가 아니다. 다중 테이블 star schema 가 필요하면 원 `fpna.dims` 사용.
- ✗ **receipt/스파인 우회 차단**(빌드시점) — 프리핸드엔 스파인이 없으므로 PostToolUse 훅이
  그 역할을 *근사*한다(동등하진 않음).

---

## 재작성 완료 판정

다음이 모두 참이면 재작성을 완료로 본다:
1. `python scripts/xlsx_doctor.py <파일>` 에서 **[1][2] 0건, [6] FAIL 0건**.
2. **[7] 커버리지 완전**(모든 합계·비율이 ties/ratios 로 선언됨).
3. 위 A·C·D 의 ✅ 항목이 계약에 빠짐없이 선언됨.
4. (가능시) Excel COM 재계산으로 에러셀 0 재확인.
5. E·F 의 ◐ 항목은 작성 규칙으로 적용됨(자동검증은 아니므로 작성자가 의식적으로 확인).
