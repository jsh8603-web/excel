# CLAUDE.md 스니펫 — 고정비 테이블 작업 시 상시 적재

> 폐쇄망 repo 루트 `CLAUDE.md` 에 붙여, 회사 Claude 가 고정비 테이블을 만들 때
> 데이터 엔지니어링 기본값으로 빠지지 않게 계약을 상시 환기한다.

## 고정비 테이블 셋 — 절대 규율 (fpna-fixed-cost-tables)

- **테이블 = fact + conformed dimension** (Kimball). 평면 더미 금지. 먼저 **grain 선언**(`fpna.dims.Fact`) — 미선언 빌드 실패.
- **시간축은 `fpna.dims.AccountingCalendar` 에서만** 생성(역월/4-4-5). 데이터에서 날짜 뽑아 열 만들기 금지. 결측 기간도 행 유지 + `NO_DATA`(R1).
- **조인 기본 FULL OUTER**(`view_contract.full_outer`) — match_status 부여, anti-join only 금지(R2). **변동은 시나리오 축 차이로만**, 한쪽만 있는 cost center 0 처리 금지(R9).
- **모든 시트에 `_RECON` 블록**(`view_contract.recon_block`) + 원천합 vs 출력합 차이=0(R3, tol=0). 계약·자산 ↔ GL 대사, 배부 전·후 합 tie-out(R11).
- ⛔ **sample/head/top_n/preview·"대표 1일"·예외-only·임의 합산 축소 금지**(R6, Default-Deny). 필터는 사용자 명시 + 헤더 선언 + 제외건수 표기 시에만(R4).
- 생성: `py main.py render fc_depreciation_schedule|fc_variance_bridge out\x.xlsx` → **QC 게이트(R1~R11) 통과 후에만 저장**. 회귀: `py -S main.py selftest` + `py -m unittest tests.test_fc`.
- openpyxl 외 런타임 라이브러리 금지(pandas/numpy/dbt/GE). 합성 재무수치 금지(구조 골든만 예외). 실데이터 숫자 반출 금지.
