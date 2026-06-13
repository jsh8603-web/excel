# Sources — 출처 기록

런타임 동봉(vendor)과 학습 참고(런타임 제외)를 구분한다. 학습 참고는 **개념/패턴만** 파이썬으로
재구현했고 **소스 복붙은 없다**. 영감 전용(저별점) repo는 코드·파일·시트 복사 금지, 아이디어만.

## 런타임 동봉 (vendor/ — 커밋, 회사가 이걸로 동작)

| 패키지 | 버전 | 라이선스 | 역할 |
|---|---|---|---|
| openpyxl | 3.1.5 | MIT (`vendor/LICENSES/openpyxl-LICENCE.rst`) | xlsx 읽기/쓰기 런타임 |
| et_xmlfile | 2.0.0 | MIT (`vendor/LICENSES/et_xmlfile-LICENCE.rst`) | openpyxl 의존(XML 스트리밍) |

순수 파이썬 검증: `vendor/`에 `.pyd/.so/.dll` 0개(`py fpna/_bootstrap.py`로 확인).

## 학습 참고 — 정형화(ingest) 알고리즘 (개념만 재구현)

| repo / 출처 | 라이선스 | 역할 | 배운 점 |
|---|---|---|---|
| nacnudus/unpivotr | MIT | 학습참고 | 시트를 셀좌표 평면테이블로 환원 후 behead/enhead 로 "방향상 최근접 헤더"를 데이터에 부착 = 다중헤더 언피벗 일반해 |
| nacnudus/spreadsheet-munging-strategies | (미명시, 추정 CC/MIT) | 학습참고 | data_type(numeric vs character)이 헤더/데이터 분리의 가장 robust한 1차 신호, 서식(이탤릭·들여쓰기·테두리)은 보조 |
| tidyr — Tidy Data (Wickham, JSS 2014) | CC | 학습참고 | messy 5유형 중 헤더가-값·변수가-행+열 이 우리 언피벗 타깃, fix = pivot_longer |
| USPTO 11341322 (Table detection in spreadsheets) | 특허(아이디어 재구현) | 영감 | 영역탐지 정량 기준: connected-component density≥0.7 ∧ ≥2col ∧ ≥5row |
| Pytheas (VLDB 2020) | (repo 미확인) | 학습참고 | "컬럼 값 coherency"로 데이터행 먼저 확정 후 헤더 역추적 (data/header/metadata/aggregate 분류) |
| ptype (arXiv 1911.10081) | 학술 | 학습참고 | missing/anomaly를 별도 생성과정으로 모델 → 센티넬이 타입 vote 오염 방지 |
| 타입추론 survey (arXiv 2411.11891) | 학술 | 학습참고 | per-type regex + majority vote(임계 초과 시 컬럼타입 확정) |
| SpreadsheetLLM (EMNLP 2024) | 학술(LLM부 제외) | 영감 | 인접 행/열 type-pattern 변화점=anchor=헤더/경계 후보(LLM 없이 차용) |
| MS "Inferring Units in Spreadsheets" (VL/HCC 2020) | 학술 | 영감 | 헤더 단위라벨 추론 + 단위 constraint propagation(이종단위 덧셈 차단) |
| 한국 사용자지정 서식(천원/백만 trailing comma) | 공개자료 | 학습참고 | `#,##0,`=천·`#,##0,,`=백만 = 표시 스케일(저장값은 실수치), `;` 음수섹션 `(...)` |

## 학습 참고 — 엑셀 룩 / 제로에러 패턴 (내 코드로 재현)

| repo / 출처 | 라이선스 | 역할 | 배운 점 |
|---|---|---|---|
| anthropics/skills (xlsx) | (미확인) | 학습참고 | openpyxl은 수식 평가 못함 → `data_only=True`는 마지막 저장 캐시만. 값+수식 동시수집 시 두 번 로드. 에러센티넬 `#REF!/#DIV0!/...` 스캔 |
| tfriedel/claude-office-skills | (미확인) | 학습참고 | 엑셀 룩/제로에러 패턴(서식·검증 게이트) 아이디어 |
| CFI / Macabacus / Vertex42 포맷 규약 | 공개 관찰 | 학습참고 | house_style 토큰화: 음수괄호·입력셀 파랑/계산셀 검정·gridlines off·합계 단일보더·항목좌/숫자우 |

## 영감 전용 (저별점 — 코드/파일/시트 복사 금지, 아이디어만)

FinPulse-FPA-Model, Financial-Planning-and-Analysis-Project-, financial-analyst-portfolio,
financial-dashboard-excel, DCF, Copilot-Excel-Finance — 템플릿 유형 아이디어만 참고.

## 학습 참고 — 다중시트 연동 / COA / 3-statement (pack 배선, 2026-06-13)

분류: [V]=런타임 vendor(내 파생 정적데이터) / [L]=학습참고(개념 재구현, 복붙 0) / [I]=영감 / [기각].

| ref | 분류 | repo 위치 | 취한 것 / 이유 |
|---|---|---|---|
| FAST Standard | [L] | `fpna/pack.py` Control(Index) 시트 | 색/부호 규약은 house_style 에 이미 있음 → 차용 0. Control 개념만. |
| Damodaran spreadsheets (stern, 무료·수정가능; ginzu 다탭) | [L] | `fpna/pack.py` 시트 연동 구조 | 다탭 연동 구조 학습. (옵션 `refdata/damodaran_defaults.json` = 집에서 받는 산업 디폴트, 미반입) |
| 3-statement 순환/리볼버/cash-sweep (WSP·ModelReef·IBA·FME·CFI) | [L] | `fpna/finance.py solve_revolver`(개념) | BS-항등·현금-tie 는 view_contract(tie_out/R11)·pnl_3statement linked 에 이미 있음. 순환 솔버 메커닉만 재구현. |
| 표준 COA — SEC US-GAAP XBRL Taxonomy(공개도메인) + ifrs-gaap.com(IFRS 구조/명칭) | [L] | `fpna/coa.py` | US-GAAP element 명칭=공개도메인 → tag 사용. IFRS=명칭 참조만(파일 벌크복사 0). |
| SEC EDGAR XBRL Financial Statement Data Sets(공개도메인) | [V] | `fpna/refdata/coa_us_gaap.json` | 표준 IS/BS/CFS 라인 골격(구조 메타만, ★재무수치 0). 내 파생물. 회사↔집 규율 적합. |
| PFRAM (IMF·World Bank, 공개 다중시트 연동) | [L] | `fpna/pack.py` 시트 배선/Control · `packs.md` | 시트 연동/Control 구조 학습(매크로 실행 X, 구조만). |
| Doubletalk (O'Reilly OO 복식부기) | [L] | `fpna/pack.py` `ledger_mode`(옵션, 기본 off) | 원장 posting→BS 구성상 balance 로 plug 회피하는 대안 아이디어. |
| ModelForge (Whatsonyourmind/modelforge) | [기각/0] | sources.md 기록만 | pydantic/SQLite/anthropic 금지스택 + view_contract/conserve 가 우월. 코드 차용 0. |
| Flevy / eFinancialModels 통합 PF·타당성 템플릿(유료) | [I] | — | 영감 전용, 차용 0. |
| (기존·불변) unpivotr/tidyr/Pytheas/ptype/SpreadsheetLLM → ingest, anthropics-skills/CFI/Macabacus/Vertex42 → house_style | [기존] | 위 표 참조 | 이미 기록·구현 완료. 재작업 금지. |

## 비고

- 학습참고 repo의 실제 clone은 `_sources/`(`.gitignore`, 커밋 금지)에서 수행하며 repo에 포함하지 않는다.
- 라이선스 "미확인" 항목은 개념/패턴 차용 수준이며 코드·텍스트 복제는 하지 않음.
- 리서치 raw 보존(홈): `~/.claude/docs/archive/research-raw/unpivotr-tidy-ingest-native-20260609.txt`,
  `~/.claude/docs/archive/research-raw/messy-excel-tidy-ingest-native-20260609.txt`.
