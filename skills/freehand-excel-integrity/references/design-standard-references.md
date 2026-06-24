# Excel 디자인 표준 — 레퍼런스 & 채택 규칙

산출물의 *디자인* 품질(정렬·폰트 위계·라벨/주석 규율·색상 역할)을 권위 표준에 근거해
규정한다. 정확성(값/수식)은 게이트 [1]~[12] 가, 디자인은 [14] 가 자문으로 점검한다.

## 1. 레퍼런스 (권위 순)

**FAST Standard** — fast-standard.org. Flexible/Appropriate/Structured/Transparent.
ICAEW 가 최초로 공식 인정한 금융모델링 표준. 핵심 규칙:
- 일관된 열 구조. 열 A/B/C 는 제목 계층(섹션/하위/하위하위)으로 예약, 한 열=한 목적.
- 시간은 좌→우, 계산은 위→아래. 타임라인 열의 수식은 전부 동일(검토 용이).
- "positive as normal" 부호규약. 라벨 열의 가감 표시는 우측정렬.
- 단순성: 짧은 수식, 제한된 함수, 불필요한 detail/허위 정밀(spurious precision) 금지.

**ICAEW — Twenty Principles for Good Spreadsheet Practice** (4th ed.). 핵심:
- 입력→처리→출력의 명확한 흐름. 입력은 *한 번만*, 가능한 한 *같은 구역*에.
- 한 시트에 서로 다른 수식 종류는 최소화하고 그룹을 명확히 분리.
- 조직 표준 서식을 cell styles/themes 로 일관 적용. 감사 대상이면 체크/알림 내장.
- **청중을 식별하고 그에 맞게 설계.** 비-모델러가 볼 거면 문서/주석을 구조적으로.

**Macabacus 서식 규약** (help.macabacus.com) — IB/FP&A 사실상 표준:
- 폰트 색 역할: 입력=파랑, 하드코드/상수=흑/회(색 안 씀), 동일시트 수식=흑,
  타시트·타북 링크=초록, 외부데이터 함수=별색. 색은 *의미*에만 쓰고 남용 금지.
- 숫자는 우측정렬, % 는 이탤릭, 통화 형식 일관(과다 number format 은 Excel 한계).
- 섹션 헤더/탭은 *옅은 회색* 음영으로 배경에 가라앉히고, 카테고리 헤더는 파랑.
- sum bar(합계 위 테두리), leader dots(라벨↔숫자 정렬 보조). 제목 텍스트는 title case.

## 2. 채택 표준 (이 repo)

| 영역 | 규칙 | 근거 |
|---|---|---|
| 정렬 | 숫자=우측, 라벨=좌측, 가감 인디케이터=우측. 숫자에 left/center 금지 | FAST, Macabacus |
| 라벨 | 간결하게. **별표·마크다운 강조·장식선(***/■) 금지.** 제목은 군더더기 없이 | FAST 단순성 |
| 메인 제목 | 간결한 제목 1개. (조직 관례에 따라 영문 우선 가능) title case | Macabacus |
| 폰트 위계 | 제목 ≤18pt, 헤더 ≤14pt, 본문 ~10–11pt. 시트당 폰트 크기 ≤~5종 | ICAEW cell styles |
| 색상 역할 | 입력 파랑·수식 흑·링크 초록·하드코드 흑회. 색=의미 전용 | Macabacus |
| 주석/설명 | 헤더·데이터 근처에 장문 금지. 전용 cover/notes 시트나 셀 주석으로 분리 | ICAEW 구조 |
| 구조 | 입력 1회·같은 구역. 입력→계산→출력 흐름. 시간 좌→우 | FAST, ICAEW |
| 헤더 영역 | 표 헤더는 옅은 회색 음영, 과한 색/볼드 자제 | Macabacus |

## 3. 린터 매핑 (게이트 [14], 자문·보수적)

| 불만(사용자) | [14] 검출 | 표준 |
|---|---|---|
| 별표 많은 제목, fpna 스럽지 않음 | 장식문자(별표/강조) | FAST 단순성 |
| 헤더 폰트 너무 큼 | 폰트 ≥18pt 과대 + 크기 종수 과다 | ICAEW cell styles |
| 열 가운데/왼쪽 정렬 | 숫자 셀 left/center 정렬 | FAST/Macabacus |
| 설명 헤더 근처 과다 | 상단 5행 장문(>150자) | ICAEW 구조 |
| 설명 첫장 몰아쓰기 | (상단 장문 누적) → 전용 notes 권장 | ICAEW |

[14] 는 **자문**(치명 아님) 이며 보수적이라 실데이터(평평한 표·정상 리포트)에는 발화하지
않는다(검증 완료). 색상 역할은 house_style 의 color role 로 생성 단계에서 적용한다.

## 4. 원문 링크

- FAST Standard: https://fast-standard.org/the-fast-standard/ , 자료실 https://fast-standard.org/fast-resources/
- ICAEW 20 Principles: https://www.icaew.com/technical/technology/excel-community/20-principles-for-good-spreadsheet-practice-2024-edition
- Macabacus Colors/Numbers: https://help.macabacus.com/article/318-colors , https://help.macabacus.com/article/281-number-formats
