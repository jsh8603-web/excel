---
title: FP&A 산출물 디자인 표준 (house_style 레이아웃 규칙)
tags: [fpna, house-style, design-standard, layout, view-contract, qc]
status: stable
ssot: fpna/house_style.py
related: [rules/normalize_rules.md, fpna/view_contract.py, fpna/templates/base.py]
---

# 디자인 표준 (design_standard.md)

FP&A 산출물(.xlsx)의 **룩·레이아웃 규칙**을 문서화한다. 코드 SSOT 는
`fpna/house_style.py` 이며, 이 문서는 그 설계 근거와 적용 규약이다. 빌더
(`fpna/templates/*.py`)는 색·폰트·서식을 직접 하드코딩하지 않고 house_style
헬퍼만 사용한다(룩 변경은 house_style 한 곳).

## 1. 표준 레이아웃 골격 (6단)

모든 템플릿은 다음 골격으로 통일한다(일관성 = 신뢰성).

| 단 | 내용 | 헬퍼 |
|---|---|---|
| ① 제목 블록 | title / subtitle | `report_frame` (title_block) |
| ② 메타 헤더 | 단위 · 통화 · 회계기준 · 기준일 | `report_frame` (meta_header) |
| ③ 본문 | 표 / 차트 | 빌더가 직접 |
| ④ _RECON 대사 | completeness / accuracy / cutoff | `view_contract.recon_block` |
| ⑤ 출처 footer | Source / Note / Prepared by | `report_footer` (source_footer) |
| ⑥ page_setup | 반복 헤더 · 가로폭 맞춤 · 페이지/날짜 푸터 | `report_footer` (page_setup_report) |

빌더는 `frame → 본문 → footer` 순서만 지킨다. gridlines off · freeze_panes(헤더
고정)는 `report_frame` 이 일괄 적용한다.

## 2. 색 규약 (IB 4색 + 무채색 본문)

- 입력셀 = 파랑 글씨(`INPUT_FG`), 계산셀 = 검정(`CALC_FG`), 시트간 링크 =
  초록(`LINK_FG`), 외부 참조 = 빨강. `set_cell(role=...)` 으로 표현한다.
- 본문은 무채색 + 단일 액센트 1색(`ACCENT` 네이비블루). 폰트는 회사 PC 확정
  보유분만(맑은 고딕 / Calibri).
- 양(개선) = 초록(`POS_FG`), 음(악화) = 빨강(`NEG_FG`).

## 3. 숫자 서식 (SSOT 상수)

- 음수 = 괄호 `#,##0;(#,##0)`(`FMT_INT`). 빨강 음수는 선택(`FMT_INT_RED`).
- 백만 스케일은 원값을 두고 표시만 ÷1e6(`FMT_INT_MN` = `#,##0,,`). 헤더에 단위 병기.
- 비율 = `0.0%`(`FMT_PCT1`), 배수 = `0.0"x"`(`FMT_MULT`).
- KPI 신호 화살표 = `[Green]"▲"0.0%;[Red]"▼"0.0%`(`FMT_KPI_ARROW`).

## 4. 합계·tie-out 시각화

- 합계 행은 상단 단일 보더(`BORDER_TOP_STRONG`). 최종 합계는 위 single·아래
  double(`BORDER_TOTAL`).
- tie-out 체크 셀은 `check_cell` 로 그린다. 값이 허용오차 `tol` 을 벗어나면
  조건부 서식(CellIsRule)으로 적색 강조 — BS 균형·브리지 합·소스 일치 등.

## 5. 조건부 서식 (openpyxl native)

- 히트맵 = `apply_heatmap`(ColorScaleRule), 데이터바 = `apply_databar`,
  아이콘셋 = `apply_iconset`. 짝수행 음영밴딩 = `apply_zebra`(FormulaRule
  `MOD(ROW(),2)=0` — 테이블스타일 대신 banding 으로 차트 충돌 회피).
- native 미지원 항목은 대안으로 처리한다. 스파크라인 = `REPT("█",x)` 의사
  스파크라인. What-If Data Table = 파이썬 선계산 후 값 기입.

## 6. 차트 규약

- 워터폴(브리지) = 누적막대 + base 시리즈 투명(`add_waterfall`, fill=none 트릭).
- 추이/포캐스트 = 라인(`add_line_chart`), 규모 비교·토네이도 = 막대
  (`add_bar_chart`). 차트 데이터 영역은 1행 위에 시리즈명을 둔다.
- 채우기 모드(fill)에서 openpyxl 라운드트립은 차트를 떨굴 수 있으므로 차트는
  코드로 재생성한다(베이스는 표·서식 위주).

## 7. 입력 안전장치 (selector / 가정)

- 시나리오 selector 등은 `DataValidation`(드롭다운)으로 입력 도메인을 강제하고,
  CHOOSE/INDEX 로 케이스를 전환한다(정적 컬럼 나열 금지 — 동적 스위치).
- 가정 셀은 입력 역할(파랑)로 분리하고, 산출은 계산 역할(검정 수식)로 둔다.

## 8. View Contract 게이트 연계

레이아웃과 별개로, 산출의 정직성은 `fpna/view_contract.py` 의 불변식(R1~R17)을
템플릿 `qc()` 에서 호출해 강제한다(상세는 view_contract 모듈 docstring).

- 시계열 = R1 시간축 전수성(`assert_time_ruler`) — 침묵 갭 차단.
- 계층 합 = R10 roll-up tie(`assert_hierarchy_ties` / `assert_tie_out`) — 누락 차단.
- 시나리오·민감도 = R9 base 정합(`assert_scenario_aligned` 정신) — 거짓 헤지 차단.
- 분개·브리지 = R3 tie-out(`assert_tie_out`) — 누락·중복 차단.

QC 미통과 산출물은 `render` 게이트가 저장을 보류한다(부정직 산출 차단).

## 6. 디자인 *검사* 계층 (design_audit.py) — 2026-06 추가

생성 측(house_style)이 표준을 *적용*한다면, `fpna/design_audit.py` 는 그 표준을 *어겼는지*
잡는다(freehand·외부입수 대비). house_style 토큰(ALLOWED_SIZES/SIZE_TITLE/역할색)을 직접
읽어 SSOT 와 정합 — 표준대로 생성된 산출엔 침묵(오탐 0).

- `assert_design_standard(rep, wb)` — run_report 스파인(`_base_owned_gate`)에 연결. 장식문자
  (별표/마크다운 강조)는 hard-fail, 숫자 좌/가운데 정렬·비표준/과대 폰트·헤더근처 장문은 보고.
- `edit_cell(ws, ref, value)` (house_style) — 세션 중 수정은 직접 `ws[ref]=v` 대신 이걸로.
  기존 role/서식/정렬을 유지해 **편집 드리프트**를 차단한다.
- `restyle_inplace(wb)` / `tools/restyle.py` — **외부 입수** 파일을 표준으로 *비파괴* 정규화
  (서식만, 숫자·수식 값 불변을 저장 전 단언). 외부 파일엔 golden diff 를 쓰지 않는다.

근거 표준(FAST/ICAEW/Macabacus)은 freehand-excel-integrity 스킬의
`references/design-standard-references.md` 참조.
