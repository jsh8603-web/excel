# Excel backend routing — 신호 기반 라우팅 + 규약 게이트

`.xlsx` 를 만들거나 고치는 **쓰기 백엔드는 3종**: openpyxl / xlsxwriter / xlwings.
pywin32 **COM 은 독립 백엔드가 아니라 xlwings 의 탈출구**(`xlwings.api`) — 피벗·슬라이서
같은 Excel 고유기능이 필요할 때 xlwings 가 내부에서 COM 으로 강하한다. 라우팅은 퍼센트가
아니라 **신호 결정함수**(`route()` in `scripts/verify_workbook.py`, `--route` 로 확인)로 정한다.

> 일반 가이드다. `jsh8603-web/excel` 런타임(`fpna/`,`main.py`)은 openpyxl 전용·pandas 미사용이
> 하드 제약이라 거기선 openpyxl 로 수렴한다. 다른 프로젝트(예: `fixed_trans`)에선 3종 다 선택지다.

## 라우팅 결정함수 (신호 → 백엔드)

```
excel_feature(피벗/슬라이서)?           → xlwings (필요 시 COM 강하)
elif 기존 파일 편집:  열림(라이브) → xlwings  /  닫힘 → openpyxl
else (신규):          대량 AND 차트 → xlsxwriter  /  그 외 → openpyxl
```

## 작성 기준(basis) — 정적 vs 수식, 그리고 폴백

모든 작성·수정은 워크북 mutation 이고, 옳으려면 **작성 기준**이 있어야 한다. 기준이 없으면
다른 방법으로 mutation 을 만든다(폴백). 먼저 **정적/수식**을 가른다 — 게이트가 갈린다:

| 기준 종류 | 작성 방식 | 계약(contract) 요구 | 재계산 |
|---|---|---|---|
| **정적 값** (코드계산 주입) | 셀에 숫자 직접 기입 | `ties[].expected`(소스 독립총계) + **값모드 `ratios`**(`{cell,num,den}`) 필수 | **불요** — doctor 가 오프라인 완결 |
| **수식 구동** | `=SUM(...)`·`=A-B` 작성 후 fill-down | `ties`(SUM 이 parts 커버) | **필요** — 재계산 체인 + fill-down 린터 |

| # | 백엔드 (mutation) | 작성 기준 | 기준 부재 시 폴백 |
|---|---|---|---|
| 1 | **openpyxl** (신규/닫힌 편집) | 프리핸드=`*.contract.json`, 정형=템플릿 `qc()` | 템플릿 없음→프리핸드+계약 / 깨진 파일→원본서 의도 복원 |
| 2 | **xlsxwriter** (신규 대량+차트) | `expected_n` + `ties` | 행수·합계 기준 없음→source 에서 뽑아 contract emit |
| 3 | **xlwings** (열린 라이브 편집·Excel 고유기능) | 기존 구조 + 편집셀 계약 + 부품 보존 기준 | 기준문서 없음→`roundtrip_gate.fingerprint(원본)` 로 자동 추출. Excel 고유기능→`xlwings.api`(COM) |

## 규약 게이트 (verify_workbook.py — no free-pass)

1. **계약 커버리지** — `xlsx_doctor` green **+ `[7]` 완전**(모든 합계·비율이 ties/ratios 선언).
   자문성 `[7]` 을 fatal 로 격상 — "계약 존재"만으론 부족, 전부 선언돼야 통과.
2. **roundtrip (편집 경로 전용)** — 기존 파일 편집은 `--before` 스냅샷 필수, `roundtrip_gate` 가
   부품(외부링크/차트/피벗/Table) 소실 0 증명. **신규 생성은 면제.**
3. **재계산 폴백 체인** (수식 워크북만):
   `pywin32 → verify_xlsx` / 없고 `LibreOffice → soffice --headless` / 없고 `formulas 패키지`(커버리지
   한계 경고) / 정적 워크북 → skip(doctor 오프라인 검증).
4. **수식 일관성 린터** — `formula_lint` 가 같은 열 수식의 행별 상대참조 형태 불일치를 적발
   (`=An-Bn` 사이의 `=A4-B2` 같은 fill-down 파손).

## 다운그레이드 투명성 (침묵 폴백 금지)

환경이 백엔드를 강등하면(예: 무설치 PC 라 xlwings→openpyxl, 피벗 불가) 게이트가
`DOWNGRADE: <요청기능> → <대체결과>` 로 명시한다. `resolve()` 가 요청 기능 capability 와 실효
백엔드를 대조해 잃은 기능을 출력 — 조용히 넘어가지 않는다.

## 환경별 제외 규칙

| 환경 | 제외 | 이유 | 대안 |
|---|---|---|---|
| 무설치 PC / Linux CI | xlwings(+COM 강하) | Excel·pywin32 필요 | openpyxl, `tools/format-ooxml.py` |
| `jsh8603-web/excel` 런타임 | pandas(xlsxwriter·xlwings), COM | `-S` 무설치 + vendor 동봉 | openpyxl. `tests/test_runtime_purity.py` 강제(이 repo 전용) |
| 외부링크/Power Query 워크북 | openpyxl `save`, `format-in-place` | 저장 시 부품 떨굼 | `format-ooxml`(부품 바이트 복사) 또는 xlwings |

## portability — `jsh8603-web/excel` repo 밖에서 쓸 때 (예: fixed_trans)

**스킬과 함께 이동(자립, 어디서나):** `scripts/` 의 검증 6종 —
`verify_workbook.py`(라우터+게이트+preflight), `xlsx_doctor.py`(값·계약), `roundtrip_gate.py`(부품
보존), `recalc_check.py`(재계산 체인), `formula_lint.py`(fill-down), `verify_xlsx.py`(COM 재계산,
Excel 있을 때만). 전부 stdlib+openpyxl, `formulas`/pywin32/soffice 는 '있으면 쓰는' 선택 엔진.

**repo 전용(안 따라감):** `tools/excel_doctor.py`·`content-gate.py`(`fpna.*` import), 런타임 차단
규칙 + `tests/test_runtime_purity.py`(`fpna/` 전용 — fixed_trans 엔 미적용).

**과발동 주의:** Authoring Rules 는 openpyxl 갈래 기준. "무조건 openpyxl"이 아니라 "라우터 + 게이트"로
읽을 것 — xlwings/xlsxwriter 작업은 그 백엔드로 쓰되 위 게이트는 그대로 통과시킨다.
