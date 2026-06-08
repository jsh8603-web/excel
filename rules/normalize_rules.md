# 정형화 규칙 (normalize_rules.md)

누더기 엑셀 → tidy long 변환의 **결정적 규칙**. 코드(`fpna/ingest`)가 SSOT이며 이 문서는 설계 근거다.
규칙은 코드로 고정되어 있어 같은 입력 → 같은 출력(재현 가능). 갱신본 재실행도 결정적.

> 설계는 unpivotr/tidy 이론 + 표 탐지 연구 개념을 **파이썬으로 재구현**한 것이다(소스 복붙 없음, [sources.md](../sources.md)).

## 산출 스키마 (tidy.csv)

| 컬럼 | 의미 |
|---|---|
| `entity` | 행 주체(제품/부서/계정 등, 가장 바깥 행헤더) |
| `period` | 기간(연/월/분기 — period 패턴 헤더에서) |
| `metric` | 지표(매출/비용 등, 열헤더·안쪽 행헤더 결합 `>`) |
| `value` | 값(숫자/None). **원값 보존** — 단위 스케일은 곱하지 않음 |
| `unit` | 단위 라벨(천원/백만 등). `(단위: …)` 또는 number_format 에서 |
| `row_role` | `data` / `subtotal` / `total` / `header` |
| `level` | 들여쓰기 계층(계정과목 트리 깊이) |
| `src_row`, `src_col` | 원본 셀 좌표(추적성) |

## 8단계 (코드 매핑)

### 1. 셀 평면화 — `cells.as_cells`
- `data_only=False`(수식)와 `data_only=True`(캐시값)로 **두 번 로드** → 값+수식 동시 수집.
- ⚠ openpyxl 한계: `data_only` 캐시는 "Excel이 마지막 저장한 값". 미개봉 생성파일은 None일 수 있음.
- 셀 1개 = `Cell`(row,col,value,formula,data_type,is_blank,fmt,merge_anchor). 모든 후속 단계가 이 평면 위에서 동작(unpivotr 통찰).

### 2. 블록 탐지 — `detect.detect_blocks`
- 채워진 셀 → union-find **connected-component**(8방향). bbox 의 **density ≥ 0.55** ∧ `n_cols≥2` ∧ `n_rows≥3` 게이트.
- `GAP_TOL=0`(직접 인접만) — 표 사이 1칸 구분 행/열을 경계로 인식해 **다중 표 분리** 우선. 표 내부 빈 구분열로 과분할되면 `gap=1`(단, 인접 표 병합 위험).
- 근거: USPTO 11341322 / Pytheas coherency(density·최소크기 정량 기준).

### 2.5 비데이터 행 격리 — `detect.strip_title_rows` / `strip_footnote_rows`
- **제목/단위 행**: 블록 폭 상대 기준 — `n_cols≥4`인 넓은 표에서 채움 셀 ≤2 이고 값처럼 보이는 셀이 없는 상단 행. (좁은 2~3열 표는 헤더 오인 방지를 위해 보호.)
- **단위 라벨**: `(단위: 천원)` 정규식 스캔(행 제거 여부 무관). 시트 단위 폴백도 둠.
- **각주**: 하단의 `주1)/※/*` 시작 또는 1셀 character 행.

### 3. 병합 fill + 헤더 해소 — `headers.unmerge_fill` / `classify_cells`
- `unmerge_fill`: 병합 anchor 값을 영역 내 빈 셀에 전파 → 헤더 방향매칭이 직상/직좌만으로 해소(2방향 단순화).
- **value-like 판정**(핵심): 실제 숫자/날짜 = 값. 텍스트라도 괄호/△▲/%/콤마/소수점 **마커 동반**이면 값(`(50)`,`△30`,`85%`,`1,234`). **bare 정수("2024")는 비값**(연도/헤더 가능성) → 연도 헤더와 데이터 숫자를 구분.
- 2-pass 분류: ① 위에서부터 value-like 비율 ≥0.5 행 직전까지 = top 헤더밴드 ② 데이터행 기준 좌측 비값 열 = left 헤더밴드 ③ 데이터영역만으로 열 타입 재집계.

### 4. behead 언피벗 — `headers.unpivot_block`
- 헤더 셀을 데이터 셀에 부착하고 그 헤더를 제거(반복). 다중레벨 = 바깥→안.
- **방향 매칭**(`nearest_header`): `up`(동일열 직상 최근접), `left`(동일행 직좌), `up-left`(스팬 열헤더), `left-up`(스팬 행헤더).
- ⚠ up-left/left-up 정렬 우선순위는 리서치 일부 추론분 — 골든샘플 테스트로 검증함.

### 5. 센티넬·단위·스케일 정규화 — `normalize.normalize_value`
- 센티넬(`…/-/–/N/A/없음/*` 등) → `value=None`.
- 음수표기: 괄호 `(1,234)`→-1234, 선행 `△▲−`→음수, 후행 `(-)`→음수.
- `%` 텍스트 → 비율(0.85). 콤마 제거.
- **스케일 주의**: number_format trailing comma(`#,##0,`=천 표시)는 **표시 축약**일 뿐 저장값은 이미 실수치 → **곱하지 않음**(이중계상 방지). 단위는 `unit` 컬럼으로만 보존.
- 타입 추론(`infer_column_type`): 센티넬을 vote에서 **제외**(ptype 정수) → `…/-` 가 숫자열을 텍스트로 오판 안 함.

### 6. 검증 — `validate.validate_rows`
- dataclass `TidyRow` + 수동 제약(entity/metric 결측, period 필수 옵션, value 에러리터럴). 위반 행 reject + 리포트.

### 7. 수식 스멜 스캔 — `validate.scan_formula_smells`
- 정규식으로: **하드코딩 상수**(셀참조 아닌 매직넘버), **과도 중첩**(괄호 깊이>4), **시트간 참조 과다**(>3), **외부파일 참조**(`[book.xlsx]` — 경로 마스킹 권장).

### 8. 산출
- `tidy.csv`(**utf-8-sig** BOM — 한글 Excel 더블클릭 호환), `schema.json`(컬럼/블록/타입), `smell_report.md`.

## 알려진 한계 / 향후 보강 후보 (리서치 갭)

코드에 미구현이거나 단순화한 항목(실 데이터로 검증하며 점진 보강):
- transpose 자동 감지(period가 행축인 전치 표)
- 반복 헤더(페이지 분할표 중간 재출력 헤더) 제거
- centre-aligned 스팬 헤더의 justify 재배치
- 혼합 타입 열의 열 단위 다수결(부분 구현 — value-like 기반)
- 단위 constraint propagation(이종 단위 덧셈 차단)

> 보강 시 골든/픽스처 회귀(`py main.py selftest`) 통과 + 실데이터 1건 검증 후 규칙 코드화.
