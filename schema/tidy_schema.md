# tidy 출력 스키마 (schema/tidy_schema.md)

`fpna.ingest`가 산출하는 `tidy.csv` / `schema.json` 의 계약. 인코딩 = **utf-8-sig**(한글 Excel 호환).

## tidy.csv 컬럼

| 컬럼 | 타입 | 설명 |
|---|---|---|
| `entity` | TEXT | 행 주체(가장 바깥 행헤더). 예: 제품A, 영업부, 매출채권 |
| `period` | TEXT | 기간(period 패턴 헤더). 예: 2025, 2025-06, Q1 |
| `metric` | TEXT | 지표. 열헤더 + 안쪽 행헤더를 ` > ` 로 결합 |
| `value` | NUM\|빈칸 | 값. 센티넬은 빈칸(None). **원값 보존**(단위 미곱) |
| `unit` | TEXT | 단위 라벨(천원/백만/₩mn 등) |
| `row_role` | TEXT | `data` \| `subtotal` \| `total` \| `header` |
| `level` | NUM | 들여쓰기 계층(0=최상위) |
| `src_row` | NUM | 원본 셀 행 |
| `src_col` | NUM | 원본 셀 열 |

## schema.json 구조

```json
{
  "columns": { "entity": "TEXT", "period": "TEXT", "metric": "TEXT",
               "value": "NUM|MIXED|...", "unit": "TEXT", "row_role": "TEXT",
               "level": "NUM", "src_row": "NUM", "src_col": "NUM" },
  "blocks": [ { "sheet": "...", "block": 0, "range": "R4C1:R9C5",
                "unit": "천원", "n_long_rows": 16 } ],
  "n_rows": 18,
  "n_rejected": 0,
  "generated_by": "fpna.ingest.pipeline"
}
```

## row_role 활용 (이중집계 방지)

- 집계/검증 시 `row_role == "data"` 만 합산. `subtotal/total` 은 **교차검증용**으로만 사용
  (자식 합 == 소계 인지 확인). 그냥 다 더하면 이중계상.

## 값 해석 주의

- `value` 는 셀 저장 원값. number_format 의 천/백만 축약은 표시일 뿐 곱하지 않았으므로,
  `unit` 컬럼을 보고 소비측에서 스케일 해석한다.
- `(50)`→ -50, `△30`→ -30, `85%`→ 0.85, 센티넬(`-`,`…`,`N/A`)→ 빈칸.
