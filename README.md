# FP&A Excel 시스템

사내 재무계획·분석(FP&A)용 무설치·리스크0 Excel 자동화. **(A) 누더기 엑셀 → tidy 정형화** +
**(B) 분석결과를 맥킨지/회계법인 룩 단일 템플릿으로 렌더 + 디스패처 + QC**.

회사 PC는 **GitHub pull + 파이썬 실행만** 가능하고 설치/네트워크가 막혀 있다는 전제로 설계했다.
따라서 런타임 외부 의존성은 `vendor/`에 소스로 동봉한 **openpyxl 하나뿐**이며, 회사에서는
설치 스텝 0으로 `git pull` 후 `py main.py` 만으로 동작한다.

## 빠른 시작 (회사 PC)

```powershell
git pull
py main.py selftest                          # 골든샘플 전부 + ingest 픽스처 회귀 (설치 0)
py main.py ingest "C:\path\누더기.xlsx" out\ingest   # 정형화 → tidy.csv/schema.json/smell_report.md
py main.py dispatch "예실 변동 분석 브리지"             # 어느 템플릿 쓸지 판정
py main.py render variance out\variance.xlsx          # 템플릿 렌더(QC 게이트 통과 시 저장)
py main.py list                                       # 구현된 템플릿 유형
```

## 확정 제약 (협상 불가)

- 런타임 라이브러리 = **openpyxl + et_xmlfile** 만. `vendor/`에 **순수 파이썬 소스**로 동봉(설치 금지).
  컴파일 산물(`.pyd/.so/.dll`) 0개 — `py fpna/_bootstrap.py`가 검증.
- **pandas/numpy/pydantic/XlsxWriter/formulas 금지.** 표 변환·검증·재무계산은 표준 라이브러리 + `dataclass`로 직접 구현.
- 중간데이터는 **csv/json(stdlib)**. parquet 금지. 차트는 openpyxl(워터폴 = stacked-bar + 투명 base 트릭).
- 합성데이터 금지(재무 수치), PIT 규율, 출력은 **QC 통과 후 확정**.
- COM/xlwings/Excel MCP 등은 **집-전용 검증**일 뿐 회사 런타임 의존이 아니다(`tools/` 격리).

## 저장소 구조

```
vendor/openpyxl/  vendor/et_xmlfile/  vendor/LICENSES/   ← 순수파이썬 소스 동봉(커밋)
fpna/
  _bootstrap.py          # sys.path 에 vendor/ 주입 (모든 진입점이 최우선 import)
  house_style.py         # 맥킨지/회계법인 룩 SSOT (색·폰트·숫자서식·보더·워터폴)
  finance.py             # 순수파이썬 NPV/IRR/payback/CAGR/variance/비율
  ingest/                # 누더기 엑셀 → tidy long 파이프라인 (작업2)
  dispatcher.py          # 요청+컬럼 단서 → 템플릿 유형 라우팅 (작업5)
  render.py              # 빌드 + QC 게이트 + 저장 (작업6)
  templates/             # 유형별 dataclass+빌더+골든샘플+QC (작업4)
rules/normalize_rules.md # 정형화 규칙(코드로 고정, 결정적)
schema/                  # tidy 출력 스키마 문서
skills/fpna-excel/SKILL.md  # Claude 스킬 패키징 (작업7)
templates_base/          # (선택) 채우기모드용 베이스 .xlsx — 솔로 사용 전제
samples/                 # 내가 만든 골든샘플(.xlsx)
tests/                   # 회귀 테스트 + 픽스처 생성기
tools/verify_xlsx.py     # [집-전용] 실제 Excel(COM) 재계산·룩 검증
main.py  dispatch.md  sources.md  README.md  .gitignore
```

## 정형화 파이프라인 (작업2) — 누더기 → tidy

`fpna/ingest`는 다음을 순수 파이썬으로 결정적으로 처리한다(상세: [rules/normalize_rules.md](rules/normalize_rules.md)).

1. **셀 평면화**(`as_cells`): 값+수식(data_only 두 번 로드)·서식·병합 anchor 를 한 `Cell`로.
2. **블록 탐지**: connected-component + density 임계로 한 시트의 다중 표 분리.
3. **비데이터 행 격리**: 제목/단위주석/각주(블록 폭 상대 기준).
4. **병합 fill + 헤더 해소**: 병합 anchor 전파 후 2-pass 분류(헤더밴드/데이터영역).
5. **behead 언피벗**: 방향 최근접(up/left/up-left/left-up) 매칭으로 (entity, period, metric, value) long 변환.
6. **소계/합계 플래그**: 라벨 키워드 → `row_role=subtotal` (이중집계 방지).
7. **센티넬·단위 정규화**: `…/-/N/A` → None, 괄호음수 `(50)`→-50, 세모 `△30`→-30, % 텍스트→비율, 단위 컬럼 분리.
8. **검증 + 스멜 스캔 + 산출**: dataclass 제약 검증 → reject 리포트, 수식 스멜(하드코딩 상수/과중첩/시트참조 과다/외부참조) → `tidy.csv`(utf-8-sig)·`schema.json`·`smell_report.md`.

> 한국 재무 엑셀 특수성(천원/백만 단위, 괄호·세모 음수, 소계행, 들여쓰기 계정 계층)을 반영했다.
> 설계 보강은 unpivotr/tidy 이론 + 표 탐지 연구(USPTO/Pytheas/ptype 등) 개념을 파이썬으로 재구현한 것이며,
> 외부 소스 복붙은 없다([sources.md](sources.md)).

## 템플릿 유형 (작업4)

| 유형 | 용도 |
|---|---|
| `variance` | 예실(Plan vs Actual) 워터폴 + 변동표 + 코멘터리 |
| `investment_appraisal` | NPV·IRR·할인 회수기간(break-even) |
| `period_trend` | MoM/QoQ/YoY 추이 + 라인차트 |
| `rolling_forecast` | 실적+전망 롤링 포캐스트 |
| `budget_build` | 예산·인건비(headcount) 수립 |
| `cashflow_13w` | 13주 단기 현금/유동성 |
| `unit_economics` | CAC·LTV·LTV/CAC·회수개월 |
| `scenario_sensitivity` | 시나리오 + 토네이도(민감도) |
| `pnl_3statement` | 손익계산서(3-statement 골격) |
| `board_kpi_pack` | 이사회 KPI 팩/대시보드 |

각 유형 = (a) dataclass 입력 스키마 (b) 빌더(생성/채우기 모드) (c) 골든샘플 (d) QC 체크.

## QC 게이트 (작업6)

`render`는 QC 미통과 시 **산출을 보류**한다. 체크: 수식에러(`#REF!` 등) 0건,
합계 교차검증(파이썬 재계산 vs 셀 의도), 부호규약, 단위/포맷 일관.
홈에서는 `tools/verify_xlsx.py`로 **진짜 Excel COM 재계산**까지 교차검증한다
(NPV/IRR 등 셀 수식이 파이썬 `fpna.finance`와 일치하는지).

## 배포 방식 (미확인 — 첫 회차 확인 필요)

- **옵션 A**: 각 워크북에 모듈 import (자기완결).
- **옵션 B**: `Personal.xlsb`(개인 매크로 통합문서)에 공통층 배치 — 전역 반영.
  회사 PC에서 `XLSTART` 쓰기 권한 가능 여부 확인 후 결정. (본 파이썬 시스템은 둘과 독립)

> repo는 **private**(솔로 내부 사용). 라이선스 보유/직접 제작 베이스 `.xlsx`는 `templates_base/`에 커밋 가능.
> 기본은 `house_style` 코드 생성모드. ⚠ public 재전환 시 외부 템플릿(Vertex42/CFI 등 재배포 금지)은 다시 제외.
