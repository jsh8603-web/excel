---
name: fpna-excel
description: >
  사내 FP&A Excel 작업 스킬. (A) 누더기/비정형 엑셀을 tidy 정형 데이터로 변환하고,
  (B) 분석결과를 맥킨지/회계법인 룩 단일 템플릿(예실 variance·투자 NPV/IRR·기간추이·
  롤링포캐스트·예산/인건비·13주현금·유닛이코노믹스·시나리오민감도·손익·이사회KPI)으로
  렌더한다. 디스패처가 요청을 유형으로 라우팅하고 QC 게이트가 출력을 검증한다.
  무설치·openpyxl 단일 의존(vendor 동봉). "예실/변동분석, 현금흐름, NPV/IRR 투자타당성,
  포캐스트, 누더기 엑셀 정형화, 손익, 이사회 KPI" 요청 시 사용.
---

# FP&A Excel Skill

회사 PC(설치0·외부망 차단·AI 없음)에서 `git pull` 후 **`py main.py`만으로** 동작하는
FP&A 자동화. 런타임 의존성은 `vendor/`에 동봉한 openpyxl 하나뿐.

## 트리거

- "이 엑셀 정형화/tidy 로 바꿔줘", "누더기 표 정리" → **ingest**
- "예실/변동/브리지", "NPV/IRR/투자 타당성", "MoM/추이", "포캐스트 갱신",
  "예산/인건비", "13주 현금/유동성", "CAC/LTV", "시나리오/민감도", "손익", "이사회 KPI" → **dispatch → render**

## 진입점 (작업2~6 호출법)

모든 명령은 repo 루트에서 실행. 진입점은 `fpna._bootstrap`을 최우선 import 해 `vendor/`를 path에 주입한다.

```powershell
# 0) 회귀 확인 (설치 0)
py main.py selftest

# 2) 정형화: 누더기 엑셀 → tidy.csv / schema.json / smell_report.md
py main.py ingest "<파일.xlsx>" out\ingest [--sheet 시트명]

# 5) 디스패치: 요청 텍스트 → 템플릿 유형 판정
py main.py dispatch "<요청 텍스트>"

# 6) 렌더: 유형 빌드 + QC 게이트 통과 시 저장
py main.py render <type> out\<type>.xlsx
py main.py list      # 유형 목록
```

## 워크플로 (회사 ↔ 집 아티팩트 왕복)

1. **회사 → 집**: 구조 메타(컬럼/시트/포맷/규모)와 에러 메시지만 텍스트로. 실데이터 숫자는 반출 금지.
2. **집**: `ingest`로 구조 처리 검증 → 적합한 템플릿 유형의 3층 데이터 바인딩 래퍼 작성 → 골든샘플+QC로 검증 → `tools/verify_xlsx.py`(집-전용 Excel COM)로 수식 재계산·룩 확인.
3. **집 → 회사**: 코드/수식/런북(전부 텍스트). 회사에서 실데이터로 `render` 실행, 결과는 회사 PC에서 마무리.

## 프로그래매틱 사용

```python
import fpna._bootstrap                       # 최우선 (vendor 주입)
from fpna.ingest import run_ingest
from fpna.dispatcher import dispatch
from fpna.render import render
from fpna.templates import get_template

res = run_ingest("messy.xlsx", "out/ingest")          # 정형화
d = dispatch("예실 변동 분석", columns=res.schema["columns"])  # 라우팅
data = get_template(d.template).golden_sample()        # 또는 실데이터 바인딩 dataclass
r = render(d.template, data, "out/variance.xlsx")      # QC 게이트 후 저장
print(r.qc.summary(), r.saved)
```

## 제약 (반드시 준수)

- openpyxl 외 런타임 라이브러리 추가 금지(pandas/numpy/pydantic/XlsxWriter/formulas 금지).
- 중간데이터 csv/json(stdlib)만. 차트는 openpyxl(워터폴=stacked-bar+투명 base).
- 합성 재무수치 금지, PIT 규율, **QC 통과 후 출력 확정**.
- COM/xlwings/MCP는 집-전용 검증(`tools/`)이며 회사 런타임 의존 아님.

## 관련 문서

- 정형화 규칙: `rules/normalize_rules.md`
- 디스패처: `dispatch.md`
- 출처/라이선스: `sources.md`
- tidy 스키마: `schema/tidy_schema.md`
