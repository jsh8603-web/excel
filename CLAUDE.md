# CLAUDE.md — FP&A Excel 시스템 작업 지침

이 repo에서 작업하는 Claude Code(또는 다른 AI 에이전트)가 **반드시 먼저 읽어야 하는** 운영 매뉴얼.
사람용 개요는 [README.md](README.md), 스킬 트리거는 [skills/fpna-excel/SKILL.md](skills/fpna-excel/SKILL.md) 참조.

## 0. 한 줄 요약

회사 PC(무설치·외부망 차단·AI 없음)에서 `git pull` 후 **`py main.py`만으로** 도는 FP&A 자동화.
(A) 누더기 엑셀 → tidy 정형화, (B) 분석결과를 맥킨지/회계법인 룩 템플릿으로 렌더 + 디스패처 + QC.

## 1. 절대 제약 (위반 = 시스템 붕괴)

- **런타임 외부 의존성 = openpyxl + et_xmlfile 뿐.** `vendor/`에 순수 파이썬 소스로 동봉돼 있다.
  - ⛔ **pandas / numpy / pydantic / XlsxWriter / formulas / 기타 pip 패키지 import 금지.**
    표 변환·검증·재무계산은 **표준 라이브러리 + dataclass로 직접 구현**한다.
  - ⛔ `vendor/`에 컴파일 산물(`.pyd/.so/.dll`) 반입 금지(순수 파이썬 보장). `py fpna/_bootstrap.py`가 검증.
  - 새 라이브러리가 꼭 필요하면: 순수 파이썬인지 확인 → `pip install --target vendor <pkg>` → 컴파일 산물 0 확인 → LICENSE를 `vendor/LICENSES/`에 복사 → sources.md 기록. **이 절차 없이는 import 추가 금지.**
- **중간 데이터 = csv/json(stdlib)만.** parquet 등 컴파일 의존 금지. 차트는 openpyxl(워터폴=stacked-bar+투명 base).
- **모든 진입점은 `import fpna._bootstrap`을 최우선**으로 한다(vendor/를 sys.path에 주입). 이게 빠지면 회사 PC에서 ImportError.
- **합성 재무수치 금지**(테스트 *구조* 픽스처는 예외 — 의미 없는 더미임을 명시). PIT 규율. **출력은 QC 통과 후에만 확정.**
- **COM/xlwings/Excel MCP는 검증 보조 도구**(`tools/`)다 — pywin32+Excel 있으면 회사/집 무관하게 쓸 수 있다(Python312 site-packages 에 pywin32 포함된 환경이면 회사에서도 산출물 COM 검증 가능). ⛔ 단 **런타임 코드**(`fpna/`, `main.py`)는 절대 이것들에 의존하면 안 된다 — pywin32 는 `.pyd` 컴파일 산물이라 `vendor/` 동봉 불가 → `py -S`(site-packages 차단) 무설치 검증이 깨진다. "집-전용"이 아니라 "런타임 비의존"이 제약의 본질.

## 2. 실행·검증 (작업 후 반드시)

```powershell
py main.py selftest          # 골든샘플 10종 QC + ingest 픽스처 회귀 → ALL PASS 여야 함
py -m unittest tests.test_fpna   # stdlib 회귀 (pytest 불필요)
py main.py ingest "<파일.xlsx>" out\ingest [--sheet 시트명]   # 누더기 엑셀 → tidy
py main.py profile "<마트.csv>" out\profile_spec.yaml         # 정제 마트테이블 → 차원없는 SHAPE 스키마(회사→집)
py main.py encrypt <평문> [out] --mail [--max-lines N]        # 메일 본문 텍스트로 암호화(길면 part 분할)
py main.py decrypt <암호문> [out]                             # 복호화(part 마커 자동 합본·정렬)
py main.py dispatch "<요청 텍스트>"
py main.py render <type> out\<type>.xlsx
py main.py list
```

**회사 무설치 재현 검증** (코드 바꾼 뒤 의무):
```powershell
py -S main.py selftest       # -S = site-packages 차단 → vendor/ 동봉본으로만 도는지 확인
```
홈 site-packages에도 openpyxl이 있으므로 `-S` 없이는 vendor 검증이 무의미하다.

**실제 Excel 검증(pywin32+Excel 필요, 권장)**: `py tools/verify_xlsx.py out\x.xlsx`
→ 진짜 Excel로 전체 재계산해 ①에러셀 0 ②NPV/IRR 등 셀 수식이 `fpna.finance`와 일치하는지 확인.

### 환경 메모 (홈 PC)
- `py`/`python`이 **PATH에 없다.** bash에서는 `/c/Users/jsh86/AppData/Local/Programs/Python/Python312/python.exe` 사용.
- bash 콘솔이 cp949라 한글이 깨져 보임 → `PYTHONIOENCODING=utf-8` 권장.
- Excel(Office16)+pywin32 설치돼 있어 COM 검증 가능.

## 3. 아키텍처 (수정 시 영향 범위)

```
fpna/_bootstrap.py    vendor/ 주입. 모든 진입점이 최우선 import.
fpna/house_style.py   룩 SSOT(색·폰트·숫자서식·보더·차트). 룩 변경은 여기 한 곳만.
fpna/finance.py       순수파이썬 NPV/IRR/payback/CAGR/variance/비율 + solve_revolver(이자↔부채↔현금 순환 고정점). QC 재계산도 여기.
fpna/coa.py           표준 계정과목 taxonomy(IS/BS/CFS·sign·us_gaap/ifrs tag). refdata/coa_us_gaap.json(공개도메인 명칭, 재무수치 0).
fpna/ingest/          누더기→tidy. cells→detect→headers→normalize→validate→pipeline 순.
fpna/profile.py       정제 마트테이블 → 차원없는 SHAPE 스키마(yaml, 8축). ⚠ 누더기는 ingest, 정제 마트는 profile(다른 단계).
fpna/crypto.py        텍스트 대칭 암복호화(ChaCha20-Poly1305+scrypt). _chacha.py=RFC8439 이식. 메일 본문 운반·part 분할.
fpna/dispatcher.py    요청 텍스트 → stage(pack/report/ingest/profile/transport/analysis) + 분석표 유형. pack 게이트=resolve_pack.
fpna/render.py        build → QC 게이트 → (통과 시만) 저장.
fpna/templates/       유형별 모듈. 각각 INPUT(dataclass)/golden_sample()/build()/qc().
                      __init__.py 의 _MODULES 레지스트리에 등록.
fpna/pack.py          다중 exhibit 연동 팩(graft 합본 + Control 시트 + run_report 스파인). PackSpec/build_pack.
fpna/packs/           팩 카탈로그 레지스트리. make_spec()→PackSpec. feasibility(사업타당성) 구현. 가이드=packs.md.
```

### 라우팅 레이어 (단일 vs 팩)
- **단일 의도** → `dispatch.md`(L1): 한 요청 → 한 템플릿. `dispatcher.dispatch` cascade.
- **연동 묶음** → `packs.md`(L2): 여러 장표가 공유 가정·크로스시트 tie 로 묶일 때. `dispatcher.classify_stage` 가 pack 게이트로 선분류 → `pack.build_pack`(스파인 경유, 모델체크 A=L+E·현금 tie). 단일 exhibit 이면 팩으로 부풀리지 말 것(과투자).

### ingest 파이프라인 단계 (상세: rules/normalize_rules.md)
셀 평면화(값+수식 두 번 로드) → 블록 탐지(connected-component+density) → 비데이터 행 격리(제목/단위/각주)
→ 병합 fill + 2-pass 헤더 분류 → behead 언피벗(방향 매칭) → 소계 플래그 → 센티넬/괄호·세모음수/단위 정규화
→ dataclass 검증 + 수식 스멜 → tidy.csv(utf-8-sig)/schema.json/smell_report.md.

## 4. 확장 방법

### 새 템플릿 유형 추가
1. `fpna/templates/<type>.py` 생성. 다음을 노출(덕 타이핑):
   - `TYPE` 상수, `INPUT` dataclass, `golden_sample()->INPUT`,
     `build(data, *, mode="create", base_path=None)->openpyxl.Workbook`, `qc(wb, data)->QCReport`.
2. `build`는 `fpna.house_style`만 써서 서식 지정(직접 색/폰트 하드코딩 금지).
   실제 모델 수식은 셀에 수식 문자열로 기입(`=NPV(...)` 등) — 별도 라이브러리 불필요.
3. `qc`는 `fpna.finance`로 파이썬 재계산해 셀 의도와 대조 + `qc_no_formula_errors` 호출.
4. `fpna/templates/__init__.py`의 `_MODULES`에 등록.
5. `py main.py golden <type>` + `py tools/verify_xlsx.py`로 검증. dispatcher 키워드도 `fpna/dispatcher.py`에 추가.

### ingest 규칙 수정
- 규칙은 **코드가 SSOT**(rules/normalize_rules.md는 설계 근거). 결정성 유지 — 같은 입력→같은 출력.
- 수정 후 `py main.py selftest`로 픽스처 회귀 통과 + 실데이터 1건 확인 후 확정.
- 새 엣지케이스는 `tests/make_fixtures.py`에 구조 픽스처로 추가하고 `tests/test_fpna.py`에 단언 추가.

### 머지 게이트 (커밋 전)
- `py main.py selftest` = ALL PASS, `py -m unittest tests.test_fpna` = OK, `py -S main.py selftest` = ALL PASS.
- 새 런타임 import 0건(§1). QC 미통과 산출물은 커밋 금지.

## 5. 두 모드 (생성 vs 채우기)

- **생성모드(create, 기본)**: house_style로 처음부터 작성. 차트·서식 전부 코드. **권장.**
- **채우기모드(fill)**: `templates_base/`의 베이스 .xlsx를 열어 데이터 셀만 채움.
  ⚠ openpyxl 라운드트립은 차트/이미지/피벗을 떨굴 수 있다 → 베이스는 표·서식 위주, 차트는 코드 재생성.

## 6. 회사 ↔ 집 통신 규율

- **회사→집**: 구조 메타(컬럼/시트/포맷/규모)·에러 메시지만 텍스트로. ⛔ 실데이터 숫자·거래처명·원본파일·스크린샷 반출 금지.
- **집→회사**: 코드/수식/런북(전부 텍스트). 회사에서 실데이터로 실행, 결과는 회사 PC에서 마무리.

## 7. Git / 저장소 규율

- 커밋: 내 코드 전부 + `vendor/`(.py만) + SKILL.md + 골든샘플(`samples/`) + 문서.
- ⛔ 커밋 금지: `_sources/`(학습 clone), 남의 소스 복붙, `vendor/` 컴파일 산물, `archive/`, `out/`, `plan/progress.md`.
- repo는 **private**(솔로 내부 사용) → 라이선스 보유/직접 제작 베이스 .xlsx를 `templates_base/`에 커밋 가능. ⚠ public 재전환 시 외부 템플릿(Vertex42/CFI 등 재배포 금지)은 다시 제외.
- PAT는 `~/.claude/keys/github-pat.md`. push 후 remote URL에서 PAT 제거.
- 커밋 메시지 = conventional commits(feat/fix/chore/refactor).

## 8. 미확정·임의결정 (확인 필요)

- repo private(2026-06-09 전환 완료) → 베이스 .xlsx 커밋 가능. 기본은 house_style 코드 생성. public 재전환 시 외부 템플릿 재제외.
- 폰트 맑은 고딕+Calibri 회사 PC 보유 가정.
- 배포 방식 A(워크북별 import) vs B(`Personal.xlsb`) 미정 — 파이썬 시스템과 독립이라 영향 없음.
- behead의 `up-left`/`left-up` 정렬 우선순위는 리서치 일부 추론분 — 골든 테스트 통과했으나 실데이터 재확인 권장.
- 로컬 검증은 pywin32 COM(`tools/verify_xlsx.py`). xlwings/Excel MCP로 교체 가능.
