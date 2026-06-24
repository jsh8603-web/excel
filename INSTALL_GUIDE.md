# 적용 가이드 — 우리가 구현한 것을 회사 시스템에 물리기

회사 시스템 = repo대로(4도구 + fpna 파이프라인), 엑셀 스킬 로드 시 구동, 전역 CLAUDE.md엔
"4도구 사용 여부"만. 우리가 만든 디자인 게이트·편집/외부 가드를 **3층**으로 물린다.

```
① 전역 CLAUDE.md   → "언제 어디로"만(라우팅 스위치). 규칙 본문 X → 가볍게·항상 읽힘
② 스킬 SKILL.md    → "어떻게"(절차·게이트). 발동 시에만 로드 → 전역 오염 없음
③ 파이프라인 코드   → "우회 불가 강제"(_base_owned_gate 의 design_audit). 진짜 enforcement
```

핵심: **전역엔 스위치, 강제는 게이트, 절차는 스킬.** 전역에 규칙을 잔뜩 적으면 매 세션
토큰만 먹고 안 지켜진다 — 그게 지금 드리프트의 근본 원인.

---

## 1단계 — repo 패치 (③ 강제 + 분기)

`fpna_excel_integration.patch` 하나로 끝. 깨끗한 main 에 바로 적용된다(선행 패치 불필요;
layout_audit/formula_audit 은 이미 main 에 있음).

```bash
cd <repo>
git apply --check fpna_excel_integration.patch   # 충돌 없는지 먼저 확인
git apply fpna_excel_integration.patch
py -m pytest tests/ -q                            # 294 + 신규 5 통과 확인
```

패치가 넣는 것:
- `fpna/design_audit.py` — 디자인 표준 게이트(장식문자 hard-fail·정렬/폰트/주석 보고) + 비파괴 restyle. house_style 토큰 직접 읽음(SSOT 정합).
- `fpna/house_style.py` — `ALLOWED_SIZES` + `edit_cell`(편집 드리프트 차단).
- `fpna/pipeline.py` — `_base_owned_gate` 스파인에 design 게이트 연결(run_report 산출 자동 통과).
- `tools/restyle.py` — 외부 입수 파일 비파괴 정규화 CLI.
- `skills/fpna-excel/SKILL.md` — §9 경로 분기(정형/비정형/외부/편집).
- `rules/design_standard.md` — §6 검사 계층 문서.
- `tests/test_design_audit.py` — 회귀 5종.

이걸로 **정형 산출물은 생성 시 디자인 게이트를 우회 없이 통과**한다(LLM 준수에 안 기댐).

## 2단계 — freehand-integrity 스킬 설치 (② 비정형/외부/편집 절차)

`freehand-excel-integrity.skill` 을 스킬 디렉터리에 푼다. 비정형 생성·외부 입수·세션 편집을
받치는 절차(xlsx_doctor [1]~[15], house_style_min, recalc, router, restyle, golden)가 들어있다.

repo 안에서 돌면 xlsx_doctor 가 vendored 미러 대신 **진짜 fpna.house_style** 토큰을 읽는다
(PYTHONPATH 에 repo 루트가 있을 때). 회사 PC는 repo 루트에서 구동하므로 자동.

## 3단계 — 전역 CLAUDE.md (① 라우팅 스위치)

`claude_md_excel_snippet.md` 내용을 ~/.claude/CLAUDE.md(또는 repo 루트 CLAUDE.md)의 기존
"4도구" 섹션 옆에 붙인다. **규칙 본문은 넣지 말 것** — 어디로 보낼지만.

---

## 적용 후 동작 (경로별)

| 입수 | 무엇이 강제하나 | 명령 |
|---|---|---|
| 정형 생성 | 파이프라인(③) — design_audit 게이트 우회 불가 | `run_report`(스킬이 호출) |
| 비정형 생성 | 스킬(②) — house_style 적용 + doctor | `xlsx_doctor <f>` (contract 선언) |
| 외부 입수 | 스킬(②) — 표준 대조·비파괴 | `xlsx_doctor <f> --external` → `tools/restyle.py <f>` |
| 세션 편집 | 스킬(②) — 역할유지 + 드리프트 대조 | `house_style.edit_cell` + `xlsx_doctor <f> --golden` |

## 검증 체크리스트

```bash
# 1) 정형: 표준 산출이 디자인 게이트 통과(침묵)
py -m pytest tests/test_design_audit.py -q          # 5 passed

# 2) 외부: 나쁜 파일 검수 → 비파괴 정규화
py tools/restyle.py <외부.xlsx>                      # "값·수식 불변 검증 통과"

# 3) 편집: 직접 덮어쓰기 → 드리프트 잡힘 / edit_cell → 안 잡힘
py <skill>/scripts/xlsx_doctor.py <f> --golden       # 기준선→대조
```

## 유지보수 (드리프트 방지)

- 룩 변경은 **repo `fpna/house_style.py` 한 곳**(SSOT). design_audit·린터는 거기서 토큰을 읽으므로 자동 따라온다.
- 독립 스킬의 `house_style_min.py` 는 repo 밖(회사 외)용 폴백 — repo 토큰과 주기적 동기화만.
- 새 디자인 규칙은 design_audit 에 추가하고 tests/test_design_audit.py 에 케이스 1개씩.
