<!--
전역 CLAUDE.md 에 붙여넣는 "Excel 라우팅 스위치" 스니펫.
원칙: 전역엔 *어디로 보낼지*만(가볍게). 규칙 본문·절차는 스킬이, 강제는 파이프라인 코드가 소유.
( ~/.claude/CLAUDE.md 또는 repo 루트 CLAUDE.md 의 기존 "4도구 사용 여부" 섹션 옆에 추가 )
-->

## Excel 작업 (라우팅)

- "excel/엑셀" 트리거 → **fpna-excel 스킬로 진입**. 직접 openpyxl freehand 금지.
  파이프라인이 grain→contract→QC→**design 게이트**를 우회 없이 소유(run_report 스파인).
- 경로 판정(상세 라우팅·절차는 스킬이 가짐):
  - 정형 반복 산출물(주간 CME 등) → **템플릿 재실행(run_report)**. 편집 아님 → 드리프트 원천 차단.
  - 비정형/일회성 → **freehand-excel-integrity 스킬**(house_style 적용 + xlsx_doctor + contract).
  - 외부 입수 파일(우리가 안 만든 .xlsx) → `xlsx_doctor --external` + `tools/restyle.py`(비파괴).
  - 세션 중 수정 → `house_style.edit_cell` + `xlsx_doctor --golden`. **직접 ws[ref]=v 금지.**
- 4도구 선택은 신호 기반(router.decide): 재계산필요=xlwings · 신규대량=xlsxwriter ·
  편집/기본=openpyxl · 피벗/슬라이서=xlwings(.api COM). 침묵 폴백 금지(DOWNGRADE 명시).

<!-- 끝. 규칙 설명을 여기 더 적지 말 것 — 전역이 무거워지면 매 세션 토큰만 먹고 안 지켜진다.
     강제는 코드(_base_owned_gate 의 design_audit), 절차는 SKILL.md 가 소유한다. -->
