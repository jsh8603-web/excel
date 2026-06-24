---
next-action: 100+라인 zone 혼합 대규모 성능·정확성 테스트 실행 → 결과 보고 → push(cfdf6d6·c57b6e4) 사용자 대기
tags: [handoff, design-zones, excel]
date: 2026-06-24
---

# Handoff — design-zones (한 시트 정형블록+freehand 혼합 게이트)

## 1. 현재 상태 · 첫 행동
- design-zones 기능 **구현 완료 + 전 게이트 GREEN**. 커밋 `cfdf6d6`·`c57b6e4` (★미push). 그 전 push분: `62cd91a`(excel-router), `63c9926`(design-zones 코어), `1d00a09`,`a23d06f`.
- **첫 행동**: 100+라인 zone 혼합 대규모 테스트(아래 §5 미해결) 실행 → 결과 보고 → 사용자에게 push 여부 물어 push.

## 2. 진행 맵 (plan-design-zones.md / progress-design-zones.md)
- Phase 1~3 전부 `[x]`. 코어: `fpna/design_zones.py`(신규), `fpna/design_audit.py`(zone-aware 확장, contract=None 하위호환), `fpna/styles_interp.py`(applyX 해석기), `fpna/excel_router.py`(파일검사 자동분기), `fpna/pipeline.py`(_base_owned_gate wb._fpna_zone_contract 연결).
- 진입점: `py main.py excel <파일> [--fix]` — 코드가 파일 검사로 zone/external 자동 분기.
- 설계 SSOT: `.consult-design-zones/DESIGN.md` (자문 9R + COM 실측). raw 전문 `.consult-design-zones/raw/`.

## 3. 사용자 박제 (대화 고유)
- "엑셀" 트리거 + 기존 파일(사람 제작 포함)이면 **코드가 자동 분기**해 돌아가야 함(에이전트 판단 X). → `main.py excel` 구현으로 충족.
- 회사 PULL → 능력 ready. 표준 산출=자동, 혼합=draw_house_block 사용 시 발동.
- 결함 없는 본은 루트 저장+전송 의무 → `example_design_zones_mixed.xlsx`(풀 P&L) 전송 완료.
- 통계청式 대규모 dirty로 결함 찾기 지시 → 진행 중.

## 4. 파일 inventory
- 신규: `fpna/design_zones.py` `fpna/styles_interp.py` `fpna/excel_router.py` `tools/styles_calibrate.py` `tools/zone_regression.py` `tests/test_design_zones.py` `example_design_zones_mixed.xlsx(.contract.json)` `.consult-design-zones/`
- 수정: `fpna/design_audit.py` `fpna/pipeline.py` `main.py`(excel 커맨드+os import) `skills/fpna-excel/SKILL.md`(§0 ⓪·§9) `skills/freehand-excel-integrity/scripts/xlsx_doctor.py`([16]) `CLAUDE.md`(§3)
- 테스트 산출: `out/zone/`, `out/mock/`, `out/dirty/`, `out/big/`(gitignore)

## 5b. ★ingest KOSIS 2-키 결함 (정밀 진단 완료 — 다음 세션 1-shot)
원본 `kosis_원본_기타금융기관총자산.xlsx`(국가1=대륙 merged / 국가2=국가 + 2행헤더). tidy metric 망가짐(국가명+값+"-" concat), entity=대륙.
**Root 2개(복합)**:
1. `fpna/ingest/headers.py` `classify_cells` 의 `_is_data_row`+MIN_DATA_RUN lookahead(L179~202): 데이터행 직후 **전부 "-" 센티넬 행**(아제르바이잔 등)이 value_frac=0 → data-run 끊김 → 실데이터 첫행(아르메니아)이 top 헤더밴드로 흡수 → row3 값·row4 "-"가 열헤더 체인 유입. **Fix**: 센티넬-only 행(셀 대부분이 "-"/dash sentinel)을 data-run 에서 데이터로 인정(_is_data_row 가 sentinel 행도 data 취급) 또는 lookahead 가 sentinel 행을 skip 하지 말고 data 로 카운트.
2. `fpna/ingest/pipeline.py` attr→매핑(L111~116): `entity=row_labels[0]`(대륙) + `metric_parts = row_labels[1:] + metric_parts` → 안쪽 키(국가)가 metric 으로 접힘. **Fix**: 다중 키컬럼 시 안쪽 라벨을 metric 이 아니라 별도 dim(예: entity=innermost 국가, 대륙=region dim) 또는 entity=keys join. tidy 스키마(entity 단일)와 정합 필요 — region 컬럼 추가 또는 entity="대륙·국가" 결합.
**검증**: 수정 후 `py main.py selftest`(29픽스처 회귀 0) + `py main.py ingest kosis_원본...xlsx out/k --sheet 데이터` → metric=순수 지표명, entity=국가, region=대륙 확인. 신규 픽스처(2키+merged+sentinel행) tests/make_fixtures.py 추가 권장.

## 5. 미해결 · 실패 (삽질 위험)
- **100+라인 zone 혼합 대규모 미실행**: draw_house_block로 120+ 라인아이템 P&L(actual/fcst) 생성 → `main.py excel`로 zone 검수, 타이밍·drift 0 확인. resolve_blocks가 cells = row_map×col_map 전수라 대규모 시 O(rows×cols) — 성능 관찰 필요.
- 대규모 dirty 실측(완료): 162행 통계청 → external 753 숫자텍스트 검출 482ms / ingest 810 tidy reject0 511ms. (소규모서 SUBTOTAL_DETECTED 떴는데 대규모선 smell 카테고리 빈 출력 — "계" 행 소계검출 누락 가능성, 확인 거리.)
- push 미실행(cfdf6d6·c57b6e4).

## 6. 결함 7건 (이번 세션 발견·수정, 커밋 분산)
1 라우터 external ok 너무 엄격(자문성까지 fail) 2 zone drift 빈셀 오탐 3 RLE row/col-band max까지 과확장(_trim_to_data) 4 external 숫자-텍스트 미탐지(_numbers_as_text) 5 헤더밴드 미선언 시 unsealed(사용성) 6 골든 16pt ALLOWED_SIZES 밖(잠재, 보류) 7 band 시각비교 dtype가 값↔수식 혼재 band 오탐(dtype 제외).

## 7. 자문 종합
9R(설계5+잔여리스크2+미해결2) gemini-web+claude-web + COM 실측. 핵심: NamedStyle 드롭(set_cell 직접포맷 충돌·role 폰트색 중복) → 숨김 값 마커 2트랙(행→row-band×열→col-band, RLE) + 좌표-free band-id 매니페스트 + resolved 재계산. 전문 `.consult-design-zones/DESIGN.md` §10.
