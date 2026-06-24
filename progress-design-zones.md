# progress — design-zones 구현 (plan-design-zones.md)

> Layer B. plan = `plan-design-zones.md`, 설계 SSOT = `.consult-design-zones/DESIGN.md`. 세션 = btn-excel(Opus).
> 한번에 전 Phase 진행(사용자 지시 2026-06-24). 각 step 완료 시 회귀 게이트 통과 후 체크.

## Phase 1 — 준비/스캐폴딩
- [x] 1.1 `fpna/design_zones.py` 헬퍼(_dfam/_color_key/_resolved/마커 RLE 파서) — model: opus
- [x] 1.2 `.contract.json` 로더/검증(좌표 0 단언) — model: opus
- [x] 1.3 `tests/test_design_zones.py` 구조 픽스처(2-D actual/fcst) — model: opus

## Phase 2 — 구현
- [x] 2.1 `draw_house_block`(set_cell 재사용 + 숨김 행/열 마커 + 불변식 6) — model: opus
- [x] 2.2 `resolve_blocks` 파서(마커 2트랙→block_id, RLE, unsealed) — model: opus
- [x] 2.3 `design_audit.design_findings(contract=None)` zone-aware 확장(하위호환) — model: opus
- [x] 2.4 `design_audit.restyle_inplace(contract=None)` zone-aware — model: opus
- [x] 2.5 `pipeline._base_owned_gate` 연결 — model: opus
- [x] 2.6 `xlsx_doctor [14]`/SKILL §9 혼합 경로 — model: opus
- [x] 2.7 COM 캘리브(`tools/styles_calibrate.py`) + `fpna/styles_interp.py` — model: opus (오프라인, 분리 가능)

## Phase 3 — 검증
- [x] 3.1 `py main.py selftest` ALL PASS(정형 29종 회귀 0)
- [x] 3.2 `py -m unittest` test_fpna/test_design_audit/test_design_zones OK
- [x] 3.3 `py -S main.py selftest` ALL PASS
- [x] 3.4 편집-생존 COM 매트릭스(픽스처) = 실측표 일치
- [x] 3.5 AND/unsealed/RLE 시뮬 loud 정확
- [x] 3.6 `py tools/verify_xlsx.py` 에러셀 0
- [x] 3.7 런타임 import 0

> 인계: [handoff-design-zones-20260624.md](./handoff-design-zones-20260624.md)

## Working Notes
> [ckpt-202606242045:btn-excel] design-zones 전 Phase 완료+커밋 cfdf6d6·c57b6e4(미push). 결함 7건 발견·수정(라우터 ok/빈셀/RLE과확장 trim/숫자텍스트/헤더밴드/dtype + 잠재16pt). 결함없는 본=example_design_zones_mixed.xlsx(풀 P&L 10라인+소계3+달성%, 루트, 라우터 clean·verify 에러0). 대규모 실측: 162행 통계청 dirty→external 753 숫자텍스트 검출 482ms / ingest 810 tidy reject0 511ms. **다음 의도**: 100+라인 zone 혼합 대규모 성능·정확성 테스트(미실행) → 결과 보고 → push(cfdf6d6,c57b6e4) 사용자 대기. **동기화**: push 미실행.
> [ckpt-202606242010:btn-excel] design-zones 코어 완료 (fpna/design_zones.py 신규 + design_audit.py zone-aware 확장 + tests/test_design_zones.py 7종). 게이트: selftest ALL PASS(정형29종 회귀0) / unittest 83 OK / py -S ALL PASS / 런타임 import 0. draw_house_block 생성↔zone_findings 검증 대칭 실증. **남은 것**: 2.5 pipeline 자동 contract 연결(run_report 가 혼합시트 contract 생성하는 wiring — mixed 생성경로 실사용 시) / 2.6 xlsx_doctor [14] strict 토글 + SKILL §9 문서 / 2.7 COM 캘리브(오프라인 진단도구, styles.xml applyX 해석기) / 3.4~3.6 COM 회귀매트릭스·verify_xlsx(혼합 생성파일 필요). 모두 코어 위 통합/진단 레이어 — 코어는 commit 가능 상태. 미커밋(사용자 commit 지시 대기).
