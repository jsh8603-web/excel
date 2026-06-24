# 설계 — 한 시트 내 "정형 블록 + freehand" 혼합 디자인 게이트 (design-zones)

> 출처: 2026-06-24 btn-excel 세션. 자문 3R(설계) + 2R(잔여리스크) × (gemini-web + claude-web) + 로컬 COM 실측.
> raw 전문 = `./raw/`. 본 문서 = 채택 전 **설계 SSOT(prior)**. ⚠ 코드화 전 falsify 게이트(consult-adoption-gate) 통과 필요.

## 0. 문제

`fpna-excel` 라우팅에서 한 시트/워크북의 **일부만 표준(정형)**이고 나머지는 freehand일 때:
- `dispatcher.dispatch()`는 무매칭도 freehand로 안 보내고 `pnl_3statement` 강제(dispatcher.py:137) → freehand 선택은 100% LLM 판단(SKILL §0.5).
- `design_audit.design_findings(wb)`·`restyle_inplace(wb)`는 **워크북 전체 순회**(design_audit.py:43,109) — range/region 개념 0.
- freehand `xlsx_doctor [14]` 디자인은 `[자문성]` WARN뿐(xlsx_doctor.py:800), fatal 미반영.
→ 정형 표가 freehand로 오분류되면 design 강제(hard)가 사라지고 룩이 깨진 채 통과. (사용자 관찰)

## 1. 수렴된 아키텍처 (3자 합의)

**전제 재구성(Claude)**: strict zone = 감사 대상이 아니라 **생성 계약**. 룩은 캐노니컬 생성기로 by-construction 보장, 감사는 회귀 방어(defense-in-depth).

| 축 | 결정 | 근거 |
|---|---|---|
| 외부 JSON 좌표 range | **폐기** | 실측 T2: openpyxl insert_rows가 좌표 자동갱신 X |
| 정적 카운트(expected_rows/max_row) | **폐기** | Claude: insert 후 계약=N/파일=N+3 staleness, Gemini R3 수용 |
| ~~정체성 운반체(in-band) NamedStyle~~ | **드롭 확정(falsify+미해결 3R)** | openpyxl `cell.style=name`은 룩까지 적용→우리 set_cell 직접 Font와 충돌, "정체성 전용 라벨" 불가(Claude). role은 이미 폰트색 resolved-추정(`_infer_role`)이라 중복. M6 리네임도 드롭하면 소멸 |
| **정체성/바인딩 운반체** | **숨김 마커 트랙(값)** — 1-D=숨김 열(행→row-band id) / 2-D=숨김 열 × 숨김 행(열→col-band id) | 실측: 값이라 ClearFormats·전체붙여넣기(해당 트랙 외)·열삽입 신규열 빈칸 모두 robust. block_id(r,c)=(row-band[r],col-band[c]) 계산 |
| extent(경계) | **named ∪ Z 의 live bounding box (OR)** | OR이라 한 앵커 소실해도 행 생존 → S2 가장자리수축 닫음 |
| block_id 라벨 | **per-row AND**: named 함의 block_id == Z block_id, 불일치 loud | Claude: union만으론 행수보존 mislabel(sort desync) 침묵 |
| 구조 검증 | **밴드 모델**(열구조+밴드순서+밴드별 cardinality+role), 좌표 0 | Claude. Gemini의 정적카운트 대체 |
| 준수 검증 | **태그 불신 → resolved 재계산** (`cell.font/fill/border/align/numFmt/quotePrefix/data_type`) vs house_style 스펙 | T4 실측. S6(텍스트화 SUM깸)·formula석화 차단 위해 numFmt/quotePrefix/dtype/formula 필수 포함 |
| 수선 | `cell.style="role__bid"` 재할당(override 통째 제거+태그 복원). **insert_rows 금지** | openpyxl 스타일 미전파 버그 회피 |
| 3차 앵커(tabular 한정) | **Excel Table(ListObject)** ref | 실측: openpyxl작성 표를 Excel이 손상없이 열고 행삽입 시 ref C6→C7 자동확장. 단 헤더필수·겹침불가·병합불가 |
| golden_compare | strict=정보성 강등, **태그 전소 복구 오라클**(유니크 시그니처, 모호시 거부·escalate) + freehand 위치검사 | 양쪽 |

## 2. 구멍(hole) vs 정당한 신규행 — 2인자 판별 (Claude)

bbox 안 태그 없는 데이터 행 R: 위·아래 동일 block_id 최근접 role(role_above/below) 조회.
- 같은 밴드 런 내부 + 밴드 [1,None] 가변 + colsig(상대열→dtype family) 일치 → **unadopted_row**(정당 신규행, restyle 입양)
- 밴드 [1,1]/[K,K] 고정 → **cardinality_violation**
- colsig 불일치 → **intrusion**(flag만, 자동입양 금지)
- 밴드 이음새 → colsig 일치 시 입양후보 / 둘 다 불일치 → **seam_intrusion** flag
- 부분 태그 행 → **partial_hole**(해당 열 role 재주입)
- **non-data hole 정밀정의 = 빈셀 ∧ 수식 range 미참조 ∧ Table data body 밖** → 관용. 합계참여 신규 하드값 hole = 무조건 flag(loud)
- cardinality="삽입 허용?", colsig="그 밴드 데이터?". 둘 다 통과해야 자동입양.

## 3. 침묵 모드 목록 + 방어 (실측 검증)

| 모드 | 실측 | 방어 |
|---|---|---|
| (a) 직접 폰트변경 | 태그 생존, override 읽힘 | resolved_drift 검출 → 재할당 복원 |
| (b) ClearFormats | 태그 소실→Normal, 값 생존 | Z마커 생존 → extent 유지. bbox 내 Normal = hole loud |
| (f) 전체붙여넣기 중간행 | 그 행만 소실, 상하 생존 | hole로 검출 |
| (g) 마지막행 ClearFormats | **bbox 축소 침묵** | Z마커가 extent 유지(OR union) |
| sort desync | (Gemini 제기) | per-row AND-agreement loud + 중간 이빨 hole |
| insert후 전체붙여넣기 | named만 상속·Z공백 행에 paste = 양앵커 동시소실 | named-only(Z공백)=unsealed loud+backfill, insert시점 loud화 |
| Z 광역 paste(Z열 포함) | Z 오염 | Z를 데이터서 격리 + Z 비-block값 유입 loud |
| formula 석화 | 값/서식 통과, 기능사망 | resolved에 data_type=='f' 포함, data_only=False |
| Table 하향 auto-grow | 표 밑 타이핑→ref 흡수 침묵성장 | 3차 앵커가 ref 행수 vs 기대 대조, 예상외 확장 loud |
| Z열 전체삭제 | 전역소실 | 오히려 loud(스키마 실패), 기록 |

## 4. 3함수 개조 (코드 스케치 = raw/design-R3-*.txt)

- `design_findings(wb, contract) -> findings`: 태그 클러스터 수집 → fragmentation / 밴드순서·cardinality / resolved-drift(numFmt/quotePrefix/dtype/formula 포함) / 구멍(colsig 2인자) / **per-row named↔Z AND-agreement** / unsealed(named-only) 검출. 좌표 0.
- `xlsx_doctor.design_lint`/`golden_compare`: strict=밴드모델 hard, golden=태그전소 복구오라클+freehand 위치검사+hole census.
- `restyle_inplace(wb, contract)`: drift/partial_hole→`cell.style=name` 재할당, unadopted_row→안전입양, intrusion/seam/frag/cardinality/AND-mismatch→인간 flag. insert_rows 미사용.
- 신규 헬퍼: `_id_of`(named style 파싱), `_dfam`(dtype family), `_color_key`(theme/indexed/rgb 정규화 — false drift 방지), `_resolved`(numFmt/quotePrefix/dtype 포함).

## 5. 생성 대칭 `draw_house_block(ws, origin, band_model, block_id)` — 불변식 6개 (Claude)

① NamedStyle로만 정체성(직접포맷 금지) — 최중요 ② 빈 셀 포함 직사각형 전체 태깅(bbox 안정) ③ colsig 방출 ④ `_resolved()`와 동일 canonical 직렬화(false drift 방지) ⑤ 블록 행 연속·비인터리브 ⑥ `__` 예약. + 숨김 열 block_id 동시 기록(out-of-band 앵커).

## 6. COM 경계 (Claude, 봉인)

COM은 **런타임/CI 절대 미포함**(결정성·openpyxl-only 위반). 역할 = **1회 오프라인 캘리브레이션**: 순수파이썬 `xl/styles.xml` 해석기(cellStyleXfs→cellXfs xf체인 + applyFont/Fill/NumFmt/Border/Align 상속 자체 해소)를 Excel-effective와 1회 대조해 진리표 확정. 런타임은 그 결정적 해석기만. (`tools/`에 캘리브레이션 스크립트, `fpna/`엔 해석기.)

## 7. .contract.json 스키마 (좌표 0)
```json
{"blocks":{"pl_main":{
  "bands":[{"role":"header","card":[1,1]},{"role":"data","card":[1,null]},{"role":"total","card":[1,1]}],
  "order":["header","data","total"],
  "colsig":{"data":{"0":"text","1":"number","2":"number"}},
  "house_style":{"header__pl_main":{...canonical resolved...},"data__pl_main":{...}},
  "table":"tbl_pl_main"   // 선택, tabular 블록만
}}}
```

## 8. 채택 전 falsify 게이트 (consult-adoption-gate — 다음 단계)
- [ ] `draw_house_block` 생성↔`design_findings` 검증 대칭이 우리 house_style.py 토큰과 정합하는지 코드 Read 후 확인
- [ ] AND-agreement(named↔Z) 행별 대조 + unsealed 검출을 실파일로 1회 시뮬
- [ ] COM 오프라인 캘리브레이션: openpyxl `cell.font` resolved vs Excel COM effective 1회 대조(applyX 상속 진리표)
- [ ] hole census 좌표/카운트 결정성 확인

## 9. 미해결 (측정으로 닫을 것)
- M6(동명 스타일 cross-workbook 전체붙여넣기 → `header__pl_main 2` 리네임 변종) 미측정 — rsplit 파싱 깨질 수 있음, 측정 필요.
- openpyxl resolved ≠ Excel-effective(applyX 상속) 정도 — §6 캘리브레이션에서 정량화.

---

## 10. 미해결 자문 (falsify + 1-D/2-D) 해소 — NamedStyle 드롭 + 2-D 이중 마커 (2026-06-24 추가)

### 10.1 NamedStyle 드롭 (3자 수렴 + falsify)
- **근거(Claude, 결정적)**: openpyxl `cell.style="role__bid"`는 font/fill/border/numFmt/align **전체를 셀에 적용** → 우리 `set_cell`의 직접 `cell.font=role_font`와 충돌. "정체성 전용 라벨"로 못 씀. (a)룩까지 운반=전 호출부 마이그레이션 / (b)복제 병존=중복+드리프트. 둘 다 회귀 위험 큼.
- role 정체성은 이미 `_infer_role`(폰트색 resolved-추정)로 존재 → NamedStyle은 role-절반 재구현. block_id만 새로 필요.
- **드롭 시 잃는 침묵모드 = 정확히 1개**: 동role·타블록 단일셀 전체붙여넣기(저빈도, 비지배벡터). formula resolved 비교로 부분보상.
- **얻는 것**: ClearFormats 내구성(값 마커 생존), M6 cross-workbook 리네임 소멸, 경계삽입 loud안전, 마이그레이션 0.

### 10.2 2-D 레이아웃 확정 (사용자: ACTUAL/FCST 한 행에 좌우 병렬)
바인딩(어느 셀=어느 블록) vs 의미(계약) **레이어 분리**(Claude reframe):
- **바인딩 레이어** = in-cell 숨김 값 마커, 엔진이 셀 따라 이동.
  - 숨김 **열**(데이터 좌측 밀착) = 행→**row-band id**(라인아이템, 좌우 공유)
  - 숨김 **행**(데이터 상단 밀착, row1 회피=헤더머지 churn) = 열→**col-band id**(label/actual/fcst/variance)
  - `block_id(r,c) = (row-band[r], col-band[c])` — 계산, per-cell 저장 0. 공유-row 구조라 텐서곱 faithful(rank-1 separable).
- **의미 레이어** = **좌표-free 매니페스트**(숨김 `_Manifest` 시트 또는 .contract.json), **band-id 키**로 {role, numFmt, dtype, formula-template, quotePrefix}. 좌표 0 → stale 0. (★매니페스트에 col-range/row-anchor 절대 금지 — 그게 Q3 stale 버그.)
- **교차고정** = "마커 문자열 == 매니페스트 키". 파서가 마커 스캔→band-id→매니페스트 lookup. 좌표는 read-time 물리위치서 도출.
- **sparse band-start + RLE**(Claude): 밴드 시작 행/열만 마킹, 다음 마커까지 연속. 중간삽입 빈칸=자동 continuation. band-id=콘텐츠유래 안정 id(라벨 해시), 위치無.
- **ragged 2-D**(ACTUAL 5-20 / FCST 5-25 등, 우리 공유-row regime엔 거의 無): (row-band,col-band)→block_id map을 매니페스트에 추가.
- **가시 헤더**("실적/전망")=fallback 검증만(편집·머지 취약), 권위 바인딩 아님. (색=role fallback과 대칭.)

### 10.3 실측으로 닫은 measure 항목 (raw/col_marker.py)
- 열삽입(D앞): 기존 마커 셀 따라 이동, 신규 D열 **빈칸**(unsealed/RLE continuation). 행삽입 unsealed-row와 대칭.
- 데이터 ClearFormats: 다른 행/열의 마커 무영향(값 생존).
- ★마커가 **값**이라 Excel auto-fill(서식 한정)에 안 휘말림 → Claude의 "auto-fill 방향 비대칭" 우려는 값 마커엔 무관. 새 행/열 = 항상 빈 마커 = 결정적 continuation/unsealed, **침묵 오라벨 0**.

### 10.4 테마색 falsify
색 토큰=explicit hex("111418" 등), `set_cell`이 명시 ARGB 기입 → 우리 생성셀 `_infer_role` 정상. 사용자 Excel 테마색 선택 시 rgb 비-str→fallback(crash 無), band-model[위치]가 role SoT라 봉인. (`_infer_role` 가드 `isinstance(rgb,str)` 유지.)

### 10.5 갱신된 채택 결론
정수 = **숨김 마커 2트랙(바인딩) + 좌표-free band-key 매니페스트(의미) + 기존 set_cell/_infer_role role 기구 + 확장 resolved 비교(numFmt/quotePrefix/dtype/formula) + 침묵모드 방어 + RLE + Table 3차(tabular)**. NamedStyle 미사용. 모든 set_cell 호출부 불변.
