#!/usr/bin/env python3
"""
fpna/design_zones.py — 한 시트 내 "정형 블록 + freehand" 혼합의 영역(zone) 게이트.

설계 SSOT: .consult-design-zones/DESIGN.md (자문 9R + COM 실측 수렴).

핵심(좌표 0 / 행·열 이동 면역):
  · 바인딩(어느 셀=어느 블록) = 숨김 *값* 마커 2트랙
      - 숨김 열 = 행→row-band id (라인아이템, 좌우 공유)   ← ZRB 센티넬로 위치 발견
      - 숨김 행 = 열→col-band id (label/actual/fcst/...)    ← 같은 센티넬 코너
      - sparse band-start + RLE: band 시작에만 마킹, 다음 마커까지 연속.
      - 마커는 *값*이라 Excel auto-fill(서식)에 안 휘말림. 신규 행/열=빈 마커=continuation(중간) / unsealed(검출).
  · 의미(계약) = band-id 키 매니페스트(좌표 free). block_id(r,c)=(row-band[r], col-band[c]).
  · 준수 = 태그 불신 → resolved 재계산(font/fill/numFmt/quotePrefix/data_type) vs house_style 스펙.

NamedStyle 미사용(우리 set_cell 직접포맷과 충돌 + role 이미 폰트색 추정). 모든 기존 set_cell 호출부 불변.
런타임 openpyxl 단일 의존 + stdlib. COM 미사용.
"""
from __future__ import annotations

# 센티넬: 마커 트랙의 코너를 위치 무관하게 찾는 토큰(값이라 셀 이동 따라감).
ZONE_ANCHOR = "​__zone__"          # zero-width prefix(가시성 0) + 마커 코너
SEP = "::"                               # row-band::col-band 결합 키


# --------------------------------------------------------------------------- #
# resolved 값 추출(태그 불신) — 색/서식 정규화로 false drift 방지              #
# --------------------------------------------------------------------------- #
def _dfam(v) -> str:
    """dtype family(결정적). formula 석화 침묵모드 차단 위해 'f' 구분."""
    if v is None:
        return "empty"
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, str):
        return "formula" if v.startswith("=") else "text"
    if isinstance(v, (int, float)):
        return "number"
    return "date"


def _color_key(color):
    """theme/indexed/rgb 정규화 — 같은 색이 표현만 달라 false drift 나는 것 차단."""
    if color is None:
        return None
    t = getattr(color, "type", None)
    if t == "rgb" and getattr(color, "rgb", None):
        return ("rgb", str(color.rgb)[-6:].upper())
    if t == "theme":
        return ("theme", getattr(color, "theme", None), round(float(getattr(color, "tint", 0) or 0), 4))
    if t == "indexed":
        return ("idx", getattr(color, "indexed", None))
    return None


def canon(x):
    """JSON 왕복 안정형(tuple→list 정규화). .contract.json 직렬화 후 비교 시 false drift 차단."""
    import json
    return json.loads(json.dumps(x, ensure_ascii=False, sort_keys=True))


# Excel 재저장이 폰트명을 로케일 별칭으로 정규화(예: "Malgun Gothic"→"맑은 고딕"). 캘리브로
# 확인된 동일 폰트 별칭은 한 canonical 로 접어 false drift 차단(탐지 정확도는 유지).
_FONT_ALIAS = {"맑은 고딕": "Malgun Gothic", "맑은고딕": "Malgun Gothic"}


def _font_name(f):
    n = getattr(f, "name", None)
    return _FONT_ALIAS.get(n, n)


def _norm_numfmt(code):
    """Excel 재저장이 리터럴 괄호를 이스케이프(`(`→`\\(`)하는 것 정규화(재저장 내성)."""
    if not isinstance(code, str):
        return code
    return code.replace("\\(", "(").replace("\\)", ")")


def resolved(cell) -> dict:
    """셀의 실제 렌더링 속성. numFmt/quotePrefix/data_type 포함(S6·석화 차단).
    폰트명은 로케일 별칭 정규화(Excel 재저장 내성)."""
    f, fl, al = cell.font, cell.fill, cell.alignment
    return {
        "font": (_font_name(f), float(getattr(f, "sz", 0) or 0),
                 bool(getattr(f, "b", False)), bool(getattr(f, "i", False)),
                 _color_key(getattr(f, "color", None))),
        "fill": (getattr(fl, "patternType", None),
                 _color_key(getattr(fl, "fgColor", None)) if getattr(fl, "patternType", None) else None),
        "num": _norm_numfmt(cell.number_format),
        "align": (getattr(al, "horizontal", None), getattr(al, "vertical", None)),
        "quote": bool(getattr(cell, "quotePrefix", False)),
        "dtype": cell.data_type,
    }


# --------------------------------------------------------------------------- #
# 바인딩 — 숨김 마커 트랙 stamp/read (sparse band-start + RLE)                  #
# --------------------------------------------------------------------------- #
def stamp_zone(ws, *, origin, row_bands, col_bands, marker_col=None, marker_row=None):
    """블록 영역에 숨김 값 마커 2트랙 기입(생성 측). 좌표는 *기입 시점*에만 쓰고 저장 안 함.

    origin = (row, col) 데이터 좌상단. row_bands = [(start_row, rowband_id), ...](band-start만).
    col_bands = [(start_col, colband_id), ...]. marker_col/row 미지정 시 origin 직전 행/열.
    """
    o_r, o_c = origin
    mc = marker_col if marker_col is not None else o_c - 1
    mr = marker_row if marker_row is not None else o_r - 1
    if mc < 1 or mr < 1:
        raise ValueError("marker col/row 가 시트 밖(origin 을 (2,2) 이상으로)")
    ws.cell(mr, mc).value = ZONE_ANCHOR
    for sr, rid in row_bands:
        ws.cell(sr, mc).value = str(rid)
    for sc, cid in col_bands:
        ws.cell(mr, sc).value = str(cid)
    ws.column_dimensions[ws.cell(mr, mc).column_letter].hidden = True
    ws.row_dimensions[mr].hidden = True
    return {"marker_col": mc, "marker_row": mr}


def _find_anchor(ws):
    """ZONE_ANCHOR 센티넬 위치 발견(좌표 무관, 셀 이동 따라감). 없으면 None."""
    for row in ws.iter_rows():
        for c in row:
            if c.value == ZONE_ANCHOR:
                return c.row, c.column
    return None


def _rle_down(ws, col, start_row, end_row):
    """marker 열을 아래로 RLE 해소 → {row: band_id}. 빈칸=직전 band 연속."""
    out, cur = {}, None
    for r in range(start_row, end_row + 1):
        v = ws.cell(r, col).value
        if v not in (None, ""):
            cur = str(v)
        if cur is not None:
            out[r] = cur
    return out


def _rle_right(ws, row, start_col, end_col):
    out, cur = {}, None
    for c in range(start_col, end_col + 1):
        v = ws.cell(row, c).value
        if v not in (None, ""):
            cur = str(v)
        if cur is not None:
            out[c] = cur
    return out


def read_band_maps(ws):
    """센티넬 기준 마커 2트랙 RLE 해소 → (row_map{row:rowband}, col_map{col:colband}, anchor).
    마커 없으면 ({}, {}, None) — zone 미선언 시트."""
    anchor = _find_anchor(ws)
    if anchor is None:
        return {}, {}, None
    mr, mc = anchor
    row_map = _rle_down(ws, mc, mr + 1, ws.max_row)
    col_map = _rle_right(ws, mr, mc + 1, ws.max_column)
    return row_map, col_map, anchor


# --------------------------------------------------------------------------- #
# 의미 — 좌표 free 계약 + block 해소                                            #
# --------------------------------------------------------------------------- #
def load_contract(obj) -> dict:
    """dict 또는 json 문자열 → 검증된 계약. ⛔ 좌표(셀참조/range) 키 발견 시 거부(stale 원천차단)."""
    import json
    if isinstance(obj, str):
        obj = json.loads(obj)
    blocks = obj.get("blocks", {})
    _COORD = ("range", "ref", "cell", "row_anchor", "col_range", "origin")
    for bid, spec in blocks.items():
        for k in spec:
            if k in _COORD:
                raise ValueError("계약 %r 에 좌표 키 %r — 좌표 free 위반(매니페스트는 band-id 키만)" % (bid, k))
        spec.setdefault("bands", [])
        spec.setdefault("colsig", {})
        spec.setdefault("house_style", {})
    obj.setdefault("band_map", {})       # (rowband SEP colband) → block_id, ragged 전용
    return obj


def block_of(rowband, colband, contract) -> str:
    """(row-band, col-band) → block_id. ragged band_map 우선, 없으면 결합키."""
    key = "%s%s%s" % (rowband, SEP, colband)
    return contract.get("band_map", {}).get(key, key)


def _trim_to_data(ws, row_map, col_map):
    """RLE 가 max_row/col 까지 과확장하는 것 bound — 블록 컬럼/행에 데이터 있는 마지막
    행·열까지로 잘라 블록 밖(아래/오른쪽) freehand 흡수를 차단. trailing 빈 영역 제거."""
    if not row_map or not col_map:
        return row_map, col_map
    rows, cols = sorted(row_map), sorted(col_map)
    last_r = rows[0]
    for r in rows:
        if any(ws.cell(r, c).value not in (None, "") for c in cols):
            last_r = r
    last_c = cols[0]
    for c in cols:
        if any(ws.cell(r, c).value not in (None, "") for r in rows if r <= last_r):
            last_c = c
    return ({r: b for r, b in row_map.items() if r <= last_r},
            {c: b for c, b in col_map.items() if c <= last_c})


def resolve_blocks(ws, contract):
    """시트 내 각 데이터 셀의 (rowband, colband, block_id) 해소.
    반환 (cells{(r,c):(rowband,colband,block_id)}, row_map, col_map). 마커 교차점만.
    RLE 는 데이터 extent 까지만 bound(trailing 빈 영역·블록 밖 흡수 차단)."""
    row_map, col_map, anchor = read_band_maps(ws)
    if anchor is None:
        return {}, {}, {}
    row_map, col_map = _trim_to_data(ws, row_map, col_map)
    cells = {}
    for r, rb in row_map.items():
        for c, cb in col_map.items():
            cells[(r, c)] = (rb, cb, block_of(rb, cb, contract))
    return cells, row_map, col_map


def draw_house_block(ws, *, origin, rows, col_bands, bands_by_col, row_band="block",
                     marker_col=None, marker_row=None, contract=None):
    """생성 primitive — 한 블록을 set_cell(role)로 그리고 숨김 마커 2트랙 stamp + 계약 캡처.

    불변식(DESIGN §5): 룩은 set_cell 직접포맷으로(NamedStyle 미사용), 마커는 값, 빈셀 포함
    직사각 전체 커버, 계약은 _resolved() 동일 직렬화로 캡처. row_band 단일(공유 라인아이템).
      origin=(r0,c0). rows = [[v,v,...], ...] (행별 값 리스트, 열 순서 = col_bands 순서 평탄화).
      col_bands = [(start_col, colband_id), ...]. bands_by_col = {colband_id: role}.
    """
    from fpna import house_style as hs
    r0, c0 = origin
    cols = [c for (c, _id) in col_bands]
    span = (max(cols) + 1) - c0 if cols else 0
    colband_of = {}
    cur = None
    for cc in range(c0, c0 + span):
        for (sc, cid) in col_bands:
            if sc == cc:
                cur = cid
        colband_of[cc] = cur
    for i, rowvals in enumerate(rows):
        r = r0 + i
        for j, v in enumerate(rowvals):
            c = c0 + j
            role = bands_by_col.get(colband_of.get(c), "calc")
            hs.set_cell(ws, r, c, v, role=role, number_format=hs.FMT_INT if isinstance(v, (int, float)) else None)
    stamp_zone(ws, origin=origin, row_bands=[(r0, row_band)], col_bands=col_bands,
               marker_col=marker_col, marker_row=marker_row)
    return capture_contract(ws, bands_by_col, contract=contract)


def capture_contract(ws, bands_by_col, *, contract=None):
    """생성 직후 블록에서 block_id(=rowband::colband)별 resolved 스펙을 캡처(생성↔검증 대칭, 불변식 ④).

    bands_by_col = {colband_id: role_name}(set_cell role). block_id 마다 대표 셀의 resolved 를
    house_style 스펙으로 저장. _resolved() 와 *동일 직렬화*라 false drift 0.
    """
    contract = contract or {"blocks": {}, "band_map": {}}
    cells, row_map, col_map = resolve_blocks(ws, contract)
    blocks = contract.setdefault("blocks", {})
    for (r, c), (rb, cb, bid) in cells.items():
        if cb in bands_by_col and bid not in blocks:
            blocks[bid] = {"spec": resolved(ws.cell(r, c)), "role": bands_by_col[cb]}
    contract.setdefault("band_map", {})
    return contract
