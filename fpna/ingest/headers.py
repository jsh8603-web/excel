"""
fpna.ingest.headers — 셀 분류 + 병합 fill + behead 언피벗.

unpivotr behead/enhead 개념을 파이썬으로 재구현(소스 복붙 아님):
- 셀 분류: 타입(numeric=data / character=header) 1차 + 위치 보조.
- unmerge_fill: 병합 anchor 값을 영역 내 전파 → 이후 up/left 2방향 매칭으로 단순화.
- behead: 가장 바깥 헤더 1줄을 데이터 셀에 부착하고 그 헤더를 제거(반복).
- nearest_header: 방향(up/left/up-left/left-up) 기하 최근접 매칭.

⚠ up-left/left-up 정렬 우선순위는 리서치에서 일부 추론 — 골든샘플 테스트로 검증.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .cells import (Cell, DATA, HEADER, LABEL, CORNER, BLANK,
                    T_NUMERIC, T_DATE, T_CHARACTER, T_ERROR)
from .detect import Block, UNIT_RE


def unmerge_fill(cells: list[Cell]) -> None:
    """병합영역 anchor 의 값/타입을 영역 내 빈 셀에 전파(in-place).

    openpyxl 은 병합셀 중 anchor 만 값 보유 → 나머지에 같은 값 채워
    헤더 방향매칭이 직상/직좌만으로 해소되게 한다.
    """
    pos = {(c.row, c.col): c for c in cells}
    anchors: dict[tuple[int, int], Cell] = {}
    for c in cells:
        if c.merge_anchor and c.merge_anchor == (c.row, c.col):
            anchors[c.merge_anchor] = c
    for c in cells:
        if c.merge_anchor and c.merge_anchor != (c.row, c.col):
            a = anchors.get(c.merge_anchor) or pos.get(c.merge_anchor)
            if a is not None and not a.is_blank:
                c.value = a.value
                c.data_type = a.data_type
                c.is_blank = False
                c.fmt = a.fmt


import re as _re

# 숫자값 '마커'(괄호/세모/%/콤마/소수점/부호). bare 4자리(연도)와 데이터 숫자를 구분.
_MARKER_RE = _re.compile(r"[(),%△▲\-+.]|\d{1,3},\d{3}")


def fill_down_ditto(cells: list[Cell], cols: list[int],
                    row_range: tuple[int, int]) -> set[tuple[int, int]]:
    """카테고리 열의 빈칸(상동)을 직전 값으로 단방향(위→아래) 충전(in-place).

    무음 손상 방어 ③: ditto 빈칸 미충전 → 행 의미 깨짐.
    - 결정적: 같은 입력 → 같은 출력. 같은 열에서 위→아래로만 전파.
    - 첫 비빈칸 이전의 선두 빈칸은 채우지 않음(부모 없음).
    - 채운 셀의 (row,col) 집합 반환 → 호출측이 DITTO_FILLED 플래그 부착.
    """
    pos = {(c.row, c.col): c for c in cells}
    filled: set[tuple[int, int]] = set()
    r0, r1 = row_range
    for col in cols:
        last = None  # (value, data_type, fmt)
        for r in range(r0, r1 + 1):
            c = pos.get((r, col))
            if c is None:
                continue
            if not c.is_blank:
                last = (c.value, c.data_type)
            elif last is not None:
                c.value, c.data_type = last
                c.is_blank = False
                filled.add((r, col))
    return filled


def _value_like(c: Cell) -> bool:
    """이 셀이 '데이터 값'처럼 보이는가?

    - 실제 저장 숫자/날짜 → True
    - 텍스트지만 숫자마커(괄호/△/%/콤마/소수점)를 동반 → True (예: (50), △30, 85%, 1,234)
    - bare 정수 텍스트(예: '2024') → False (연도/헤더 가능성, 모호하므로 헤더 취급)
    - 그 외 텍스트 → False
    """
    if c.is_blank:
        return False
    # 오류값(#DIV/0! 등)도 데이터 슬롯을 차지 → 열을 데이터 열로 유지(보존 위해).
    # 정수 연도 헤더 보호는 셀 단위가 아니라 행 맥락(_is_year_header_row)에서 처리한다 —
    # 데이터 값이 우연히 연도 범위(예 1932)여도 값으로 유지하기 위함.
    if c.data_type in (T_NUMERIC, T_DATE, T_ERROR):
        return True
    if c.data_type == T_CHARACTER:
        s = str(c.value).strip()
        if UNIT_RE.search(s):              # '(단위: ...)' 안내문 = 비값
            return False
        # @(텍스트)서식 + 숫자 내용 = text-as-num 데이터(Excel SUM 누락 지점). 값으로 인식.
        if c.fmt.number_format == "@" and _re.fullmatch(r"-?\d[\d,]*\.?\d*", s):
            return True
        if _re.fullmatch(r"\d{1,4}", s):   # bare 정수(연도 포함) = 비값
            return False
        if _MARKER_RE.search(s):
            return True
    return False


def _col_dominant_type(block_cells: list[Cell], b: Block,
                       data_start_row: int | None = None) -> dict[int, str]:
    """열별 다수결 타입. data_start_row 지정 시 그 행 이상(데이터 영역)만 집계."""
    out: dict[int, str] = {}
    for col in range(b.min_col, b.max_col + 1):
        num = chr_ = 0
        for c in block_cells:
            if c.col != col or c.is_blank:
                continue
            if data_start_row is not None and c.row < data_start_row:
                continue
            if _value_like(c):
                num += 1
            elif c.data_type == T_CHARACTER:
                chr_ += 1
        out[col] = "num" if num > chr_ and num > 0 else ("chr" if chr_ else "blank")
    return out


def _row_value_frac(block_cells: list[Cell], row: int) -> float:
    rc = [c for c in block_cells if c.row == row and not c.is_blank]
    if not rc:
        return -1.0  # 빈 행
    return sum(1 for c in rc if _value_like(c)) / len(rc)


def _is_year_header_row(block_cells: list[Cell], row: int) -> bool:
    """행의 정수 숫자셀이 2개 이상이고 전부 연도범위(1900~2100)면 연도 헤더행.

    정수 연도 헤더(2024,2025)는 _value_like 상 값이라 데이터로 오인되기 쉬운데,
    '숫자가 전부 연도' 패턴으로 헤더임을 식별한다. 데이터행은 값 중 일부만 우연히
    연도범위(예 1932)라 all() 에서 걸러져 값으로 유지된다.
    """
    nums = [c.value for c in block_cells
            if c.row == row and c.data_type == T_NUMERIC
            and isinstance(c.value, int) and not isinstance(c.value, bool)]
    return len(nums) >= 2 and all(1900 <= v <= 2100 for v in nums)


def _col_value_frac(block_cells: list[Cell], col: int, data_start_row: int) -> float:
    cc = [c for c in block_cells if c.col == col and not c.is_blank
          and c.row >= data_start_row]
    if not cc:
        return -1.0
    return sum(1 for c in cc if _value_like(c)) / len(cc)


def classify_cells(block_cells: list[Cell], b: Block) -> tuple[int, int]:
    """블록 내 셀을 corner/header/data/label/blank 로 분류(in-place).

    반환: (top_header_rows, left_header_cols) = 헤더 밴드 두께.
    2-pass:
      1) top band = 위에서부터 'value_like 비율 ≥ 0.5' 인 행 직전까지(헤더는 텍스트 우세).
      2) left band = 왼쪽에서부터, 데이터행 기준 value_like 비율 < 0.5 인 열.
      3) 데이터 영역만으로 열 타입 재집계 → 값 분류.
    """
    # 1) top header band
    top_rows = 0
    for r in range(b.min_row, b.max_row + 1):
        frac = _row_value_frac(block_cells, r)
        # 연도 헤더행(숫자가 전부 1900~2100)은 frac 높아도 헤더로 — 정수연도 헤더 보호.
        # 데이터행은 값 중 일부만 우연히 연도라 _is_year_header_row 에서 걸러진다.
        if frac >= 0.5 and not _is_year_header_row(block_cells, r):
            break
        top_rows += 1          # frac < 0.5(헤더) 또는 빈 행
    if top_rows > b.n_rows - 1:    # 전부 헤더로 잡히면(데이터 없음) 마지막 1행은 데이터로
        top_rows = max(0, b.n_rows - 1)
    data_start_row = b.min_row + top_rows
    header_row_end = data_start_row - 1

    # 2) left header band (데이터행 기준)
    left_cols = 0
    for col in range(b.min_col, b.max_col + 1):
        frac = _col_value_frac(block_cells, col, data_start_row)
        if frac >= 0.5:
            break
        left_cols += 1
    if left_cols > b.n_cols - 1:
        left_cols = max(0, b.n_cols - 1)
    header_col_end = b.min_col + left_cols - 1

    # 3) 데이터 영역 열 타입
    col_type = classify_cells._col_type = _col_dominant_type(
        block_cells, b, data_start_row=data_start_row)

    for c in block_cells:
        if c.is_blank:
            c.cls = BLANK
            continue
        in_top = c.row <= header_row_end
        in_left = c.col <= header_col_end
        if in_top and in_left:
            c.cls = CORNER
        elif in_top:
            c.cls = HEADER          # 열헤더
        elif in_left:
            c.cls = LABEL           # 행헤더/라벨
        else:
            # 데이터 영역: 값처럼 보이거나(괄호음수/△/%/콤마 포함) 열이 숫자우세면 DATA.
            if _value_like(c) or col_type.get(c.col) == "num":
                c.cls = DATA
            else:
                c.cls = LABEL
    return top_rows, left_cols


# --------------------------------------------------------------------------
# behead 방향 매칭
# --------------------------------------------------------------------------
def nearest_header(d: Cell, headers: list[Cell], direction: str) -> Cell | None:
    if direction == "up":
        cand = [h for h in headers if h.col == d.col and h.row < d.row]
        return max(cand, key=lambda h: h.row, default=None)
    if direction == "left":
        cand = [h for h in headers if h.row == d.row and h.col < d.col]
        return max(cand, key=lambda h: h.col, default=None)
    if direction == "up-left":
        cand = [h for h in headers if h.col <= d.col and h.row < d.row]
        cand.sort(key=lambda h: (-h.col, -h.row))
        return cand[0] if cand else None
    if direction == "left-up":
        cand = [h for h in headers if h.row <= d.row and h.col < d.col]
        cand.sort(key=lambda h: (-h.row, -h.col))
        return cand[0] if cand else None
    raise ValueError("unknown direction: %s" % direction)


@dataclass
class LongRow:
    value: object
    attrs: dict = field(default_factory=dict)
    row: int = 0
    col: int = 0
    unit: str | None = None
    is_subtotal: bool = False
    level: int = 0
    # G3/G5 신호(결정적). 가장 안쪽 행라벨 셀에서 수집.
    label_bold: bool = False          # 행라벨 폰트 볼드(소계 신호)
    label_indent: int = 0             # alignment.indent + 선행공백 환산 레벨(계층)
    cell_red: bool = False            # 데이터 셀 빨강폰트(색 음수 신호)
    number_format: str = "General"    # 데이터 셀 표시서식(text-as-num/mixed 탐지용)


def unpivot_block(block_cells: list[Cell], b: Block, *,
                  col_header_names: list[str] | None = None,
                  row_header_names: list[str] | None = None) -> list[LongRow]:
    """블록을 long 행 리스트로 변환.

    col_header_names : 열헤더 레벨별 attr 이름(바깥→안). 미지정 시 hdr_c{n}.
    row_header_names : 행헤더 레벨별 attr 이름(바깥→안). 미지정 시 hdr_r{n}.
    """
    unmerge_fill(block_cells)
    top_rows, left_cols = classify_cells(block_cells, b)

    headers_col = [c for c in block_cells if c.cls == HEADER]
    headers_row = [c for c in block_cells if c.cls == LABEL and c.col <= b.min_col + left_cols - 1]
    data_cells = [c for c in block_cells if c.cls == DATA]

    from .normalize import leading_space_level, strip_footnote_marker

    import bisect as _bisect

    # ⚠ 루프 불변식 선계산: 헤더 레벨/밴드를 데이터 셀마다 재계산하면 O(셀×헤더)=O(n²)
    #   (행라벨은 행당 1개라 큰 시트에서 폭주). 밴드는 1회만 만든다.
    col_levels = sorted({h.row for h in headers_col})
    col_bands = {hr: [h for h in headers_col if h.row == hr] for hr in col_levels}
    row_levels = sorted({h.col for h in headers_row})
    # 행라벨 밴드: 열 고정 → row 오름차순 1회 정렬 + bisect 로 'row<=d.row 최대' O(log n).
    row_bands: dict = {}
    for hc in row_levels:
        band = sorted((h for h in headers_row if h.col == hc), key=lambda h: h.row)
        row_bands[hc] = (band, [h.row for h in band])

    out: list[LongRow] = []
    for d in data_cells:
        attrs: dict = {}
        level = 0
        # 열헤더 레벨(위→아래 여러 행) — 밴드가 작아 선형 매칭 유지.
        for i, hr in enumerate(col_levels):
            h = nearest_header(d, col_bands[hr], "up-left")
            name = (col_header_names[i] if col_header_names and i < len(col_header_names)
                    else "hdr_c%d" % i)
            # G7: 열헤더 각주마커 제거 → 동일 논리열 키 통일.
            hv = h.value if h else None
            if isinstance(hv, str):
                hv, _ = strip_footnote_marker(hv)
            attrs[name] = hv
        # 행헤더 레벨(왼→오 여러 열) — left-up = col<d.col, row<=d.row 중 최대 row.
        inner_label = None     # 가장 안쪽(오른쪽) 행라벨 셀
        for i, hc in enumerate(row_levels):
            h = None
            if hc < d.col:                      # nearest_header 의 h.col < d.col 조건과 동일
                band, brows = row_bands[hc]
                j = _bisect.bisect_right(brows, d.row) - 1   # row <= d.row 중 최대
                if j >= 0:
                    h = band[j]
            name = (row_header_names[i] if row_header_names and i < len(row_header_names)
                    else "hdr_r%d" % i)
            attrs[name] = h.value if h else None
            if h is not None:
                level = max(level, h.fmt.indent)
                inner_label = h
        # G3: 가장 안쪽 라벨의 indent + 선행공백 → 계층 레벨.
        # G5: 라벨 볼드 / 데이터 셀 빨강폰트 신호.
        label_bold = False
        label_indent = 0
        if inner_label is not None:
            label_bold = bool(inner_label.fmt.bold)
            sp_level = leading_space_level(inner_label.value)
            label_indent = inner_label.fmt.indent + sp_level
            level = max(level, label_indent)
        cell_red = bool(d.fmt.font_color and str(d.fmt.font_color).upper()[-6:]
                        in ("FF0000", "C00000"))
        out.append(LongRow(value=d.value, attrs=attrs, row=d.row, col=d.col,
                           level=level, label_bold=label_bold,
                           label_indent=label_indent, cell_red=cell_red,
                           number_format=d.fmt.number_format))
    return out


def no_header_suspect(block_cells: list[Cell], b: Block, top_rows: int) -> bool:
    """no-header 의심(smell 신호용 — 격하/정제변경 안 함, 오탐 회피).

    MetaCollector GuessHeaderRange 흡수: 헤더 후보 마지막 행과 첫 데이터행의 열별
    data_type 이 ≥50% 일치하고, 헤더 후보행에 숫자/날짜가 섞이면(헤더답지 않음) True.
    가드: 헤더행이 전부 텍스트면(정상 헤더) False → 정상 표 오탐 차단.
    """
    if top_rows < 1 or top_rows >= b.n_rows:
        return False
    hdr_r = b.min_row + top_rows - 1
    dat_r = b.min_row + top_rows
    pos: dict[tuple[int, int], str] = {}
    for c in block_cells:
        if c.row in (hdr_r, dat_r) and not c.is_blank:
            pos[(c.row, c.col)] = c.data_type
    cols = sorted({col for (_r, col) in pos})
    if not cols:
        return False
    match = 0
    hdr_has_value = False
    for col in cols:
        ht = pos.get((hdr_r, col))
        dt = pos.get((dat_r, col))
        if ht in (T_NUMERIC, T_DATE):
            hdr_has_value = True
        if ht is not None and ht == dt:
            match += 1
    return hdr_has_value and (match / len(cols)) >= 0.5


__all__ = ["unmerge_fill", "classify_cells", "nearest_header", "no_header_suspect",
           "unpivot_block", "LongRow", "fill_down_ditto"]
