"""
fpna.ingest.detect — 블록(다중 표) 탐지 + 비데이터 행 격리.

기법 출처(개념만 재구현):
- 영역탐지: connected-component + density 임계 (USPTO 11341322 / Pytheas coherency).
  채워진 셀을 그래프로 보고 인접 연결요소를 찾되, 직사각 bbox 의 채움밀도와
  최소 크기로 진짜 표만 채택. 빈행/빈열 gap tolerance 로 과분할 방지.
- 비데이터 행 격리: 제목/단위주석/각주/반복헤더 (Pytheas 4-way 분류 단순화).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from .cells import Cell, T_CHARACTER, T_NUMERIC, T_DATE, T_ERROR

# 영역탐지 파라미터(보수적 기본값)
MIN_COLS = 2
MIN_ROWS = 3
MIN_DENSITY = 0.55          # bbox 내 채움 비율 하한(USPTO 0.7 보다 완화 — 재무표 빈칸 많음)
# gap=0 = 직접 인접(8방향)만 연결. 표 사이 1칸 구분열/행을 경계로 인식(다중표 분리 우선).
# 표 내부에 완전 빈 구분열이 있어 과분할되면 gap=1 로 올린다(단, 인접 표 병합 위험).
GAP_TOL = 0

UNIT_RE = re.compile(r"\(?\s*단위\s*[:：]?\s*([^\)\]]+)\)?")
SUBTOTAL_KW = ("합계", "합 계", "소계", "소 계", "총계", "총 계", "누계",
               "계", "total", "subtotal", "grand total", "sum")


@dataclass
class Block:
    min_row: int
    max_row: int
    min_col: int
    max_col: int

    def contains(self, r: int, c: int) -> bool:
        return self.min_row <= r <= self.max_row and self.min_col <= c <= self.max_col

    @property
    def n_rows(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def n_cols(self) -> int:
        return self.max_col - self.min_col + 1


def _occupied(cells: list[Cell]) -> set[tuple[int, int]]:
    return {(c.row, c.col) for c in cells if not c.is_blank}


def _union_find_components(occ: set[tuple[int, int]], gap: int) -> list[set]:
    """gap 칸까지 인접(8-방향 + gap dilation)으로 연결요소 분할."""
    parent: dict[tuple[int, int], tuple[int, int]] = {p: p for p in occ}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    occ_list = sorted(occ)
    # gap 범위 내 이웃을 연결 (대각 포함)
    for (r, c) in occ_list:
        for dr in range(-gap - 1, gap + 2):
            for dc in range(-gap - 1, gap + 2):
                if dr == 0 and dc == 0:
                    continue
                nb = (r + dr, c + dc)
                if nb in parent:
                    union((r, c), nb)
    comps: dict[tuple[int, int], set] = {}
    for p in occ:
        root = find(p)
        comps.setdefault(root, set()).add(p)
    return list(comps.values())


def detect_blocks(cells: list[Cell], *, gap: int = GAP_TOL,
                  min_cols: int = MIN_COLS, min_rows: int = MIN_ROWS,
                  min_density: float = MIN_DENSITY) -> list[Block]:
    """다중 표 블록 탐지. 연결요소 → bbox → 크기·밀도 게이트."""
    occ = _occupied(cells)
    if not occ:
        return []
    blocks: list[Block] = []
    for comp in _union_find_components(occ, gap):
        rows = [p[0] for p in comp]
        cols = [p[1] for p in comp]
        b = Block(min(rows), max(rows), min(cols), max(cols))
        area = b.n_rows * b.n_cols
        density = len(comp) / area if area else 0.0
        # 채움밀도 게이트. 단 좌측 ditto 라벨컬럼(반복 빈칸)이 많은 큰 표는 전체 density
        # 가 낮아도 정당한 표다(값열·최내측 키열은 매행 채워짐). USPTO 단일 density 가
        # '4키 ditto + 값열 1개' 류를 통째로 탈락시켜 tidy 0 무음손실 → 커버리지로 보강:
        #   행 커버(채워진 셀 있는 행 비율) 높고 + 매행 가까이 채워진 dense 컬럼 존재 →
        #   산발 노이즈가 아닌 columnar 표로 인정.
        nrows = b.n_rows
        row_cover = len(set(rows)) / nrows if nrows else 0.0
        col_fill: dict = {}
        for cc in cols:
            col_fill[cc] = col_fill.get(cc, 0) + 1
        dense_cols = sum(1 for n in col_fill.values() if n >= 0.5 * nrows)
        columnar = row_cover >= 0.9 and dense_cols >= 1
        if (b.n_cols >= min_cols and b.n_rows >= min_rows
                and (density >= min_density or columnar)):
            blocks.append(b)
    # 위→아래, 좌→우 정렬(결정성)
    blocks.sort(key=lambda x: (x.min_row, x.min_col))
    return blocks


def absorb_header_bands(cells: list[Cell], blocks: list[Block], *,
                        max_gap: int = 4, max_absorb: int = 6) -> list[Block]:
    """데이터블록 위쪽에 빈행 gap 으로 떨어진 헤더밴드를 블록에 흡수(min_row 확장).

    동기(ONS류 실데이터): 다층 메타(제목/발행일/지역) + 멀티헤더(지표/단위/코드) +
    빈행 1줄 + 시계열 데이터. 헤더밴드는 텍스트 위주라 bbox 채움밀도가 MIN_DENSITY
    미달 → 별도 블록조차 안 잡히고, 빈행이 데이터블록과 분리 → 데이터블록에 헤더 0 →
    behead 컬럼명 None → tidy 0(무음 실패). 데이터블록 위 인접 헤더밴드를 끌어와 해소.

    흡수 조건(모두 충족해야 한 행 흡수):
      (a) 빈행은 gap 으로 건너뜀(누적 ≤ max_gap, 초과 시 중단).
      (b) 그 행의 채워진 셀 중 블록 열범위[min_col,max_col] 내 개수가 블록폭의 절반↑
          (제목/단위 같은 1~2칸 메타행은 폭 미달로 자동 배제 — 다중표 회귀 안전).
      (c) 텍스트 위주(블록 내 셀의 숫자/날짜 비율 ≤ 0.5) — 숫자 데이터행 오흡수 차단.
      (d) 위가 다른 블록 영역이면 중단(가로/세로 다중표 보호).
    블록 수는 불변(min_row 만 낮춤). 결정적.
    """
    if not blocks:
        return blocks
    occ_by_row: dict[int, set] = {}
    numlike: set = set()
    for c in cells:
        if c.is_blank:
            continue
        occ_by_row.setdefault(c.row, set()).add(c.col)
        if c.data_type in (T_NUMERIC, T_DATE):
            numlike.add((c.row, c.col))
    # 행 → 소속 블록 인덱스(unmerge 후엔 헤더밴드도 별도 블록일 수 있어 '병합'이 필요).
    owner: dict[int, int] = {}
    for idx, b in enumerate(blocks):
        for r in range(b.min_row, b.max_row + 1):
            owner[r] = idx
    removed: set = set()

    for idx, b in enumerate(blocks):
        if idx in removed:
            continue
        # 헤더밴드 정렬 폭 하한. 하한 3 → 좁은 표에서 'A2:라벨 / B2:값' 2칸 메타행
        # (발행일·연락처) 배제. //3 = 계층병합 밴드(통계청 산업분류: 행마다 채움 듬성)도
        # 흡수하되 1~2칸 메타는 여전히 배제(ONS 19칸·g13 발행일 2칸 모두 정합).
        half = max(3, b.n_cols // 3)
        gap = 0
        rr = b.min_row - 1
        new_min = b.min_row
        absorbed = 0
        while rr >= 1 and absorbed < max_absorb:
            cols = occ_by_row.get(rr)
            if not cols:                         # (a) 빈행 — gap 건너뜀
                gap += 1
                if gap > max_gap:
                    break
                rr -= 1
                continue
            oidx = owner.get(rr)
            if oidx == idx:                      # 자기 블록 도달 — 중단
                break
            in_cols = [c for c in cols if b.min_col <= c <= b.max_col]
            if len(in_cols) < half:              # (b) 폭 미달(메타행) — 중단
                break
            nnum = sum(1 for c in in_cols if (rr, c) in numlike)
            if nnum > len(in_cols) * 0.5:        # (c) 숫자 위주 = 데이터행
                break
            if oidx is not None and oidx != idx:
                # (d) 위가 헤더성 별도 블록 → 데이터 열범위와 정렬되는 '하단 연속 행'만
                #     흡수한다. 위쪽 메타행(제목/발행일/문의 — 폭 미달)은 경계에서 제외해
                #     metric/period 오염(발행일이 period 로 새는 것)을 막는다. 블록은 제거.
                ob = blocks[oidx]
                m = ob.max_row
                while m >= ob.min_row:
                    mcols = occ_by_row.get(m, set())
                    m_in = [c for c in mcols if b.min_col <= c <= b.max_col]
                    m_num = sum(1 for c in m_in if (m, c) in numlike)
                    if len(m_in) >= half and m_num <= len(m_in) * 0.5:
                        m -= 1
                    else:
                        break
                band_top = m + 1                 # 정렬 헤더밴드 시작(메타행 위로 제외)
                if band_top <= ob.max_row:        # 정렬 행이 하나라도 있으면 흡수
                    new_min = min(new_min, band_top)
                removed.add(oidx)
                break
            # 블록 미소속 occupied 헤더밴드 → 한 행씩 흡수.
            new_min = rr
            absorbed += 1
            gap = 0
            rr -= 1
        if new_min < b.min_row:
            b.min_row = new_min
    out = [b for i, b in enumerate(blocks) if i not in removed]
    out.sort(key=lambda x: (x.min_row, x.min_col))
    return out


def cells_in_block(cells: list[Cell], b: Block) -> list[Cell]:
    return [c for c in cells if b.contains(c.row, c.col)]


# --------------------------------------------------------------------------
# 비데이터 행 격리 (제목 / 단위 / 각주)
# --------------------------------------------------------------------------
def _row_cells(block_cells: list[Cell], row: int) -> list[Cell]:
    return [c for c in block_cells if c.row == row and not c.is_blank]


_MARKER_RE = re.compile(r"[(),%△▲]|\d{1,3},\d{3}")


def _looks_value(c: Cell) -> bool:
    if c.data_type in (T_NUMERIC, T_DATE, T_ERROR):
        return True
    if isinstance(c.value, str):
        s = c.value.strip()
        # '(단위: 백만원)' 류 안내문은 괄호 때문에 값처럼 보이나 비데이터 행.
        if UNIT_RE.search(s):
            return False
        if re.fullmatch(r"\d{1,4}", s):
            return False
        return bool(_MARKER_RE.search(s))
    return False


def strip_title_rows(block_cells: list[Cell], b: Block) -> tuple[list[Cell], list[int], dict]:
    """블록 상단의 제목/단위 안내 행 격리(블록 폭 상대 기준).

    제목 행 = 위에서부터, '채워진 셀이 블록 폭에 비해 극히 적고(≤2 그리고 폭의 1/3 미만)
    값처럼 보이는 셀이 없는' 행. 좁은 표(n_cols<4)는 제목 오인 방지를 위해 폭 기준만 적용.
    단위 안내'(단위: ...)'는 행 제거 여부와 무관하게 스캔.
    반환: (남은 셀, 제거한 행번호, {"unit": ...})
    """
    title_rows: list[int] = []
    unit_meta: dict = {}
    width_thresh = max(1, b.n_cols // 3)
    for r in range(b.min_row, b.max_row + 1):
        rc = _row_cells(block_cells, r)
        n_filled = len(rc)
        has_value = any(_looks_value(c) for c in rc)
        for c in rc:
            if isinstance(c.value, str):
                m = UNIT_RE.search(c.value)
                if m:
                    unit_meta["unit"] = m.group(1).strip()
        # 제목 판정: 폭 넓은 표에서 1~2칸짜리 비값 행만(좁은 표 보호)
        is_title = (n_filled > 0 and not has_value
                    and n_filled <= 2 and n_filled <= width_thresh
                    and b.n_cols >= 4)
        if is_title:
            title_rows.append(r)
        elif has_value:
            break
    remaining = [c for c in block_cells if c.row not in title_rows]
    return remaining, title_rows, unit_meta


def strip_footnote_rows(block_cells: list[Cell], b: Block) -> tuple[list[Cell], list[int]]:
    """블록 하단의 각주 행 격리(주1)/* 로 시작하거나 1셀 character 행)."""
    foot_rows: list[int] = []
    for r in range(b.max_row, b.min_row - 1, -1):
        rc = _row_cells(block_cells, r)
        if not rc:
            continue
        n_numeric = sum(1 for c in rc if c.data_type == T_NUMERIC)
        first_text = next((c.value for c in sorted(rc, key=lambda x: x.col)
                           if isinstance(c.value, str)), "")
        is_note = bool(re.match(r"^\s*(주\s*\d|\*|注|※|Note)", str(first_text)))
        if (len(rc) <= 1 and n_numeric == 0) or is_note:
            foot_rows.append(r)
        else:
            break
    remaining = [c for c in block_cells if c.row not in foot_rows]
    return remaining, foot_rows


def _row_signature(block_cells: list[Cell], row: int) -> tuple:
    """행의 (col, 정규화된 텍스트) 시그니처. 헤더 동일성 비교용(값셀은 제외)."""
    out = []
    for c in sorted(_row_cells(block_cells, row), key=lambda x: x.col):
        if isinstance(c.value, str):
            out.append((c.col, c.value.strip().lower().replace(" ", "")))
    return tuple(out)


def _index_by_row(block_cells: list[Cell]) -> dict:
    """행번호 → 비빈칸 셀 리스트(1회 구축). 행마다 전체 셀 재스캔(O(rows×cells)) 회피."""
    idx: dict = {}
    for c in block_cells:
        if not c.is_blank:
            idx.setdefault(c.row, []).append(c)
    return idx


def _row_signature_of(rc: list[Cell]) -> tuple:
    """행 셀들에서 (col, 정규화 텍스트) 시그니처. 값셀 제외(_row_signature 와 동일 로직)."""
    out = []
    for c in sorted(rc, key=lambda x: x.col):
        if isinstance(c.value, str):
            out.append((c.col, c.value.strip().lower().replace(" ", "")))
    return tuple(out)


def strip_repeated_header_rows(block_cells: list[Cell], b: Block,
                               header_rows: list[int]) -> tuple[list[Cell], list[int]]:
    """G8: 페이지브레이크로 표 중간 재삽입된 헤더행(첫 헤더와 동일) 제거.

    header_rows = 블록 상단 헤더 밴드 행번호들. 그 시그니처(텍스트만)와 동일한
    데이터영역 행을 찾아 제거 → 데이터에서 배제. 반환: (남은셀, 제거행번호).
    숫자가 1개라도 있는 행은 진짜 데이터로 보고 보존(헤더 오인 방지).
    """
    if not header_rows:
        return block_cells, []
    by_row = _index_by_row(block_cells)          # ⚠ 행 인덱스 1회 → 루프 내 재스캔 제거(O(n²)→O(n))
    header_sigs = {_row_signature_of(by_row.get(r, [])) for r in header_rows}
    header_sigs.discard(())
    if not header_sigs:
        return block_cells, []
    data_start = max(header_rows) + 1
    dropped: list[int] = []
    for r in range(data_start, b.max_row + 1):
        rc = by_row.get(r)
        if not rc:
            continue
        if any(c.data_type == T_NUMERIC for c in rc):
            continue  # 숫자 동반 = 데이터 행
        if _row_signature_of(rc) in header_sigs:
            dropped.append(r)
    if not dropped:
        return block_cells, []
    dropped_set = set(dropped)
    remaining = [c for c in block_cells if c.row not in dropped_set]
    return remaining, dropped


def label_is_subtotal(label: str) -> bool:
    if not label:
        return False
    s = str(label).strip().lower().replace(" ", "")
    return any(kw.replace(" ", "") in s for kw in SUBTOTAL_KW)


# --------------------------------------------------------------------------
# G5 색/볼드 소계 + 색 음수
# --------------------------------------------------------------------------
# 빨강 폰트 화이트리스트(음수 신호). _color_hex 는 끝 6자리 대문자 → FF0000 정규화.
RED_FONT_WHITELIST = {"FF0000", "C00000", "FF0000FF"[-6:]}


def is_red_font(color_hex) -> bool:
    """폰트색이 '회계 빨강 음수' 화이트리스트인가. (FFFF0000→FF0000 정규화 가정)."""
    if not color_hex:
        return False
    return str(color_hex).upper()[-6:] in RED_FONT_WHITELIST


def subtotal_signal_score(label, *, bold: bool = False,
                          arith_match: bool = False) -> tuple[int, dict]:
    """소계 3신호(label/bold/arith) 교집합 점수.

    G5: label(소계/합계/계 정규식) + bold(font.bold) + 산술(row합==sibling합).
    반환: (score 0~3, {signal:bool}). score≥2 이면 호출측이 is_subtotal 판정.
    """
    sig = {
        "label": label_is_subtotal(label),
        "bold": bool(bold),
        "arith": bool(arith_match),
    }
    return sum(1 for v in sig.values() if v), sig


__all__ = [
    "Block", "detect_blocks", "absorb_header_bands", "cells_in_block",
    "strip_title_rows", "strip_footnote_rows", "strip_repeated_header_rows",
    "label_is_subtotal",
    "subtotal_signal_score", "is_red_font", "RED_FONT_WHITELIST",
    "UNIT_RE", "SUBTOTAL_KW",
]
