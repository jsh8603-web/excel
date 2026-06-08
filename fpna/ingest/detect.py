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

from .cells import Cell, T_CHARACTER, T_NUMERIC, T_DATE

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
        if b.n_cols >= min_cols and b.n_rows >= min_rows and density >= min_density:
            blocks.append(b)
    # 위→아래, 좌→우 정렬(결정성)
    blocks.sort(key=lambda x: (x.min_row, x.min_col))
    return blocks


def cells_in_block(cells: list[Cell], b: Block) -> list[Cell]:
    return [c for c in cells if b.contains(c.row, c.col)]


# --------------------------------------------------------------------------
# 비데이터 행 격리 (제목 / 단위 / 각주)
# --------------------------------------------------------------------------
def _row_cells(block_cells: list[Cell], row: int) -> list[Cell]:
    return [c for c in block_cells if c.row == row and not c.is_blank]


_MARKER_RE = re.compile(r"[(),%△▲]|\d{1,3},\d{3}")


def _looks_value(c: Cell) -> bool:
    if c.data_type in (T_NUMERIC, T_DATE):
        return True
    if isinstance(c.value, str):
        s = c.value.strip()
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


def label_is_subtotal(label: str) -> bool:
    if not label:
        return False
    s = str(label).strip().lower().replace(" ", "")
    return any(kw.replace(" ", "") in s for kw in SUBTOTAL_KW)


__all__ = [
    "Block", "detect_blocks", "cells_in_block",
    "strip_title_rows", "strip_footnote_rows", "label_is_subtotal",
    "UNIT_RE", "SUBTOTAL_KW",
]
