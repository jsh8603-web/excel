"""
fpna.ingest.reconcile — tidy 충실도 게이트(A1: per-cell reconciliation).

목적: parse 결과(tidy long)가 원본 셀을 1:1로 충실히 옮겼는지 **출처 좌표 단위**로
검증한다. scalar 합-diff 는 병합셀 다값손실·열 오배치를 못 잡으므로 per-cell 로 본다.

핵심 설계(순환 회피):
- ground truth = parse 와 **독립한 단순·관대 스캔**. ⛔ detect_blocks/classify_cells
  재사용 금지(같은 로직으로 검증하면 같은 실수를 못 본다).
- 대상 = data-class 셀(detail행 × measure열). 소계/총계/헤더/각주 행은 제외.
  (분류 휴리스틱이 보수적이라 오탐을 줄이는 방향 — 모호하면 ground truth 에서 뺀다.)
- 비교는 **정규화 前 raw 셀값**으로(동일 셀 비교 → 단위/스케일 무관).
  fidelity 체크 ≠ unit 체크. 둘을 분리한다.

산출(ReconReport):
- covered_once / missing / duplicate 좌표 집합(coverage map)
- value_mismatch: tidy 가 보유한 출처셀 값이 원본 셀과 어긋난 좌표
- 모두 결정적(좌표 정렬). 같은 입력 → 같은 출력.

본 모듈은 smell/리포트성이며 reject 와 독립이다(충실도 관측). 호출측이 결과를
schema/smell 로 노출한다.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .cells import Cell, T_NUMERIC, T_ERROR
from .detect import label_is_subtotal, UNIT_RE


# --------------------------------------------------------------------------
# 독립 ground-truth 스캐너
# --------------------------------------------------------------------------
def _is_year_like(v) -> bool:
    """bare 4자리 정수(연도/헤더 후보)는 measure 값이 아니라 헤더로 본다."""
    if isinstance(v, int) and not isinstance(v, bool):
        return 1000 <= v <= 9999
    if isinstance(v, str):
        s = v.strip()
        return s.isdigit() and len(s) == 4
    return False


def _is_groundtruth_value(c: Cell) -> bool:
    """이 셀이 '옮겨졌어야 할 measure 데이터 값'인가(독립·보수 판정).

    - 숫자/오류값 = 값.  (bool 은 logical 이라 numeric 아님 → 제외)
    - bare 4자리(연도) = 헤더 후보 → 제외(모호하면 ground truth 에서 뺀다).
    - 날짜셀 = period 헤더일 가능성 높음 → 제외(measure 아님).
    - 텍스트는 measure 가 아님 → 제외.
    """
    if c.is_blank:
        return False
    if c.data_type == T_ERROR:
        return True
    if c.data_type == T_NUMERIC:
        if _is_year_like(c.value):
            return False
        return True
    return False


def _row_left_label(cells_by_pos: dict, row: int, min_col: int, max_col: int):
    """행의 가장 왼쪽 텍스트 라벨(소계/총계 판정용). 없으면 ''."""
    for col in range(min_col, max_col + 1):
        c = cells_by_pos.get((row, col))
        if c is not None and not c.is_blank and isinstance(c.value, str):
            s = c.value.strip()
            if s and not UNIT_RE.search(s):
                return s
    return ""


def groundtruth_cells(cells: list[Cell]) -> set[tuple[int, int]]:
    """원본 셀 평면에서 data-class measure 셀 좌표를 독립 수집.

    ⚠ detect_blocks/classify 와 무관한 단순 스캔. 소계/총계 라벨이 붙은 행은
    통째로 제외(이중집계 셀을 fidelity 대상에서 빼 오탐 회피).
    """
    if not cells:
        return set()
    by_pos = {(c.row, c.col): c for c in cells}
    rows = [c.row for c in cells if not c.is_blank]
    cols = [c.col for c in cells if not c.is_blank]
    if not rows:
        return set()
    min_col, max_col = min(cols), max(cols)

    # 소계/총계 라벨 행 = 제외 대상.
    subtotal_rows: set[int] = set()
    for r in set(rows):
        label = _row_left_label(by_pos, r, min_col, max_col)
        if label and label_is_subtotal(label):
            subtotal_rows.add(r)

    gt: set[tuple[int, int]] = set()
    for c in cells:
        if c.row in subtotal_rows:
            continue
        if _is_groundtruth_value(c):
            gt.add((c.row, c.col))
    return gt


# --------------------------------------------------------------------------
# 값 비교(정규화 前 raw 동일성)
# --------------------------------------------------------------------------
def _raw_numeric(v):
    """비교용 raw 수치 추출. 셀에 저장된 실수만(텍스트 음수표기 등은 별도).

    동일 셀 raw 비교이므로 스케일/단위 무관. 부동소수 허용오차로 흡수.
    """
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


_VAL_TOL = 1e-6


def _tidy_reconstruct_raw(tr):
    """tidy 행에서 '원본 셀 raw 수치'를 복원.

    - scale_applied>1 이면 value 는 환산 후이므로 raw_value 가 원본.
    - SIGN_FROM_COLOR 면 부호를 뒤집어 색보정 전(=원본 셀) 값 복원.
    - 그 외엔 value 가 곧 원본 셀 수치.
    """
    flags = tr.flags or ""
    # 색 음수 보정: 원본 셀은 양수였다 → 절대값(부호 환원).
    if "SIGN_FROM_COLOR" in flags:
        base = tr.raw_value if tr.raw_value is not None else tr.value
        n = _raw_numeric(base)
        return abs(n) if n is not None else None
    if tr.scale_applied and tr.scale_applied > 1 and tr.raw_value is not None:
        return _raw_numeric(tr.raw_value)
    return _raw_numeric(tr.value)


# --------------------------------------------------------------------------
# 리포트
# --------------------------------------------------------------------------
@dataclass
class ReconReport:
    sheet: str
    n_groundtruth: int = 0
    covered: set = field(default_factory=set)       # GT 중 정확히 1회 덮인 좌표
    missing: set = field(default_factory=set)        # GT 인데 tidy 가 안 옮긴 좌표
    duplicate: dict = field(default_factory=dict)    # 좌표 → tidy 매핑 횟수(2+)
    value_mismatch: list = field(default_factory=list)  # (coord, gt_raw, tidy_raw)
    extra: set = field(default_factory=set)          # GT 아닌데 tidy 가 옮긴 좌표(정보용)

    @property
    def ok(self) -> bool:
        return not self.missing and not self.duplicate and not self.value_mismatch


def reconcile_sheet(sheet: str, gt: set, raw_values: dict, tidy_rows: list) -> ReconReport:
    """한 시트의 ground truth 와 tidy(data 행)를 좌표 단위로 대조.

    gt         : groundtruth_cells() 가 unmerge 前 raw cells 에서 뽑은 좌표 집합
                 (parse 와 독립 — 순환 회피, 호출측이 미리 계산해 전달).
    raw_values : (row,col) → 원본 셀값(정규화/병합전파 前). 값 비교용.
    tidy_rows  : 그 시트 소속 TidyRow 전체(소계/오류 포함). coverage 는 GT 기준.
    """
    rep = ReconReport(sheet=sheet, n_groundtruth=len(gt))

    # tidy 의 (src_row, src_col) → 매핑 횟수. coverage = 출처셀이 tidy 로 옮겨졌는가
    # (좌표 점유) — **역할 무관**. 소계/총계 행도 자기 출처셀을 보존하고, ERROR_CELL 도
    # 좌표를 점유한다(값은 null). ⛔ role 필터 금지: groundtruth 독립 스캔은 ditto 를
    # 적용 안 해 다중 키컬럼 *안쪽*의 소계('계' 가 ditto 로 비어있음)를 못 가려내므로,
    # tidy 만 소계로 판정(전체 라벨)하면 그 셀이 false-missing 으로 뜬다(다중키 KOSIS).
    cover: dict[tuple[int, int], int] = {}
    coord_to_rows: dict[tuple[int, int], list] = {}
    for tr in tidy_rows:
        if not tr.src_row or not tr.src_col:
            continue
        key = (tr.src_row, tr.src_col)
        cover[key] = cover.get(key, 0) + 1
        coord_to_rows.setdefault(key, []).append(tr)

    for coord in sorted(gt):
        n = cover.get(coord, 0)
        if n == 0:
            rep.missing.add(coord)
        elif n == 1:
            rep.covered.add(coord)
        else:
            rep.duplicate[coord] = n

    # GT 아닌데 덮인 좌표(헤더/연도/소계 오인 등 — 정보용, 실패로 보지 않음).
    for coord in cover:
        if coord not in gt:
            rep.extra.add(coord)

    # 값 동일성: GT 이면서 1:1로 덮인 좌표만(중복/결측은 위에서 별도 보고).
    for coord in sorted(rep.covered):
        if coord not in raw_values:
            continue
        gt_raw = _raw_numeric(raw_values[coord])
        if gt_raw is None:
            continue  # 오류값/비수치 — 좌표 점유만 확인(값비교 생략)
        trs = coord_to_rows.get(coord, [])
        if not trs:
            continue
        tidy_raw = _tidy_reconstruct_raw(trs[0])
        if tidy_raw is None:
            # tidy 가 값을 못 살림(센티넬 등)인데 GT 는 수치 → mismatch.
            rep.value_mismatch.append((coord, gt_raw, None))
            continue
        if abs(gt_raw - tidy_raw) > _VAL_TOL * (abs(gt_raw) + 1):
            rep.value_mismatch.append((coord, gt_raw, tidy_raw))

    return rep


def recon_to_smells(rep: ReconReport):
    """ReconReport → Smell 리스트(smell_report 에 노출).

    충실도 위반을 표면화한다(reject 와 독립). kind:
      RECON_MISSING_CELL     : 원본 data 셀이 tidy 에 누락(무음 손실).
      RECON_DUPLICATE_CELL   : 한 출처셀이 tidy 행 2+로 중복.
      RECON_VALUE_MISMATCH   : tidy 값이 원본 셀과 불일치(다값손실/오배치).
    """
    from .validate import Smell
    out = []
    for (r, c) in sorted(rep.missing):
        out.append(Smell("R%dC%d" % (r, c), "RECON_MISSING_CELL",
                         "원본 data 셀이 tidy 에 누락(coverage=0)"))
    for (r, c), n in sorted(rep.duplicate.items()):
        out.append(Smell("R%dC%d" % (r, c), "RECON_DUPLICATE_CELL",
                         "출처셀이 tidy 행 %d개로 중복 매핑" % n))
    for (r, c), gt_raw, tidy_raw in rep.value_mismatch:
        out.append(Smell("R%dC%d" % (r, c), "RECON_VALUE_MISMATCH",
                         "tidy 값 %r != 원본 셀 %r" % (tidy_raw, gt_raw)))
    return out


__all__ = ["ReconReport", "reconcile_sheet", "recon_to_smells",
           "groundtruth_cells"]
