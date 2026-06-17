"""
fpna.ingest.tidy_mart — ingest tidy(long) → profile 입력 mart(wide) 어댑터.

ingest(누더기→long tidy)와 profile(wide fact→SHAPE) 사이의 포맷 격차를 메운다.
직접 연결하면 measure 가 value 1개로 뭉개지고 메타 컬럼(src_row/level/scale_*)이
통계에 끼어든다 → 이 어댑터가 data 행만 추려 metric 을 컬럼으로 pivot 한다.

⚠ 두 원본(ingest pipeline / profile) 무수정. 순수 stdlib·결정적(입력 등장 순 보존).
"""
from __future__ import annotations

import csv

# TidyRow 의 비-차원 메타(통계 무관) — pivot 시 제거
_META_COLS = {"unit", "row_role", "level", "src_sheet", "src_row", "src_col",
              "scale_applied", "scale_source", "raw_value", "flags"}


def _read_csv(path: str) -> tuple[list[str], list[list[str]]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return [], []
    return [h.strip() for h in rows[0]], rows[1:]


def is_long_tidy(header: list[str], *, metric_col: str = "metric",
                 value_col: str = "value") -> bool:
    """헤더가 ingest long tidy(metric/value 컬럼 보유)인지 — profile 직행 전 판정용."""
    return metric_col in header and value_col in header


def tidy_to_mart(tidy_path: str, mart_path: str, *, metric_col: str = "metric",
                 value_col: str = "value", role_col: str = "row_role",
                 data_role: str = "data") -> tuple[list[str], list[str], int]:
    """long tidy → wide mart csv. 반환 (dim_cols, metrics, n_rows).

    - data 행만(소계/합계/헤더 제외). 메타 컬럼 제거.
    - dim = header 에서 metric/value/메타 제외(예: entity, period).
    - 같은 (dim 조합, metric) 중복 시 마지막 값(결정적, 입력 순).
    """
    header, body = _read_csv(tidy_path)
    if not header:
        raise ValueError("빈 tidy: %s" % tidy_path)
    if not is_long_tidy(header, metric_col=metric_col, value_col=value_col):
        raise ValueError("long tidy 아님(%s/%s 컬럼 없음): %s"
                         % (metric_col, value_col, tidy_path))

    idx = {h: i for i, h in enumerate(header)}
    dim_cols = [h for h in header
                if h not in (metric_col, value_col) and h not in _META_COLS]
    mi, vi = idx[metric_col], idx[value_col]
    ri = idx.get(role_col)

    def _cell(row, i):
        return row[i] if i < len(row) else ""

    keys: list[tuple] = []
    key_seen: set = set()
    metrics: list[str] = []
    metric_seen: set = set()
    table: dict[tuple, dict] = {}

    for row in body:
        if ri is not None and _cell(row, ri) != data_role:
            continue
        m = _cell(row, mi)
        if not m:
            continue
        key = tuple(_cell(row, idx[d]) for d in dim_cols)
        if key not in key_seen:
            key_seen.add(key)
            keys.append(key)
        if m not in metric_seen:
            metric_seen.add(m)
            metrics.append(m)
        table.setdefault(key, {})[m] = _cell(row, vi)

    out_header = dim_cols + metrics
    with open(mart_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(out_header)
        for key in keys:
            w.writerow(list(key) + [table[key].get(m, "") for m in metrics])
    return dim_cols, metrics, len(keys)


__all__ = ["tidy_to_mart", "is_long_tidy"]
