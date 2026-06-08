"""
fpna.ingest.pipeline — 누더기 엑셀 → tidy 변환 오케스트레이터.

단계(리서치 보강 반영, 결정적·재현가능):
  1) load: data_only False/True 두 번 열어 값+수식 동시 수집(as_cells)
  2) detect_blocks: connected-component + density 로 다중 표 분리
  2.5) 비데이터 행 격리: 제목/단위/각주
  3) unmerge_fill + classify + 헤더 밴드
  4) behead 언피벗 → long
  4.5) 소계/합계 행 플래그(이중집계 방지)
  5) 센티넬·단위·스케일 정규화
  6) dataclass 스키마 검증 → reject 리포트
  7) 수식 스멜 스캔
  8) 산출: tidy.csv(utf-8-sig) + schema.json + smell_report.md

회사 PC: `py main.py ingest <파일.xlsx>` 로 호출.
"""
from __future__ import annotations

import csv
import datetime as _dt
import json
import os
import re
from dataclasses import asdict, dataclass, field

import fpna._bootstrap  # noqa: F401

import openpyxl

from .cells import as_cells, T_NUMERIC, T_DATE
from .detect import (detect_blocks, cells_in_block, strip_title_rows,
                     strip_footnote_rows, label_is_subtotal, Block, UNIT_RE)
from .headers import unpivot_block, unmerge_fill, LongRow
from .normalize import normalize_value, parse_unit_label, infer_column_type
from .validate import TidyRow, validate_rows, scan_formula_smells

_PERIOD_RE = re.compile(r"(\d{4}\s*[-/.년]?\s*(\d{1,2})?\s*(월|분기|Q|H)?|\dQ|[1-4]분기|상반기|하반기)",
                        re.IGNORECASE)


@dataclass
class IngestResult:
    tidy_rows: list = field(default_factory=list)
    schema: dict = field(default_factory=dict)
    smells: list = field(default_factory=list)
    report: object = None
    n_blocks: int = 0


def _looks_like_period(v) -> bool:
    if isinstance(v, (_dt.datetime, _dt.date)):
        return True
    if v is None:
        return False
    return bool(_PERIOD_RE.search(str(v)))


def _map_long_to_tidy(longs: list[LongRow], unit: str | None) -> list[TidyRow]:
    """LongRow.attrs(hdr_c*/hdr_r*) → (entity, period, metric, value) 휴리스틱 매핑.

    - period: period 패턴인 헤더값 우선(보통 열헤더).
    - metric: period 아닌 열헤더 또는 가장 안쪽 행헤더.
    - entity: 가장 바깥 행헤더.
    나머지 헤더는 metric 에 ' > ' 로 결합 보존.
    """
    out: list[TidyRow] = []
    for lr in longs:
        col_vals = [(k, v) for k, v in lr.attrs.items() if k.startswith("hdr_c")]
        row_vals = [(k, v) for k, v in lr.attrs.items() if k.startswith("hdr_r")]

        period = None
        metric_parts: list[str] = []
        for _, v in col_vals:
            if v is None:
                continue
            if period is None and _looks_like_period(v):
                period = str(v)
            else:
                metric_parts.append(str(v))

        entity = None
        row_labels = [str(v) for _, v in row_vals if v is not None]
        if row_labels:
            entity = row_labels[0]
            if len(row_labels) > 1:
                metric_parts = row_labels[1:] + metric_parts
        metric = " > ".join(metric_parts) if metric_parts else None

        # 정규화
        num, sentinel, _ = normalize_value(lr.value)
        role = "data"
        label_for_role = " ".join(filter(None, [entity, metric or ""]))
        if label_is_subtotal(label_for_role):
            role = "subtotal"
        out.append(TidyRow(
            entity=entity, period=period, metric=metric,
            value=(None if sentinel is not None else num),
            unit=unit, row_role=role, level=lr.level,
            src_row=lr.row, src_col=lr.col,
        ))
    return out


def ingest_workbook(path: str, *, sheet: str | None = None) -> IngestResult:
    """엑셀 파일을 tidy long 으로 변환."""
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    wb_f = openpyxl.load_workbook(path, data_only=False, read_only=False)
    wb_v = openpyxl.load_workbook(path, data_only=True, read_only=False)

    sheets = [sheet] if sheet else wb_f.sheetnames
    all_tidy: list[TidyRow] = []
    all_smells = []
    schema_blocks = []
    n_blocks = 0

    for sn in sheets:
        ws_f = wb_f[sn]
        ws_v = wb_v[sn]
        cells = as_cells(ws_f, ws_v)
        all_smells.extend([asdict_smell(s, sn) for s in scan_formula_smells(cells)])

        # 병합셀 값을 영역 내 전파(헤더 해소·블록탐지·제목격리 정확도↑)
        unmerge_fill(cells)
        # 시트 단위 라벨 폴백: '(단위: ...)' 를 시트 전체에서 1회 스캔
        sheet_unit = None
        for c in cells:
            if isinstance(c.value, str):
                m = UNIT_RE.search(c.value)
                if m:
                    sheet_unit = m.group(1).strip()
                    break

        blocks = detect_blocks(cells)
        for bi, b in enumerate(blocks):
            n_blocks += 1
            bc = cells_in_block(cells, b)
            bc, _titles, unit_meta = strip_title_rows(bc, b)
            bc, _foots = strip_footnote_rows(bc, b)
            unit = unit_meta.get("unit") or sheet_unit
            # 블록 bbox 재계산(격리 후)
            if not bc:
                continue
            rows = [c.row for c in bc]
            cols = [c.col for c in bc]
            b2 = Block(min(rows), max(rows), min(cols), max(cols))
            longs = unpivot_block(bc, b2)
            tidy = _map_long_to_tidy(longs, unit)
            all_tidy.extend(tidy)
            schema_blocks.append({
                "sheet": sn, "block": bi,
                "range": "R%dC%d:R%dC%d" % (b.min_row, b.min_col, b.max_row, b.max_col),
                "unit": unit, "n_long_rows": len(tidy),
            })

    kept, rep = validate_rows(all_tidy, numeric_metric=False)

    # 컬럼 타입 추론(스키마 문서화용)
    schema = {
        "columns": {
            "entity": "TEXT", "period": "TEXT",
            "metric": "TEXT",
            "value": infer_column_type([r.value for r in kept]),
            "unit": "TEXT", "row_role": "TEXT", "level": "NUM",
            "src_row": "NUM", "src_col": "NUM",
        },
        "blocks": schema_blocks,
        "n_rows": len(kept),
        "n_rejected": rep.n_rejected,
        "generated_by": "fpna.ingest.pipeline",
    }
    return IngestResult(tidy_rows=kept, schema=schema, smells=all_smells,
                        report=rep, n_blocks=n_blocks)


def asdict_smell(s, sheet) -> dict:
    return {"sheet": sheet, "cell": s.cell, "kind": s.kind, "detail": s.detail}


# --------------------------------------------------------------------------
# 산출 writer (utf-8-sig: 한글 Excel 더블클릭 호환)
# --------------------------------------------------------------------------
def write_tidy_csv(rows: list[TidyRow], path: str) -> None:
    cols = ["entity", "period", "metric", "value", "unit",
            "row_role", "level", "src_row", "src_col"]
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(cols)
        for r in rows:
            d = asdict(r)
            w.writerow([d[c] for c in cols])


def write_schema_json(schema: dict, path: str) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(schema, fh, ensure_ascii=False, indent=2, default=str)


def write_smell_report(smells: list[dict], path: str) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("# 수식 스멜 리포트\n\n")
        if not smells:
            fh.write("스멜 없음.\n")
            return
        by_kind: dict[str, list] = {}
        for s in smells:
            by_kind.setdefault(s["kind"], []).append(s)
        for kind, items in sorted(by_kind.items()):
            fh.write("## %s (%d건)\n\n" % (kind, len(items)))
            for s in items[:50]:
                fh.write("- `%s` %s — %s\n" % (s["sheet"], s["cell"], s["detail"]))
            if len(items) > 50:
                fh.write("- … 외 %d건\n" % (len(items) - 50))
            fh.write("\n")


def run_ingest(path: str, out_dir: str, *, sheet: str | None = None) -> IngestResult:
    """파일 정형화 + 3종 산출 기록. out_dir 에 tidy.csv/schema.json/smell_report.md."""
    os.makedirs(out_dir, exist_ok=True)
    res = ingest_workbook(path, sheet=sheet)
    write_tidy_csv(res.tidy_rows, os.path.join(out_dir, "tidy.csv"))
    write_schema_json(res.schema, os.path.join(out_dir, "schema.json"))
    write_smell_report(res.smells, os.path.join(out_dir, "smell_report.md"))
    return res


__all__ = ["IngestResult", "ingest_workbook", "run_ingest",
           "write_tidy_csv", "write_schema_json", "write_smell_report"]
