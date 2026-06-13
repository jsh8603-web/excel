"""fpna.ingest — 누더기 엑셀 → tidy long 정형화 파이프라인.

순수 파이썬(openpyxl + stdlib). 결정적·재현가능.
주 진입점: run_ingest(path, out_dir) / ingest_workbook(path).
"""
from __future__ import annotations

from .pipeline import (IngestResult, ingest_workbook, run_ingest,
                       write_tidy_csv, write_schema_json, write_smell_report)
from .validate import TidyRow, validate_rows, scan_formula_smells
from .cells import as_cells, Cell
from .reconcile import (ReconReport, reconcile_sheet, recon_to_smells,
                        groundtruth_cells)

__all__ = [
    "IngestResult", "ingest_workbook", "run_ingest",
    "write_tidy_csv", "write_schema_json", "write_smell_report",
    "TidyRow", "validate_rows", "scan_formula_smells",
    "as_cells", "Cell",
    "ReconReport", "reconcile_sheet", "recon_to_smells", "groundtruth_cells",
]
