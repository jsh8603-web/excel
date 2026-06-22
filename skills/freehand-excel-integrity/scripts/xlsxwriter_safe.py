#!/usr/bin/env python3
"""
xlsxwriter_safe.py — xlsxwriter 생성 버킷의 '생성 규칙'을 코드로 강제.

xlsxwriter 는 write-only(되읽기·재계산 불가)이고, write_formula 에 캐시값을 안 주면
소비자(pandas/PowerBI/data_only)가 0/빈칸으로 본다(stress test 확인). 이 래퍼는:
  · write_formula 에 value= 를 **필수**로 강제(누락 시 raise) → 캐시 동봉.
  · 신규 전용임을 명시(기존 파일 편집 요청은 openpyxl/xlwings 로 라우팅해야 함).
  · close 후 doctor + 재계산으로 외부검증(자기검증 불가하므로).

사용:
  sw = SafeWorkbook("out.xlsx")
  ws = sw.add_worksheet("S")
  ws.write_number(0,1,100); ws.write_number(1,1,200)
  ws.write_formula(2,1, "=SUM(B1:B2)", value=300)   # value 필수
  sw.close(verify=True)                              # 닫고 doctor+recalc
"""
from __future__ import annotations

import subprocess
import sys
import os

try:
    import xlsxwriter
    from xlsxwriter.worksheet import Worksheet as _WS
except ImportError:
    raise SystemExit("xlsxwriter 필요: pip install xlsxwriter")


class SafeWorksheet:
    """write_formula 에 value= 를 강제하는 얇은 프록시."""
    def __init__(self, ws: _WS):
        self._ws = ws

    def write_formula(self, *args, value=None, **kw):
        # 시그니처: (row, col, formula, cell_format=None) 또는 (a1, formula, ...)
        if value is None:
            raise ValueError(
                "xlsxwriter write_formula 는 value=<계산값> 필수 — "
                "캐시값 없으면 소비자가 0/빈칸으로 본다. Python 에서 계산한 값을 넘겨라.")
        return self._ws.write_formula(*args, value=value, **kw)

    def __getattr__(self, name):
        # 그 외 write_number/write/merge_range 등은 그대로 위임
        return getattr(self._ws, name)


class SafeWorkbook:
    def __init__(self, path: str, options: dict | None = None):
        self.path = path
        self._wb = xlsxwriter.Workbook(path, options or {})

    def add_worksheet(self, name=None):
        return SafeWorksheet(self._wb.add_worksheet(name))

    def add_format(self, *a, **k):
        return self._wb.add_format(*a, **k)

    def close(self, verify: bool = True):
        self._wb.close()
        if verify:
            here = os.path.dirname(os.path.abspath(__file__))
            doctor = os.path.join(here, "xlsx_doctor.py")
            if os.path.exists(doctor):
                # write-only 라 자기검증 불가 → openpyxl 재오픈 doctor + 재계산
                subprocess.run([sys.executable, doctor, self.path, "--recalc"])


__all__ = ["SafeWorkbook", "SafeWorksheet"]
