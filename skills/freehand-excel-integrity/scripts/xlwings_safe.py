#!/usr/bin/env python3
"""
xlwings_safe.py — xlwings(라이브 Excel) 생성 버킷의 규칙을 코드로 강제. [Windows/Mac+Excel 전용]

헤드리스가 못 하는 가치를 살린다:
  · 작성 후 app.calculate() 로 **실제 재계산** → .value 되읽어 에러셀/결과를 그 자리에서 검증.
  · 상태 위생: screen_updating/calculation off→복원, 끝나면 반드시 quit + 참조해제(고아 Excel 차단).
  · 날짜는 datetime 유지(부동소수 epoch 로 깨지지 않게).

Excel 없으면 깔끔히 SKIP(이 버킷은 Excel 필요 = 규칙).

사용:
  from xlwings_safe import excel_app, generate_and_verify
  errs = generate_and_verify("out.xlsx", build_fn)   # build_fn(sheet) 가 셀을 채움
"""
from __future__ import annotations

import contextlib

try:
    import xlwings as xw
    _HAVE = True
except Exception:
    _HAVE = False

_ERR = {"#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!"}


def available() -> bool:
    return _HAVE


@contextlib.contextmanager
def excel_app(visible=False):
    """상태 위생 보장 컨텍스트: 끝나면 무조건 quit(고아 프로세스 차단)."""
    if not _HAVE:
        raise RuntimeError("xlwings/Excel 미가용 — 이 버킷은 Excel 필요")
    app = xw.App(visible=visible, add_book=False)
    prev_calc = app.calculation
    try:
        app.screen_updating = False
        app.calculation = "manual"
        yield app
    finally:
        with contextlib.suppress(Exception):
            app.calculation = prev_calc
            app.screen_updating = True
        with contextlib.suppress(Exception):
            app.quit()          # 반드시 종료(참조해제 포함)


def scan_error_cells(sheet) -> list:
    """used range 를 되읽어 에러값 셀 수집."""
    errs = []
    used = sheet.used_range
    vals = used.value
    if not isinstance(vals, list):
        vals = [[vals]]
    elif vals and not isinstance(vals[0], list):
        vals = [vals]
    r0, c0 = used.row, used.column
    for i, row in enumerate(vals):
        for j, v in enumerate(row):
            if isinstance(v, str) and v in _ERR:
                errs.append("%s!R%dC%d=%s" % (sheet.name, r0 + i, c0 + j, v))
    return errs


def generate_and_verify(path: str, build_fn, sheet_name="Sheet1") -> list:
    """build_fn(sheet) 로 작성 → 재계산 → 에러셀 되읽기 → 저장 → quit.
    반환: 발견된 에러셀 리스트(비어야 정상)."""
    with excel_app() as app:
        book = app.books.add()
        sheet = book.sheets[0]
        sheet.name = sheet_name
        build_fn(sheet)
        app.calculate()                 # 실제 재계산
        errs = scan_error_cells(sheet)  # 되읽기 검증
        book.save(path)
        book.close()
    return errs


__all__ = ["available", "excel_app", "scan_error_cells", "generate_and_verify"]


if __name__ == "__main__":
    print("xlwings 가용:", available(), "— 미가용 시 이 버킷은 Excel 런타임에서만 실행")
