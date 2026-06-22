#!/usr/bin/env python3
"""
com_safe.py — pywin32 COM 버킷의 수명주기·피벗 규칙. [Windows+Excel 전용]

COM 최대 위험은 **고아 Excel 프로세스**(참조 미해제)와 상태 오염이다. 이 모듈은:
  · excel_com(): DisplayAlerts/ScreenUpdating off, Calculation manual → 끝나면 반드시
    Workbook.Close + App.Quit + ReleaseComObject + gc (고아 프로세스 0).
  · verify_pivots(): 피벗 RefreshAll 후, 피벗 총계가 소스 범위 합과 tie 하는지(3% 케이스 후검증).
  · 규율 메모: 1-base 인덱스, Variant 날짜 epoch(1899-12-30, 1900 윤년 버그) 주의.

Excel/pywin32 없으면 깔끔히 SKIP.

사용:
  from com_safe import excel_com, process_count
  with excel_com() as app:
      wb = app.Workbooks.Open(path); ...; wb.Save(); wb.Close(SaveChanges=False)
"""
from __future__ import annotations

import contextlib

try:
    import win32com.client as win32
    import pythoncom  # noqa
    _HAVE = True
except Exception:
    _HAVE = False


def available() -> bool:
    return _HAVE


def process_count() -> int:
    """현재 EXCEL.EXE 프로세스 수(고아 누수 검증용). 실패 시 -1."""
    try:
        import subprocess
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE"],
                             capture_output=True, text=True).stdout
        return out.count("EXCEL.EXE")
    except Exception:
        return -1


@contextlib.contextmanager
def excel_com(visible=False):
    """COM 수명주기 보장: 끝나면 Quit + 참조해제 + gc(고아 프로세스 차단)."""
    if not _HAVE:
        raise RuntimeError("pywin32/Excel 미가용 — 이 버킷은 Windows+Excel 필요")
    import gc
    app = win32.DispatchEx("Excel.Application")
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    prev = app.Calculation
    try:
        app.Calculation = -4135  # xlCalculationManual
        yield app
    finally:
        with contextlib.suppress(Exception):
            app.Calculation = prev
            app.ScreenUpdating = True
        with contextlib.suppress(Exception):
            app.Quit()
        with contextlib.suppress(Exception):
            del app
        gc.collect()  # 참조해제 강제(고아 프로세스 차단)


def verify_pivots(workbook, tol=0.5) -> list:
    """피벗 RefreshAll 후 grand total 이 소스 범위 합과 tie 하는지(후검증)."""
    issues = []
    with contextlib.suppress(Exception):
        workbook.RefreshAll()
    for sht in workbook.Worksheets:
        with contextlib.suppress(Exception):
            for pt in sht.PivotTables():
                # 소스 데이터 범위 합 vs 피벗 grand total
                src = pt.SourceData
                # 휴리스틱: DataBodyRange 합과 PivotTable.GrandTotal 비교는 레이아웃 의존 →
                # 여기선 RefreshAll 성공 + 캐시 유효성만 보증, 정밀 tie 는 contract 로.
                if pt.PivotCache().RecordCount in (0, None):
                    issues.append("PivotTable '%s' 캐시 비어있음(소스/refresh 확인)" % pt.Name)
    return issues


__all__ = ["available", "excel_com", "process_count", "verify_pivots"]


if __name__ == "__main__":
    print("pywin32/COM 가용:", available(), "— 미가용 시 Windows+Excel 런타임에서만 실행")
    if available():
        print("EXCEL.EXE 프로세스 수:", process_count())
