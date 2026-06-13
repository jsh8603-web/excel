#!/usr/bin/env python3
"""
format-xlwings.py — 실제 Excel(xlwings/win32com)로 서식 적용. 연결을 '완벽히' 보존.

왜 이게 완벽한가: 실제 Excel 이 파일을 쓰므로 외부링크·외부 DB 연결(connections)·Power Query·
데이터모델·피벗·VBA·Excel Table 이 모두 보존된다 → 다운스트림 Power BI/Tableau 바인딩도 안전.
openpyxl 의 부품 소실 문제가 원천적으로 없다.

라우팅:
  needs_excel(path) 가 True(연결/쿼리/외부링크/Table/모델/VBA 보유)면 이 경로(실 Excel) 사용,
  아니면 탭간 전용 openpyxl 경로(format-in-place)로도 충분.

런타임: Windows(+Mac) + Excel 설치 필요. 적용 후 roundtrip-gate 로 무결성 자동 검증.
쓰임: python3 tools/format-xlwings.py <workbook.xlsx> [--engine xlwings|win32com] [--numfmt "#,##0;(#,##0)"]
"""
from __future__ import annotations
import sys, os, zipfile, argparse, importlib.util

NUMFMT = "#,##0;(#,##0)"
# 실 Excel 보존이 필요함을 의미하는 부품(있으면 openpyxl 금지)
_NEED = ("xl/connections.xml", "xl/externalLinks/", "customXml/", "xl/model",
         "xl/queryTables/", "xl/tables/", "xl/pivotTables/", "xl/vbaProject.bin")

def needs_excel(path: str) -> bool:
    """연결/쿼리/외부링크/Table/모델/VBA 가 하나라도 있으면 실 Excel 경로가 필요."""
    parts = zipfile.ZipFile(path).namelist()
    return any(p == s or p.startswith(s) for p in parts for s in _NEED)

# ── 엔진 1: xlwings (권장 — Win/Mac, 앱 수명관리 깔끔) ─────────────────────────
def format_xlwings(path: str, numfmt: str = NUMFMT) -> None:
    import xlwings as xw
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        wb = app.books.open(os.path.abspath(path))
        try:
            for sht in wb.sheets:
                ur = sht.used_range
                if ur.count and ur.last_cell.row >= 1:
                    ur.number_format = numfmt           # 값 미변경, 표시서식만
                    sht.range((1, 1), (1, ur.last_cell.column)).font.bold = True  # 헤더행
            wb.save()                                   # 같은 .xlsx 로 저장(실 Excel → 전부 보존)
        finally:
            wb.close()
    finally:
        app.quit()

# ── 엔진 2: win32com (Windows 전용, xlwings 없이) ────────────────────────────
def format_win32com(path: str, numfmt: str = NUMFMT) -> None:
    import win32com.client as win32  # pywin32 (Windows only)
    xl = win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(os.path.abspath(path))
    try:
        for ws in wb.Worksheets:
            ur = ws.UsedRange
            ur.NumberFormat = numfmt
            ws.Rows(1).Font.Bold = True
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)
        xl.Quit()

def _load_gate():
    p = os.path.join(os.path.dirname(__file__), "roundtrip-gate.py")
    s = importlib.util.spec_from_file_location("rtg", p); m = importlib.util.module_from_spec(s); s.loader.exec_module(m)
    return m

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--engine", choices=["xlwings", "win32com"], default="xlwings")
    ap.add_argument("--numfmt", default=NUMFMT)
    ap.add_argument("--no-verify", action="store_true")
    a = ap.parse_args()

    if not needs_excel(a.workbook):
        print("참고: 이 워크북엔 연결/Table 류가 없어 openpyxl(format-in-place)로도 안전합니다.")

    rtg = None if a.no_verify else _load_gate()
    before = rtg.fingerprint(a.workbook) if rtg else None

    engine = format_xlwings if a.engine == "xlwings" else format_win32com
    try:
        engine(a.workbook, a.numfmt)
    except ImportError as e:
        print(f"엔진 미설치/환경 아님: {e}\n→ Windows+Excel 에서 실행하세요(xlwings 또는 pywin32 필요).")
        return 2
    except Exception as e:
        print(f"Excel 서식 적용 실패: {type(e).__name__}: {e}")
        return 1

    if rtg:
        issues = rtg.compare(before, rtg.fingerprint(a.workbook))
        if issues:
            print("경고: 실 Excel 저장 후에도 무결성 차이 발견(예상 밖):")
            for i in issues:
                print(f"  ✗ {i}")
            return 1
        print("서식 적용 + 무결성 통과(연결/Table/링크 보존).")
    else:
        print("서식 적용 완료(검증 생략).")
    return 0

if __name__ == "__main__":
    sys.exit(main())
