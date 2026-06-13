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

회복성 보강(2026-06-14):
  - 거대 used_range(셀 수 > CELL_THRESHOLD) — 행 단위 청크 분할로 OLE timeout 회피
  - 시트 보호(ProtectContents=True) — NumberFormat 거부되므로 skip + 로그(운영 의도 보존)
  - 시트별 try/except — 한 시트 실패해도 다른 시트는 계속, 끝에 종합 rc
  - UTF-8 stdout — 한국어 시트명/메시지 안전 출력
"""
from __future__ import annotations
import sys, os, zipfile, argparse, importlib.util

try:
    sys.stdout.reconfigure(encoding="utf-8")    # Python 3.7+; 한국어 시트명 출력 안전
except Exception:
    pass

NUMFMT = "#,##0;(#,##0)"
CELL_THRESHOLD = 50_000     # 셀 수 임계 — 이상이면 청크 분할
CHUNK_ROWS = 500            # 청크당 행 수
# 실 Excel 보존이 필요함을 의미하는 부품(있으면 openpyxl 금지)
_NEED = ("xl/connections.xml", "xl/externalLinks/", "customXml/", "xl/model",
         "xl/queryTables/", "xl/tables/", "xl/pivotTables/", "xl/vbaProject.bin")

def needs_excel(path: str) -> bool:
    """연결/쿼리/외부링크/Table/모델/VBA 가 하나라도 있으면 실 Excel 경로가 필요."""
    parts = zipfile.ZipFile(path).namelist()
    return any(p == s or p.startswith(s) for p in parts for s in _NEED)

# ── 엔진 1: xlwings (권장 — Win/Mac, 앱 수명관리 깔끔) ─────────────────────────
def _apply_numfmt_xlwings(sht, nrows, ncols, numfmt):
    """청크 분할/일괄 NumberFormat 적용. 거대 시트는 행 단위로 쪼개 OLE timeout 회피."""
    if nrows * ncols > CELL_THRESHOLD:
        for start in range(1, nrows + 1, CHUNK_ROWS):
            end = min(start + CHUNK_ROWS - 1, nrows)
            sht.range((start, 1), (end, ncols)).number_format = numfmt
    else:
        sht.used_range.number_format = numfmt


def format_xlwings(path: str, numfmt: str = NUMFMT) -> int:
    """반환: 0 = 전부 OK, 1 = 일부 시트 실패(저장은 진행), 2 = 워크북 단위 치명."""
    import xlwings as xw
    app = xw.App(visible=False, add_book=False)
    rc = 0
    try:
        app.display_alerts = False
        wb = app.books.open(os.path.abspath(path))
        try:
            for sht in wb.sheets:
                ur = sht.used_range
                if not ur.count or ur.last_cell.row < 1:
                    print(f"  [{sht.name}] 빈 시트 — skip")
                    continue
                nrows, ncols = ur.last_cell.row, ur.last_cell.column
                # 시트 보호 — NumberFormat 거부됨, 운영 의도 보존을 위해 skip
                if sht.api.ProtectContents:
                    print(f"  [{sht.name}] {nrows}x{ncols} PROTECTED — skip")
                    continue
                try:
                    _apply_numfmt_xlwings(sht, nrows, ncols, numfmt)
                    sht.range((1, 1), (1, ncols)).font.bold = True
                    print(f"  [{sht.name}] {nrows}x{ncols} OK")
                except Exception as e:
                    print(f"  [{sht.name}] {nrows}x{ncols} FAIL: {type(e).__name__}: {str(e)[:140]}")
                    rc = 1
            wb.save()                                   # 같은 .xlsx 로 저장(실 Excel → 전부 보존)
        finally:
            wb.close()
    finally:
        app.quit()
    return rc

# ── 엔진 2: win32com (Windows 전용, xlwings 없이) ────────────────────────────
def _apply_numfmt_win32com(ws, nrows, ncols, numfmt):
    if nrows * ncols > CELL_THRESHOLD:
        for start in range(1, nrows + 1, CHUNK_ROWS):
            end = min(start + CHUNK_ROWS - 1, nrows)
            ws.Range(ws.Cells(start, 1), ws.Cells(end, ncols)).NumberFormat = numfmt
    else:
        ws.UsedRange.NumberFormat = numfmt


def format_win32com(path: str, numfmt: str = NUMFMT) -> int:
    import win32com.client as win32  # pywin32 (Windows only)
    xl = win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(os.path.abspath(path))
    rc = 0
    try:
        for ws in wb.Worksheets:
            ur = ws.UsedRange
            if ur.Cells.Count == 0:
                print(f"  [{ws.Name}] 빈 시트 — skip")
                continue
            nrows, ncols = ur.Rows.Count, ur.Columns.Count
            if ws.ProtectContents:
                print(f"  [{ws.Name}] {nrows}x{ncols} PROTECTED — skip")
                continue
            try:
                _apply_numfmt_win32com(ws, nrows, ncols, numfmt)
                ws.Rows(1).Font.Bold = True
                print(f"  [{ws.Name}] {nrows}x{ncols} OK")
            except Exception as e:
                print(f"  [{ws.Name}] {nrows}x{ncols} FAIL: {type(e).__name__}: {str(e)[:140]}")
                rc = 1
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)
        xl.Quit()
    return rc

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
        engine_rc = engine(a.workbook, a.numfmt)
    except ImportError as e:
        print(f"엔진 미설치/환경 아님: {e}\n→ Windows+Excel 에서 실행하세요(xlwings 또는 pywin32 필요).")
        return 2
    except Exception as e:
        print(f"Excel 서식 적용 실패(워크북 단위 치명): {type(e).__name__}: {e}")
        return 1

    if rtg:
        issues = rtg.compare(before, rtg.fingerprint(a.workbook))
        if issues:
            print("경고: 실 Excel 저장 후에도 무결성 차이 발견(예상 밖):")
            for i in issues:
                print(f"  ✗ {i}")
            return 1
        if engine_rc == 0:
            print("서식 적용 + 무결성 통과(연결/Table/링크 보존).")
        else:
            print("부분 서식 적용 + 무결성 통과(일부 시트는 위 로그대로 skip/fail).")
    else:
        print("서식 적용 완료(검증 생략).")
    return engine_rc

if __name__ == "__main__":
    sys.exit(main())
