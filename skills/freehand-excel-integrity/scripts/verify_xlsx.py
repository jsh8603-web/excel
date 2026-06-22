"""
verify_xlsx.py — 실제 Excel(COM) 로 .xlsx 를 강제 재계산해 에러셀을 검증.

⚠ 런타임 의존성이 *아니다*. Windows + Excel + pywin32(win32com) 가 있는 환경에서만
동작하며, 없으면 rc=2 로 깔끔히 빠진다(무설치 회사 PC·Linux CI 에서는 호출 불가).
freehand-excel-integrity 스킬과 함께 이동한다(fpna/·vendor 무관, standalone).

언제 쓰나(backend routing): **xlwings live-edit** 와 **pywin32 COM** 백엔드의 산출물처럼,
셀에 라이브 수식이 들어가는 경우. xlsx_doctor 는 디스크의 수식 텍스트만 보므로
#VALUE!/#DIV/0! 가 실제 재계산에서만 드러나는 케이스를 놓친다 — 이 도구가 진짜 Excel 로
재계산해 그 갭을 메운다. 정적 값만 쓴 openpyxl/xlsxwriter 산출은 불필요(게이트가 오프라인 완결).

사용:
  py scripts/verify_xlsx.py <파일.xlsx>             # 에러셀 스캔 + 재계산
  py scripts/verify_xlsx.py <파일.xlsx> --pdf out.pdf
"""
from __future__ import annotations

import argparse
import os
import sys

ERROR_VALUES = {-2146826281: "#DIV/0!", -2146826246: "#N/A",
                -2146826259: "#NAME?", -2146826288: "#NULL!",
                -2146826252: "#NUM!", -2146826265: "#REF!",
                -2146826273: "#VALUE!"}


def verify(path: str, *, pdf: str | None = None, png: str | None = None) -> int:
    try:
        import win32com.client as win32
    except ImportError:
        print("pywin32 필요: py -m pip install pywin32 (Excel 있는 환경 전용 검증)")
        return 2
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        print("파일 없음:", path)
        return 2

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    rc = 0
    try:
        wb = excel.Workbooks.Open(path)
        excel.CalculateFull()         # 강제 전체 재계산
        n_err = 0
        for ws in wb.Worksheets:
            used = ws.UsedRange
            for row in used.Rows:
                for cell in row.Cells:
                    v = cell.Value
                    if isinstance(v, int) and v in ERROR_VALUES:
                        n_err += 1
                        print("  에러셀 %s!%s = %s" % (ws.Name, cell.Address, ERROR_VALUES[v]))
        if n_err:
            print("FAIL: 수식 에러 %d건" % n_err)
            rc = 1
        else:
            print("OK: Excel 재계산 후 에러셀 0건 (%s)" % os.path.basename(path))

        if pdf:
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf))  # 0=xlTypePDF
            print("PDF:", pdf)
        if png:
            ws = wb.Worksheets(1)
            ws.UsedRange.CopyPicture(Format=2)   # 2=xlBitmap
            # 클립보드→파일 저장은 추가 라이브러리 필요 → PDF 권장.
            print("PNG 캡처는 PDF 경로 권장(클립보드 저장 생략):", png)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    return rc


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--pdf")
    ap.add_argument("--png")
    args = ap.parse_args(argv)
    return verify(args.path, pdf=args.pdf, png=args.png)


if __name__ == "__main__":
    sys.exit(main())
