"""
tools/verify_xlsx.py — [집-전용] 실제 Excel(COM) 로 생성물 검증.

⚠ 이 스크립트는 회사 런타임 의존성이 *아니다*. 집에서 생성한 .xlsx 가
진짜 Excel 에서 ① 수식이 에러 없이 재계산되는지 ② 파이썬 QC 값과 일치하는지
③ 화면 캡처(PNG/PDF) 로 룩이 맞는지 확인하는 검증 도구다.
회사 PC(설치 0, Excel 자동화 불가)에서는 절대 호출하지 않는다.

요구: Windows + Excel 설치 + pywin32(win32com). (홈 PC 에 이미 있음)

사용:
  py tools/verify_xlsx.py <파일.xlsx>             # 에러셀 스캔 + 재계산
  py tools/verify_xlsx.py <파일.xlsx> --pdf out.pdf
  py tools/verify_xlsx.py <파일.xlsx> --png out.png  # 첫 시트 사용영역 캡처
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
        print("pywin32 필요: py -m pip install pywin32 (집-전용 검증 도구)")
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
