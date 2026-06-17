"""
tools/excel_shape.py — 통합 Shape 추출기 (구조 + 통계).

dev/검증 전용 (xlwings·win32com·pandas — vendored 아님).
⚠ fpna/ 런타임은 이 파일을 절대 import 하지 않는다 (무설치 규율 불변).
풀스택 환경(win32com·xlwings 0.36.5·pandas 3.0.3)에서 동일 작동.

두 검증 자산을 원본 보존한 채 '연결'만 한다:
  - 구조 shape: tools/metacollector.bas (MetaCollector v1.2.1) 의 CollectReport() COM 호출
  - 통계 shape: fpna.profile.profile_table (8축 SHAPE) 재사용 (CSV 경유, 자산 무수정)

산출: schema-only 단일 텍스트(yaml). 실제 셀 값 없음 — 구조/분포 메타만.
export: py main.py encrypt out/shape_bundle.yaml --mail  (crypto.py part 분할)

CLI:
  py tools/excel_shape.py <xlsx> [-o out/shape_bundle.yaml] [--sheets A,B] [--bas tools/metacollector.bas]
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile

# repo root 를 path 에 주입 → fpna.profile import (자산2 재사용)
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from fpna.profile import profile_table, emit_yaml  # noqa: E402  (자산2, 무수정)

# 데이터 시트가 아닌 것(리포트/메타 시트) 제외 힌트
_NON_DATA_SHEETS = ("report", "metacollector", "control")


def _is_data_sheet(name: str, used_rows: int, used_cols: int) -> bool:
    if name.strip().lower() in _NON_DATA_SHEETS:
        return False
    return used_rows >= 2 and used_cols >= 1


def collect_structure(app, book, bas_path: str) -> str:
    """MetaCollector .bas import → CollectReport() COM 호출 → 구조 리포트 텍스트."""
    bas_abs = os.path.abspath(bas_path)
    try:
        book.api.VBProject.VBComponents.Import(bas_abs)
    except Exception as e:  # noqa: BLE001
        return ("STRUCTURE: (skipped) VBA import 실패 — Excel 옵션 > 보안 센터 > "
                "매크로 설정 > 'VBA 프로젝트 개체 모델 접근 신뢰' 체크 필요. (%s)" % e)
    try:
        book.activate()
        return app.api.Run("CollectReport")  # CollectReport(ActiveWorkbook)
    except Exception as e:  # noqa: BLE001
        return "STRUCTURE: (error) CollectReport 실패 — %s" % e


def profile_sheet(sheet, tmp_dir: str, idx: int) -> dict | None:
    """시트 데이터 → 임시 CSV → profile_table(자산2) → 통계 shape dict."""
    import pandas as pd

    df = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
    if df is None or getattr(df, "empty", True):
        return None
    # datetime 컬럼 → 날짜 문자열 정규화. profile 의 date regex 는 끝앵커($)라
    # 'YYYY-MM-DD HH:MM:SS'(시간부) 를 탈락시켜 date 컬럼을 오분류한다 → 자정/비자정 모두 통일.
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d")
    # 파일명은 시트명 대신 인덱스 — Windows 예약어(NUL/CON)·끝 점·공백 시트명이 경로를 깨뜨림.
    tmp = os.path.join(tmp_dir, "_tmp_%d.csv" % idx)
    df.to_csv(tmp, index=False, encoding="utf-8-sig")
    try:
        return profile_table(tmp, table_name=sheet.name)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


def build_bundle(xlsx_path: str, *, bas_path: str, sheets: list[str] | None = None) -> dict:
    """워크북 1개 → {source, structure_report, tables} 번들."""
    import xlwings as xw

    app = xw.App(visible=False, add_book=False)
    structure, tables = "", {}
    tmp_dir = tempfile.mkdtemp(prefix="excel_shape_")
    try:
        app.display_alerts = False
        app.screen_updating = False
        book = app.books.open(os.path.abspath(xlsx_path))
        try:  # book.close 를 app.quit 과 분리 — 루프 예외 시에도 워크북 정리 보장
            structure = collect_structure(app, book, bas_path)
            for idx, sht in enumerate(book.sheets):
                if sheets is not None and sht.name not in sheets:
                    continue
                ur = sht.used_range
                rows = ur.rows.count if ur else 0
                cols = ur.columns.count if ur else 0
                # --sheets 로 명시 지정한 시트는 NON_DATA 필터 우회(사용자 의도 우선),
                # 단 빈 시트(2행 미만)는 명시여도 skip.
                explicit = sheets is not None and sht.name in sheets
                if not explicit and not _is_data_sheet(sht.name, rows, cols):
                    continue
                if rows < 2 or cols < 1:
                    continue
                spec = profile_sheet(sht, tmp_dir, idx)
                if spec:
                    tables.update(spec.get("tables", {}))  # 시트명 유일 → 충돌 불가
        finally:
            book.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        try:
            app.quit()
        except Exception:  # noqa: BLE001  (quit 실패해도 예외 전파 차단)
            pass
    return {"source": os.path.basename(xlsx_path),
            "structure_report": structure, "tables": tables}


def bundle_to_text(bundle: dict) -> str:
    """메일 본문 운반용 단일 텍스트 — 구조 리포트 + 통계 SHAPE(yaml). 실데이터 값 없음."""
    out = [
        "# === EXCEL SHAPE BUNDLE (schema-only, no data values) ===",
        "# source: %s" % bundle.get("source"),
        "",
        "## --- STRUCTURE (MetaCollector v1.2.1) ---",
        bundle.get("structure_report") or "(none)",
        "",
        "## --- STATISTICAL SHAPE (profile, 8 axes) ---",
        emit_yaml({"tables": bundle.get("tables", {})}),
    ]
    return "\n".join(out)


def main(argv=None):
    p = argparse.ArgumentParser(description="구조+통계 통합 Shape 추출 (dev 전용)")
    p.add_argument("xlsx", help="대상 워크북(.xlsx/.xlsm)")
    p.add_argument("-o", "--out", default="out/shape_bundle.yaml")
    p.add_argument("--sheets", help="콤마구분 시트명(미지정=데이터 시트 자동 감지)")
    p.add_argument("--bas", default=os.path.join(os.path.dirname(__file__), "metacollector.bas"))
    args = p.parse_args(argv)

    sheets = [s.strip() for s in args.sheets.split(",")] if args.sheets else None
    bundle = build_bundle(args.xlsx, bas_path=args.bas, sheets=sheets)
    text = bundle_to_text(bundle)

    out_abs = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out_abs), exist_ok=True)
    with open(out_abs, "w", encoding="utf-8") as f:
        f.write(text)
    print("shape bundle -> %s (%d tables, structure %d chars)"
          % (args.out, len(bundle["tables"]), len(bundle["structure_report"])))
    print("transport: py main.py encrypt %s --mail" % args.out)


if __name__ == "__main__":
    main()
