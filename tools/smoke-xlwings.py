#!/usr/bin/env python3
"""
smoke-xlwings.py — [Windows+Excel 전용] 실 Excel 서식 패스가 연결/Table/링크를 보존함을 증명.

이 검증은 Excel 이 있어야만 의미가 있다(리눅스 CI 에선 skip). 실제 Excel 로 열고 서식 적용 후 저장한 뒤,
roundtrip-gate 로 before==after(연결/Table/시트/헤더/외부링크 소실 0)를 확인한다.

Run (Windows, Excel 설치):
  python tools/smoke-xlwings.py                 # Table 있는 샘플 자동 생성·검증
  python tools/smoke-xlwings.py <your.xlsx>     # 실제 연결 워크북으로 검증(사본 권장)
Exit: 0 = 보존 확인, 1 = 위반, 2 = 환경(Excel/xlwings) 아님.
"""
import sys, os, tempfile, shutil, importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))

def _load(name, fn):
    s = importlib.util.spec_from_file_location(name, os.path.join(HERE, fn))
    m = importlib.util.module_from_spec(s); s.loader.exec_module(m); return m

def make_sample(path):
    sys.path.insert(0, os.path.join(HERE, "..", "vendor"))
    from openpyxl import Workbook
    from openpyxl.worksheet.table import Table
    wb = Workbook(); ws = wb.active; ws.title = "Ledger"
    ws.append(["계정", "2024", "2025"]); ws.append(["매출", 1000, 1200]); ws.append(["비용", 600, 700])
    ws.add_table(Table(displayName="tblLedger", ref="A1:C3")); wb.save(path)

def main():
    fx = _load("fx", "format-xlwings.py"); rtg = _load("rtg", "roundtrip-gate.py")
    if len(sys.argv) > 1:
        target = os.path.join(tempfile.mkdtemp(), os.path.basename(sys.argv[1]))
        shutil.copy(sys.argv[1], target)        # 원본 보호: 사본에서 검증
    else:
        target = os.path.join(tempfile.mkdtemp(), "sample.xlsx"); make_sample(target)
        print("샘플 생성:", target)

    before = rtg.fingerprint(target)
    print("적용 전: tables=%s sheets=%s 연결부품=%d" % (list(before["tables"]), before["sheets"], len(before["sig_parts"])))
    try:
        fx.format_xlwings(target)
    except ImportError as e:
        print(f"환경 아님(Excel/xlwings 필요): {e}"); return 2
    except Exception as e:
        print(f"Excel 실행 실패(이 머신에 Excel 없음?): {type(e).__name__}: {e}"); return 2

    issues = rtg.compare(before, rtg.fingerprint(target))
    if issues:
        print("FAIL: 실 Excel 저장 후 무결성 위반")
        for i in issues: print("  ✗", i)
        return 1
    print("PASS: 실 Excel 서식 후 연결/Table/시트/헤더/외부링크 전부 보존 ✓")
    return 0

if __name__ == "__main__":
    sys.exit(main())
