import fpna._bootstrap, openpyxl, json, os, shutil
from fpna import design_zones as dz, design_audit as da
import win32com.client as win32
WD="out/zone/regress"; os.makedirs(WD, exist_ok=True)

def gen(path):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="P&L"
    contract=dz.draw_house_block(ws, origin=(6,2),
        rows=[[1,2,3,4],[5,6,7,8],[9,10,11,12]],
        col_bands=[(2,"actual"),(4,"fcst")], bands_by_col={"actual":"calc","fcst":"calc"},
        marker_col=1, marker_row=5)
    wb.save(path); return contract

def findings(path, contract):
    wb=openpyxl.load_workbook(path)
    return da.zone_findings(wb, contract)

contract=gen(os.path.join(WD,"base.xlsx"))
xl=win32.gencache.EnsureDispatch("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
results={}
try:
    # 1) clean (편집 없음) → drift 0
    results["clean"]=findings(os.path.join(WD,"base.xlsx"), contract)
    # 2) ClearFormats on B6 (strict) → drift 검출
    p=os.path.join(WD,"clear.xlsx"); shutil.copy(os.path.join(WD,"base.xlsx"),p)
    wb=xl.Workbooks.Open(os.path.abspath(p)); wb.Worksheets("P&L").Range("B6").ClearFormats(); wb.Save(); wb.Close()
    results["clearformats_B6"]=findings(p, contract)
    # 3) 외부 전체붙여넣기 → B7 → drift
    p=os.path.join(WD,"paste.xlsx"); shutil.copy(os.path.join(WD,"base.xlsx"),p)
    wb=xl.Workbooks.Open(os.path.abspath(p)); ws=wb.Worksheets("P&L")
    ws.Range("H1").Value="ext"; ws.Range("H1").Copy(); ws.Range("B7").PasteSpecial(Paste=-4104)
    xl.CutCopyMode=False; wb.Save(); wb.Close()
    results["fullpaste_B7"]=findings(p, contract)
finally:
    xl.Quit()

print("="*60)
ok=True
for k,z in results.items():
    nd=len(z["resolved_drift"]); nu=len(z["unsealed"])
    print("%-18s drift=%d unsealed=%d %s" % (k, nd, nu, z["resolved_drift"][:2]))
clean_ok = len(results["clean"]["resolved_drift"])==0
clear_ok = len(results["clearformats_B6"]["resolved_drift"])>=1
paste_ok = len(results["fullpaste_B7"]["resolved_drift"])>=1
print("\n판정: clean drift0=%s / ClearFormats 검출=%s / fullpaste 검출=%s" % (clean_ok, clear_ok, paste_ok))
print("3.4 PASS" if (clean_ok and clear_ok and paste_ok) else "3.4 FAIL")
