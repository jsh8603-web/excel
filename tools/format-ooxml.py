#!/usr/bin/env python3
"""
format-ooxml.py — 연결을 끊지 않고 서식만 바꾸는 '외과적' 패처 (openpyxl 미사용).

배경: openpyxl 은 워크북을 통째로 재작성하며 connections.xml(외부 DB)·Power Query(customXml)·
externalLinks·data model 을 떨군다. 통합도 어렵고 Excel COM 도 못 쓰는 환경(리눅스 CI 등)을 위한
최후 수단: **xl/styles.xml 만 다시 쓰고 나머지 모든 부품은 zip 에서 바이트 그대로 복사**한다.
→ 연결/쿼리/외부링크/모델이 전혀 손상되지 않는다(부품을 건드리지 않으므로).

한계: 셀별 정밀 타깃은 sheet XML 외과수술이 필요(여기선 styles.xml 의 numFmt/cellXfs 수준 적용).
Windows+Excel 이 있으면 COM/xlwings(실제 Excel)로 제자리 편집하는 게 더 풍부하고 안전하다.

쓰임: python3 tools/format-ooxml.py <in.xlsx> <out.xlsx> [--numfmt "#,##0;(#,##0)"]
"""
from __future__ import annotations
import sys, zipfile, argparse, re
import xml.etree.ElementTree as ET

SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", SS)
Q = lambda t: f"{{{SS}}}{t}"

def patch_styles_xml(xml_bytes: bytes, numfmt: str, fmt_id: int = 164) -> bytes:
    root = ET.fromstring(xml_bytes)
    # 1) numFmts 에 우리 포맷 추가(없으면 생성)
    numFmts = root.find(Q("numFmts"))
    if numFmts is None:
        numFmts = ET.Element(Q("numFmts"))
        root.insert(0, numFmts)  # styles.xml 첫 자식이어야 스키마상 안전
    if not any(nf.get("numFmtId") == str(fmt_id) for nf in numFmts.findall(Q("numFmt"))):
        nf = ET.SubElement(numFmts, Q("numFmt"))
        nf.set("numFmtId", str(fmt_id)); nf.set("formatCode", numfmt)
    numFmts.set("count", str(len(numFmts.findall(Q("numFmt")))))
    # 2) cellXfs 의 각 xf 에 숫자포맷 적용(텍스트/숫자 구분은 styles 수준에서 불가 → 전체 적용; 거친 적용)
    cellXfs = root.find(Q("cellXfs"))
    if cellXfs is not None:
        for xf in cellXfs.findall(Q("xf")):
            xf.set("numFmtId", str(fmt_id)); xf.set("applyNumberFormat", "1")
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")

def format_ooxml(in_path: str, out_path: str, numfmt: str = "#,##0;(#,##0)") -> dict:
    zin = zipfile.ZipFile(in_path)
    copied, patched = 0, 0
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                data = patch_styles_xml(data, numfmt); patched += 1
            else:
                copied += 1            # ← 연결/쿼리/외부링크/모델 전부 바이트 그대로
            zout.writestr(item, data)
    return {"copied_verbatim": copied, "patched": patched}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inp"); ap.add_argument("out")
    ap.add_argument("--numfmt", default="#,##0;(#,##0)")
    a = ap.parse_args()
    r = format_ooxml(a.inp, a.out, a.numfmt)
    print(f"styles.xml 패치 {r['patched']}개, 나머지 {r['copied_verbatim']}개 부품 바이트복사 → {a.out}")
    print("연결/쿼리/외부링크/모델 부품은 건드리지 않음. roundtrip-gate 로 검증 권장.")

if __name__ == "__main__":
    main()
