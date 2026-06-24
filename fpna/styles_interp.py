#!/usr/bin/env python3
"""
fpna/styles_interp.py — 순수파이썬 xl/styles.xml 해석기(applyX 상속 자체 해소).

배경(DESIGN §6): openpyxl 고수준 `cell.font` 등은 OOXML 의 "cellStyleXfs(base) + applyX
override" 상속을 항상 충실히 풀지는 않는다(버전 의존). 이 모듈은 styles.xml 을 직접 읽어
셀의 style index(s) → effective 포맷을 결정적으로 해소한다. 런타임 COM 미사용(stdlib zip+xml).

용도: (1) Excel 수동편집 파일의 effective 값 진단, (2) tools/styles_calibrate.py 가 openpyxl
resolved 와 effective 를 대조해 충실도맵을 1회 빌드. 우리 *생성* 파일(set_cell 직접 Font,
xfId=0 Normal base, applyX 명시)은 openpyxl-resolved == effective 라 보정 불요.
"""
from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _tag(e):
    return e.tag.split("}", 1)[-1]


class StylesIndex:
    """xl/styles.xml 파싱 + xfId 상속 해소."""

    def __init__(self, xlsx_path):
        with zipfile.ZipFile(xlsx_path) as z:
            xml = z.read("xl/styles.xml")
        root = ET.fromstring(xml)
        self.num_fmts = {0: "General"}        # builtin 은 code 생략(0=General 등)
        self.cell_style_xfs = []              # named style base xf
        self.cell_xfs = []                    # 셀 직접 xf
        self.cell_styles = {}                 # name -> xfId
        self.xfid_to_name = {}
        for child in root:
            t = _tag(child)
            if t == "numFmts":
                for nf in child:
                    self.num_fmts[int(nf.get("numFmtId"))] = nf.get("formatCode")
            elif t == "cellStyleXfs":
                self.cell_style_xfs = [self._xf(x) for x in child]
            elif t == "cellXfs":
                self.cell_xfs = [self._xf(x) for x in child]
            elif t == "cellStyles":
                for cs in child:
                    name, xfid = cs.get("name"), int(cs.get("xfId"))
                    self.cell_styles[name] = xfid
                    self.xfid_to_name[xfid] = name

    @staticmethod
    def _xf(x):
        g = x.get
        return {
            "numFmtId": int(g("numFmtId", 0)),
            "fontId": int(g("fontId", 0)),
            "fillId": int(g("fillId", 0)),
            "borderId": int(g("borderId", 0)),
            "xfId": int(g("xfId", 0)) if g("xfId") is not None else None,
            "applyNumberFormat": g("applyNumberFormat") == "1",
            "applyFont": g("applyFont") == "1",
            "applyFill": g("applyFill") == "1",
            "applyBorder": g("applyBorder") == "1",
            "applyAlignment": g("applyAlignment") == "1",
        }

    def effective(self, s_index: int) -> dict:
        """셀 style index → effective {numFmtId, fontId, fillId, borderId, named_style}.

        OOXML 실무: cellXfs[s] 의 값이 셀의 effective(Excel 렌더 = COM 검증). applyX=0 인데
        해당 attr 가 default(0) 이고 base(cellStyleXfs[xfId]) 가 non-default 일 때만 base 상속.
        (writer 가 applyX 플래그를 생략해도 cellXfs 의 명시값은 적용됨 — calib 으로 확인.)"""
        if s_index < 0 or s_index >= len(self.cell_xfs):
            return {}
        xf = self.cell_xfs[s_index]
        base = self.cell_style_xfs[xf["xfId"]] if (xf["xfId"] is not None and xf["xfId"] < len(self.cell_style_xfs)) else {}
        out = {}
        for attr, flag in (("numFmtId", "applyNumberFormat"), ("fontId", "applyFont"),
                           ("fillId", "applyFill"), ("borderId", "applyBorder")):
            v = xf[attr]
            if v == 0 and not xf[flag] and base.get(attr, 0):   # 명시 안 했고 default → base 상속
                v = base[attr]
            out[attr] = v
        out["num_code"] = self.num_fmts.get(out["numFmtId"], "?builtin%d" % out["numFmtId"])
        out["named_style"] = self.xfid_to_name.get(xf["xfId"]) if xf["xfId"] is not None else None
        return out


def _sheet_path(z, sheet_name):
    """workbook.xml + rels 로 sheet_name → xl/worksheets/*.xml 경로(결정적)."""
    wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
    rid = None
    for sh in wb_xml.iter("%ssheet" % _NS):
        if sh.get("name") == sheet_name:
            rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            break
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    for rel in rels:
        if rel.get("Id") == rid:
            tgt = rel.get("Target")
            return "xl/" + tgt if not tgt.startswith("/") else tgt.lstrip("/")
    return None


def cell_s_index(xlsx_path, sheet_name, coord):
    """시트 xml 에서 셀의 style index(s 속성)을 직접 읽음(openpyxl 비의존, 결정적)."""
    with zipfile.ZipFile(xlsx_path) as z:
        sp = _sheet_path(z, sheet_name)
        if sp is None:
            return None
        root = ET.fromstring(z.read(sp))
    for c in root.iter("%sc" % _NS):
        if c.get("r") == coord:
            return int(c.get("s", 0))
    return 0


def effective_of_cell(xlsx_path, cell, sheet_name) -> dict:
    """셀 → effective 포맷. 시트 xml 의 s 속성 → styles.xml applyX 해소. 진단/캘리브용."""
    s = cell_s_index(xlsx_path, sheet_name, cell.coordinate)
    return StylesIndex(xlsx_path).effective(s) if s is not None else {}
