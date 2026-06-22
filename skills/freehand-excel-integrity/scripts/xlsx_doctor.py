#!/usr/bin/env python3
"""
xlsx_doctor.py — 클로드 코드가 *직접 작성한* .xlsx 를 검사(+서식 자동수리)하는 단일파일.

의존성: openpyxl 만(템플릿/repo 불필요). 클로드 코드가 어떤 작업 폴더에서 엑셀을
프리핸드로 만들든, 끝에 이 스크립트를 돌려 "깨진 산출"을 스스로 잡고 고치게 한다.
per-change subagent 검수를 대체하는 deterministic 게이트.

무엇을 잡나 (W26 스크린샷 버그 클래스 기준)
  [1] 수식 에러 리터럴(#REF!/#VALUE!/#DIV/0! …)        신뢰 · 수리불가→재생성
  [2] 숫자영역 텍스트 침투(헤더 "AUG" 누수 / 주석 누수)  신뢰 · 수리불가→재생성
  [3] 열 내 서식 불균형(엉뚱한 셀만 다른 서식)            휴리스틱 · --fix 수리
  [4] 가드 없는 나눗셈(#VALUE!/#DIV/0! 후보, CPU/CPP)     자문성
  [5] cross-foot 의심(소계 ≠ 자식합)                      자문성(부호규약 모름)

수리 경계(정직)
  ▶ 서식(3)은 결정적 수리 가능.
  ▶ 값/텍스트(1·2·4·5)는 dead 파일만으로 자동수리 불가 — 올바른 값은 원천에서
    다시 와야 한다. 클로드 코드가 *재작성*하는 게 정답(아래 가이드 참조).

사용
  python xlsx_doctor.py report.xlsx            # 검사만(exit 0/1)
  python xlsx_doctor.py report.xlsx --fix      # 서식 수리 → report.fixed.xlsx
  python xlsx_doctor.py report.xlsx --fix --inplace
"""
from __future__ import annotations

import argparse
import os
import sys
from collections import Counter, defaultdict

try:
    from openpyxl import load_workbook
except ImportError:
    print("openpyxl 필요: pip install openpyxl", file=sys.stderr)
    sys.exit(2)

ERROR_LITERALS = {"#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!",
                  "#NUM!", "#SPILL!", "#CALC!", "#GETTING_DATA"}
HEADER_TOKENS = {"APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
                 "JAN", "FEB", "MAR", "Q1", "Q2", "Q3", "Q4", "FY",
                 "1Q", "2Q", "3Q", "4Q"}


def _isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _isf(v):
    return isinstance(v, str) and v.startswith("=")


# --------------------------------------------------------------------------- #
def scan_error_cells(wb):
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERROR_LITERALS:
                    bad.append("%s!%s=%s" % (ws.title, c.coordinate, c.value))
    return bad


def text_in_numeric_columns(wb, min_majority=3):
    """숫자-다수 열의 텍스트 침투. 헤더토큰은 데이터 시작행보다 아래면 누수,
    그 외 텍스트는 데이터 구간(시작행 이상)이면 누수."""
    out = []
    for ws in wb.worksheets:
        bycol = defaultdict(lambda: {"num_rows": [], "txt": []})
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None or _isf(v):
                    continue
                if _isnum(v):
                    bycol[c.column]["num_rows"].append(c.row)
                elif isinstance(v, str) and v.strip():
                    bycol[c.column]["txt"].append((c.row, c.coordinate, v))
        for col, d in bycol.items():
            if len(d["num_rows"]) < min_majority:
                continue
            first = min(d["num_rows"])
            for r, coord, v in d["txt"]:
                tok = v.strip().upper() in HEADER_TOKENS
                if tok and r > first:
                    out.append((ws.title, coord, v, "헤더토큰 누수"))
                elif not tok and r >= first:
                    out.append((ws.title, coord, v, "주석/텍스트 누수"))
    return out


def column_format_outliers(wb, min_majority=3):
    out = []
    for ws in wb.worksheets:
        bycol = defaultdict(list)
        for row in ws.iter_rows():
            for c in row:
                if _isnum(c.value):
                    bycol[c.column].append((c.coordinate, c.number_format))
        for col, items in bycol.items():
            if len(items) < min_majority + 1:
                continue
            cnt = Counter(nf for _, nf in items)
            (top_nf, top_n), = cnt.most_common(1)
            if top_n < min_majority:
                continue
            for coord, nf in items:
                if nf != top_nf and cnt[nf] <= max(1, top_n // 3):
                    out.append({"sheet": ws.title, "col": col, "coord": coord,
                                "found": nf, "expected": top_nf})
    return out


def unguarded_divisions(wb):
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=") and "/" in v):
                    continue
                up = v.upper()
                if "IFERROR" in up or "ISNUMBER" in up or "IFNA" in up:
                    continue
                out.append("%s!%s=%s" % (ws.title, c.coordinate, v))
    return out


def crossfoot_suspects(wb, abs_tol=0.5):
    susp = []
    for ws in wb.worksheets:
        bycol = defaultdict(list)
        for row in ws.iter_rows():
            for c in row:
                if _isnum(c.value):
                    bycol[c.column].append((c.row, c.value, c))
        for col, items in bycol.items():
            items.sort()
            for i, (r, v, cell) in enumerate(items):
                totalish = bool(getattr(cell.font, "bold", False)) and \
                    getattr(cell.border, "top", None) is not None and \
                    getattr(cell.border.top, "style", None) is not None
                if not totalish or i == 0:
                    continue
                run, pr = [], r
                for rr, vv, _ in reversed(items[:i]):
                    if rr == pr - 1:
                        run.append(vv); pr = rr
                    else:
                        break
                if len(run) >= 2 and abs(sum(run) - v) > abs_tol:
                    susp.append("%s!%s(=%g, 위 %d셀 합=%g)" %
                                (ws.title, cell.coordinate, v, len(run), sum(run)))
    return susp


# --------------------------------------------------------------------------- #
# 사이드카 계약(생성시 _fpna_meta 의 프리핸드 버전) — 값 차원 불변식 포팅       #
#   <파일>.contract.json 로 클로드 코드가 "약속"을 선언하면 doctor 가 대조.      #
# --------------------------------------------------------------------------- #
def _split_ref(ref, default_sheet):
    """'SHEET!A1' / 'A1' → (sheet, addr)."""
    if "!" in ref:
        s, a = ref.split("!", 1)
        return s, a
    return default_sheet, ref


def _cells_in(ws, a1):
    """'A1' 또는 'A1:B5' → 셀 객체 리스트(평탄)."""
    from openpyxl.utils import range_boundaries
    c0, r0, c1, r1 = range_boundaries(a1)
    out = []
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            out.append(ws.cell(row=r, column=c))
    return out


def _totalish_cells(wb):
    """굵게+상단테두리 숫자셀(합계 관습) 좌표 집합."""
    s = set()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if not _isnum(c.value):
                    continue
                if bool(getattr(c.font, "bold", False)) and \
                        getattr(c.border, "top", None) is not None and \
                        getattr(c.border.top, "style", None) is not None:
                    s.add("%s!%s" % (ws.title, c.coordinate))
    return s


def _division_cells(wb):
    s = set()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and "/" in c.value:
                    s.add("%s!%s" % (ws.title, c.coordinate))
    return s


def contract_coverage(wb, contract):
    """계약 커버리지(약한 고리 보완): 합계형 셀·나눗셈 셀이 있는데 ties/ratios 에
    선언이 없으면 WARN. 계약이 *무엇을 안 보는지*를 드러내 작성자가 빠뜨리지 않게.
    반환 [(WARN, msg), ...]."""
    out = []
    contract = contract or {}
    declared_t = set()
    for t in contract.get("ties", []):
        ds = contract.get("sheet") or (wb.worksheets[0].title if wb.worksheets else "")
        s, a = _split_ref(t["total"], ds)
        declared_t.add("%s!%s" % (s, a))
    declared_r = set()
    for r in contract.get("ratios", []):
        ds = contract.get("sheet") or (wb.worksheets[0].title if wb.worksheets else "")
        ref = r["cell"] if isinstance(r, dict) else r
        s, a = _split_ref(ref, ds)
        declared_r.add("%s!%s" % (s, a))

    totals = _totalish_cells(wb)
    divs = _division_cells(wb)
    uncovered_t = totals - declared_t
    uncovered_r = divs - declared_r
    if uncovered_t:
        out.append(("WARN", "합계형 셀 %d개가 ties 미선언: %s (소계 정합 미검증)"
                    % (len(uncovered_t), ", ".join(sorted(uncovered_t)[:6]))))
    if uncovered_r:
        out.append(("WARN", "나눗셈 셀 %d개가 ratios 미선언: %s (#VALUE! 가드 미검증)"
                    % (len(uncovered_r), ", ".join(sorted(uncovered_r)[:6]))))
    return out


def check_contract(wb, contract):
    """사이드카 계약 검증. 반환 [(레벨, 메시지), ...]. 레벨 ∈ {FAIL, WARN}."""
    out = []
    by_sheet = {ws.title: ws for ws in wb.worksheets}
    default_sheet = contract.get("sheet") or (wb.worksheets[0].title if wb.worksheets else "")

    # 1) tie-out (R3/R10/R11/R14): total == Σparts (값) 또는 == expected(독립 N-version).
    for t in contract.get("ties", []):
        ts, ta = _split_ref(t["total"], default_sheet)
        name = t.get("name", t["total"])
        if ts not in by_sheet:
            out.append(("FAIL", "tie '%s': 시트 없음" % name)); continue
        total_cell = _cells_in(by_sheet[ts], ta)[0]
        tv = total_cell.value
        tol = t.get("tol", 0.5)
        # (a) 작성자가 소스에서 독립 계산한 기대총계 — 정적 값 워크북의 진짜 N-version
        if "expected" in t:
            if _isnum(tv):
                if abs(tv - t["expected"]) > tol:
                    out.append(("FAIL", "tie '%s': 렌더 total=%g ≠ expected(독립)=%g"
                                % (name, tv, t["expected"])))
            else:
                out.append(("FAIL", "tie '%s': total 셀이 숫자 아님(%r) — 정적 값 기대" % (name, tv)))
        # (b) parts 와의 내부 정합(렌더 합계 == 렌더 자식합)
        if t.get("parts"):
            ps, pa = _split_ref(t["parts"], default_sheet)
            if ps in by_sheet:
                parts = [c.value for c in _cells_in(by_sheet[ps], pa) if _isnum(c.value)]
                if _isf(tv):
                    rng = pa.split("!")[-1].upper().replace(" ", "")
                    if not ("SUM(" in tv.upper().replace(" ", "") and rng in tv.upper().replace(" ", "")):
                        out.append(("FAIL", "tie '%s': total 수식이 parts(%s) 를 SUM 으로 안 덮음"
                                    % (name, pa.split("!")[-1])))
                elif _isnum(tv):
                    if abs(sum(parts) - tv) > tol:
                        out.append(("FAIL", "tie '%s': total=%g ≠ Σparts=%g (%d개)"
                                    % (name, tv, sum(parts), len(parts))))
                elif "expected" not in t:
                    out.append(("FAIL", "tie '%s': total 셀이 숫자/수식 아님(%r)" % (name, tv)))

    # 2) grain 유일성(R8): 선언 key 범위에 중복 라벨 금지(침묵 병합 차단)
    g = contract.get("grain")
    if g and g.get("region"):
        gs, ga = _split_ref(g["region"], default_sheet)
        ws = by_sheet.get(gs)
        if ws:
            vals = [c.value for c in _cells_in(ws, ga)
                    if isinstance(c.value, str) and c.value.strip()]
            dup = [k for k, n in Counter(vals).items() if n > 1]
            if dup:
                out.append(("FAIL", "grain 중복(침묵 병합 위험): %s" % ", ".join(dup[:6])))

    # 3) ratio 완전성(R17): 정적 값이면 num/den 재계산 대조, 수식이면 가드 확인.
    for ref in contract.get("ratios", []):
        if isinstance(ref, dict):
            # 값 모드: {cell, num, den, tol} — .py 계산 정적 비율의 독립 재계산
            cs, ca = _split_ref(ref["cell"], default_sheet)
            ws = by_sheet.get(cs)
            if not ws:
                continue
            cv = _cells_in(ws, ca)[0].value
            ns, na = _split_ref(ref["num"], default_sheet)
            ds, da = _split_ref(ref["den"], default_sheet)
            num = _cells_in(by_sheet[ns], na)[0].value if ns in by_sheet else None
            den = _cells_in(by_sheet[ds], da)[0].value if ds in by_sheet else None
            tol = ref.get("tol", 0.01)
            if not (_isnum(den) and den != 0):
                # 분모 무효 → 셀은 NA 센티넬(문자열)이거나 빈칸이어야(숫자/에러 금지)
                if _isnum(cv) or (isinstance(cv, str) and cv in ERROR_LITERALS):
                    out.append(("FAIL", "ratio %s: 분모 무효인데 값/에러(%r) — 'NA' 센티넬 권장"
                                % (ref["cell"], cv)))
            elif _isnum(num):
                if not _isnum(cv):
                    out.append(("FAIL", "ratio %s: 숫자 아님(%r), 기대 %g" % (ref["cell"], cv, num/den)))
                elif abs(cv - num/den) > max(tol, abs(num/den)*tol):
                    out.append(("FAIL", "ratio %s: %g ≠ num/den=%g (독립 재계산)"
                                % (ref["cell"], cv, num/den)))
            continue
        # 문자열 ref: 수식 가드 모드(라이브 수식 워크북)
        rs, ra = _split_ref(ref, default_sheet)
        ws = by_sheet.get(rs)
        if not ws:
            continue
        v = _cells_in(ws, ra)[0].value
        if _isf(v):
            up = v.upper()
            if not ("ISNUMBER" in up or "IFERROR" in up or "IFNA" in up):
                out.append(("FAIL", "ratio %s 가드 없음(#VALUE! 위험): %s" % (ref, v)))
        elif isinstance(v, str) and v in ERROR_LITERALS:
            out.append(("FAIL", "ratio %s 이미 에러: %s" % (ref, v)))

    # 4) FieldSpec(metric_table): 범위·부호·허용값 (binding/metric_table 포팅)
    for fs in contract.get("fields", []):
        s, a = _split_ref(fs["region"], default_sheet)
        ws = by_sheet.get(s)
        if not ws:
            continue
        cells = _cells_in(ws, a)
        nums = [(c.coordinate, c.value) for c in cells if _isnum(c.value)]
        sign = fs.get("sign")
        for coord, v in nums:
            if sign in ("+", "nonneg") and v < 0:
                out.append(("FAIL", "field %s!%s 음수(기대 %s): %g" % (s, coord, sign, v)))
            elif sign in ("-", "nonpos") and v > 0:
                out.append(("FAIL", "field %s!%s 양수(기대 %s): %g" % (s, coord, sign, v)))
            if "min" in fs and v < fs["min"]:
                out.append(("FAIL", "field %s!%s < min %s: %g" % (s, coord, fs["min"], v)))
            if "max" in fs and v > fs["max"]:
                out.append(("FAIL", "field %s!%s > max %s: %g" % (s, coord, fs["max"], v)))
        if "accepted_values" in fs:
            allowed = set(fs["accepted_values"])
            for c in cells:
                if c.value is not None and not _isf(c.value) and c.value not in allowed:
                    out.append(("FAIL", "field %s!%s 허용값 밖: %r" % (s, c.coordinate, c.value)))

    # 5) 시나리오 정합(R9): actual/budget 라벨 집합이 같은 모집단인지
    sc = contract.get("scenario")
    if sc and sc.get("actual") and sc.get("budget"):
        def _labels(ref):
            s, a = _split_ref(ref, default_sheet)
            ws = by_sheet.get(s)
            return {c.value for c in _cells_in(ws, a)
                    if isinstance(c.value, str) and c.value.strip()} if ws else set()
        a_keys, b_keys = _labels(sc["actual"]), _labels(sc["budget"])
        ao, bo = a_keys - b_keys, b_keys - a_keys
        if ao or bo:
            out.append(("FAIL", "scenario 모집단 불일치(0처리 금지, LEFT/RIGHT_ONLY 노출): "
                        "actual-only %s, budget-only %s"
                        % (sorted(ao)[:4], sorted(bo)[:4])))

    # 6) 행수 보존(R7 no_silent_drop): 선언 region 의 비어있지 않은 라벨 수 == n
    en = contract.get("expected_n")
    if en and en.get("region") and "n" in en:
        s, a = _split_ref(en["region"], default_sheet)
        ws = by_sheet.get(s)
        if ws:
            got = sum(1 for c in _cells_in(ws, a)
                      if c.value is not None and str(c.value).strip())
            if got != en["n"]:
                out.append(("FAIL", "행수 불일치(누락/추가): 기대 %d ≠ 실제 %d" % (en["n"], got)))

    # 7) 수식 방향/칼럼 참조(E): region 의 각 행 수식이 =<left><행><op><right><행>
    from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
    for fc in contract.get("formula_refs", []):
        s, a = _split_ref(fc["region"], default_sheet)
        ws = by_sheet.get(s)
        if not ws:
            continue
        op = fc.get("op", "-")
        left = fc["left"] if isinstance(fc["left"], int) else column_index_from_string(fc["left"])
        right = fc["right"] if isinstance(fc["right"], int) else column_index_from_string(fc["right"])
        c0, r0, c1, r1 = range_boundaries(a)
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                v = ws.cell(row=r, column=c).value
                if not _isf(v):
                    continue
                want = "=%s%d%s%s%d" % (get_column_letter(left), r, op,
                                       get_column_letter(right), r)
                if v.replace("$", "").replace(" ", "").upper() != want.upper():
                    out.append(("FAIL", "formula %s!%s 참조 불일치: 기대 %s ≠ %s"
                                % (s, ws.cell(row=r, column=c).coordinate, want, v)))

    return out


# --------------------------------------------------------------------------- #
def doctor(path, fix=False, inplace=False, contract_path=None):
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        print("파일 없음:", path); return 2
    wb = load_workbook(path)

    err = scan_error_cells(wb)
    stray = text_in_numeric_columns(wb)
    outliers = column_format_outliers(wb)
    unguarded = unguarded_divisions(wb)
    crossfoot = crossfoot_suspects(wb)

    # 사이드카 계약 자동탐색: <파일>.contract.json
    if contract_path is None:
        cand = path.rsplit(".", 1)[0] + ".contract.json"
        contract_path = cand if os.path.isfile(cand) else None
    contract_findings = []
    coverage = []
    contract_obj = {}
    if contract_path and os.path.isfile(contract_path):
        import json
        with open(contract_path, encoding="utf-8") as f:
            contract_obj = json.load(f)
        contract_findings = check_contract(wb, contract_obj)
    coverage = contract_coverage(wb, contract_obj)

    print("=" * 60)
    print("XLSX DOCTOR:", os.path.basename(path))
    print("=" * 60)
    fatal = False

    print("\n[1] 수식 에러 리터럴: %s" % ("없음" if not err else "%d건" % len(err)))
    for x in err[:12]:
        print("    ✗", x)
    if err:
        fatal = True
        print("    → 자동수리 불가. 참조/분모 오류 — 재작성 필요.")

    print("\n[2] 숫자영역 텍스트 침투: %s" % ("없음" if not stray else "%d건" % len(stray)))
    for s, coord, v, kind in stray[:12]:
        print("    ✗ %s!%s = %r  [%s]" % (s, coord, v, kind))
    if stray:
        fatal = True
        print("    → #VALUE! 의 상류 원인(분모 오염). 값 자동수리 불가 — 재작성 필요.")

    print("\n[3] 열 내 서식 불균형: %s" % ("없음" if not outliers else "%d건" % len(outliers)))
    for o in outliers[:12]:
        print("    ✗ %s!%s = %r (열 다수서식 %r)" %
              (o["sheet"], o["coord"], o["found"], o["expected"]))

    print("\n[4] 가드 없는 나눗셈: %s  [자문성]" %
          ("없음" if not unguarded else "%d건" % len(unguarded)))
    for x in unguarded[:12]:
        print("    ⚠", x)
    if unguarded:
        print("    → CPU/CPP 류는 =IF(OR(NOT(ISNUMBER(분모)),분모=0),\"NA\",분자/분모) 로.")

    print("\n[5] cross-foot 의심: %s  [자문성·부호규약 모름→오탐 가능]" %
          ("없음" if not crossfoot else "%d건" % len(crossfoot)))
    for x in crossfoot[:12]:
        print("    ⚠", x)

    print("\n[6] 사이드카 계약(tie/grain/ratio/field/scenario/formula): %s" % (
        "선언 없음" if contract_path is None else
        ("통과 ✓" if not contract_findings else "%d건 위반" % len(contract_findings))))
    for lvl, msg in contract_findings[:12]:
        print("    %s %s" % ("✗" if lvl == "FAIL" else "⚠", msg))
        if lvl == "FAIL":
            fatal = True

    print("\n[7] 계약 커버리지(미선언 합계/비율): %s  [자문성]" %
          ("완전" if not coverage else "%d건" % len(coverage)))
    for _lvl, msg in coverage[:12]:
        print("    ⚠", msg)
    if coverage:
        print("    → 합계/비율은 ties/ratios 로 선언해야 값 검증이 걸린다(미선언=무검증).")

    n_fixed = 0
    if fix and outliers:
        bysheet = {ws.title: ws for ws in wb.worksheets}
        for o in outliers:
            bysheet[o["sheet"]][o["coord"]].number_format = o["expected"]
            n_fixed += 1
        out = path if inplace else path.rsplit(".", 1)[0] + ".fixed.xlsx"
        wb.save(out)
        print("\n[FIX] 서식 %d건 수리 → %s (값/텍스트 이슈는 미수리)" %
              (n_fixed, os.path.basename(out)))

    print("\n" + "-" * 60)
    if not (fatal or outliers):
        print("진단: 치명 이슈 없음 ✓"); return 0
    if fix and not fatal and n_fixed:
        print("진단: 서식 수리 완료. 값/텍스트 치명 이슈 없음 ✓"); return 0
    print("진단: 이슈 발견 — 위 항목 확인.")
    return 1


def main(argv=None):
    ap = argparse.ArgumentParser(description="클로드 코드가 만든 .xlsx 검사/수리(openpyxl만 필요)")
    ap.add_argument("path")
    ap.add_argument("--fix", action="store_true")
    ap.add_argument("--inplace", action="store_true")
    ap.add_argument("--contract", default=None, help="사이드카 계약 JSON 경로(기본: <파일>.contract.json 자동탐색)")
    a = ap.parse_args(argv)
    return doctor(a.path, fix=a.fix, inplace=a.inplace, contract_path=a.contract)


if __name__ == "__main__":
    sys.exit(main())
