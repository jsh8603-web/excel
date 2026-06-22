#!/usr/bin/env python3
"""
verify_workbook.py — freehand-excel-integrity 신호 라우터 + 규약 게이트 (no free-pass).

백엔드는 3종(쓰기 툴): **openpyxl / xlsxwriter / xlwings**. pywin32 COM 은 독립 백엔드가
아니라 **xlwings 의 탈출구**(xlwings.api)로만 쓴다 — 피벗/슬라이서 같은 Excel 고유기능이
필요할 때 xlwings 가 내부에서 COM 으로 강하한다.

[라우팅] 퍼센트가 아니라 **신호 결정함수**(route)로 정한다:
  - 기존 .xlsx 편집:  열려있음(라이브) → xlwings / 닫힘 → openpyxl
  - 신규:             대량+차트 → xlsxwriter / 그 외 → openpyxl
  - 피벗·슬라이서 등 Excel 고유기능 → xlwings (필요 시 COM 강하)

[게이트] (1) 계약 커버리지([7] 완전: 모든 합계·비율이 ties/ratios 선언) (2) 편집 경로면
roundtrip 부품 보존 필수(신규 생성 제외) (3) 재계산 폴백 체인(pywin32→LibreOffice→formulas
→정적 skip) (4) 수식 일관성 린터(fill-down 파손). 다운그레이드는 침묵 금지 — DOWNGRADE 명시.

Exit: 0 = 통과 / 1 = 규약 미충족·검증 실패 / 2 = 사용 오류.
"""
from __future__ import annotations

import argparse
import importlib.util
import os
import subprocess
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))

# 백엔드별 요구: import 모듈 후보[any], 필요한 sibling 스크립트[all], 부재 시 폴백.
_BACKEND_REQS = {
    "openpyxl":   {"modules": ["openpyxl"],                  "scripts": ["xlsx_doctor.py", "formula_lint.py"],
                   "fallback": None, "note": "(vendor 동봉 — 없을 일 없음)"},
    "xlsxwriter": {"modules": ["xlsxwriter"],                "scripts": ["xlsx_doctor.py", "formula_lint.py"],
                   "fallback": "openpyxl", "note": "openpyxl 로 작성 (차트도 openpyxl 지원, bulk 성능만 하락)"},
    "xlwings":    {"modules": ["xlwings", "win32com.client"], "scripts": ["roundtrip_gate.py", "verify_xlsx.py"],
                   "fallback": "openpyxl", "note": "무설치/CI 에선 xlwings 불가 → openpyxl(+tools/format-ooxml.py)"},
}

# 백엔드 기능 capability (다운그레이드 탐지용).
_CAP = {
    "openpyxl":   {"chart", "bulk", "formula"},
    "xlsxwriter": {"chart", "bulk", "formula"},
    "xlwings":    {"chart", "bulk", "formula", "pivot", "slicer", "live"},
}


def _module_available(name: str) -> bool:
    for cand in ("../vendor", "../../vendor", "../../../vendor"):
        p = os.path.join(_HERE, cand)
        if os.path.isdir(p) and p not in sys.path:
            sys.path.insert(0, p)
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError, ModuleNotFoundError):
        return False


# ── 라우팅 (신호 결정함수) ───────────────────────────────────────────────────
def route(*, editing: bool, file_open: bool, bulk: bool, chart: bool, excel_feature: bool) -> str:
    """신호 → 백엔드. 순수 함수(환경 무관). 환경 가용성은 resolve() 가 별도로 본다."""
    if excel_feature:                      # 피벗/슬라이서 → xlwings (내부서 COM 강하)
        return "xlwings"
    if editing:                            # 기존 파일 편집
        return "xlwings" if file_open else "openpyxl"
    return "xlsxwriter" if (bulk and chart) else "openpyxl"   # 신규


def resolve(backend: str, wants: set[str]) -> tuple[str, list[str]]:
    """환경 가용성 적용 → (실효 백엔드, DOWNGRADE 메시지들). 침묵 폴백 금지."""
    downgrades: list[str] = []
    eff = backend
    req = _BACKEND_REQS.get(backend)
    if req and req["modules"] and not any(_module_available(m) for m in req["modules"]):
        fb = req["fallback"] or backend
        eff = fb
        lost = (wants & _CAP.get(backend, set())) - _CAP.get(eff, set())
        for feat in sorted(lost):
            downgrades.append(f"DOWNGRADE: {feat} 요청 → {backend} 불가({req['note']}) → {eff} 로 대체(해당 기능 없음)")
        if not lost:
            downgrades.append(f"DOWNGRADE: {backend} 모듈 부재 → {eff} 로 대체({req['note']})")
    # 실효 백엔드가 못 주는 요청 기능(폴백 아니어도) 직접 점검
    for feat in sorted(wants - _CAP.get(eff, set())):
        msg = f"DOWNGRADE: {feat} 요청 → {eff} 미지원(기능 생략)"
        if msg not in downgrades and not any(feat in d for d in downgrades):
            downgrades.append(msg)
    return eff, downgrades


# ── PREFLIGHT (진입 보장 + 재계산 엔진 탐지) ─────────────────────────────────
def preflight(backend: str) -> tuple[bool, list[str]]:
    req = _BACKEND_REQS.get(backend)
    if not req:
        return True, [f"알 수 없는 backend={backend}, skip"]
    lines, ok = [], True
    mods = req["modules"]
    if mods and not any(_module_available(m) for m in mods):
        ok = False
        lines.append(f"필수 모듈 부재: {' 또는 '.join(mods)} → 폴백: {req['fallback']} {req['note']}")
    else:
        lines.append(f"모듈 OK: {mods}")
    missing = [s for s in req["scripts"] if not os.path.isfile(os.path.join(_HERE, s))]
    if missing:
        ok = False
        lines.append(f"검증 스크립트 부재(번들 손상): {missing}")
    else:
        lines.append(f"검증 스크립트 OK: {req['scripts']}")
    # 재계산 폴백 체인 가용성(게이트3 분기 결정)
    try:
        sys.path.insert(0, _HERE)
        import recalc_check
        eng = recalc_check.detect_engines()
        chosen = "pywin32" if eng["pywin32"] else ("LibreOffice" if eng["soffice"]
                 else ("formulas" if eng["formulas"] else "없음(수식 결과 미검증)"))
        lines.append(f"재계산 엔진: {chosen}  {eng}")
    except Exception as e:
        lines.append(f"재계산 엔진 탐지 실패: {e}")
    return ok, lines


def _run(script: str, *script_args: str) -> int:
    return subprocess.call([sys.executable, os.path.join(_HERE, script), *script_args])


def _run_capture(script: str, *script_args: str) -> tuple[int, str]:
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    p = subprocess.run([sys.executable, os.path.join(_HERE, script), *script_args],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", env=env)
    sys.stdout.write(p.stdout or "")
    return p.returncode, (p.stdout or "")


def _is_formula_workbook(path: str) -> bool:
    sys.path.insert(0, _HERE)
    import recalc_check
    return recalc_check.is_formula_workbook(path)


# ── 게이트 ───────────────────────────────────────────────────────────────────
def verify(path: str, *, backend: str, mode: str, before: str | None, contract: str | None,
           wants: set[str], do_recalc: bool, allow_no_contract: bool) -> int:
    if not os.path.isfile(path):
        print("파일 없음:", path)
        return 2

    failures: list[str] = []
    warns: list[str] = []
    ran: list[str] = []

    editing = (mode == "edit") or (mode == "auto" and bool(before)) or ("live" in wants)
    if backend == "auto":
        backend = route(editing=editing, file_open=("live" in wants),
                        bulk=("bulk" in wants), chart=("chart" in wants),
                        excel_feature=bool(wants & {"pivot", "slicer"}))

    # 환경 적용 + 다운그레이드 투명성
    eff_backend, downgrades = resolve(backend, wants)
    for d in downgrades:
        warns.append(d)

    # 규약 0: preflight (실효 백엔드 기준)
    pf_ok, pf_lines = preflight(eff_backend)
    for ln in pf_lines:
        ran.append(f"preflight: {ln}")
    if not pf_ok:
        failures.append(f"preflight 실패: {eff_backend} 갈래 필수 툴 부재 — 진입 불가.")

    # 규약 1: 계약 커버리지 ([7] 완전이어야 함 — advisory → fatal 격상)
    doctor_args = [path] + (["--contract", contract] if contract else [])
    rc, out = _run_capture("xlsx_doctor.py", *doctor_args)
    ran.append("xlsx_doctor (값·계약)")
    if rc != 0:
        failures.append(f"xlsx_doctor 실패(rc={rc}): 에러셀/텍스트누수/계약 위반 — 해당 셀 재작성.")
    cov_line = next((l for l in out.splitlines() if "[7] 계약 커버리지" in l), "")
    if cov_line and ("완전" not in cov_line):
        if allow_no_contract:
            warns.append(f"계약 커버리지 미완 (--allow-no-contract): {cov_line.strip()}")
        else:
            failures.append(f"계약 커버리지 미완: {cov_line.strip()} — 합계·비율을 ties/ratios 에 전부 선언.")

    # 규약 2: 편집 경로면 roundtrip 부품 보존 필수 (신규 생성 제외)
    if editing:
        if not before:
            failures.append("규약 누락: 편집 경로인데 --before 스냅샷 없음 — 부품 보존 증명 불가.")
        elif not os.path.isfile(before):
            failures.append(f"--before 파일 없음: {before}")
        else:
            rc = _run("roundtrip_gate.py", before, path)
            ran.append("roundtrip_gate (부품 보존)")
            if rc != 0:
                failures.append("roundtrip_gate 실패: 외부링크/차트/피벗/Table 소실 — format-ooxml/xlwings 로 교체.")
    else:
        ran.append("roundtrip: 신규 생성 → 면제")

    # 규약 3: 재계산 폴백 체인 (수식 워크북만)
    if do_recalc:
        if _is_formula_workbook(path):
            rc, _ = _run_capture("recalc_check.py", path)
            ran.append("recalc_check (재계산 체인)")
            if rc == 1:
                failures.append("recalc 실패: 재계산 후 에러셀 발견.")
            elif rc == 2:
                warns.append("recalc 미검증: 재계산 엔진 없음(pywin32/LibreOffice/formulas) — 커버리지 한계.")
        else:
            ran.append("recalc: 정적 값 → 불요(doctor 오프라인 검증)")

    # 규약 4: 수식 일관성 린터 (fill-down 파손)
    if _is_formula_workbook(path):
        rc, _ = _run_capture("formula_lint.py", path)
        ran.append("formula_lint (fill-down 일관성)")
        if rc == 1:
            failures.append("formula_lint 실패: 열 수식 상대참조 형태 불일치(fill-down 파손).")

    # ── 요약 ──
    print(f"\n[verify_workbook] requested={backend} effective={eff_backend} mode={'edit' if editing else 'new'} file={os.path.basename(path)}")
    for r in ran:
        print(f"  · {r}")
    for w in warns:
        print(f"  ⚠ {w}")
    for f in failures:
        print(f"  ✗ FAIL: {f}")
    if failures:
        print(f"\nFAIL: 규약 {len(failures)}건 미충족 — 프리패스 차단.")
        return 1
    print("\nPASS: 백엔드별 필수 검증 모두 통과." + (" (다운그레이드 있음 ↑)" if downgrades else ""))
    return 0


def _parse_wants(s: str | None) -> set[str]:
    return {w.strip() for w in (s or "").split(",") if w.strip()}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="freehand-excel-integrity 신호 라우터 + no-free-pass 게이트")
    ap.add_argument("path", nargs="?")
    ap.add_argument("--backend", default="auto", choices=["auto", "openpyxl", "xlsxwriter", "xlwings"])
    ap.add_argument("--mode", default="auto", choices=["auto", "new", "edit"])
    ap.add_argument("--before", default=None, help="편집 전 원본(roundtrip 비교용)")
    ap.add_argument("--contract", default=None)
    ap.add_argument("--want", default=None, help="요청 기능 csv: pivot,slicer,chart,bulk,live,formula")
    ap.add_argument("--no-recalc", action="store_true", help="재계산 체인 생략")
    ap.add_argument("--allow-no-contract", action="store_true")
    ap.add_argument("--preflight", action="store_true", help="진입 보장만 점검 후 종료")
    ap.add_argument("--route", action="store_true", help="신호→백엔드 라우팅 결정만 출력")
    a = ap.parse_args(argv)
    wants = _parse_wants(a.want)

    if a.route:
        editing = a.mode == "edit" or "live" in wants
        b = route(editing=editing, file_open=("live" in wants), bulk=("bulk" in wants),
                  chart=("chart" in wants), excel_feature=bool(wants & {"pivot", "slicer"}))
        eff, dg = resolve(b, wants)
        print(f"[route] signals: editing={editing} live={'live' in wants} bulk={'bulk' in wants} "
              f"chart={'chart' in wants} excel_feature={bool(wants & {'pivot','slicer'})}")
        print(f"  → backend={b}  effective={eff}")
        for d in dg:
            print(f"  ⚠ {d}")
        return 0

    if a.preflight:
        b = a.backend if a.backend != "auto" else route(
            editing=(a.mode == "edit" or "live" in wants), file_open=("live" in wants),
            bulk=("bulk" in wants), chart=("chart" in wants),
            excel_feature=bool(wants & {"pivot", "slicer"}))
        eff, dg = resolve(b, wants)
        ok, lines = preflight(eff)
        print(f"[preflight] requested={b} effective={eff}")
        for ln in lines:
            print(("  ✓ " if ok else "  · ") + ln)
        for d in dg:
            print(f"  ⚠ {d}")
        print("PASS: 진입 가능." if ok else "FAIL: 필수 툴 부재 — 폴백 라우팅.")
        return 0 if ok else 1

    if not a.path:
        ap.error("path 필요 (또는 --route/--preflight)")
    return verify(a.path, backend=a.backend, mode=a.mode, before=a.before, contract=a.contract,
                  wants=wants, do_recalc=not a.no_recalc, allow_no_contract=a.allow_no_contract)


if __name__ == "__main__":
    sys.exit(main())
