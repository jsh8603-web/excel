#!/usr/bin/env python3
"""tools/stress_check.py — ingest 산출 무결점(invariant) 판정.

KOSIS 실데이터(누더기 export)를 ingest 한 결과가 손상 없이 tidy 로 복원됐는지
객관 게이트로 판정한다. PASS/FAIL + 사유. 스트레스 루프의 합격선.

무결점 정의(KOSIS 정형 통계표 — 소계/합계행 없는 평면 분류표 기준):
  1) reject 0                       — 행 무음 손실 없음
  2) reconciliation 전 시트 ok      — covered==gt, missing/mismatch/dup 0
  3) HIGH_REJECT_RATE smell 없음
  4) metric 순수성                  — 데이터행 metric != (순수숫자 | 센티넬 | None)
  5) entity 존재                    — 데이터행 entity != None
  6) period 존재                    — 데이터행 period != None
  7) 헤더 누수 없음                 — entity/metric/period 에 '-'(센티넬)·순수숫자값 침입 0
"""
from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fpna._bootstrap  # noqa: F401
from fpna.ingest.pipeline import run_ingest

_NUMERIC = re.compile(r"^\s*[-+]?\d{1,3}(,\d{3})*(\.\d+)?\s*$")
_SENT = {"-", "–", "—", "...", "…", "N/A", "n/a", "na", "."}


def _is_numeric_str(s) -> bool:
    return isinstance(s, str) and bool(_NUMERIC.match(s)) and any(ch.isdigit() for ch in s)


def check(path: str, out_dir: str, *, sheet: str | None = None) -> tuple:
    res = run_ingest(path, out_dir, sheet=sheet)
    rows = res.tidy_rows
    fails: list[str] = []

    if res.report.n_rejected > 0:
        fails.append("reject %d행" % res.report.n_rejected)

    for rc in res.recon:
        if not rc.ok or rc.missing or rc.value_mismatch or rc.duplicate:
            fails.append("recon[%s] covered=%d/%d missing=%d mismatch=%d dup=%d" % (
                rc.sheet, len(rc.covered), rc.n_groundtruth,
                len(rc.missing), len(rc.value_mismatch), len(rc.duplicate)))

    kinds = {s["kind"] for s in res.smells}
    if "HIGH_REJECT_RATE" in kinds:
        fails.append("HIGH_REJECT_RATE smell")

    data = [r for r in rows if r.row_role == "data"]
    bad_metric = [r for r in data if r.metric is None or _is_numeric_str(r.metric)
                  or (isinstance(r.metric, str) and r.metric.strip() in _SENT)]
    if bad_metric:
        fails.append("metric 불순 %d행 예:%r" % (len(bad_metric), bad_metric[0].metric))

    no_entity = [r for r in data if r.entity is None or str(r.entity).strip() == ""]
    if no_entity:
        fails.append("entity 결측 %d행 (R%dC%d)" % (
            len(no_entity), no_entity[0].src_row, no_entity[0].src_col))

    no_period = [r for r in data if r.period is None or str(r.period).strip() == ""]
    if no_period:
        fails.append("period 결측 %d행" % len(no_period))

    leak = []
    for r in data:
        for fld in (r.entity, r.metric, r.period):
            if isinstance(fld, str) and (fld.strip() in _SENT or _is_numeric_str(fld)):
                leak.append((fld, r.src_row, r.src_col)); break
    if leak:
        fails.append("헤더 누수 %d행 예:%r@R%dC%d" % (len(leak), *leak[0]))

    ok = not fails
    summary = "%-34s %s | rows=%d reject=%d sheets_recon_ok=%s smells=%d" % (
        path.split("/")[-1].split("\\")[-1],
        "PASS" if ok else "FAIL",
        len(rows), res.report.n_rejected,
        all(rc.ok for rc in res.recon), len(res.smells))
    return ok, summary, fails


def main(argv):
    files = argv[1:]
    all_ok = True
    for f in files:
        sheet = None
        if "::" in f:
            f, sheet = f.split("::", 1)
        out = "out/stress/_chk_" + re.sub(r"[^A-Za-z0-9]", "_", f.split("/")[-1])[:20]
        try:
            ok, summary, fails = check(f, out, sheet=sheet)
        except Exception as e:                       # noqa: BLE001
            all_ok = False
            print("ERROR %s | %s" % (f, e))
            continue
        print(summary)
        for x in fails:
            print("   ✗", x)
        all_ok = all_ok and ok
    print("\n>>> 전체:", "ALL PASS" if all_ok else "FAIL 있음")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
