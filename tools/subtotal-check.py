#!/usr/bin/env python3
"""
subtotal-check.py — 소계/합계 정합 게이트 (자식합 == 소계).

발견된 갭: ingest 의 tidy 는 row_role(data/subtotal/total)을 옳게 태깅하지만, 파이프라인의
smell_report 는 '수식 스멜'만 본다. 자식합 != 소계 인 모순(전형적 FP&A 사고)은 검출되지 않는다.
이 게이트가 그걸 메운다. 순수 파이썬, tidy.csv 만 소비(신규 의존성 0).

스코핑(중요): 소계는 '같은 (period, metric) 안에서, 직전 소계 이후 ~ 이 소계 직전' 의 data 행만 합산한다.
src_row 순서로 누적/리셋하므로, 소계 뒤에 오는 별개 항목(예: '기타')을 자식으로 오인하지 않는다.

Run: python3 tools/subtotal-check.py <tidy.csv> [--abs-tol 0.5] [--rel-tol 0.001]
Exit: 0 = 정합, 1 = 불일치 있음.
"""
from __future__ import annotations
import csv, sys, argparse
from collections import defaultdict

SUBTOTAL_ROLES = {"subtotal", "total"}

def load_rows(tidy_csv):
    rows = []
    with open(tidy_csv, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            v = (r.get("value") or "").strip()
            try:
                val = float(v) if v != "" else None
            except ValueError:
                val = None
            rows.append({
                "entity": r.get("entity", ""), "period": r.get("period", ""),
                "metric": r.get("metric", ""), "value": val,
                "role": (r.get("row_role") or "data").strip(),
                "src_row": int(r.get("src_row") or 0),
            })
    return rows

def check(tidy_csv, abs_tol=0.5, rel_tol=0.0):
    rows = load_rows(tidy_csv)
    # group by (period, metric); within group walk by src_row
    groups = defaultdict(list)
    for r in rows:
        groups[(r["period"], r["metric"])].append(r)
    violations = []
    for (period, metric), grp in groups.items():
        grp.sort(key=lambda r: r["src_row"])
        running = 0.0
        had_data = False
        for r in grp:
            if r["role"] in SUBTOTAL_ROLES:
                if r["value"] is not None and had_data:
                    tol = max(abs(r["value"]) * rel_tol, abs_tol)
                    if abs(running - r["value"]) > tol:
                        violations.append({
                            "period": period, "metric": metric, "entity": r["entity"],
                            "subtotal": r["value"], "children_sum": running,
                            "diff": r["value"] - running, "src_row": r["src_row"],
                        })
                running = 0.0  # reset after each subtotal
                had_data = False
            elif r["role"] == "data" and r["value"] is not None:
                running += r["value"]
                had_data = True
    return violations

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("tidy", help="ingest 산출 tidy.csv")
    ap.add_argument("--abs-tol", type=float, default=0.5)
    ap.add_argument("--rel-tol", type=float, default=0.0)
    a = ap.parse_args()
    v = check(a.tidy, a.abs_tol, a.rel_tol)
    if not v:
        print("SUBTOTAL CHECK PASS: 모든 소계가 자식합과 일치")
        return 0
    for x in v:
        print(f"  ✗ {x['period']}/{x['metric']} '{x['entity']}'(r{x['src_row']}): "
              f"소계={x['subtotal']:g} 이나 자식합={x['children_sum']:g} (차이 {x['diff']:+g})")
    print(f"SUBTOTAL CHECK FAIL: {len(v)}건 불일치")
    return 1

if __name__ == "__main__":
    sys.exit(main())
