"""
tools/ingest_smoke.py — 디렉토리 엑셀 배치 ingest 스모크 + oracle 정확도.

dev 전용. 폴더 안 *.xlsx 를 전부 ingest 돌려 크래시 여부 / tidy 행수 / smell 종류 /
(dirty_excel_gen 의 *_expected.csv 가 있으면) 복원 정확도를 표로 낸다.

CLI:
  python ingest_smoke.py out/dirty_corpus
"""
from __future__ import annotations

import csv
import glob
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from fpna.ingest import ingest_workbook  # noqa: E402


def _load_expected(path: str) -> dict:
    out = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in list(csv.reader(f))[1:]:
            if len(row) >= 3:
                out[(row[0], row[1])] = float(row[2])
    return out


def smoke(d: str) -> list[dict]:
    rows = []
    for x in sorted(glob.glob(os.path.join(d, "*.xlsx"))):
        rec = {"file": os.path.basename(x), "status": "OK"}
        try:
            res = ingest_workbook(x)
            rec["rows"] = len(res.tidy_rows)
            rec["smells"] = sorted({s["kind"] for s in res.smells})
            exp = x[:-5] + "_expected.csv"
            if os.path.exists(exp):
                E = _load_expected(exp)
                got = {(r.entity, r.period): r.value
                       for r in res.tidy_rows if r.row_role == "data"}
                hit = sum(1 for k, v in E.items()
                          if k in got and got[k] is not None and abs(got[k] - v) < 0.5)
                rec["oracle"] = "%d/%d" % (hit, len(E))
                rec["acc"] = hit / len(E) if E else 1.0
        except Exception as ex:  # noqa: BLE001  (배치 스모크 — 크래시도 수집 대상)
            rec["status"] = "CRASH"
            rec["err"] = ("%s: %s" % (type(ex).__name__, ex))[:80]
        rows.append(rec)
    return rows


def main(argv=None):
    import argparse
    p = argparse.ArgumentParser(description="ingest 배치 스모크 (dev)")
    p.add_argument("dir")
    args = p.parse_args(argv)

    rows = smoke(args.dir)
    crashes = [r for r in rows if r["status"] == "CRASH"]
    accs = [r["acc"] for r in rows if "acc" in r]

    print("=== ingest smoke: %d files ===" % len(rows))
    print("크래시: %d / %d" % (len(crashes), len(rows)))
    if accs:
        print("oracle 복원 정확도: 평균 %.1f%% (min %.0f%%, 100%%=%d개)"
              % (100 * sum(accs) / len(accs), 100 * min(accs),
                 sum(1 for a in accs if a >= 0.999)))
    # 문제 케이스(크래시 또는 정확도 < 100%) 상위 노출
    # tidy 0(silent fail) 도 문제 케이스 — oracle 없는 실데이터에서 조용한 빈 결과 표면화.
    bad = sorted((r for r in rows if r["status"] == "CRASH"
                  or r.get("acc", 1.0) < 1.0 or r.get("rows") == 0),
                 key=lambda r: r.get("acc", 0.0))
    if bad:
        print("--- 문제 케이스 (크래시 / 정확도<100% / tidy 0) ---")
        for r in bad[:15]:
            tidy0 = "  ⚠TIDY0-SILENT-FAIL" if r.get("rows") == 0 and r["status"] == "OK" else ""
            print("  %-26s %-6s rows=%-6s oracle=%-7s %s%s"
                  % (r["file"], r["status"], r.get("rows", "-"),
                     r.get("oracle", "-"), r.get("err", ""), tidy0))
    else:
        print("문제 케이스 없음(크래시 0, tidy 생성, 정확도 100%).")
    return 1 if crashes else 0


if __name__ == "__main__":
    sys.exit(main())
