#!/usr/bin/env python3
"""
tools/restyle.py — 외부 입수 .xlsx(파이프라인 밖에서 만든 파일)를 house_style 표준으로
*비파괴* 정규화한다. 서식만 바꾸고 숫자·수식 값은 절대 건드리지 않는다(저장 전 단언).

외부 파일엔 layout_audit 의 golden diff(우리 산출 기준선)를 쓰지 말 것 — 기준선이 없어
틀어진 디자인을 정답으로 굳힌다. 대신 표준(house_style)과 직접 대조해 정규화한다.

사용:
  py tools/restyle.py <파일.xlsx>            # 검사+정규화 → <파일>.restyled.xlsx
  py tools/restyle.py <파일.xlsx> -o out.xlsx
"""
import argparse
import sys

from openpyxl import load_workbook

from fpna import design_audit


def main(argv=None):
    ap = argparse.ArgumentParser(description="외부 엑셀 비파괴 디자인 정규화(서식만)")
    ap.add_argument("path")
    ap.add_argument("-o", "--out", default=None)
    a = ap.parse_args(argv)
    out = a.out or (a.path.rsplit(".", 1)[0] + ".restyled.xlsx")
    wb = load_workbook(a.path)
    # 검수(디자인 위반 요약)
    f = design_audit.design_findings(wb)
    n = sum(len(v) for v in f.values())
    print("디자인 위반 %d건 (장식 %d·정렬 %d·폰트 %d·주석 %d)"
          % (n, len(f["decoration"]), len(f["num_align"]), len(f["font"]), len(f["annotation"])))
    # 비파괴 정규화
    changes = design_audit.restyle_inplace(wb)
    wb.save(out)
    print("정규화 %d건 → %s (값·수식 불변 검증 통과)" % (len(changes), out))
    for x in changes[:20]:
        print("  ·", x)
    return 0


if __name__ == "__main__":
    sys.exit(main())
