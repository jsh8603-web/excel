"""
tools/make_base_templates.py — house_style 기반 베이스 템플릿 생성.

각 유형의 골든샘플을 빌드한 뒤 **입력셀(파랑 글씨, role=input)의 값만 비워**
templates_base/<type>_base.xlsx 로 저장한다. 헤더·서식·수식·차트는 보존.
→ 채우기모드(fill)의 캔버스이자, 사람이 Excel 에서 스파크라인/조건부서식/슬라이서 등
  openpyxl 이 못 만드는 요소를 덧입히는 출발점.

⚠ 외부 라이선스 템플릿이 아니라 자작 베이스다(재배포 제약 없음). repo private 전제로 커밋.

사용: py tools/make_base_templates.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import fpna._bootstrap  # noqa: F401

from fpna import house_style as hs
from fpna.templates import available, get_template

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "templates_base")


def _is_input_cell(cell) -> bool:
    """입력셀(파랑 글씨) 판정 — INPUT_FG 색."""
    try:
        rgb = cell.font.color.rgb if cell.font and cell.font.color else None
        return isinstance(rgb, str) and rgb[-6:].upper() == hs.INPUT_FG.upper()
    except Exception:
        return False


def clear_inputs(wb) -> int:
    """입력셀 값 제거(수식·서식·헤더 보존). 비운 셀 수 반환."""
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and not (isinstance(c.value, str)
                                                and c.value.startswith("=")):
                    if _is_input_cell(c):
                        c.value = None
                        n += 1
    return n


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)
    for t in available():
        mod = get_template(t)
        wb = mod.build(mod.golden_sample())
        cleared = clear_inputs(wb)
        out = os.path.join(OUT_DIR, "%s_base.xlsx" % t)
        wb.save(out)
        sys.stdout.buffer.write(
            ("  %s_base.xlsx (입력셀 %d개 비움)\n" % (t, cleared)).encode("utf-8"))
    sys.stdout.buffer.write(("베이스 템플릿 생성 완료 → %s\n" % OUT_DIR).encode("utf-8"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
