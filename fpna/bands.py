"""
fpna.bands — 차원량별 plausibility band (단위혼동 방어).

자문 3R(C5): 'rate 5 vs 0.05' 같은 의미적 단위혼동은 차분오라클로 못 잡는다
(둘 다 유효 float). 유일한 stdlib 경로 = 차원량별 그럴듯한 범위를 선언하고 경계 assert.

★한계(박제): 자릿수 오류만 잡는다. 4.8%를 0.048 vs 정답 0.052 는 둘 다 밴드 안 →
'그럴듯하게 틀린 값'엔 무력. 입력단 1차 방어 + 출력단 sanity anchor 용도.
"""
from __future__ import annotations

import fpna._bootstrap  # noqa: F401

# 차원량 → (lo, hi) 그럴듯한 범위. 비율은 소수(0.05=5%) 기준.
BANDS = {
    "rate": (-0.5, 2.0),        # 금리/할인율 -50%~200%
    "pct": (-1.0, 1.0),         # 일반 비율 -100%~100%
    "margin": (-2.0, 1.0),      # 마진 -200%~100%
    "growth": (-1.0, 10.0),     # 성장률 -100%~1000%
    "multiple": (-100.0, 100.0),  # 배수
    "ratio": (0.0, 1000.0),     # 회전율 등 양의 비율
    "share": (0.0, 1.0),        # 구성비 0~100%
}


def in_band(value, dim: str) -> bool:
    """value 가 dim 의 그럴듯한 범위 안인가. 미선언 dim 은 통과(보수)."""
    if value is None or dim not in BANDS:
        return True
    lo, hi = BANDS[dim]
    try:
        return lo <= float(value) <= hi
    except (TypeError, ValueError):
        return False


def assert_band(rep, value, dim: str, *, name: str = "") -> bool:
    """rep 에 band 검증 결과 추가(QCReport.add). 통과 여부 반환."""
    ok = in_band(value, dim)
    label = "단위band:%s(%s)" % (name or dim, dim)
    rep.add(label, ok, "" if ok else "값=%s 범위=%s 이탈(단위혼동 의심)" % (value, BANDS.get(dim)))
    return ok


__all__ = ["BANDS", "in_band", "assert_band"]
