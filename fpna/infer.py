"""
fpna.infer — tidy 컬럼 의미 추론 (measure / dimension / time / id).

실무 데이터(문자 식별자 다수 + 수치 measure 희소)를 템플릿에 자동 매핑하기 위한
1차 부품. 도메인 사전 없이 **통계적·결정적**으로 컬럼 역할을 분류한다 — 같은 입력은
항상 같은 분류(임계값 고정). measure 가 0개여도 동작(전부 dimension/id = 정리표 경로).

분류 축:
  - time      : 날짜/기간 패턴(YYYY-MM[-DD], FY2024-P01 등) 비율 높음
  - measure   : 숫자 파싱 비율 높음 + 고카디널리티(연속값). 저카디 숫자는 범주코드 → dimension
  - id        : 문자 + 카디널리티 ≈ 행수(거의 유일) = 식별자
  - dimension : 그 외 문자(반복 라벨) 또는 저카디 숫자(코드)

⛔ 도메인 고유어/식별자 하드코딩 0. 컬럼명이 아니라 *값의 통계*로만 판정.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

import fpna._bootstrap  # noqa: F401

# 날짜/기간 패턴 (값 기반 판정)
_DATE_RE = re.compile(r"^\d{4}[-/.]\d{1,2}([-/.]\d{1,2})?$")
_PERIOD_RE = re.compile(r"^(fy)?\s?\d{4}[-_/ ]?[pq]?\d{1,2}$", re.IGNORECASE)
_NUM_RE = re.compile(r"^-?\(?[\d,]+(\.\d+)?\)?$")     # 천단위/괄호음수 허용

ROLES = ("time", "measure", "dimension", "id")


@dataclass
class ColumnRole:
    """1개 컬럼의 추론 결과."""
    name: str
    role: str               # ROLES 중 하나
    dtype: str              # "number" | "text" | "date"
    cardinality: int        # 비결측 고유값 수
    n: int                  # 표본 행수
    null_rate: float        # 결측 비율
    num_rate: float         # 숫자 파싱 비율
    confidence: float = 1.0


# --------------------------------------------------------------------------- #
# 값 판정 (stdlib)                                                            #
# --------------------------------------------------------------------------- #
def _looks_number(v) -> bool:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return True
    s = str(v).strip()
    return bool(s) and bool(_NUM_RE.match(s))


def _looks_date(v) -> bool:
    s = str(v).strip()
    return bool(_DATE_RE.match(s) or _PERIOD_RE.match(s))


def _columns_union(rows: list) -> list:
    """행마다 키가 달라도 등장 순서를 보존한 컬럼 합집합."""
    seen: dict = {}
    for r in rows:
        for k in r.keys():
            if k not in seen:
                seen[k] = None
    return list(seen.keys())


# --------------------------------------------------------------------------- #
# 추론                                                                         #
# --------------------------------------------------------------------------- #
def infer_columns(rows: list, *, sample: int = 200) -> list:
    """tidy rows(list[dict]) → list[ColumnRole]. 결정적.

    표본(앞 sample 행)으로 컬럼별 통계를 내고 역할을 부여한다. 저카디 숫자 임계 =
    max(1, min(12, nn//10)) — 12 이하 또는 표본의 10% 이하 고유값이면 범주코드로 본다.
    """
    if not rows:
        return []
    cols = _columns_union(rows)
    sub = rows[:sample]
    out: list = []
    for c in cols:
        vals = [r.get(c) for r in sub]
        nonnull = [v for v in vals if v not in (None, "")]
        n = len(vals)
        nn = len(nonnull)
        null_rate = (1.0 - nn / n) if n else 1.0
        card = len(set(str(v) for v in nonnull))
        num_rate = (sum(1 for v in nonnull if _looks_number(v)) / nn) if nn else 0.0
        date_rate = (sum(1 for v in nonnull if _looks_date(v)) / nn) if nn else 0.0

        if nn == 0:
            role, dtype, conf = "dimension", "text", 0.3      # 전부 결측 — 약한 dimension
        elif date_rate >= 0.8:
            role, dtype, conf = "time", "date", date_rate
        elif num_rate >= 0.8:
            # 저카디 숫자 = 범주코드(고유값이 적고 행 대비 비율 낮음). 연속 measure 는
            # 카디널리티가 행수에 가깝다. 절대 상한(20) + 비율(≤0.5) 동시 충족 시 dimension.
            if card <= 20 and (card / nn) <= 0.5:
                role, dtype, conf = "dimension", "number", 0.7
            else:
                role, dtype, conf = "measure", "number", num_rate
        else:
            if card >= 0.9 * nn:                              # 거의 유일 = 식별자
                role, dtype, conf = "id", "text", card / nn
            else:
                role, dtype, conf = "dimension", "text", 1.0 - (card / nn if nn else 0)
        out.append(ColumnRole(c, role, dtype, card, n, round(null_rate, 4),
                              round(num_rate, 4), round(conf, 4)))
    return out


def summarize(roles: list) -> dict:
    """역할별 컬럼명 묶음 + measure 유무. 템플릿 추천(dispatch)의 입력."""
    by = {r: [] for r in ROLES}
    for cr in roles:
        by[cr.role].append(cr.name)
    return {
        "time": by["time"], "measure": by["measure"],
        "dimension": by["dimension"], "id": by["id"],
        "has_measure": bool(by["measure"]),
        "n_measure": len(by["measure"]), "n_dimension": len(by["dimension"]),
    }


__all__ = ["ColumnRole", "ROLES", "infer_columns", "summarize",
           "_looks_number", "_looks_date"]
