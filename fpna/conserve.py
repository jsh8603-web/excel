"""
fpna.conserve — 독립 재구성 reconciliation (N-version) 선언 + 스윕.

자문 3R 프리미티브: "독립 재구성 reconciliation을 hard gate로". 검증 독립성 =
검증 경로(raw_sum_fn)가 빌드 경로(build 헬퍼)를 import/호출하지 않는 N-version.
비독립(qc가 build 재호출)이면 같은 버그를 두 경로가 반복 → 침묵형 오답.

ConserveSpec(선언) = 보일러플레이트 킬러. 템플릿이 (raw_sum_fn, reported_key)만
선언하면 스파인이 대조한다:
  lhs = sign · raw_sum_fn(source)   # 독립 경로(stdlib + INPUT/facts 만)
  rhs = meta[reported_key]          # build 가 _fpna_meta 에 심은 주장값
  ok  = |lhs - rhs| <= tol

★모듈경계 독립성: raw_sum_fn 본문은 fpna.templates.* 의 build 헬퍼를 부르면 안 된다.
  tests/test_conserve.py 가 ast 로 본 모듈/스펙 정의처의 import 를 검사해 강제한다.

source-resolver 제네릭(자문 R3): raw_sum_fn 은 단일 INPUT 뿐 아니라 여러 시트 facts
dict 도 source 로 받을 수 있다 → B 크로스시트 tie 가 같은 ConserveSpec 으로 재사용된다.

★검출력 한계(독립 리뷰 2026-06-13 — 정직 박제):
  - **집계 spec**(sum(parts) == 보고 grand): raw_sum_fn 과 build 의 grand 가 같은 집계식이면
    이건 N-version 이 아니라 **drift/메타 손상 가드**다. 보고값 표류·메타 오염은 잡지만,
    집계식 자체의 버그(부호·연산자)는 두 경로가 같이 틀려 침묵 통과한다.
  - **변환 spec**(상각/리스/NPV/롤포워드): build 와 다른 알고리즘 경로로 재유도하면
    진짜 formula N-version 이라 off-by-one·부호를 노출한다. 이쪽이 침묵형 오답의 해독제.
  - 변이 하니스(test)는 INPUT 만 흔들고 식을 co-mutate 하지 않으므로 위 집계 tautology 를
    구조적으로 못 잡는다 → 집계 spec 의 신뢰는 "drift 가드"로 한정해 읽어야 한다.
  후속(point reported_key at the rendered total CELL)으로 집계도 진짜 독립화 가능(plan §C).
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

import fpna._bootstrap  # noqa: F401


@dataclass(frozen=True)
class ConserveSpec:
    """보존 등식 1개 선언.

    name        : 게이트 라벨.
    raw_sum_fn  : (source) -> float. 독립 재계산(stdlib + INPUT/facts 만, build 호출 금지).
    reported_key: build 가 wb._fpna_meta 에 심은 보고값 키.
    tol         : 표시단위 반올림 흡수(기본 0.5).
    sign        : raw 측 부호(+1/-1).
    """
    name: str
    raw_sum_fn: Callable
    reported_key: str
    tol: float = 0.5
    sign: int = 1


def eval_specs(specs, source, meta) -> list:
    """specs 를 (name, lhs_independent, rhs_reported, tol) 리스트로 평가.

    스파인이 이걸 받아 assert_tie_out 한다. lhs/rhs None(raw 계산 예외 또는 키 부재)도
    그대로 전달 → 소비자가 실패로 처리(게이트 통과 금지). raw_sum_fn 내부 예외(키 오타·
    0나눗셈·None)는 여기서 격리 → 호출자(build_report/qc) 전체 크래시 방지.
    """
    out = []
    for s in specs:
        try:
            lhs = s.sign * float(s.raw_sum_fn(source))
        except Exception:
            lhs = None                      # raw 독립계산 실패 = 게이트 FAIL(크래시 X)
        rhs = meta.get(s.reported_key) if meta else None
        out.append((s.name, lhs, rhs, s.tol))
    return out


__all__ = ["ConserveSpec", "eval_specs"]
