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

    스파인이 이걸 받아 assert_tie_out 한다. rhs None(키 부재)도 그대로 전달 →
    스파인이 실패로 처리(보고 누락 = 게이트 통과 금지).
    """
    out = []
    for s in specs:
        lhs = s.sign * float(s.raw_sum_fn(source))
        rhs = meta.get(s.reported_key) if meta else None
        out.append((s.name, lhs, rhs, s.tol))
    return out


__all__ = ["ConserveSpec", "eval_specs"]
