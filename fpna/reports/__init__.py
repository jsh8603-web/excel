"""
fpna.reports — 다중시트 Report 레지스트리 (B 실행경로).

`fpna.report` 는 오케스트레이터(엔진)이고, 여기 `reports/` 는 그 엔진에 먹일
구체 ReportSpec 을 만드는 `make_spec()` 모듈들의 레지스트리다. dispatcher 의
stage="report" 라우팅과 main.py `cmd_report` 가 이 레지스트리를 경유한다.

각 모듈 계약(덕 타이핑):
  - `make_spec(data=None) -> fpna.report.ReportSpec`
"""
from __future__ import annotations

import importlib

import fpna._bootstrap  # noqa: F401

# name → 모듈 경로(지연 import: 무거운 템플릿 의존 회피).
_REPORTS = {
    "fc_boardpack": "fpna.reports.fc_boardpack",
}


def get_report(name: str):
    """레지스트리 name → make_spec 을 노출하는 모듈. 미등록이면 KeyError."""
    if name not in _REPORTS:
        raise KeyError("미등록 리포트: %s (가능: %s)" % (name, ", ".join(available())))
    return importlib.import_module(_REPORTS[name])


def available() -> list:
    """등록된 리포트 name 목록(정렬)."""
    return sorted(_REPORTS)


__all__ = ["get_report", "available"]
