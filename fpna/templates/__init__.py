"""fpna.templates — 템플릿 유형 레지스트리.

각 유형 모듈은 build/qc/golden_sample + INPUT dataclass 를 노출한다.
REGISTRY[type] = 모듈. 디스패처/렌더러가 이걸로 라우팅한다.
"""
from __future__ import annotations

import importlib

# (type 이름, 모듈경로). 구현 완료분만 등록.
_MODULES = {
    "variance": "fpna.templates.variance",
    "investment_appraisal": "fpna.templates.investment_appraisal",
    "period_trend": "fpna.templates.period_trend",
    "rolling_forecast": "fpna.templates.rolling_forecast",
    "budget_build": "fpna.templates.budget_build",
    "cashflow_13w": "fpna.templates.cashflow_13w",
    "unit_economics": "fpna.templates.unit_economics",
    "scenario_sensitivity": "fpna.templates.scenario_sensitivity",
    "pnl_3statement": "fpna.templates.pnl_3statement",
    "board_kpi_pack": "fpna.templates.board_kpi_pack",
    # 고정비 FP&A (fpna-fixed-cost-tables 스킬)
    "fc_depreciation_schedule": "fpna.templates.fc_depreciation_schedule",
    "fc_variance_bridge": "fpna.templates.fc_variance_bridge",
}


def get_template(type_name: str):
    """유형 이름 → 모듈(build/qc/golden_sample 보유)."""
    path = _MODULES.get(type_name)
    if not path:
        raise KeyError("unknown template type: %s (있음: %s)"
                       % (type_name, ", ".join(sorted(_MODULES))))
    return importlib.import_module(path)


def available() -> list[str]:
    """실제 import 가능한(구현된) 유형만."""
    ok = []
    for name, path in _MODULES.items():
        try:
            importlib.import_module(path)
            ok.append(name)
        except Exception:
            pass
    return ok


__all__ = ["get_template", "available", "_MODULES"]
