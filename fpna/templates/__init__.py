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
    "cohort_retention": "fpna.templates.cohort_retention",
    "headcount_plan": "fpna.templates.headcount_plan",
    "listing": "fpna.templates.listing",          # 범용 정리표(measure 희소 데이터 착지점)
    # 고정비 FP&A (fpna-fixed-cost-tables 스킬)
    "fc_depreciation_schedule": "fpna.templates.fc_depreciation_schedule",
    "fc_variance_bridge": "fpna.templates.fc_variance_bridge",
    "fc_maturity_wall": "fpna.templates.fc_maturity_wall",
    "fc_runrate_normalized": "fpna.templates.fc_runrate_normalized",
    "fc_cuttability_ladder": "fpna.templates.fc_cuttability_ladder",
    "fc_driver_unitcost": "fpna.templates.fc_driver_unitcost",
    "fc_forward_da": "fpna.templates.fc_forward_da",
    "fc_prepaid_rollforward": "fpna.templates.fc_prepaid_rollforward",
    "fc_lease_ifrs16": "fpna.templates.fc_lease_ifrs16",
    "fc_allocation": "fpna.templates.fc_allocation",
    # C3 빠진 템플릿 6 (자문 consult-existing-assets 2b)
    "pvm_bridge": "fpna.templates.pvm_bridge",
    "debt_schedule": "fpna.templates.debt_schedule",
    "consolidation_fx": "fpna.templates.consolidation_fx",
    "working_capital": "fpna.templates.working_capital",
    "cost_allocation": "fpna.templates.cost_allocation",
    "dupont_roic": "fpna.templates.dupont_roic",
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
