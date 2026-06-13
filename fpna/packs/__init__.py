"""fpna.packs — 다중 exhibit 팩 카탈로그(레지스트리).

각 팩 모듈은 make_spec() -> pack.PackSpec 를 노출한다. dispatcher 의 pack 게이트가
트리거(팩/통합/연동/타당성/이사회/중기계획)를 만나면 이 레지스트리로 라우팅한다.
packs.md = 사람용 판정·카탈로그 가이드(본 레지스트리의 설계 근거 R7).
"""
from __future__ import annotations

import importlib

import fpna._bootstrap  # noqa: F401

# (pack name, 모듈경로). make_spec() 노출 모듈만 등록.
_PACKS = {
    "feasibility": "fpna.packs.feasibility",          # 사업타당성·투자심사
}


def get_pack(name: str):
    """pack name → make_spec() 보유 모듈."""
    path = _PACKS.get(name)
    if not path:
        raise KeyError("unknown pack: %s (있음: %s)" % (name, ", ".join(sorted(_PACKS))))
    return importlib.import_module(path)


def available() -> list:
    """import 가능한(구현된) 팩 name 목록(정렬)."""
    ok = []
    for name, path in _PACKS.items():
        try:
            importlib.import_module(path)
            ok.append(name)
        except Exception:
            pass
    return sorted(ok)


__all__ = ["get_pack", "available", "_PACKS"]
