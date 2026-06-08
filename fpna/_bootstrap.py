"""
fpna._bootstrap — 회사 PC 무설치 런타임 부트스트랩.

모든 진입점(main.py, dispatcher, render, ingest)은 *맨 먼저* 이 모듈을
import 한다. 그러면 repo의 vendor/ 절대경로가 sys.path 맨 앞에 추가되어
`import openpyxl` 이 설치 없이 동봉 소스로 해소된다.

설계 원칙
---------
- 회사 PC: pip 설치 0. `git pull` 후 `py main.py` 만으로 동작.
- 런타임 외부 의존성 = openpyxl + et_xmlfile (둘 다 vendor/ 에 순수 .py 동봉).
- pandas/numpy/pydantic/XlsxWriter/formulas 등은 import 하지 않는다(금지).

사용
----
    import fpna._bootstrap  # noqa: F401  (sys.path 주입 side-effect)
    import openpyxl
"""
from __future__ import annotations

import os
import sys

# 이 파일: <repo>/fpna/_bootstrap.py  →  repo root = parent of fpna/
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(_THIS_DIR)
VENDOR_DIR = os.path.join(REPO_ROOT, "vendor")


def ensure_vendor_on_path() -> str:
    """vendor/ 절대경로를 sys.path 맨 앞에 (중복 없이) 추가하고 반환한다."""
    if not os.path.isdir(VENDOR_DIR):
        raise RuntimeError(
            "vendor/ 디렉터리를 찾을 수 없습니다: %s\n"
            "회사 PC에서는 repo를 통째로 git pull 해야 합니다 "
            "(vendor/ 가 openpyxl 동봉 소스를 담고 있음)." % VENDOR_DIR
        )
    if VENDOR_DIR in sys.path:
        sys.path.remove(VENDOR_DIR)
    sys.path.insert(0, VENDOR_DIR)
    return VENDOR_DIR


def verify_pure_python() -> None:
    """vendor/ 에 컴파일 산물(.pyd/.so/.dll)이 섞이지 않았는지 방어적으로 확인.

    순수 파이썬 보장이 깨지면(컴파일 확장 동봉) 회사 PC 아키텍처/버전에서
    ImportError 가 날 수 있으므로, 발견 즉시 경고만 출력한다(중단은 안 함).
    """
    bad = []
    for root, _dirs, files in os.walk(VENDOR_DIR):
        for f in files:
            if f.endswith((".pyd", ".so", ".dll")):
                bad.append(os.path.join(root, f))
    if bad:
        sys.stderr.write(
            "[fpna._bootstrap] 경고: vendor/ 에 컴파일 산물이 있습니다 "
            "(순수 파이썬 보장 위반):\n  " + "\n  ".join(bad) + "\n"
        )


# import side-effect: 진입점이 이 모듈을 import 하면 즉시 path 주입
ensure_vendor_on_path()


if __name__ == "__main__":
    p = ensure_vendor_on_path()
    verify_pure_python()
    print("vendor on path:", p)
    try:
        import openpyxl  # noqa: F401

        print("openpyxl:", openpyxl.__version__)
    except Exception as e:  # pragma: no cover
        print("openpyxl import 실패:", e)
        sys.exit(1)
