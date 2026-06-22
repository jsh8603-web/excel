"""
tests/test_runtime_purity.py — 런타임 백엔드 라우팅의 '차단 갈래'를 코드로 강제.

배경(backend routing)
---------------------
.xlsx 작성 백엔드는 4종(openpyxl / xlwings / xlsxwriter / pywin32 COM)이지만,
**런타임 트랙**(`fpna/`, `main.py`)은 openpyxl 전용이다 — 회사 PC 무설치 `-S` 재현 +
vendor/ 순수 .py 동봉이 전제이기 때문(CLAUDE.md §1, skills/freehand-excel-integrity
references/backend-routing.md '환경별 차단 규칙').

그 규칙이 그동안 docstring 주석으로만 존재했다(_bootstrap.py 는 vendor/ 컴파일 산물만
검사하고, 소스가 무엇을 import 하는지는 보지 않았다). 이 테스트는 런타임 소스 전체를
ast 로 정적 스캔해, pandas/numpy/pywin32/xlwings 등 차단 대상 import 가 단 한 줄이라도
섞이면 즉시 실패한다 — 회사 PC 에서 ImportError 로 깨지기 전에 머지 게이트에서 잡는다.

검증 도구(format-xlwings, verify_xlsx 등)는 `tools/` 에 살며 런타임 의존성이 아니므로
스캔 범위에서 제외한다.
"""
from __future__ import annotations

import ast
import os
import unittest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 런타임 트랙 = 회사 PC 에서 `py main.py` 로 실제 import 되는 코드.
# tools/ (COM·xlwings 검증 보조), tests/, archive/, vendor/ 는 런타임이 아니다.
_RUNTIME_ROOTS = [
    os.path.join(_REPO_ROOT, "fpna"),
    os.path.join(_REPO_ROOT, "main.py"),
]

# CLAUDE.md §1 + backend-routing.md: 런타임에서 import 가 허용되지 않는 서드파티.
# openpyxl, et_xmlfile 두 개만 허용(둘 다 vendor/ 동봉 순수 .py). 나머지는 stdlib.
_FORBIDDEN_TOPLEVEL = {
    "pandas",
    "numpy",
    "pydantic",
    "xlsxwriter",      # 대소문자 무시 비교
    "formulas",
    "xlwings",
    "win32com",
    "win32api",
    "win32gui",
    "pythoncom",
    "pywintypes",
}


def _iter_runtime_py():
    for root in _RUNTIME_ROOTS:
        if os.path.isfile(root) and root.endswith(".py"):
            yield root
        elif os.path.isdir(root):
            for dirpath, _dirs, files in os.walk(root):
                if "__pycache__" in dirpath:
                    continue
                for f in files:
                    if f.endswith(".py"):
                        yield os.path.join(dirpath, f)


def _imported_toplevels(path):
    """파일이 import 하는 모듈의 최상위 이름 집합(중첩 import 포함, ast 정적 스캔)."""
    with open(path, encoding="utf-8") as fh:
        tree = ast.parse(fh.read(), filename=path)
    names = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                names.add(alias.name.split(".")[0])
        elif isinstance(node, ast.ImportFrom):
            # `from X import ...` (상대 import 는 module 이 None → 건너뜀)
            if node.level == 0 and node.module:
                names.add(node.module.split(".")[0])
    return names


class RuntimePurityTest(unittest.TestCase):
    def test_no_forbidden_runtime_imports(self):
        """fpna/ + main.py 어느 모듈도 차단 대상 서드파티를 import 하지 않는다."""
        forbidden_lower = {m.lower() for m in _FORBIDDEN_TOPLEVEL}
        violations = []
        scanned = 0
        for path in _iter_runtime_py():
            scanned += 1
            for mod in _imported_toplevels(path):
                if mod.lower() in forbidden_lower:
                    rel = os.path.relpath(path, _REPO_ROOT)
                    violations.append("%s → import %s" % (rel, mod))
        self.assertGreater(scanned, 0, "런타임 .py 를 한 개도 스캔하지 못함(경로 오류?)")
        self.assertEqual(
            violations,
            [],
            "런타임 트랙(fpna/·main.py)에 차단 대상 서드파티 import 발견 "
            "— 회사 PC 무설치(-S)에서 ImportError. openpyxl/et_xmlfile + stdlib 만 허용:\n  "
            + "\n  ".join(violations),
        )

    def test_vendor_has_no_compiled_artifacts(self):
        """vendor/ 에 .pyd/.so/.dll(컴파일 산물)이 섞이면 순수 .py 동봉 전제가 깨진다."""
        vendor = os.path.join(_REPO_ROOT, "vendor")
        if not os.path.isdir(vendor):
            self.skipTest("vendor/ 없음")
        bad = []
        for dirpath, _dirs, files in os.walk(vendor):
            for f in files:
                if f.endswith((".pyd", ".so", ".dll")):
                    bad.append(os.path.relpath(os.path.join(dirpath, f), _REPO_ROOT))
        self.assertEqual(bad, [], "vendor/ 에 컴파일 산물(순수 파이썬 위반): %s" % bad)


if __name__ == "__main__":
    unittest.main()
