"""
main.py — FP&A Excel 시스템 단일 진입점.

회사 PC: `git pull` 후 설치 없이 `py main.py <명령>` 으로 동작.
(vendor/ 동봉 openpyxl 을 _bootstrap 이 sys.path 에 주입.)

명령:
  py main.py selftest                      # 골든샘플 전부 + ingest 픽스처 회귀
  py main.py ingest <파일.xlsx> [out_dir]  # 누더기 엑셀 → tidy.csv/schema.json/smell_report.md
  py main.py profile <마트.csv> [out.yaml] # 정제 마트테이블 → 차원없는 SHAPE 스키마(회사→집 운반용)
  py main.py encrypt <평문> [out] [--pass X] # 텍스트 대칭암호화(ChaCha20-Poly1305+scrypt) → armored
  py main.py decrypt <암호문> [out] [--pass X] # 복호화(passphrase 오답·변조 시 거부)
  py main.py dispatch "<요청 텍스트>"        # 어느 템플릿을 쓸지 판정
  py main.py render <type> [out.xlsx]       # 골든샘플로 템플릿 렌더(QC 게이트)
  py main.py golden [type]                  # 골든샘플 빌드+QC (type 생략 시 전부)
  py main.py list                           # 구현된 템플릿 유형
"""
from __future__ import annotations

import sys

import fpna._bootstrap  # noqa: F401  (vendor 주입, 최우선)


def _print(s=""):
    try:
        sys.stdout.buffer.write((str(s) + "\n").encode("utf-8"))
    except Exception:
        print(s)


def cmd_list(_args):
    from fpna.templates import available, _MODULES
    ok = set(available())
    _print("템플릿 유형:")
    for name in _MODULES:
        _print("  %s %s" % ("[구현]" if name in ok else "[스텁]", name))


def cmd_ingest(args):
    from fpna.ingest import run_ingest
    if not args:
        _print("사용: py main.py ingest <파일.xlsx> [out_dir] [--sheet 이름]"); return 2
    path = args[0]
    sheet = None
    rest = args[1:]
    if "--sheet" in rest:
        i = rest.index("--sheet")
        sheet = rest[i + 1] if i + 1 < len(rest) else None
        rest = rest[:i] + rest[i + 2:]
    out_dir = rest[0] if rest else "out/ingest"
    res = run_ingest(path, out_dir, sheet=sheet)
    _print("블록 %d개, tidy %d행 (reject %d). 산출: %s/{tidy.csv,schema.json,smell_report.md}"
           % (res.n_blocks, len(res.tidy_rows), res.report.n_rejected, out_dir))
    if res.smells:
        _print("수식 스멜 %d건 → smell_report.md 확인" % len(res.smells))
    return 0


def _parse_opt_list(rest, flag):
    """--flag a b c  → (값리스트|None, flag 제거된 rest). 다음 --옵션 전까지."""
    if flag not in rest:
        return None, rest
    i = rest.index(flag)
    vals = []
    j = i + 1
    while j < len(rest) and not rest[j].startswith("--"):
        vals.append(rest[j]); j += 1
    return (vals or None), rest[:i] + rest[j:]


def cmd_profile(args):
    from fpna.profile import run_profile
    if not args:
        _print("사용: py main.py profile <마트.csv> [out.yaml] "
               "[--date-col 컬럼] [--measures a b] [--grain a b] [--with-names]")
        return 2
    path = args[0]
    rest = args[1:]
    include_names = False
    if "--with-names" in rest:
        include_names = True
        rest = [x for x in rest if x != "--with-names"]
    date_col = None
    if "--date-col" in rest:
        i = rest.index("--date-col")
        date_col = rest[i + 1] if i + 1 < len(rest) else None
        rest = rest[:i] + rest[i + 2:]
    measures, rest = _parse_opt_list(rest, "--measures")
    grain, rest = _parse_opt_list(rest, "--grain")
    out = rest[0] if rest else "out/profile_spec.yaml"
    import os
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    spec = run_profile(path, out, date_col=date_col, measures=measures,
                       grain=grain, include_member_names=include_names)
    tname = next(iter(spec["tables"]))
    ncols = len(spec["tables"][tname]["columns"])
    _print("SHAPE 프로파일 추출: 테이블 '%s', 컬럼 %d개 → %s" % (tname, ncols, out))
    _print("⚠ 값·이름 미포함(차원없는 형태만). 텍스트라 메일/화면으로 운반 가능.")
    if include_names:
        _print("⚠ --with-names: 멤버 이름 포함됨 — 민감하면 제거 후 운반.")
    return 0


def _read_passphrase(args):
    """--pass X 우선, 없으면 getpass 프롬프트. (passphrase, 남은args)."""
    if "--pass" in args:
        i = args.index("--pass")
        pw = args[i + 1] if i + 1 < len(args) else ""
        return pw, args[:i] + args[i + 2:]
    import getpass
    try:
        pw = getpass.getpass("passphrase: ")
    except Exception:
        pw = input("passphrase: ")
    return pw, args


def cmd_encrypt(args):
    from fpna.crypto import encrypt_file
    if not args:
        _print("사용: py main.py encrypt <평문파일> [out.enc.txt] [--pass 암구호]"); return 2
    pw, rest = _read_passphrase(args)
    infile = rest[0]
    out = rest[1] if len(rest) > 1 else infile + ".enc.txt"
    n = encrypt_file(pw, infile, out)
    _print("암호화 → %s (%d자 armored). passphrase 는 별도 채널로 공유하세요." % (out, n))
    return 0


def cmd_decrypt(args):
    from fpna.crypto import decrypt_file
    if not args:
        _print("사용: py main.py decrypt <암호문파일> [out] [--pass 암구호]"); return 2
    pw, rest = _read_passphrase(args)
    infile = rest[0]
    out = rest[1] if len(rest) > 1 else (infile[:-8] if infile.endswith(".enc.txt") else infile + ".dec")
    try:
        n = decrypt_file(pw, infile, out)
    except ValueError as e:
        _print("복호화 실패: %s" % e); return 1
    _print("복호화 → %s (%d자)" % (out, n))
    return 0


def cmd_dispatch(args):
    from fpna.dispatcher import dispatch
    text = " ".join(args)
    res = dispatch(text)
    _print("판정: %s  (근거: %s, score=%d)" % (res.template, res.reason, res.score))
    _print("렌더: py main.py render %s out.xlsx" % res.template)
    return 0


def cmd_render(args):
    from fpna.render import render_golden
    if not args:
        _print("사용: py main.py render <type> [out.xlsx]"); return 2
    type_name = args[0]
    out = args[1] if len(args) > 1 else "out/%s.xlsx" % type_name
    from fpna.render import render
    from fpna.templates import get_template
    data = get_template(type_name).golden_sample()
    res = render(type_name, data, out)
    _print(res.qc.summary())
    _print(("저장: " + res.out_path) if res.saved else "QC 미통과 → 저장 보류")
    return 0 if res.saved else 1


def cmd_golden(args):
    from fpna.templates import available
    from fpna.render import render_golden
    types = [args[0]] if args else available()
    rc = 0
    for t in types:
        try:
            res = render_golden(t)
            _print("%s: %s%s" % (t, "PASS" if res.qc.passed else "FAIL",
                                 "" if res.saved else " (저장보류)"))
            if not res.qc.passed:
                rc = 1
        except Exception as e:
            _print("%s: 예외 %s" % (t, e)); rc = 1
    return rc


def cmd_selftest(_args):
    rc = 0
    _print("=== 골든샘플 회귀 ===")
    rc |= cmd_golden([])
    _print("\n=== ingest 픽스처 ===")
    import os
    fx = os.path.join("tests", "fixtures", "messy_sample.xlsx")
    if os.path.isfile(fx):
        rc |= cmd_ingest([fx, "out/selftest_ingest"])
    else:
        _print("픽스처 없음(py tests/make_fixtures.py 로 생성). skip")
    _print("\n결과: %s" % ("ALL PASS" if rc == 0 else "FAIL 있음"))
    return rc


_COMMANDS = {
    "list": cmd_list, "ingest": cmd_ingest, "profile": cmd_profile,
    "encrypt": cmd_encrypt, "decrypt": cmd_decrypt,
    "dispatch": cmd_dispatch, "render": cmd_render, "golden": cmd_golden,
    "selftest": cmd_selftest,
}


def main(argv=None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    if not argv or argv[0] in ("-h", "--help", "help"):
        _print(__doc__); return 0
    cmd = argv[0]
    fn = _COMMANDS.get(cmd)
    if not fn:
        _print("알 수 없는 명령: %s" % cmd); _print(__doc__); return 2
    return fn(argv[1:]) or 0


if __name__ == "__main__":
    sys.exit(main())
