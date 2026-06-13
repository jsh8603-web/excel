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
  py main.py report <name> [out.xlsx]       # 다중시트 제본(보드팩) — 크로스시트 tie QC 게이트
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
    from fpna.crypto import encrypt_file, encrypt_text, to_mail_text
    if not args:
        _print("사용: py main.py encrypt <평문파일> [out] [--pass 암구호] "
               "[--mail] [--max-lines N]"); return 2
    pw, rest = _read_passphrase(args)
    mail = "--mail" in rest
    rest = [x for x in rest if x != "--mail"]
    max_lines = 500
    if "--max-lines" in rest:
        i = rest.index("--max-lines")
        try:
            max_lines = int(rest[i + 1])
        except (IndexError, ValueError):
            pass
        rest = rest[:i] + rest[i + 2:]
    infile = rest[0]
    if not mail:
        out = rest[1] if len(rest) > 1 else infile + ".enc.txt"
        n = encrypt_file(pw, infile, out)
        _print("암호화 → %s (%d자 armored). passphrase 는 별도 채널로 공유하세요." % (out, n))
        return 0
    # --mail: 메일 본문 텍스트(들)로. 첨부 아님 — 그대로 복붙해서 보낸다.
    import os
    with open(infile, "r", encoding="utf-8", newline="") as fh:
        armored = encrypt_text(pw, fh.read())
    mid = (os.path.splitext(os.path.basename(infile))[0][:8] or "MSG")
    parts = to_mail_text(armored, max_lines=max_lines, msg_id=mid)
    base = rest[1] if len(rest) > 1 else infile + ".mail.txt"
    if len(parts) == 1:
        with open(base, "w", encoding="utf-8", newline="") as fh:
            fh.write(parts[0] + "\n")
        _print("메일 본문 1통 → %s (%d줄). 내용을 메일 본문에 붙여넣어 보내세요."
               % (base, parts[0].count(chr(10)) + 1))
    else:
        root, ext = os.path.splitext(base)
        for k, p in enumerate(parts, 1):
            with open("%s.part%d%s" % (root, k, ext), "w", encoding="utf-8", newline="") as fh:
                fh.write(p + "\n")
        _print("메일 %d통 분할 → %s.part1..%d%s (메일당 ≤%d줄). 각 part 를 별도 메일 본문에 붙여넣으세요."
               % (len(parts), root, len(parts), ext, max_lines))
        _print("받는 쪽은 메일들을 한 파일에 모두 붙여넣고 decrypt 하면 자동 정렬·복원됩니다.")
    _print("passphrase 는 별도 채널로 공유. 제목엔 [claude] 유지(암호화 안 함).")
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
    from fpna.dispatcher import route
    text = " ".join(args)
    r = route(text)          # T3: stage 분류 선행(ingest/profile/transport/analysis)
    _print("단계: %s  (근거: %s)" % (r["stage"], r.get("reason", "")))
    if r.get("template"):
        _print("템플릿: %s" % r["template"])
    _print("다음: %s" % r["next_command"])
    return 0


def cmd_render(args):
    from fpna.pipeline import run_report
    from fpna.templates import get_template
    if not args:
        _print("사용: py main.py render <type> [out.xlsx] [--csv tidy.csv]"); return 2
    type_name = args[0]
    rest = args[1:]
    csv_path = None
    if "--csv" in rest:          # 실데이터 진입(T2): tidy.csv → from_tidy 로 INPUT 바인딩
        i = rest.index("--csv")
        csv_path = rest[i + 1] if i + 1 < len(rest) else None
        rest = rest[:i] + rest[i + 2:]
    out = rest[0] if rest else "out/%s.xlsx" % type_name
    mod = get_template(type_name)
    if csv_path:
        from fpna.binding import bind_from_csv
        try:
            data = bind_from_csv(mod, csv_path)
        except NotImplementedError:
            _print("%s 는 아직 --csv 미지원(from_tidy 미선언) → 골든으로만 가능." % type_name); return 2
        except Exception as e:
            _print("실데이터 바인딩 실패: %s" % e); return 2
        _print("실데이터 바인딩: %s → %s INPUT" % (csv_path, type_name))
    else:
        data = mod.golden_sample()
    res = run_report(mod, data, out_path=out)   # 스파인 단일 통로(검증 메인 강제)
    _print(res.qc.summary())
    _print(("저장: " + res.out_path) if res.saved else "QC 미통과/우회차단 → 저장 보류")
    return 0 if res.saved else 1


def cmd_report(args):
    """다중시트 제본(B 실행경로): 레지스트리 make_spec → build_report → qc_report → 통과 시만 저장.

    사용: py main.py report <name> [out.xlsx]
    """
    from fpna.reports import get_report, available
    from fpna.report import build_report, qc_report
    if not args:
        _print("사용: py main.py report <name> [out.xlsx]")
        _print("가능: %s" % ", ".join(available())); return 2
    name = args[0]
    out = args[1] if len(args) > 1 else "out/%s.xlsx" % name
    try:
        mod = get_report(name)
    except KeyError as e:
        _print(str(e)); return 2
    spec = mod.make_spec()
    wb = build_report(spec)                        # fullCalcOnLoad 부여(build_report 내부)
    rep = qc_report(wb, spec)
    _print(rep.summary())
    if rep.passed:
        import os
        d = os.path.dirname(out)
        if d:
            os.makedirs(d, exist_ok=True)
        wb.save(out)
        _print("저장: %s  (시트 %d개)" % (out, len(wb.worksheets)))
        return 0
    _print("QC 미통과(크로스시트 tie/grain) → 저장 보류")
    return 1


def cmd_pack(args):
    """다중 exhibit 연동 팩(2-4): packs 레지스트리 make_spec → build_pack(run_report
    스파인·receipt) → 통과 시만 저장. 우회 시 저장 불가.

    사용: py main.py pack <name> [out.xlsx]
    """
    from fpna.packs import get_pack, available
    from fpna.pack import build_pack
    if not args:
        _print("사용: py main.py pack <name> [out.xlsx]")
        _print("가능: %s" % ", ".join(available())); return 2
    name = args[0]
    out = args[1] if len(args) > 1 else "out/pack_%s.xlsx" % name
    try:
        mod = get_pack(name)
    except KeyError as e:
        _print(str(e)); return 2
    import os
    d = os.path.dirname(out)
    if d:
        os.makedirs(d, exist_ok=True)
    res = build_pack(mod.make_spec(), out_path=out)
    _print(res.qc.summary())
    if res.saved:
        _print("저장: %s  (시트 %d개, receipt 발급)" % (out, len(res.wb.worksheets)))
        return 0
    _print("QC 미통과(모델체크/크로스시트 tie) → 저장 보류(스파인 우회 불가)")
    return 1


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
    "dispatch": cmd_dispatch, "report": cmd_report, "render": cmd_render,
    "pack": cmd_pack,
    "golden": cmd_golden, "selftest": cmd_selftest,
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
