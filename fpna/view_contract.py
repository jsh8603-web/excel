"""
fpna.view_contract — View Contract v2 (R1~R11 불변식).

폐쇄망 Claude 가 다중 테이블을 만날 때 데이터 엔지니어링 기본값(sparse 압축·하루치만·
스냅샷만·anti-join only)으로 빠지는 것을 **구조적으로 차단**한다. 주의 문장이 아니라
코드로 박는다 — 계약(SKILL.md)·헬퍼(이 파일)·QC 게이트(템플릿 qc())·테스트 4중 방어.

어휘는 dbt-utils(equal_rowcount/unique_combination/relationships)·Great Expectations
(expect_*)·IAASB(completeness/accuracy/cutoff) 에서 **차용만** 한다(미설치, stdlib 재구현).

모든 assert 는 `(rep: QCReport, ...)` 시그니처로 QCReport.add 에 누적한다(throw 대신).
render 게이트가 rep.passed 로 저장 여부를 결정하므로 base.py 와 정합한다.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from fpna.templates.base import QCReport
from fpna.dims import Fact, AccountingCalendar, account_leaves, rollup


# --------------------------------------------------------------------------- #
# 공통 헬퍼                                                                   #
# --------------------------------------------------------------------------- #
def _distinct(seq) -> list:
    """순서 보존 유니크."""
    seen: set = set()
    out = []
    for x in seq:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def _as_keys(on) -> tuple:
    return tuple(on) if isinstance(on, (list, tuple)) else (on,)


# --------------------------------------------------------------------------- #
# R1 — 시간축 전수성                                                          #
# --------------------------------------------------------------------------- #
def assert_time_ruler(rep: QCReport, fact: Fact, calendar: AccountingCalendar,
                      start: tuple[int, int], end: tuple[int, int], *,
                      period_key: str = "period",
                      name: str = "R1 time_ruler") -> bool:
    """표시 시간축이 캘린더 연속 ruler 의 모든 기간을 포함하는지 검증.

    결측 기간을 건너뛰면(행 누락) 실패. 결측은 NO_DATA 플래그로 행을 유지해야 한다.
    """
    ruler = calendar.ruler(start, end)
    present = {r.get(period_key) for r in fact.rows}
    missing = [lbl for lbl in ruler if lbl not in present]
    ok = not missing
    rep.add(name, ok, "" if ok else
            "결측 기간 행 누락(NO_DATA 미표기): " + ", ".join(missing[:6]))
    return ok


# --------------------------------------------------------------------------- #
# R2 — 모집단 전수성 (FULL OUTER)                                             #
# --------------------------------------------------------------------------- #
def full_outer(left: list[dict], right: list[dict], on) -> list[dict]:
    """FULL OUTER join. 행수 = |좌∪우|, 각 행에 match_status 부여.

    match_status ∈ {MAPPED, LEFT_ONLY, RIGHT_ONLY}. anti-join(미매칭만)은 진단용일 뿐
    전수 결과를 대체할 수 없다(R2).
    """
    keys = _as_keys(on)

    def kf(r: dict) -> tuple:
        return tuple(r.get(k) for k in keys)

    lidx = {kf(r): r for r in left}
    ridx = {kf(r): r for r in right}
    order = list(lidx.keys()) + [k for k in ridx if k not in lidx]
    out: list[dict] = []
    for k in order:
        lrow, rrow = lidx.get(k), ridx.get(k)
        status = "MAPPED" if (lrow and rrow) else ("LEFT_ONLY" if lrow else "RIGHT_ONLY")
        merged: dict = {}
        if lrow:
            merged.update(lrow)
        if rrow:
            for kk, vv in rrow.items():
                merged.setdefault(kk, vv)
        merged["match_status"] = status
        out.append(merged)
    return out


def assert_full_population(rep: QCReport, merged: list[dict],
                           left_keys, right_keys, *,
                           name: str = "R2 full_population") -> bool:
    """병합 결과 행수 == |좌∪우| 이고 모든 행에 match_status 가 있는지."""
    union = set(left_keys) | set(right_keys)
    has_status = all("match_status" in m for m in merged)
    ok = (len(merged) == len(union)) and has_status
    rep.add(name, ok, "" if ok else
            "행수=%d 합집합=%d status전부=%s" % (len(merged), len(union), has_status))
    return ok


# --------------------------------------------------------------------------- #
# R3 — 대사 블록 (_RECON) + tie-out                                          #
# --------------------------------------------------------------------------- #
def recon_block(*, n_input: int, n_output: int, src_sum: float, out_sum: float,
                excluded: dict | None = None,
                completeness: str = "", accuracy: str = "",
                cutoff: str = "") -> list[tuple]:
    """시트에 박을 _RECON 블록 (label, value) 행 리스트.

    머리에 IAASB completeness/accuracy/cutoff 3행을 둬 회계 유관부서가 즉시 읽게 한다.
    """
    rows: list[tuple] = [
        ("_RECON", ""),
        ("completeness", completeness),
        ("accuracy", accuracy),
        ("cutoff", cutoff),
        ("입력 행수", n_input),
        ("출력 행수", n_output),
    ]
    for reason, cnt in (excluded or {}).items():
        rows.append(("제외(%s)" % reason, cnt))
    rows += [
        ("원천 합계", src_sum),
        ("출력 합계", out_sum),
        ("합계 차이", src_sum - out_sum),
    ]
    return rows


def assert_tie_out(rep: QCReport, src_sum: float, out_sum: float, *,
                   tol: float = 0.0, name: str = "R3 tie_out") -> bool:
    """원천 합계 vs 출력 합계 차이 = 0 (기본 tol=0)."""
    if src_sum is None or out_sum is None:
        rep.add(name, False, "합계 None")
        return False
    diff = abs(src_sum - out_sum)
    ok = diff <= tol
    rep.add(name, ok, "" if ok else "합계차=%.6g (tol=%.6g)" % (diff, tol))
    return ok


# --------------------------------------------------------------------------- #
# R4 — 필터 선언 의무                                                         #
# --------------------------------------------------------------------------- #
def assert_filter_declared(rep: QCReport, filters: list[dict] | None, *,
                           name: str = "R4 filter_declared") -> bool:
    """필터는 declared + header_text + excluded_count 셋 다 있을 때만 적법."""
    bad = []
    for f in filters or []:
        if (not f.get("declared")) or (not f.get("header_text")) \
                or (f.get("excluded_count") is None):
            bad.append(f.get("field", "?"))
    ok = not bad
    rep.add(name, ok, "" if ok else "미선언 필터: " + ", ".join(map(str, bad)))
    return ok


# --------------------------------------------------------------------------- #
# R5 — 내부 tidy / 표시 wide 분리 (전수 cross-tab)                            #
# --------------------------------------------------------------------------- #
def cross_tab(fact: Fact, index_keys, column_key: str, value_key: str, *,
              all_columns: list | None = None, all_index: list | None = None,
              no_data: str = "NO_DATA") -> dict:
    """tidy fact → wide 전수 매트릭스. 결측 셀은 no_data(행·열 생략 금지, R5).

    all_columns/all_index 를 주면 (데이터에 없어도) 그 전수 축으로 전개한다.
    반환: {"columns": [...], "rows": [dict, ...]}.
    """
    idx_keys = _as_keys(index_keys)
    cols = list(all_columns) if all_columns is not None \
        else _distinct([r.get(column_key) for r in fact.rows])
    index_vals = list(all_index) if all_index is not None \
        else _distinct([tuple(r.get(k) for k in idx_keys) for r in fact.rows])
    cell: dict = {}
    for r in fact.rows:
        iv = tuple(r.get(k) for k in idx_keys)
        cell[(iv, r.get(column_key))] = r.get(value_key)
    out_rows: list[dict] = []
    for iv in index_vals:
        row = dict(zip(idx_keys, iv))
        for c in cols:
            row[c] = cell.get((iv, c), no_data)
        out_rows.append(row)
    return {"columns": list(idx_keys) + cols, "rows": out_rows}


# --------------------------------------------------------------------------- #
# R6 — 금지 휴리스틱 (Default-Deny)                                           #
# --------------------------------------------------------------------------- #
# 사용자 미요청 시 코드에 등장하면 결함. 토큰은 분해 표기해(이 파일 자체 오탐 방지).
FORBIDDEN_TOKENS: tuple[str, ...] = (
    ".sa" "mple(", ".he" "ad(", "to" "p_n", "nla" "rgest(", "nsm" "allest(",
    "pre" "view(", "rand" "om.sample",
)


def assert_no_forbidden_heuristic(rep: QCReport, source_text: str, *,
                                  user_requested: bool = False,
                                  name: str = "R6 no_heuristic") -> bool:
    """소스 텍스트에 샘플링/head/top-N 휴리스틱 토큰이 있으면 결함(미요청 시)."""
    if user_requested:
        rep.add(name, True, "사용자 명시 요청 — 허용")
        return True
    found = [t for t in FORBIDDEN_TOKENS if t in (source_text or "")]
    ok = not found
    rep.add(name, ok, "" if ok else "금지 휴리스틱 토큰: " + ", ".join(found))
    return ok


# --------------------------------------------------------------------------- #
# R8 — Grain 선언/정합                                                        #
# --------------------------------------------------------------------------- #
def assert_grain(rep: QCReport, fact: Fact, *, name: str = "R8 grain") -> bool:
    """행수 == distinct grain 조합 (중복 0) + grain key 누락(None) 0."""
    dup = fact.has_duplicate_grain()
    missing = [i for i, kt in enumerate(fact.key_tuples())
               if any(v is None for v in kt)]
    ok = (not dup) and (not missing)
    detail = []
    if dup:
        detail.append("중복 grain 존재(silent merge 위험)")
    if missing:
        detail.append("grain key 누락 행 %d개" % len(missing))
    rep.add(name, ok, "; ".join(detail))
    return ok


# --------------------------------------------------------------------------- #
# R7 — QC 게이트(Coverage) 묶음                                               #
# --------------------------------------------------------------------------- #
def assert_no_silent_drop(rep: QCReport, fact: Fact, *,
                          expected_n: int | None = None,
                          name: str = "R7 no_silent_drop") -> bool:
    if expected_n is None:
        ok = not fact.has_duplicate_grain()
        detail = "" if ok else "중복 grain(잠재 silent merge)"
    else:
        ok = len(fact.rows) == expected_n
        detail = "" if ok else "행수=%d 기대=%d" % (len(fact.rows), expected_n)
    rep.add(name, ok, detail)
    return ok


def gate(rep: QCReport, fact: Fact, calendar: AccountingCalendar,
         start: tuple[int, int], end: tuple[int, int], *,
         src_sum: float | None = None, out_sum: float | None = None,
         period_key: str = "period", tol: float = 0.0,
         expected_n: int | None = None) -> bool:
    """R7 묶음 게이트 = grain + time_ruler + tie_out + no_silent_drop."""
    assert_grain(rep, fact)
    assert_time_ruler(rep, fact, calendar, start, end, period_key=period_key)
    if src_sum is not None and out_sum is not None:
        assert_tie_out(rep, src_sum, out_sum, tol=tol)
    assert_no_silent_drop(rep, fact, expected_n=expected_n)
    return rep.passed


# --------------------------------------------------------------------------- #
# R9 — 시나리오 완전성                                                        #
# --------------------------------------------------------------------------- #
def assert_scenario_aligned(rep: QCReport, actual_keys, budget_keys, *,
                            name: str = "R9 scenario_aligned") -> bool:
    """Actual/Budget 가 같은 grain·같은 모집단인지(시나리오 제외 키 기준).

    한쪽에만 있는 키는 0 으로 버리지 말고 LEFT_ONLY/RIGHT_ONLY 행으로 노출해야 한다.
    """
    a, b = set(actual_keys), set(budget_keys)
    left_only, right_only = a - b, b - a
    ok = (not left_only) and (not right_only)
    rep.add(name, ok, "" if ok else
            "Actual-only %d, Budget-only %d (0 처리 금지 — LEFT/RIGHT_ONLY 행 노출)"
            % (len(left_only), len(right_only)))
    return ok


# --------------------------------------------------------------------------- #
# R10 — 계층 정합성                                                           #
# --------------------------------------------------------------------------- #
def assert_hierarchy_ties(rep: QCReport, accounts, leaf_values: dict, *,
                          tol: float = 0.0,
                          name: str = "R10 hierarchy_ties") -> bool:
    """rollup 합(최상위) == leaf 합. orphan(부모 없음) 진단."""
    roll = rollup(accounts, leaf_values)
    leaf_sum = sum(float(leaf_values.get(a.code, 0.0)) for a in account_leaves(accounts))
    root_sum = sum(roll[a.code] for a in accounts if a.parent is None)
    diff = abs(leaf_sum - root_sum)
    codes = {a.code for a in accounts}
    orphan = [a.code for a in accounts if a.parent and a.parent not in codes]
    ok = (diff <= tol) and (not orphan)
    detail = []
    if diff > tol:
        detail.append("leaf합=%.6g root합=%.6g 차=%.6g" % (leaf_sum, root_sum, diff))
    if orphan:
        detail.append("orphan(부모없음): " + ", ".join(orphan[:6]))
    rep.add(name, ok, "; ".join(detail))
    return ok


# --------------------------------------------------------------------------- #
# R11 — 고정비 대사                                                           #
# --------------------------------------------------------------------------- #
def assert_master_to_gl(rep: QCReport, master_sum: float, gl_sum: float, *,
                        tol: float = 0.0, reasons: list | None = None,
                        name: str = "R11 master_to_gl") -> bool:
    """계약·자산 마스터 합 ↔ GL 합. 차이는 0 또는 사유별 명세가 있어야 한다."""
    diff = abs(master_sum - gl_sum)
    ok = (diff <= tol) or bool(reasons)
    if diff <= tol:
        detail = ""
    elif reasons:
        detail = "차=%.6g (사유명세 %d건)" % (diff, len(reasons))
    else:
        detail = "차=%.6g (사유 미명세 — 일회성/조정 노출 필요)" % diff
    rep.add(name, ok, detail)
    return ok


def assert_allocation_conserves(rep: QCReport, pre_sum: float, post_sum: float, *,
                                tol: float = 0.0,
                                name: str = "R11 allocation_conserves") -> bool:
    """공통비 배부 전·후 합계 tie-out (누수 금지)."""
    diff = abs(pre_sum - post_sum)
    ok = diff <= tol
    rep.add(name, ok, "" if ok else
            "배부전=%.6g 배부후=%.6g 차=%.6g (누수)" % (pre_sum, post_sum, diff))
    return ok


def assert_commitment_conserved(rep: QCReport, recognized_cum: float, remaining: float,
                                total: float, *, cancelled: float = 0.0,
                                tol: float = 0.0,
                                name: str = "R14 commitment_conserved") -> bool:
    """R14 약정보존: Σ인식누계 + Σ잔여약정 + 취소 == 계약총액. 과대인식 차단.

    레퍼런스(차용): 정부회계 encumbrance 보존(총=소진+미소진+취소) + flow conservation
    (유입=유출+잔류). 인식누계 > 계약총액 = 리스총액보다 많은 비용화 = 사고 → 차단.
    """
    lhs = recognized_cum + remaining + cancelled
    diff = abs(lhs - total)
    over = recognized_cum > total + tol
    ok = (diff <= tol) and (not over)
    detail = []
    if diff > tol:
        detail.append("보존 불일치: 인식+잔여+취소=%.6g 총액=%.6g" % (lhs, total))
    if over:
        detail.append("과대인식: 인식누계 %.6g > 총액 %.6g" % (recognized_cum, total))
    rep.add(name, ok, "; ".join(detail))
    return ok


# --------------------------------------------------------------------------- #
# Anomaly ledger 2층 (자문 R3) — render gate 재정의                           #
#   1층 emit: anomaly = 발견. 저장 막지 않고 ledger+flag 로 노출.             #
#   2층 verify: |ledger| == surfaced flags. 은폐 시에만 passed=False.         #
#   passed=False 의 의미 = "산출물 부정직(은폐/정합/tie-out 위반)"이지         #
#   "데이터에 문제 있음"이 아니다.                                            #
# --------------------------------------------------------------------------- #
ANOMALY_TYPES: tuple[str, ...] = (
    "MISSING_ACCRUAL",    # R12 재발성 — 활성 window 내 결측(미계상)
    "REVERSAL",           # R13 — 부호 flip(환입/재분류)
    "MISLABELED_FIXED",   # R13 — 고정비로 잘못 분류된 변동비
    "SUSPECTED_MISSING",  # Contract 도입 전 provisional(저신뢰)
    "RATIO_NA",           # R17 — 결측/0분모 나눗셈
)


@dataclass
class AnomalyLedger:
    """발견(anomaly)의 1층 emit 대장. grain/period/type/magnitude."""
    rows: list = field(default_factory=list)

    def add(self, grain, period, anomaly_type: str, *,
            magnitude: float = 0.0, detail: str = "") -> None:
        if anomaly_type not in ANOMALY_TYPES:
            raise ValueError("알 수 없는 anomaly_type: %r" % anomaly_type)
        self.rows.append({"grain": grain, "period": period,
                          "anomaly_type": anomaly_type,
                          "magnitude": magnitude, "detail": detail})

    def __len__(self) -> int:
        return len(self.rows)

    def by_type(self, t: str) -> list[dict]:
        return [r for r in self.rows if r["anomaly_type"] == t]


def assert_anomaly_conserved(rep: QCReport, ledger: AnomalyLedger,
                             surfaced_flags, *,
                             name: str = "anomaly_conserved") -> bool:
    """2층 검증: 발견된 anomaly(ledger)가 산출물에 누락 없이 노출됐는지.

    surfaced_flags = 산출물에 실제로 표기된 flag 수(int) 또는 그 컬렉션.
    |ledger| != |surfaced| 이면 anomaly 은폐 → 산출물 부정직 → passed=False.
    anomaly 의 *존재*는 저장을 막지 않는다(emit). *은폐*만 막는다.
    """
    n_surf = surfaced_flags if isinstance(surfaced_flags, int) else len(surfaced_flags)
    n_led = len(ledger)
    ok = n_led == n_surf
    rep.add(name, ok, "" if ok else
            "ledger=%d surfaced=%d (anomaly 은폐 — 산출물 부정직)" % (n_led, n_surf))
    return ok


def assert_recurrence(rep: QCReport, fact: Fact, expected_grid, ledger: AnomalyLedger, *,
                      period_key: str = "period", key_keys=None,
                      name: str = "R12 recurrence") -> list:
    """R12 재발성: 활성 window 의 (key, period) 가 모두 행으로 존재하는지.

    expected_grid = caller(Contract/Asset window)가 만든 {(key_tuple, period_label)} 집합.
    결측이면 MISSING_ACCRUAL 을 ledger 에 emit(저장 막지 않음). 반환 = 결측 목록.
    이 함수는 *발견*이라 rep.passed 를 깎지 않는다 — 은폐 여부는 assert_anomaly_conserved.
    """
    kk = list(key_keys) if key_keys is not None \
        else [k for k in fact.grain_keys if k != period_key]
    present = {(tuple(r.get(k) for k in kk), r.get(period_key)) for r in fact.rows}
    missing = [pg for pg in expected_grid if pg not in present]
    for keyt, per in missing:
        ledger.add(grain=keyt, period=per, anomaly_type="MISSING_ACCRUAL")
    rep.add(name, True,
            "MISSING_ACCRUAL %d건 emit" % len(missing) if missing else "활성 window 전수 계상")
    return missing


def assert_sign_step(rep: QCReport, fact: Fact, ledger: AnomalyLedger, *,
                     value_key: str = "value", key_keys=None, period_key: str = "period",
                     reversal_ratio: float = 0.5, volume_key: str | None = None,
                     elasticity_threshold: float = 0.3,
                     name: str = "R13 sign_step") -> int:
    """R13 부호·계단: 시계열 부호 flip(REVERSAL) + (volume 주면) MISLABELED_FIXED.

    PELT piecewise-constant 의미론만 차용(코드 미반입) — 그룹별 기간순 정렬 후 직전
    trailing run-rate 대비 반대부호 & |크기| ≥ ratio·|run-rate| 이면 REVERSAL emit
    (소액 true-up 제외). volume_key 주면 Δcost%/Δvolume% > 임계 동조 시 MISLABELED_FIXED.
    발견(anomaly)이라 rep.passed 를 깎지 않는다 — 은폐는 assert_anomaly_conserved.
    """
    kk = list(key_keys) if key_keys is not None \
        else [k for k in fact.grain_keys if k != period_key]
    groups: dict = {}
    for r in fact.rows:
        key = tuple(r.get(k) for k in kk)
        groups.setdefault(key, []).append(r)

    n_rev = n_mis = 0
    for key, rows in groups.items():
        rows = sorted(rows, key=lambda r: r.get(period_key))
        # REVERSAL: trailing 평균 대비 부호 반전 + 유의 크기
        for i in range(1, len(rows)):
            prev = [r.get(value_key) for r in rows[:i] if r.get(value_key) is not None]
            cur = rows[i].get(value_key)
            if not prev or cur is None:
                continue
            run = sum(prev) / len(prev)
            if run == 0:
                continue
            if (cur * run < 0) and abs(cur) >= reversal_ratio * abs(run):
                ledger.add(grain=key, period=rows[i].get(period_key),
                           anomaly_type="REVERSAL", magnitude=cur)
                n_rev += 1
        # MISLABELED_FIXED: cost~volume 동조(elasticity) — volume 있을 때만
        if volume_key is not None and len(rows) >= 3:
            elas = _elasticity([r.get(value_key) for r in rows],
                               [r.get(volume_key) for r in rows])
            if elas is not None and elas > elasticity_threshold:
                ledger.add(grain=key, period=rows[-1].get(period_key),
                           anomaly_type="MISLABELED_FIXED", magnitude=elas)
                n_mis += 1
    rep.add(name, True,
            ("REVERSAL %d, MISLABELED %d emit" % (n_rev, n_mis))
            if (n_rev or n_mis) else "부호·분류 안정")
    return n_rev + n_mis


def _elasticity(costs: list, volumes: list) -> float | None:
    """평균 대비 변화율 회귀 기울기(Δcost% / Δvolume%) 근사. stdlib only."""
    pairs = [(c, v) for c, v in zip(costs, volumes) if c is not None and v is not None]
    if len(pairs) < 3:
        return None
    cs = [c for c, _ in pairs]
    vs = [v for _, v in pairs]
    cbar = sum(cs) / len(cs)
    vbar = sum(vs) / len(vs)
    if cbar == 0 or vbar == 0:
        return None
    # 정규화 공분산/분산 (volume% → cost% 민감도)
    sxy = sum((v - vbar) * (c - cbar) for c, v in pairs)
    sxx = sum((v - vbar) ** 2 for _, v in pairs)
    if sxx == 0:
        return None
    slope = sxy / sxx                      # Δcost / Δvolume
    return slope * (vbar / cbar)           # 탄력성으로 정규화


__all__ = [
    "assert_time_ruler",                                  # R1
    "full_outer", "assert_full_population",               # R2
    "recon_block", "assert_tie_out",                      # R3
    "assert_filter_declared",                             # R4
    "cross_tab",                                          # R5
    "FORBIDDEN_TOKENS", "assert_no_forbidden_heuristic",  # R6
    "gate", "assert_no_silent_drop",                      # R7
    "assert_grain",                                       # R8
    "assert_scenario_aligned",                            # R9
    "assert_hierarchy_ties",                              # R10
    "assert_master_to_gl", "assert_allocation_conserves",  # R11
    "ANOMALY_TYPES", "AnomalyLedger",                     # anomaly 2층
    "assert_anomaly_conserved", "assert_recurrence",       # R12
    "assert_sign_step",                                    # R13
]
