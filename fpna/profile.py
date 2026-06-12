"""
fpna.profile — 정제 마트테이블 → 차원없는 SHAPE 스키마(yaml) 추출.

⚠ ingest(누더기→tidy)와 **다른 단계**다. 이미 정제된 마트테이블(wide fact:
차원 컬럼들 + measure 컬럼들)에만 돌린다. 누더기 엑셀엔 ingest 를 먼저 쓴다.

용도: 회사→집 통신(CLAUDE.md §6). 실데이터를 '읽되', 출력으로 나가는 것은
차원 없는 형태뿐이다 — 카디널리티 / 정규화 시즌성(평균=1) / 추세율 / 상대노이즈 /
null·음수·0 률 / measure 간 상관 / AR(1) 자기상관 / 멤버 빈도(이름 제외 비율).
절대 금액·합계·평균·실제 값·멤버 이름은 출력에 일절 없다. base 는 자리표시자.

집에서 tools/synthschema.py 가 이 yaml 을 소비해 더미를 생성한다.

순수 파이썬(stdlib only). 결정적. pandas/numpy/pyyaml 미사용 — yaml 은 미니 emitter.
"""
from __future__ import annotations

import csv
import datetime as _dt
import math
import os
import re
from collections import Counter, OrderedDict, defaultdict

import fpna._bootstrap  # noqa: F401

# ----------------------------------------------------------------------------
# CSV 로딩 (stdlib)
# ----------------------------------------------------------------------------
def _read_csv(path: str) -> tuple[list[str], list[list[str]]]:
    """utf-8-sig/utf-8 자동, (header, rows) 반환."""
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    for enc in ("utf-8-sig", "utf-8", "cp949"):
        try:
            with open(path, "r", encoding=enc, newline="") as fh:
                rdr = csv.reader(fh)
                rows = list(rdr)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise UnicodeError("CSV 인코딩 해석 실패: %s" % path)
    if not rows:
        return [], []
    header = [h.strip() for h in rows[0]]
    body = rows[1:]
    return header, body


def _column(body: list[list[str]], idx: int) -> list[str]:
    out = []
    for r in body:
        out.append(r[idx] if idx < len(r) else "")
    return out


# ----------------------------------------------------------------------------
# 셀 파서
# ----------------------------------------------------------------------------
_DATE_RES = [
    re.compile(r"^(\d{4})[-/.](\d{1,2})([-/.]\d{1,2})?$"),
    re.compile(r"^(\d{4})년\s*(\d{1,2})월"),
    re.compile(r"^(\d{4})\s*[Qq분기]\s*(\d)"),
]


def _to_float(s: str):
    """숫자로 해석되면 float, 아니면 None. 콤마/괄호음수/% 처리."""
    if s is None:
        return None
    t = str(s).strip()
    if t == "":
        return None
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg, t = True, t[1:-1].strip()
    if t[:1] in ("△", "▲", "−", "-"):
        if t[:1] in ("△", "▲", "−"):
            neg = True
        t = t[1:].strip()
    pct = t.endswith("%")
    t = t.rstrip("%").replace(",", "").strip()
    try:
        v = float(t)
    except (ValueError, AttributeError):
        return None
    if neg:
        v = -v
    return v / 100.0 if pct else v


def _to_year_month(s: str):
    """(year, month) 또는 None. 정제 마트 가정(ISO 우선)."""
    if s is None:
        return None
    if isinstance(s, (_dt.datetime, _dt.date)):
        return s.year, s.month
    t = str(s).strip()
    if t == "":
        return None
    for i, rx in enumerate(_DATE_RES):
        m = rx.match(t)
        if m:
            y = int(m.group(1))
            if i == 2:  # 분기 → 대표월
                q = int(m.group(2))
                return y, min(12, max(1, (q - 1) * 3 + 1))
            mo = int(m.group(2))
            return y, min(12, max(1, mo))
    return None


def _date_hit_rate(values: list[str]) -> float:
    nonblank = [v for v in values if str(v).strip() != ""]
    if not nonblank:
        return 0.0
    hit = sum(1 for v in nonblank if _to_year_month(v) is not None)
    return hit / len(nonblank)


def _float_hit_rate(values: list[str]) -> float:
    nonblank = [v for v in values if str(v).strip() != ""]
    if not nonblank:
        return 0.0
    hit = sum(1 for v in nonblank if _to_float(v) is not None)
    return hit / len(nonblank)


# ----------------------------------------------------------------------------
# 컬럼 분류
# ----------------------------------------------------------------------------
def classify_column(values: list[str]) -> str:
    """date / measure / choice / key 중 하나."""
    if _date_hit_rate(values) >= 0.8:
        return "date"
    nonblank = [v for v in values if str(v).strip() != ""]
    if not nonblank:
        return "choice"
    nuniq = len(set(nonblank))
    if _float_hit_rate(values) >= 0.8:
        vals = [f for f in (_to_float(v) for v in nonblank) if f is not None]
        all_int = bool(vals) and all(float(f).is_integer() for f in vals)
        # 서러게이트 키: 정수 + 전부 고유 + 음수없음 + 값이 1..N 처럼 빽빽(범위≈개수).
        # 측정값(금액)은 범위가 개수보다 훨씬 넓어 여기 안 걸리고 measure 로 간다.
        if all_int and nuniq == len(nonblank) and min(vals) >= 0:
            if (max(vals) - min(vals) + 1) <= 1.5 * nuniq:
                return "key"
        return "measure"
    # 비숫자
    if nuniq > 50:
        return "key"
    return "choice"


# ----------------------------------------------------------------------------
# 통계 (stdlib 닫힌형)
# ----------------------------------------------------------------------------
def _mean(xs: list[float]) -> float:
    return sum(xs) / len(xs) if xs else 0.0


def _std(xs: list[float]) -> float:
    n = len(xs)
    if n < 2:
        return 0.0
    m = _mean(xs)
    return math.sqrt(sum((x - m) ** 2 for x in xs) / (n - 1))


def _pearson(xs: list[float], ys: list[float]) -> float | None:
    """짝지은 결측 제거 후 Pearson r. n<3 또는 분산0 이면 None."""
    pairs = [(x, y) for x, y in zip(xs, ys) if x is not None and y is not None]
    if len(pairs) < 3:
        return None
    xa = [p[0] for p in pairs]
    ya = [p[1] for p in pairs]
    mx, my = _mean(xa), _mean(ya)
    sxy = sum((x - mx) * (y - my) for x, y in pairs)
    sxx = sum((x - mx) ** 2 for x in xa)
    syy = sum((y - my) ** 2 for y in ya)
    if sxx <= 0 or syy <= 0:
        return None
    return sxy / math.sqrt(sxx * syy)


def _loglinear_trend(period_idx: list[int], period_mean: list[float]) -> float:
    """월별 평균에 log-선형 최소제곱 → 기간당 성장률(=exp(slope)-1).

    모든 평균 > 0 이고 점 3개 이상일 때만. 아니면 0.0.
    """
    pts = [(p, v) for p, v in zip(period_idx, period_mean) if v is not None and v > 0]
    if len(pts) < 3:
        return 0.0
    xs = [p for p, _ in pts]
    ys = [math.log(v) for _, v in pts]
    n = len(xs)
    sx, sy = sum(xs), sum(ys)
    sxx = sum(x * x for x in xs)
    sxy = sum(x * y for x, y in zip(xs, ys))
    denom = n * sxx - sx * sx
    if denom == 0:
        return 0.0
    slope = (n * sxy - sx * sy) / denom
    return math.exp(slope) - 1.0


def _linear_detrend(seq: list[float]) -> list[float]:
    """1차 추세(선형 드리프트) 제거. AR(1) 추정 전 잔차 정상화용."""
    n = len(seq)
    if n < 2:
        return seq
    xs = list(range(n))
    mx, my = _mean(xs), _mean(seq)
    sxx = sum((x - mx) ** 2 for x in xs)
    if sxx == 0:
        return seq
    slope = sum((x - mx) * (y - my) for x, y in zip(xs, seq)) / sxx
    inter = my - slope * mx
    return [y - (inter + slope * x) for x, y in zip(xs, seq)]


def _ar1(resid_series: list[float]) -> float | None:
    """잔차 시퀀스의 lag-1 자기상관. n<4 면 None."""
    s = [r for r in resid_series if r is not None]
    if len(s) < 4:
        return None
    a = s[:-1]
    b = s[1:]
    return _pearson(a, b)


# ----------------------------------------------------------------------------
# measure SHAPE
# ----------------------------------------------------------------------------
def _measure_shape(values: list[str], ym_pairs: list, scale: str) -> dict:
    """measure 컬럼 → 차원없는 형태. 절대값 미유출.

    values  : 원본 셀 문자열
    ym_pairs: 행별 (year, month) 또는 None (date 컬럼에서 추출, 길이 동일)
    """
    fvals = [_to_float(v) for v in values]
    n_total = len(fvals)
    n_null = sum(1 for f in fvals if f is None)
    nonnull = [f for f in fvals if f is not None]
    n_nn = len(nonnull)

    out: "OrderedDict[str, object]" = OrderedDict()
    out["type"] = "measure"
    out["base"] = scale
    out["round"] = 0
    out["null_rate"] = round(n_null / n_total, 3) if n_total else 0.0
    out["neg_rate"] = round(sum(1 for f in nonnull if f < 0) / n_nn, 3) if n_nn else 0.0
    out["zero_rate"] = round(sum(1 for f in nonnull if f == 0) / n_nn, 3) if n_nn else 0.0

    has_date = ym_pairs is not None and any(p is not None for p in ym_pairs)
    if has_date and n_nn:
        # (abs_period, month, value) — abs = year*12+month (단일 month/period)
        valid = [(ym[0] * 12 + ym[1], ym[1], f)
                 for ym, f in zip(ym_pairs, fvals) if ym is not None and f is not None]
        if len(valid) >= 3:
            pmin = min(a for a, _, _ in valid)
            by_pi: "defaultdict[int, list]" = defaultdict(list)
            for a, mo, f in valid:
                by_pi[a - pmin].append(f)
            pis_sorted = sorted(by_pi)

            def _detrend_seasonality(tr):
                """추세 tr 로 디트렌드 후 월별 시즌(평균=1) + 디트렌드 전체평균."""
                g = 1.0 + tr
                dtt = [(a, mo, (f / (g ** (a - pmin)) if g > 0 else f)) for a, mo, f in valid]
                bm: "defaultdict[int, list]" = defaultdict(list)
                for a, mo, f in dtt:
                    bm[mo].append(f)
                gm = _mean([f for _, _, f in dtt])
                s = [(_mean(bm[mo]) / gm if (gm and mo in bm) else 1.0) for mo in range(1, 13)]
                sm = _mean(s) or 1.0
                return [x / sm for x in s], gm, dtt

            # pass1: 거친 추세 → 임시 시즌
            trend = _loglinear_trend(pis_sorted, [_mean(by_pi[p]) for p in pis_sorted])
            seas_tmp, _, _ = _detrend_seasonality(trend)
            # pass2: 디시즌 후 추세 재추정(시즌 누설 제거)
            desea_pi: "defaultdict[int, list]" = defaultdict(list)
            for a, mo, f in valid:
                desea_pi[a - pmin].append(f / (seas_tmp[mo - 1] or 1.0))
            trend = _loglinear_trend(pis_sorted, [_mean(desea_pi[p]) for p in pis_sorted])
            out["trend"] = round(trend, 4)

            # 최종 시즌(정확 추세 기준)
            seas_f, gmean_dt, dt = _detrend_seasonality(trend)
            seas = [round(s, 2) for s in seas_f]
            out["seasonality"] = seas

            # 노이즈: 디트렌드·디시즌 후 상대잔차 std
            resid = [(f - gmean_dt * seas[mo - 1]) / (gmean_dt * seas[mo - 1])
                     for a, mo, f in dt if gmean_dt and seas[mo - 1]]
            out["noise"] = round(_std(resid), 3) if len(resid) >= 2 else 0.05

            # AR(1): period 평균에서 추세·시즌 제거한 잔차의 lag-1 자기상관
            g = 1.0 + trend
            resid_seq = []
            for p in pis_sorted:
                mo = ((p + pmin - 1) % 12) + 1
                pm_dt = _mean(by_pi[p]) / (g ** p) if g > 0 else _mean(by_pi[p])
                exp = gmean_dt * seas[mo - 1]
                if exp:
                    resid_seq.append(pm_dt / exp - 1.0)
            ar1 = _ar1(_linear_detrend(resid_seq))  # 잔여 드리프트 제거 후 순수 자기상관
            if ar1 is not None:
                out["ar1"] = round(ar1, 3)
        else:
            out["trend"] = 0.0
            out["noise"] = 0.05
    else:
        # 날짜 없음: 변동계수만
        m = _mean(nonnull) if n_nn else 0.0
        out["noise"] = round(_std(nonnull) / m, 3) if m else 0.05
    return out


def _choice_shape(values: list[str], include_member_names: bool) -> dict:
    """choice 컬럼 → 카디널리티 + 멤버 빈도(이름 제외 비율벡터, 내림차순)."""
    nonblank = [str(v).strip() for v in values if str(v).strip() != ""]
    cnt = Counter(nonblank)
    n = len(cnt)
    total = sum(cnt.values()) or 1
    freqs = sorted((c / total for c in cnt.values()), reverse=True)
    weights = [round(f, 4) for f in freqs]
    out: "OrderedDict[str, object]" = OrderedDict()
    out["type"] = "choice"
    out["n"] = n
    # 균등에서 유의하게 벗어날 때만 weights 박제(불필요 노이즈 회피)
    if n > 1 and (max(weights) - min(weights)) > (0.5 / n):
        out["weights"] = weights
    if include_member_names:
        out["values"] = [k for k, _ in cnt.most_common()]
    return out


# ----------------------------------------------------------------------------
# 테이블 프로파일
# ----------------------------------------------------------------------------
def profile_table(path: str, *, table_name: str | None = None,
                  date_col: str | None = None, measures: list[str] | None = None,
                  grain: list[str] | None = None,
                  include_member_names: bool = False,
                  scale: str = "EDIT_ME_scale") -> dict:
    """정제 마트 csv → {tables: {name: {columns: {...}}}} 스펙 dict.

    date_col/measures/grain 미지정 시 자동 분류. measure 간 Pearson 상관은
    첫 measure 를 기준으로 나머지에 corr_with/corr 박제(generate 소비용).
    """
    header, body = _read_csv(path)
    name = table_name or os.path.splitext(os.path.basename(path))[0]
    if not header:
        return {"tables": {name: {"columns": {}}}}

    cols = {h: _column(body, i) for i, h in enumerate(header)}
    kinds = {}
    for h in header:
        if h == date_col:
            kinds[h] = "date"
        elif measures and h in measures:
            kinds[h] = "measure"
        else:
            kinds[h] = classify_column(cols[h])

    # 날짜 컬럼(첫 date) → 행별 (year,month)
    date_h = date_col or next((h for h in header if kinds[h] == "date"), None)
    ym_pairs = None
    if date_h:
        ym_pairs = [_to_year_month(v) for v in cols[date_h]]

    measure_names = [h for h in header if kinds[h] == "measure"]
    # measure float 시퀀스(상관용)
    mfloat = {h: [_to_float(v) for v in cols[h]] for h in measure_names}

    columns: "OrderedDict[str, dict]" = OrderedDict()
    for h in header:
        k = kinds[h]
        if k == "date":
            ys = [p for p in (ym_pairs or []) if p]
            start = min(ys) if ys else (2023, 1)
            end = max(ys) if ys else (2025, 12)
            columns[h] = OrderedDict([
                ("type", "date"),
                ("start", "%04d-%02d-01" % start),
                ("end", "%04d-%02d-01" % end),
                ("grain", "month"),
            ])
        elif k == "measure":
            columns[h] = _measure_shape(cols[h], ym_pairs, scale)
        elif k == "key":
            columns[h] = OrderedDict([("type", "key"), ("prefix", (h[:3] or "KEY").upper())])
        else:
            columns[h] = _choice_shape(cols[h], include_member_names)

    # measure 간 상관(첫 measure 기준)
    if len(measure_names) >= 2:
        base_m = measure_names[0]
        for h in measure_names[1:]:
            r = _pearson(mfloat[base_m], mfloat[h])
            if r is not None and abs(r) >= 0.1:
                columns[h]["corr_with"] = base_m
                columns[h]["corr"] = round(r, 3)

    tdef: "OrderedDict[str, object]" = OrderedDict([("columns", columns)])
    if grain:
        tdef["grain"] = grain
    return {"tables": {name: tdef}}


# ----------------------------------------------------------------------------
# 미니 YAML emitter (우리 스펙 구조 전용: dict/list/scalar, 얕은 중첩)
# ----------------------------------------------------------------------------
def _yaml_scalar(v) -> str:
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, float):
        # 정수형 float 은 깔끔하게
        return ("%g" % v)
    if isinstance(v, int):
        return str(v)
    s = str(v)
    if s == "" or re.search(r"[:#\-\{\}\[\],&*!?|>%@`\"']", s) or s.strip() != s:
        return '"%s"' % s.replace('"', '\\"')
    return s


def _emit(obj, indent: int, lines: list[str]) -> None:
    pad = "  " * indent
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, dict):
                lines.append("%s%s:" % (pad, k))
                _emit(v, indent + 1, lines)
            elif isinstance(v, list):
                if v and not isinstance(v[0], (dict, list)):
                    lines.append("%s%s: [%s]" % (pad, k, ", ".join(_yaml_scalar(x) for x in v)))
                else:
                    lines.append("%s%s:" % (pad, k))
                    for item in v:
                        lines.append("%s- " % ("  " * (indent + 1)))
                        _emit(item, indent + 2, lines)
            else:
                lines.append("%s%s: %s" % (pad, k, _yaml_scalar(v)))
    else:
        lines.append("%s%s" % (pad, _yaml_scalar(obj)))


def emit_yaml(spec: dict) -> str:
    lines: list[str] = []
    _emit(spec, 0, lines)
    return "\n".join(lines) + "\n"


_HEADER = (
    "# 회사 실데이터에서 추출한 SHAPE 프로파일 (fpna.profile)\n"
    "# 포함: 카디널리티 / 시즌성(평균=1) / 추세율 / 상대노이즈 / null·음수·0률 /\n"
    "#       measure 상관 / AR(1) 자기상관 / 멤버 빈도(이름 제외) — 전부 차원 없음\n"
    "# 미포함: 실제 값·절대 금액·합계·평균·멤버 이름 (출력에 일절 없음)\n"
    "# base 는 자리표시자(EDIT_ME_scale). 집에서 임의 스케일로 교체 후 generate.\n"
)


def run_profile(path: str, out_path: str, **opts) -> dict:
    """프로파일 추출 + yaml 파일 기록. 반환: spec dict."""
    spec = profile_table(path, **opts)
    text = _HEADER + emit_yaml(spec)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(text)
    return spec


__all__ = [
    "profile_table", "classify_column", "emit_yaml", "run_profile",
    "_read_csv", "_to_float", "_to_year_month",
]
