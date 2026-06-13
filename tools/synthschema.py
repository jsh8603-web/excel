"""
synthschema.py — 스키마(SHAPE yaml)에서 합성 더미를 생성하는 dev 전용 도구.

⚠ dev 전용(pandas/numpy/pyyaml 의존 — vendored 아님). 무설치 런타임(fpna/)이 import 하면 안 된다.
fpna.profile(stdlib)이 뽑은 SHAPE yaml 을 받아 더미를 만든다 — 한 쌍.

핵심 아이디어: 민감한 건 데이터 '값'이지 '구조(스키마)'가 아니다. 스키마만 정의하면
그 위에서 그럴듯한 더미로 Power BI 모델·DAX·변환 로직을 집에서 개발하고, 회사에서는
소스 커넥션만 실데이터로 스왑한다.

fpna.profile 의 8축 SHAPE 를 소비:
  시즌성 / 추세 / 노이즈 / null_rate / **neg_rate(음수 허용)** /
  **corr_with·corr(measure 간 상관)** / **ar1(시계열 관성)** / choice weights(멤버 빈도)

CLI:
  python synthschema.py gen profile_spec.yaml -o out/ [--excel out.xlsx] [--seed 7]
"""
from __future__ import annotations

import argparse
import math
from pathlib import Path

import numpy as np
import pandas as pd

try:
    import yaml
except ImportError:
    yaml = None


# -----------------------------------------------------------------------------
# 스펙 로딩
# -----------------------------------------------------------------------------
def load_spec(path_or_dict):
    if isinstance(path_or_dict, dict):
        return path_or_dict
    if yaml is None:
        raise RuntimeError("YAML 스펙을 읽으려면 pyyaml 이 필요합니다: pip install pyyaml")
    with open(path_or_dict, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


# -----------------------------------------------------------------------------
# 의존성 정렬 (FK 기준 토폴로지 정렬)
# -----------------------------------------------------------------------------
def _table_deps(tables):
    deps = {name: set() for name in tables}
    for name, tdef in tables.items():
        for col, cdef in tdef.get("columns", {}).items():
            if cdef.get("type") == "fk":
                ref_table = str(cdef["ref"]).split(".")[0]
                if ref_table != name:
                    deps[name].add(ref_table)
    ordered, visited = [], set()

    def visit(n, stack):
        if n in visited:
            return
        if n in stack:
            raise ValueError(f"순환 참조 감지: {n}")
        stack.add(n)
        for d in deps[n]:
            if d not in tables:
                raise ValueError(f"{n} 의 FK 가 존재하지 않는 테이블 {d} 를 참조")
            visit(d, stack)
        stack.discard(n)
        visited.add(n)
        ordered.append(n)

    for n in tables:
        visit(n, set())
    return ordered


# -----------------------------------------------------------------------------
# 날짜 차원
# -----------------------------------------------------------------------------
_GRAIN_FREQ = {"day": "D", "month": "MS", "quarter": "QS", "year": "YS"}


def _build_date_dim(col, cdef, rng):
    start = pd.Timestamp(cdef.get("start", "2023-01-01"))
    end = pd.Timestamp(cdef.get("end", "2025-12-31"))
    grain = cdef.get("grain", "month")
    freq = _GRAIN_FREQ.get(grain, "MS")
    dates = pd.date_range(start, end, freq=freq)
    df = pd.DataFrame({col: dates})
    df["Year"] = dates.year
    df["Quarter"] = "Q" + dates.quarter.astype(str)
    df["Month"] = dates.month
    df["MonthName"] = dates.strftime("%b")
    df["YearMonth"] = dates.strftime("%Y-%m")
    df["Period"] = np.arange(len(dates))
    return df


# -----------------------------------------------------------------------------
# 측정값
# -----------------------------------------------------------------------------
def _seasonal_factors(spec, rng):
    if spec is True or spec is None:
        base = np.array([0.90, 0.85, 0.95, 0.98, 1.00, 1.02,
                         1.00, 0.98, 1.02, 1.08, 1.20, 1.30])
        return base
    if isinstance(spec, list) and len(spec) == 12:
        return np.array(spec, dtype=float)
    return np.ones(12)


def _coerce_base(cdef):
    """base 가 'EDIT_ME_scale' 같은 자리표시자면 기본 스케일로 대체."""
    raw = cdef.get("base", 100000)
    try:
        return float(raw)
    except (ValueError, TypeError):
        return 100000.0


def _apply_sign(vals, cdef, rng, n):
    """neg_rate 비율만큼 음수로(부호 반전). neg_rate=0/없음이면 음수 제거(클램프).

    음수 없는 데이터 흐름엔 음수가 절대 안 나오고, neg_rate>0(적자·조정 계정)이면
    그 비율만큼만 음수가 나타난다. 양방향 충실.
    """
    nr = cdef.get("neg_rate")
    if nr:
        flip = rng.random(n) < float(nr)
        return np.where(flip, -np.abs(vals), np.abs(vals))
    return np.maximum(vals, 0)


def _gen_measure(df, col, cdef, date_col, rng, generated, gen_noise):
    n = len(df)
    rounding = cdef.get("round", 0)

    derive = cdef.get("derive_from")
    if derive:
        if derive not in generated:
            raise ValueError(f"measure '{col}' 의 derive_from '{derive}' 가 아직 생성되지 않음")
        var = cdef.get("variance", 0.08)
        vals = df[derive].to_numpy() * (1.0 + rng.normal(0, var, n))
        vals = _apply_sign(vals, cdef, rng, n)
        return np.round(vals, rounding) if rounding is not None else vals

    base = _coerce_base(cdef)
    trend = float(cdef.get("trend", 0.0))
    noise = float(cdef.get("noise", 0.05))
    seasonal = _seasonal_factors(cdef.get("seasonality"), rng)

    if date_col and date_col in df.columns:
        periods = df[date_col].dt.year * 12 + df[date_col].dt.month
        periods = (periods - periods.min()).to_numpy()
        months = df[date_col].dt.month.to_numpy()
        season = seasonal[months - 1]
    else:
        periods = np.zeros(n)
        season = np.ones(n)

    # 표준정규 노이즈 성분 → AR(1) 관성 → corr_with 상관 부여 (순서 유의)
    eps = rng.normal(0, 1, n)
    ar1 = cdef.get("ar1")
    if ar1 is not None and abs(float(ar1)) < 1:
        a = float(ar1)
        k = math.sqrt(1.0 - a * a)
        e = np.empty(n)
        e[0] = eps[0]
        for t in range(1, n):
            e[t] = a * e[t - 1] + k * eps[t]
        eps = e
    cw, cr = cdef.get("corr_with"), cdef.get("corr")
    if cw and cr is not None and cw in gen_noise and len(gen_noise[cw]) == n:
        r = float(cr)
        eps = r * gen_noise[cw] + math.sqrt(max(0.0, 1.0 - r * r)) * eps
    gen_noise[col] = eps

    vals = base * (1 + trend) ** periods * season * (1 + noise * eps)
    vals = _apply_sign(vals, cdef, rng, n)   # neg_rate=0 → 음수제거 / >0 → 비율만큼 음수
    return np.round(vals, rounding) if rounding is not None else vals


# -----------------------------------------------------------------------------
# 컬럼 생성 (비-측정, 비-fk)
# -----------------------------------------------------------------------------
def _gen_simple_column(col, cdef, n, rng):
    t = cdef["type"]
    if t == "id":
        return np.arange(1, n + 1)
    if t == "key":
        prefix = cdef.get("prefix", col[:3].upper())
        width = cdef.get("width", 3)
        return [f"{prefix}{i:0{width}d}" for i in range(1, n + 1)]
    if t == "choice":
        if "values" in cdef:
            vals = cdef["values"]
        else:
            cnt = cdef.get("n", 5)
            prefix = cdef.get("prefix", col)
            vals = [f"{prefix}_{i}" for i in range(1, cnt + 1)]
        weights = cdef.get("weights")
        if weights and len(weights) == len(vals):
            p = np.array(weights, dtype=float) / np.sum(weights)
        else:
            p = None
        return rng.choice(vals, size=n, p=p)
    if t == "int":
        return rng.integers(cdef.get("min", 0), cdef.get("max", 100) + 1, n)
    if t == "float":
        lo, hi = cdef.get("min", 0.0), cdef.get("max", 1.0)
        vals = rng.uniform(lo, hi, n)
        r = cdef.get("round", 2)
        return np.round(vals, r) if r is not None else vals
    if t == "bool":
        return rng.random(n) < cdef.get("p", 0.5)
    raise ValueError(f"알 수 없는 컬럼 타입: {t} (컬럼 {col})")


# -----------------------------------------------------------------------------
# 테이블 생성
# -----------------------------------------------------------------------------
def _gen_table(name, tdef, generated, rng):
    cols = tdef.get("columns", {})
    date_cols = [c for c, d in cols.items() if d.get("type") == "date"]
    measure_cols = [c for c, d in cols.items() if d.get("type") == "measure"]
    grain = tdef.get("grain")
    if len(date_cols) > 1:
        raise ValueError(f"{name}: 테이블당 date 컬럼은 하나만 지원")

    if date_cols and not measure_cols and not grain:
        dc = date_cols[0]
        df = _build_date_dim(dc, cols[dc], rng)
        for c, d in cols.items():
            if c == dc or c in df.columns:
                continue
            df[c] = _gen_simple_column(c, d, len(df), rng)
        return df

    if date_cols and measure_cols and not grain:
        dc = date_cols[0]
        date_df = _build_date_dim(dc, cols[dc], rng)
        choice_cols = [c for c, d in cols.items() if d.get("type") in ("choice", "key")]
        axes = {dc: date_df[dc].to_numpy()}
        for c in choice_cols:
            ncnt = cols[c].get("n") or (len(cols[c].get("values", [])) or 5)
            axes[c] = _gen_simple_column(c, cols[c], ncnt, rng)
        idx = pd.MultiIndex.from_product(list(axes.values()), names=list(axes.keys()))
        df = idx.to_frame(index=False)
        max_rows = tdef.get("max_rows")
        if max_rows and len(df) > max_rows:
            df = df.sample(max_rows, random_state=int(rng.integers(0, 1_000_000))).reset_index(drop=True)
        df = df.merge(date_df, on=dc, how="left")
        df[dc] = pd.to_datetime(df[dc])
        # 시간순 정렬 → AR(1) 관성이 의미를 갖도록
        df = df.sort_values(dc).reset_index(drop=True)
        date_col = dc

    elif grain:
        key_arrays = []
        for gcol in grain:
            cdef = cols[gcol]
            ref_t, ref_c = str(cdef["ref"]).split(".")
            keys = generated[ref_t][ref_c].unique()
            key_arrays.append((gcol, keys))
        idx = pd.MultiIndex.from_product([k for _, k in key_arrays],
                                         names=[c for c, _ in key_arrays])
        df = idx.to_frame(index=False)
        max_rows = tdef.get("max_rows")
        if max_rows and len(df) > max_rows:
            df = df.sample(max_rows, random_state=int(rng.integers(0, 1_000_000))).reset_index(drop=True)
        date_col = None
        for gcol in grain:
            ref_t, ref_c = str(cols[gcol]["ref"]).split(".")
            ref_df = generated[ref_t]
            ref_date = [c for c in ref_df.columns
                        if pd.api.types.is_datetime64_any_dtype(ref_df[c])]
            if ref_date:
                date_col = "_period_dt"
                mapping = dict(zip(ref_df[ref_c], ref_df[ref_date[0]]))
                df[date_col] = df[gcol].map(mapping)
                break
    else:
        n = tdef.get("rows")
        if n is None:
            lens = [len(d["values"]) for d in cols.values()
                    if d.get("type") == "choice" and "values" in d]
            n = max(lens) if lens else 10
        df = pd.DataFrame(index=range(n))
        date_col = None

    measure_defs = []
    for c, d in cols.items():
        t = d.get("type")
        if c in df.columns:
            continue
        if t == "fk":
            ref_t, ref_c = str(d["ref"]).split(".")
            df[c] = rng.choice(generated[ref_t][ref_c].unique(), size=len(df))
        elif t == "measure":
            measure_defs.append((c, d))
        else:
            df[c] = _gen_simple_column(c, d, len(df), rng)

    # 측정값: derive_from / corr_with 의존성 때문에 base 먼저 (위상 정렬)
    pending = list(measure_defs)
    generated_measures = set()
    gen_noise = {}
    safety = 0
    while pending:
        safety += 1
        if safety > len(measure_defs) + 2:
            raise ValueError(f"{name}: measure 의존성 해결 실패 (순환?)")
        nxt = []
        for c, d in pending:
            dep = d.get("derive_from") or d.get("corr_with")
            if dep and dep not in df.columns and dep not in generated_measures:
                nxt.append((c, d))
                continue
            df[c] = _gen_measure(df, c, d, date_col, rng,
                                 set(df.columns) | generated_measures, gen_noise)
            generated_measures.add(c)
        pending = nxt

    df = df.drop(columns=[c for c in df.columns if c.startswith("_period_dt")], errors="ignore")
    ordered_cols = [c for c in cols if c in df.columns] + \
                   [c for c in df.columns if c not in cols]
    return df[ordered_cols].reset_index(drop=True)


# -----------------------------------------------------------------------------
# 공개 API
# -----------------------------------------------------------------------------
def generate(spec, seed=42):
    spec = load_spec(spec)
    tables = spec["tables"]
    rng = np.random.default_rng(spec.get("seed", seed))
    order = _table_deps(tables)
    generated = {}
    for name in order:
        generated[name] = _gen_table(name, tables[name], generated, rng)
    return generated


def to_csv_dir(frames, out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    for name, df in frames.items():
        df.to_csv(out / f"{name}.csv", index=False, encoding="utf-8-sig")
    return out


def to_excel(frames, path):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df in frames.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
    return path


# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------
def _main(argv=None):
    p = argparse.ArgumentParser(description="SHAPE 스키마 기반 합성 더미 생성기 (dev 전용)")
    sub = p.add_subparsers(dest="cmd", required=True)

    g = sub.add_parser("gen", help="SHAPE yaml 에서 더미 생성")
    g.add_argument("spec")
    g.add_argument("-o", "--out", default="out", help="CSV 출력 디렉토리")
    g.add_argument("--excel", help="단일 엑셀 워크북 경로(시트=테이블)")
    g.add_argument("--seed", type=int, default=42)

    args = p.parse_args(argv)
    if args.cmd == "gen":
        frames = generate(args.spec, seed=args.seed)
        to_csv_dir(frames, args.out)
        msg = [f"{name}: {len(df)} rows -> {args.out}/{name}.csv" for name, df in frames.items()]
        if args.excel:
            to_excel(frames, args.excel)
            msg.append(f"엑셀: {args.excel}")
        print("\n".join(msg))


if __name__ == "__main__":
    _main()
