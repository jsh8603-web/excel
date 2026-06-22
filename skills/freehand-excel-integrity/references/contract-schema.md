# Sidecar contract schema (`<file>.contract.json`)

The contract is the freehand equivalent of a build-time metadata stamp: you
declare the workbook's intended invariants, and `xlsx_doctor.py` checks the
rendered file against them. The gate auto-discovers `<file>.contract.json` next
to the workbook (or pass `--contract <path>`). Declare only the keys that apply;
unspecified checks are skipped. Any `FAIL` from this section is fatal (exit 1).

References (`SHEET!A1` or `SHEET!A1:B5`). A bare `A1` uses the top-level `sheet`.

> All sheet/cell names below (`SUMMARY`, `Risk Subtotal`, `F42`, CPU/CPP, …) are
> **examples** — substitute your workbook's actual sheet, ranges, and labels. The
> checks are structural, not domain-specific.

## Static vs formula basis (decide first)

How the numbers get into cells decides which keys you **must** declare and how the recalc
gate behaves (`backend-routing.md`):

- **Static values** (computed in Python, written as plain numbers — no `=`): the file can't
  prove a total by itself, so declare **`ties[].expected`** (an *independent* source total, true
  N-version) and **value-mode `ratios`** (`{cell,num,den}`). `xlsx_doctor` re-derives and compares
  fully **offline** — no recalculation needed.
- **Formula-driven** (`=SUM(...)`, `=A-B` written then filled down): declare `ties` with `parts`
  (the gate confirms the `SUM` covers them); the **recalc gate** (pywin32 → LibreOffice →
  `formulas`) confirms the formulas evaluate without `#VALUE!`/`#DIV/0!`, and the **fill-down
  linter** confirms each column's formulas share one relative-reference shape.

A workbook may mix both; declare per cell accordingly.

## Top-level keys

| key | type | enforces | repo perspective |
|---|---|---|---|
| `sheet` | string | default sheet for bare refs | — |
| `header_rows` | int[] | header tokens (APR..FY, Q1..Q4) only on these rows; elsewhere = leak | R5 / content-type |
| `numeric_regions` | `[[sheet,r0,r1,c0,c1], …]` | cells in the box are number or formula only; stray text = FAIL | R5 / content-type, upstream of #VALUE! |
| `ties` | object[] | declared total == Σ parts (or formula SUM covers parts) | R3 tie-out, R10 hierarchy, R11 master/alloc, R14 commitment |
| `grain` | object | key range has no duplicate labels (silent-merge guard) | R8 grain, binding grain_unique |
| `ratios` | string[] | each ratio cell formula is guarded (ISNUMBER/IFERROR/IFNA) | R17 ratio_na |
| `fields` | object[] | numeric range obeys sign/min/max; cells ∈ accepted_values | metric_table FieldSpec, R13 sign |
| `scenario` | object | Actual/Budget label sets are the same population | R9 scenario_aligned, R2 full population |
| `expected_n` | object | non-empty rows in region == declared n | R7 no_silent_drop |
| `formula_refs` | object[] | each formula in region is `=<left><row><op><right><row>` | formula direction/column (wrong-column guard) |

## Field details

### `ties` — total reconciles to its parts (R3/R10/R11/R14)
```json
{"name": "Risk Subtotal", "total": "SUMMARY!F42", "parts": "SUMMARY!F37:F41", "tol": 0.5}
```
- If `total` is a number → compares it to the sum of numeric cells in `parts` (within `tol`, default 0.5).
- If `total` is a formula → checks the formula is a `SUM` covering the declared `parts` range (the value itself is confirmed when Excel recalculates).
- **`expected` (number, optional)** — value-fed workbooks: the total you computed *independently from the source*. Gate compares the rendered total cell to it (true N-version — catches a wrong total even when rendered children are also wrong). Use with `parts`: `{"total": "SUMMARY!F52", "expected": 49732, "parts": "SUMMARY!F13:F31"}`.
- Catches: subtotal stuck at 0, subtotal pointing at the wrong range, parent ≠ Σ children, total ≠ source.

### `grain` — unique keys (R8)
```json
{"region": "SUMMARY!A13:A31"}
```
Duplicate non-empty text labels in the range = FAIL (a duplicated key silently merges/doublecounts downstream).

### `ratios` — ratio completeness (R17)
Two forms (mix freely):
- **Formula-guard** (live formula cells) — a ref string; the formula must contain `ISNUMBER`/`IFERROR`/`IFNA`, else FAIL. Error literal also FAILs.
  ```json
  ["SUMMARY!G54", "SUMMARY!G55"]
  ```
- **Value mode** (Python-computed static ratios) — `{"cell","num","den","tol"?}`. Gate recomputes `num/den` from the cells and compares to the rendered ratio; if `den` missing/zero the cell must be `"NA"` or blank (a number or `#VALUE!` = FAIL). Catches wrong/error CPU/CPP.
  ```json
  [{"cell": "SUMMARY!G54", "num": "SUMMARY!F52", "den": "SUMMARY!F53"}]
  ```

### `fields` — value constraints (FieldSpec / R13)
```json
[
  {"region": "SUMMARY!F13:F31", "sign": "nonneg"},
  {"region": "SUMMARY!F13:F31", "min": 0, "max": 1000000},
  {"region": "SUMMARY!B13:B31", "accepted_values": ["Risk", "Opportunity"]}
]
```
- `sign`: `"+"`/`"nonneg"` (no negatives) or `"-"`/`"nonpos"` (no positives).
- `min`/`max`: numeric bounds on each numeric cell.
- `accepted_values`: every non-empty, non-formula cell must be in the set (works for text or numbers).

### `scenario` — population alignment (R9/R2)
```json
{"actual": "SUMMARY!A13:A20", "budget": "BUDGET!A13:A20"}
```
Compares the distinct text labels of the two ranges. Keys present in only one side = FAIL (don't zero one-sided keys; surface them as LEFT_ONLY/RIGHT_ONLY rows).

### `expected_n` — row-count preservation (R7)
```json
{"region": "SUMMARY!A13:A31", "n": 13}
```
Counts non-empty cells in the region; ≠ `n` = FAIL (silent drop or duplication of rows).

### `formula_refs` — formula direction/column (wrong-column guard)
```json
[{"region": "VAR!D13:D40", "op": "-", "left": "C", "right": "B", "name": "Δ = Actual - Plan"}]
```
Each formula cell in the region must be exactly `=<left><row><op><right><row>`
(`$` and spaces ignored). `left`/`right` accept a column letter or 1-based index.
Catches wrong-column subtraction and flipped direction (e.g. `=B-C` where `=C-B`
was intended) — the class that value re-derivation tautologies miss.

## Full example

```json
{
  "sheet": "SUMMARY",
  "header_rows": [12],
  "numeric_regions": [["SUMMARY", 13, 57, 2, 9]],
  "ties": [
    {"name": "TOTAL == Σ cost lines", "total": "SUMMARY!F32", "parts": "SUMMARY!F13:F31"},
    {"name": "Risk Subtotal",        "total": "SUMMARY!F42", "parts": "SUMMARY!F37:F41"},
    {"name": "R&O Net = Risk - Opp", "total": "SUMMARY!F50", "parts": "SUMMARY!F42:F49"}
  ],
  "grain":   {"region": "SUMMARY!A13:A31"},
  "ratios":  ["SUMMARY!G54", "SUMMARY!G55", "SUMMARY!G56", "SUMMARY!G57"],
  "fields":  [{"region": "SUMMARY!F13:F31", "sign": "nonneg"}],
  "expected_n": {"region": "SUMMARY!A13:A31", "n": 19}
}
```

## Perspectives NOT in the contract (authoring rules / heuristics instead)

Some repo perspectives are author-discipline (no machine check) or covered by the
no-contract heuristics in the gate:
- Design skeleton (title/meta/body/_RECON/source), color roles, total borders,
  single accent, fonts — **authoring rules** (SKILL.md), partially the `[3]` format check.
- Forbidden heuristics / default-deny (no sparse-compress, no snapshot-only),
  filter-declaration, no synthetic financials, PIT — **authoring rules**.
- Error cells, format imbalance, stray text, unguarded division, cross-foot —
  **heuristics**, run with no contract needed.
