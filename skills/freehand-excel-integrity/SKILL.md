---
name: freehand-excel-integrity
description: >-
  Guardrails and a validation gate for writing correct, well-formed Excel
  (.xlsx) files by hand with openpyxl — no template engine required. Use this
  skill WHENEVER you are creating, generating, editing, or fixing any .xlsx
  spreadsheet freehand (FP&A reports, cost summaries, forecasts, variance
  tables, KPI/CPU/CPP metrics, board packs, reconciliations, anything with
  totals/subtotals/ratios/headers), even if the user just says "make a
  spreadsheet" or "build this report in Excel". It prevents the recurring
  failure modes of freehand authoring — text/annotations landing in numeric
  cells, header tokens leaking into data rows, #VALUE!/#DIV/0! from unguarded
  ratios, subtotals that don't tie to their parts, wrong-column subtractions,
  and inconsistent formatting — by enforcing authoring rules, a sidecar
  correctness contract, and a deterministic post-write check (scripts/xlsx_doctor.py)
  that must pass before the file is considered done.
---

# Freehand Excel Integrity

> **When this skill fires — creating OR rewriting any .xlsx — you already have the
> full brief. Do not ask the user for a long prompt; just execute the loop below.**

## Operating directive (do this by default)

1. **Author / rebuild** the workbook following the Authoring Rules. If rewriting a
   broken file: recover the intended structure, but **do not copy broken cells
   forward** — move every stray annotation out of value cells into comments/notes,
   keep header tokens (month/quarter) in header rows only, and write numbers/formulas
   only in value cells.
2. **Ratios:** guard live formulas as `=IF(OR(NOT(ISNUMBER(<den>)),<den>=0),"NA",<num>/<den>)`.
   If numbers are computed in code and written statically, write `"NA"` when the
   denominator is invalid — never `#VALUE!`, `None`, or text in a value cell.
3. **Emit `<file>.contract.json`** declaring **every** total (`ties` — add `expected`
   from an *independent* source computation when numbers are code-computed), **every**
   ratio (value-mode `{cell,num,den}`), the `grain`, and `scenario`/`expected_n` where
   they apply. Schema: `references/contract-schema.md`.
4. **Run the gate and fix until clean:**
   ```bash
   python scripts/xlsx_doctor.py <file.xlsx>        # auto-finds <file>.contract.json
   ```
   Definition of done: **`[1]=0, [2]=0, [6] FAIL=0, [7] coverage complete`**, exit 0.
   `[3]` format → `python scripts/xlsx_doctor.py <file> --fix`. `[4]/[5]` advisory.
5. If Excel + pywin32 is available, also open the file to force recalculation (the
   gate reads cells, not recomputed formulas). For static (code-written) values this
   is unnecessary — the gate verifies everything offline.
6. Before declaring done, sanity-check against `references/repo-intent-checklist.md`.

All concrete names here (CPU/CPP, Risk Subtotal, cost lines, `SUMMARY!F42`) are
**examples** — substitute your workbook's actual content. See Scope below.

---

When you author an .xlsx directly (openpyxl), there is no template spine to
enforce correctness, so the same bugs recur: annotation text written into value
cells, month/quarter headers leaking into data rows, ratios that throw #VALUE!,
subtotals that don't equal their children, wrong-column formulas, and ragged
formatting. This skill ports the discipline of a hardened FP&A pipeline into the
freehand workflow via three layers: **authoring rules**, a **sidecar contract**,
and a **deterministic gate** (`scripts/xlsx_doctor.py`).

The goal: after writing any spreadsheet, the gate is green — without a human or a
second agent re-reading the cells.

## Scope (general-purpose)

This skill is **not tied to any specific report**. It applies to *any* .xlsx whose
cells are numbers/formulas with totals, subtotals, ratios, or headers — financial
or not. All concrete names in this skill and its references (CPU, CPP, "Risk
Subtotal", cost lines, R&O, cell refs like `SUMMARY!F42`) are **illustrative
examples only** — substitute whatever your workbook actually contains. The checks
are structural: a "ratio" is any value defined as `num/den`; a "tie" is any total
that should equal its parts; a "field" is any numeric column with a sign/range/set
constraint. Whether the numbers come from live formulas or are computed in code and
written as static values, the same gate applies (see the value-fed section below).

## Workflow (follow every time you write an .xlsx)

1. **Author** the workbook following the Authoring Rules below.
2. **Emit a sidecar contract** `<file>.contract.json` next to the file, declaring
   the totals, grain, ratios, and field constraints you intend (see
   `references/contract-schema.md` for the full schema). This is how value-level
   correctness becomes checkable — the file alone can't prove a subtotal is right.
3. **Run the gate** and read the output:
   ```bash
   python scripts/xlsx_doctor.py <file.xlsx>          # auto-finds <file>.contract.json
   ```
4. **Fix and re-run** until clean:
   - `[1] error cells`, `[2] text in numeric region`, `[6-FAIL] contract` → must be 0.
     These are value/structure bugs: **rewrite the offending cells** (the gate names
     each one). They are not auto-fixable because the correct value comes from the
     source, not the dead file.
   - `[3] format imbalance` → run `python scripts/xlsx_doctor.py <file> --fix`
     (writes `<file>.fixed.xlsx`; `--inplace` to overwrite).
   - `[4] unguarded division`, `[5] cross-foot suspect` → advisory; resolve `[4]`
     by using guarded ratios (Rule 3).
5. Only call the task done when the gate exits 0.

Dependencies: `openpyxl` only. No pandas/numpy. The gate runs in any directory.

## Authoring Rules (apply while writing)

### Cell content & formulas
1. **Cells hold numbers or formulas (`=…`) only.** Put explanations, specs, and
   notes in a *cell comment* or a dedicated notes column — **never** in a value or
   denominator cell. Text in a value cell is the upstream cause of #VALUE!.
2. **Header tokens (APR..DEC, Q1..Q4, FY) belong only in header rows.** Never write
   a month/quarter label as a value in a data row.
3. **Ratios are guarded.** If the cell holds a live formula, write
   `=IF(OR(NOT(ISNUMBER(<den>)),<den>=0),"NA",<num>/<den>)`. **If the value is computed
   in Python and written as a static number** (no formula), then when the denominator
   is missing/zero write the string `"NA"` (or leave NO_DATA) — never write `#VALUE!`,
   `None`, or text into the cell, and never write a number you didn't actually divide.
4. **Totals/subtotals/hierarchy use `=SUM(<children range>)`** and are declared in
   the contract `ties` so the gate can confirm they actually equal their parts.
8. **Sum only leaf/data rows.** Subtotal and total rows are for cross-checking, not
   for re-summing — re-adding them double-counts.

### Shape & completeness (don't fall into data-engineering defaults)
5. **Don't drop one-sided keys.** A key present in only Actual or only Budget must
   be shown (as a row), not silently zeroed. Missing periods become a NO_DATA row,
   not a skipped row. Declare `scenario` / `expected_n` in the contract.
9. **One row = one item (declare your grain).** Decide what a single row represents
   before writing; don't mix grains. Duplicate keys silently double-count — declare
   `grain` so the gate catches it.
10. **Time axis comes from a calendar, not from the data.** Build period columns
    from the fiscal calendar (e.g. APR..MAR, Q1..Q4, FY), not by scraping whatever
    dates happen to appear. Every calendar period gets a column even if empty.
11. **No forbidden shortcuts (default-deny):** no sparse compression (dropping
    zero/empty rows), no "one period only", no snapshot-only, no anti-join-only
    (showing only mismatches). Show the full population; flag, don't hide.

### Values & provenance (critical for real work files)
6. **Preserve source values.** No synthetic/placeholder financial numbers. Keep the
   raw value; show units in the header and via number format — never mutate the value
   (don't pre-divide by 1,000,000 — use a display format like `#,##0,,`).
12. **Point-in-time discipline.** If you don't have a number, leave NO_DATA — never
    invent one. When the file is a real business artifact, don't fabricate vendor
    names, cost centers, or amounts to fill gaps.

### Look (design consistency — prevents "imbalance")
7. **One number_format per column.** Negatives in parentheses (`#,##0;(#,##0)`),
   percent `0.0%`, multiple `0.0"x"`. Total rows get a top border; the final total a
   single-top + double-bottom border.
13. **Consistent structure top to bottom:** title → meta (unit/currency/basis/date)
    → body → reconciliation block → source footer. One accent color; input cells
    blue, calculated black, cross-sheet links green; positive green, negative red.
    Freeze the header row; turn gridlines off.

## Rewriting an existing broken file

When asked to rewrite a file that already has bugs (e.g. annotations in cells,
#VALUE!, subtotals that don't tie):
1. Read the broken file to recover the intended structure (labels, periods, which
   cells are totals/ratios). The bugs tell you what was attempted.
2. Rebuild following the Authoring Rules — do **not** copy the broken cells forward.
   In particular, move every stray annotation out of value cells into comments/notes,
   and rewrite ratios as guarded formulas.
3. Write the contract declaring **every** total (`ties`), ratio (`ratios`), the grain,
   and scenario/expected_n where applicable. Use the gate's `[7] coverage` output to
   confirm you didn't leave a total or ratio undeclared.
4. Run the gate to green before returning. If Excel + pywin32 is available, also open
   the file to force a recalculation — the gate reads formulas, not recomputed values,
   so live `#VALUE!`/`#DIV/0!` only surface when Excel recalculates.

> **Cross-check against repo intent:** if you are porting from a hardened pipeline,
> read `references/repo-intent-checklist.md` and confirm each core intent is honored
> (or consciously out of scope) before declaring the rewrite done.

## The sidecar contract (how value-correctness is enforced)

The contract is the freehand equivalent of a build-time metadata stamp: you
declare your intent, the gate checks the rendered file against it. Minimal example:

```json
{
  "sheet": "SUMMARY",
  "header_rows": [12],
  "numeric_regions": [["SUMMARY", 13, 57, 2, 9]],
  "ties": [
    {"name": "TOTAL == Σ cost lines", "total": "SUMMARY!F32", "parts": "SUMMARY!F13:F31"},
    {"name": "Risk Subtotal",        "total": "SUMMARY!F42", "parts": "SUMMARY!F37:F41"}
  ],
  "grain":  {"region": "SUMMARY!A13:A31"},
  "ratios": ["SUMMARY!G54", "SUMMARY!G55"]
}
```

It also supports `fields` (sign/min/max/accepted_values), `scenario`
(actual/budget population alignment), `expected_n` (row-count preservation), and
`formula_refs` (verify a column of formulas references the intended columns and
direction). **Read `references/contract-schema.md` for the full schema, every
field, and which correctness perspective each one enforces.** Declare only what
applies to the workbook you are building — unspecified checks are skipped.

### Value-fed workbooks (Python-computed numbers, no live formulas)

If the cells hold **static numbers your code computed** (the common case — no `=`
formulas linking cells), the gate is actually *stronger*: it can fully verify the
file offline, with no Excel recalculation. Two contract features make this an
**independent recomputation** (the N-version idea):

- **`ties[].expected`** — declare the total you computed **from the source via an
  independent path** (not by reusing the same sum). The gate compares the rendered
  total cell to it. This catches a wrong total even when the rendered children also
  happen to be wrong (which a rendered-total-vs-rendered-parts check would miss).
- **value-mode `ratios`** — declare `{"cell": …, "num": …, "den": …}`. The gate
  recomputes `num/den` from those cells and compares to the rendered ratio, and flags
  a `#VALUE!`/text/`None` where a number was expected. This is exactly the CPU/CPP
  failure in real reports: a wrong or error ratio that no formula-guard would catch.

So for value-fed files, prefer: every total gets `expected` (independent) **and**
`parts` (internal), every ratio gets value-mode num/den. The guarded-formula rule
above does not apply (there are no formulas); the discipline moves into the Python
that computes the numbers — and the gate re-derives them from the rendered cells.

## What the gate catches

| # | Check | Severity | Auto-fix |
|---|---|---|---|
| 1 | Error cells (#REF!/#VALUE!/#DIV/0!) | fatal | no → rewrite |
| 2 | Text in numeric region (header/annotation leak) | fatal | no → rewrite |
| 3 | Column format imbalance | issue | yes (`--fix`) |
| 4 | Unguarded division (#VALUE! risk) | advisory | no → guard |
| 5 | Cross-foot suspect (subtotal ≠ children) | advisory | no |
| 6 | Sidecar contract (tie/grain/ratio/field/scenario/formula) | fatal on FAIL | no → rewrite |
| 7 | Contract coverage (totals/ratios present but undeclared) | advisory | declare them |

## Automating the gate in Claude Code (recommended)

Add a PostToolUse hook so the gate runs automatically after any .xlsx write and
blocks (exit 2) until clean, prompting same-turn self-correction. Put in
`.claude/settings.json`:

```json
{ "hooks": { "PostToolUse": [ { "matcher": "Write|Edit", "hooks": [
  { "type": "command",
    "command": "f=$(jq -r '.tool_input.file_path // empty'); case \"$f\" in *.xlsx) python scripts/xlsx_doctor.py \"$f\" >&2 || exit 2;; esac" }
] } ] } }
```

(Adjust the script path to where you install the skill.)
