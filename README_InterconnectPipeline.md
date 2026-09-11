# Interconnection Site-Scoring Pipeline (`modInterconnectPipeline.bas`)

One importable Excel VBA module with a **single public entry point,
`RunInterconnectPipeline`**, that runs the whole workflow end to end from a
blank workbook on one **Alt+F8** call. A second public procedure, `SelfTest`,
is the test harness (in-memory checks plus a live-Excel parity round-trip and a
2,000-substation scale smoke test).

Import via the VBA editor (**Alt+F11 → File → Import File…**) and run with
**Alt+F8 → `RunInterconnectPipeline`**. Plain `.bas`, `Option Explicit`, private
helpers `pl_`-prefixed, no external references (`Scripting.Dictionary` is used
via late binding), no `.Select` / `.Activate` / `Selection`.

---

## The four-step flow (run order)

1. **Cost / results files → `Cost Data`.** Multi-select the cost workbooks
   (picker rooted at `ThisWorkbook.Path`, falling back to
   `Application.DefaultFilePath`). Sheet 3 carries the data; the substation name
   is parsed from its **tab name** (last plausible numeric token = voltage, the
   rest = name; bare tab = name only). The output header is written **once**
   (the metadata prefix `Substation Name | Voltage (kV) | Source File |
   Source Sheet`, then the source columns verbatim — including
   `Size Overload Occurs (MW)` (trigger) and `Proposed Project Allocation ($)`
   (allocation)); from each source only the **data rows** (row 2 down of its
   used range) are appended. Blank rows and any row that re-states the source
   header (matches the header signature) are skipped, so `Cost Data` never
   accumulates stray header rows that would corrupt the downstream
   `MINIFS`/`SUMIFS`/grouping scans.
2. **Site files → `Site Data`.** From each `Summary` tab (name = A, state = C,
   voltage = G, **row 2 down** — header taken once, data rows only, a
   header-matching row skipped), de-duplicated on the full
   **(name, voltage, state) triple** (a different voltage or state is a
   different substation).
3. **Join → `Matrix`.** Intersection only, matched by name (plus voltage when
   the cost tab carried one). Cost-per-MW per (substation, MW).
4. **Analysis + ranking + weighted scoring + headroom → `Cost Curve Analysis`.**

Also produced: a run log **`_Pipeline Log`** (per-file status, dropped
duplicate triples, join alignment errors) and a **`_Pipeline State`** checkpoint
sheet (see below).

---

## Durable & resumable (checkpoint every stage)

Each stage **writes its output sheet and saves the workbook before the next
stage begins**:

```
cost → write Cost Data → Save → site → write Site Data → Save →
matrix → write Matrix → Save → analysis → write → Save
```

So a crash, a runtime error, or even a VBA **compile break** in a later stage
leaves every earlier sheet intact and saved on disk. The **sheet is the source
of truth** — a stage never keeps its only copy in a module-level array across
stages, and a resumed run reads `Cost Data` / `Site Data` / `Matrix` back from
their sheets instead of recomputing them.

**Per-stage failure isolation.** Each stage has its own handling: on failure it
saves what exists, restores application state, logs the reason, and exits with a
message naming the completed stages — e.g. *“Stage 2 failed: &lt;err&gt;.
Completed and saved: Cost Data. Fix and re-run — it will resume.”* One stage’s
failure never rolls back earlier sheets.

**Resume vs Restart (explicit).** At entry the completed stages are detected
(each output sheet exists and is non-empty; `_Pipeline State` records the last
completed stage, its row count, and a timestamp). When valid checkpoints exist
you get an explicit prompt naming exactly what was found and where a resume
would begin, e.g.:

> Found completed stages: Cost Data (42,013 rows), Site Data (1,987 rows).
> Resume from Stage 3 (Join/Matrix), or Restart from Stage 1?
> [Yes = Resume] [No = Restart] [Cancel]

**Resume** skips the completed stages and reads their data back from the
persisted sheets — it does **not** re-open any source files for a stage already
checkpointed, so the expensive 2,000-file Stage 1 is never repeated because a
later stage had a typo. **Restart** clears the output sheets and `_Pipeline
State`, then runs from Stage 1. Resume is the default. If no valid checkpoints
exist, the run simply starts at Stage 1 with no prompt (and still creates
`_Pipeline State` at the first checkpoint).

**First save of a blank workbook.** If the host has never been saved (no path),
the first checkpoint prompts once for a location, falling back to a timestamped
`InterconnectPipeline_<ts>.xlsm` in `Application.DefaultFilePath`. The
**macro-enabled format is required** (`xlOpenXMLWorkbookMacroEnabled`, 52) since
the module lives in the workbook. `ScreenUpdating`/`EnableEvents` stay off across
the save; per-stage saving is cheap relative to the work and is the price of
durability.

**Compile-safety (block `If` rule).** A VBA compile error is a project-load
failure that halts the whole run regardless of checkpoints, so the module is
kept compile-clean. In particular, **no single-line `If` uses `ElseIf` or chains
more than one statement after `Then` with colons** — any `If` needing more than
one action, or any `Else`/`ElseIf`, is a multi-line block
`If … Then` / `ElseIf` / `Else` / `End If`. (Simple single-statement guards like
`If x > 0 Then y = 1` are left as one line.) One `Attribute VB_Name`, balanced
terminators, no duplicate procedures. The saved per-stage sheets protect the
data even across a compile break; a clean compile prevents the halt — both are
needed. The module was verified structurally (balanced
`Sub`/`Function`/`If`/`For`/`With`/`Do`, single `Attribute VB_Name`, no
duplicate procedure names, call-site arities); run **Debug → Compile
VBAProject** once on import to confirm zero compile errors (the definitive
check; `SelfTest` prints the same reminder).

---

## Scaling to ~2,000 substations (what changed; results unchanged)

The optimization changes **how** the work is done, not what is produced. Global
run settings are set once at entry and restored in `Cleanup` on every exit
(including the error handler): `ScreenUpdating=False`, `EnableEvents=False`,
`DisplayAlerts=False`, `Calculation=xlManual`, `AskToUpdateLinks=False`, a live
`Application.StatusBar`, and a **single `Application.CalculateFull` at the very
end** (no intermediate recalcs; the chart is built after it).

| Stage | Before (bottleneck) | After |
|-------|---------------------|-------|
| 1 Consolidation | per-cell reads/writes | every workbook opened `UpdateLinks:=0, ReadOnly:=True, AddToMru:=False` and closed immediately; used range read in one `Range.Value`; output written one `Range.Value = array` per file; one `DoEvents` + status line per file; a locked file is logged and skipped, never aborting the batch |
| 2 Dedupe | O(N²) pairwise compare | `Scripting.Dictionary` keyed on `LCase(name)|voltage|LCase(state)` — one O(N) pass |
| 3 Join | rescans Cost Data per substation | one-pass dictionaries (identity key → id; name → records) resolved by O(1) lookup |
| 4 Matrix | 42,000 live `SUMIFS` | cost columns loaded once, grouped by name; each substation's trigger→allocation records scanned once to accumulate cumulative allocation `T` at each MW and divide by MW (the exact `SUMIFS` definition); written as one values block |
| 5 Ranking | ~32M-op rank-by-scanning | stable **mergesort** once per column (O(N log N)), competition ranks assigned in one walk |
| 6 Percentiles & score | ~22,000 volatile `PERCENTRANK.EXC`/`CEILING` | computed in VBA (`PERCENTRANK.EXC` implemented as `k/(N+1)`), written as value blocks |
| 7 Headroom | live `MINIFS` | per-substation minimum trigger from the already-loaded cost data |

Every sheet write is a single `Range.Value`/`Range.Formula = array` per block;
number formats are applied per column once after the values land.

### Two switches (both default `False` = computed values)

- **`USE_LIVE_SUMIFS`** — `True` makes Stage 4 write live `SUMIFS` bound to the
  used rows (`$D$2:$D$<last>`, never whole-column `$D:$D`). `False` writes a
  values block.
- **`USE_LIVE_FORMULAS`** — `True` makes the Stage 6/7 percentiles, Weighted
  Score and Headroom live formulas over **bounded** population ranges. `False`
  writes computed value blocks.

**Parity guarantee.** The values path and the live-formula path produce the
same numbers. The VBA implements `PERCENTRANK.EXC` exactly (first-occurrence
`k/(N+1)` positioning) and `CEILING(x,1)`; `SelfTest` round-trips a fixture
population through **real Excel formulas** on a scratch sheet and asserts the
VBA buckets match cell-for-cell, for both the raw-is-better and rank-based
definitions.

---

## Weighted score & percentiles (definitions unchanged; ceiling 165)

- **Raw-is-better** columns — flattening, headroom, weighted-score percentiles:
  `CEILING(PERCENTRANK.EXC(pop, x) * 5, 1)`.
- **Rank-based** columns — the eight band percentiles (breadth + slope over four
  thresholds): `IFERROR(CEILING((1 - PERCENTRANK.EXC(pop, rank)) * 5, 1), 0)`
  (rank 1 = best; non-qualifiers score 0).
- **Weighted Score** `= Flattening + $B$2 · (sum of the eight band percentiles)`.
  A single live weight `$B$2` (default 4) multiplies **every** band — the
  threshold axis as a whole is worth 4× the flattening axis; the weight does
  **not** grade `$25 > $50 > $75 > $100`. If graded weights are wanted later,
  each band needs its own weight cell.
- **Composite maximum = `5 + 4·(8·5)` = 165.** Practical maximum today = **125**
  (nothing prices under \$25MM, so that band scores 0 for everyone).

### Standalone Headroom (not scored)

**Headroom (MW)** is the minimum overload-trigger size (`MINIFS` over the Cost
Data trigger column, keyed like the matrix rows), computed from the trigger
column directly so triggers below the smallest or above the largest sampled MW
are captured, not clipped. **Headroom Percentile (1–5)** is raw-is-better (more
headroom is better, no inversion). Reading it: a first trigger at/below the
smallest sampled MW (100) means effectively **no headroom** in the practical
range; above the largest (300) means headroom **exceeds the studied range**.
Headroom is a display-and-rank lens **only — never added to the Weighted
Score**, and the composite maximum stays 165.

---

## Self-test / verification (`SelfTest`)

- **Parity** — VBA buckets equal Excel `PERCENTRANK.EXC`/`CEILING` on a scratch
  sheet, for raw-is-better and rank-based columns.
- **Ranking** — sort-based competition ranks match the naive definition,
  including the ties (Cunningham/Hobbs, Pleasant Hill/Roosevelt, the 13-way
  \$100MM breadth tie) asserted on the 17-substation fixture.
- **Dedupe/join** — dictionary results match the triple rule and the
  intersection rule (name / name+voltage / bare; site-without-cost and
  cost-without-site excluded and logged).
- **Scale** — 2,000 synthetic substations are ranked and bucketed under a
  bounded wall-clock (mergesort, no O(N²) blow-up).
- **Durability / resume** — simulates Stage 1 then Stage 2 completing and
  asserts `Cost Data` is present, non-empty, and (when the host has a path)
  saved on disk; that a resumed run’s first incomplete stage is **2** with only
  Cost done, **3** with Cost + Site done, and **1** with nothing done — i.e.
  completed stages are skipped and no source files are re-opened. Runs only on a
  clean workbook so it never clobbers a real pipeline’s sheets.
- **Header guard** — a source block containing a data row, a duplicated header
  row, and a blank row keeps only the data rows (one header total).
- **Tab parsing, matrix shape, headroom, and the 165 ceiling.**

Results print to the Immediate window (**Ctrl+G**). `SelfTest` also prints a
reminder to run **Debug → Compile VBAProject** (the one check that can only be
done in the VBA editor).
