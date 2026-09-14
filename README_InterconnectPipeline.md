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

> **Re-import correctly every time you get a new `.bas`** (a clean workbook does
> not guarantee a clean module): in the VBA editor, right-click
> `modInterconnectPipeline` → **Remove** → *No* (don't export), then **File →
> Import File…** the new `.bas`. Importing without removing first creates a
> second module (`modInterconnectPipeline1`) and Alt+F8 may keep running the old
> one. Then **Debug → Compile VBAProject**, confirm exactly **one**
> `modInterconnectPipeline` in the Project Explorer, and run.
>
> 30-second check that the new code is running: the cost consolidator must
> **detect** the header row by matching the field names
> (`pl_DetectHeaderRow` / `pl_RowHasCostHeaders`), **not** a hardcoded
> `firstDataRow = 2`. If you still see a fixed row-1 header, the old build is
> loaded.

---

## The four-step flow (run order)

1. **Cost / results files → `Cost Data`.** Multi-select the cost workbooks
   (picker rooted at `ThisWorkbook.Path`, falling back to
   `Application.DefaultFilePath`). Sheet 3 carries the data. The substation name
   is **parsed from the (messy) tab name into clean `(name, voltage)`**, in this
   precedence: (a) strip a trailing standalone integer — a duplicate-file
   counter — so `Cecelia 138kV 1` and `Cecelia 138kV 2` both lose the `1`/`2`;
   (b) extract voltage tolerantly — a number, then optional punctuation/spaces,
   then `kV` (case-insensitive): `138kV`, `230 kV`, `230. kV` → `138`/`230`;
   (c) clean name = the remainder with nbsp/multiple spaces collapsed to one,
   trimmed. So `Cecelia 138kV 1` → (`Cecelia`, 138), `Chalkley 230. kV` →
   (`Chalkley`, 230), bare `Cunningham` → (`Cunningham`, blank). Voltage is
   recognised **only** with a `kV` marker; a bare trailing integer is always a
   counter. Name / voltage / state are stored as **separate fields** everywhere
   — the concatenated string is never used as identity.
   **The header row is detected, not assumed:** source sheets carry a *band* row
   above the real header (e.g. `Monitored Element` / `Worst Case Contingency` /
   `Cost Allocation`), so the tool scans the first rows and picks the header as
   the first row that *contains* the field names `Size Overload Occurs (MW)`
   **and** `Proposed Project Allocation ($)` (whitespace/case-normalized); data
   starts the row **after** it. The output header is written **once** (the
   metadata prefix `Substation Name | Voltage (kV) | Source File | Source Sheet`,
   then the detected source header verbatim); from each source only the **data
   rows** are appended, with blank rows and any header-signature row skipped.
2. **Site files → `Site Data`.** From each `Summary` tab (name = A, state = C,
   voltage = G, header on row 1, **data row 2 down**, a header-matching row
   skipped), de-duplicated on the full **(name, voltage, state) triple** built
   from the parsed fields — key `LCase(name)|voltageNumeric|LCase(state)`. With
   the counter stripped and voltage parsed, `Cecelia 138kV 1` and
   `Cecelia 138kV 2` (same state, 138 kV) produce the same key and collapse to
   one entry; a different *real* voltage or state stays separate. First
   occurrence kept, drops logged.
3. **Join → `Matrix`.** Intersection only, matched by name (plus voltage when
   the cost tab carried one). The `Matrix` identity is **three separate
   columns** — `Substation` (clean name), `Voltage (kV)`, `State` — followed by
   the MW × cost body (never a concatenated `Chalkley 230. kV (Louisiana)`
   label). The body is **live `SUMIFS` by default** (`USE_LIVE_SUMIFS = True`):
   each cell is a bounded `SUMIFS` over `Cost Data` — resolving
   `Proposed Project Allocation ($)`, `Substation Name`,
   `Size Overload Occurs (MW)` (and `Voltage (kV)` for multi-voltage) by
   normalized header name, matched to that row's parsed name (+ voltage), with
   MW taken from the header cell so the formula re-drives if the header changes.
   Ranges are bound to used rows, never whole columns.
4. **Analysis + ranking + weighted scoring + headroom → `Cost Curve Analysis`.**
   Before Stage 4 runs, the Matrix is validated against the input contract:
   ≥1 substation row, ≥3 numeric strictly-ascending MW columns, and every body
   cell numeric and **≥ 0**. A cost of **`$0` is valid** — it means no upgrade
   was priced at or under that MW, i.e. the size sits **below the substation's
   first trigger (headroom)**. Only a **negative** cost, or a non-numeric cell,
   is an error. The "can it run at all" guard still stops a genuinely empty or
   malformed matrix (0 rows, non-ascending / non-positive MW header) with a
   **specific message** (e.g. *“Stage 4 cannot run: Matrix has 0 valid
   substation rows”*) rather than a silent no-op. Any Stage-4 runtime error is
   written to `_Pipeline Log` (routine, `Err.Number`, `Err.Description`,
   offending substation/MW when known) **and** shown in a `MsgBox`; on success
   the analysis sheet's row count is logged.

   **Zeros are handled, not just permitted.** A leading run of `$0` cells forms
   its own `T = 0` segment (the segmentation walk treats `0→0` as one segment and
   `0→nonzero` as a boundary), so the fitted function reads e.g.
   `y = 0 / x [100-120 MW] | y = 6,500,000 / x [130-170 MW]`. Every computation
   that would divide by a cost or take `Log(cost)` is guarded against zero — the
   Kneedle knee falls back to the closed-form geometric mean on an all-zero
   (flat) segment, and the power-law fallback fit skips zero-cost points — so a
   zero cell yields a **defined** result instead of `#DIV/0!` / overflow.
   Zero-cost substations are **kept** in the ranking and percentiles; this lines
   up with **Headroom (MW)**, which equals the first trigger MW.

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

### Two switches

- **`USE_LIVE_SUMIFS`** — **default `True`**: the Matrix body is live `SUMIFS`
  bound to the used rows (`$D$2:$D$<last>`, never whole-column `$D:$D`). `False`
  writes a values block.
- **`USE_LIVE_FORMULAS`** — default `False` (computed value blocks): `True`
  makes the Stage 6/7 percentiles, Weighted Score and Headroom live formulas
  over **bounded** population ranges.

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
- **Header detection** — a scratch sheet with a band row above the field-name
  header asserts `pl_DetectHeaderRow` picks the field-name row (row 2), not the
  band row (row 1), with a whitespace/case-normalized match.
- **Tab parsing** — `Cecelia 138kV 1` → (`Cecelia`, 138), `Cecelia 138kV 2` →
  (`Cecelia`, 138), `Chalkley 230. kV` → (`Chalkley`, 230), `Cunningham` →
  (`Cunningham`, blank); the two Cecelia tabs share one cost-identity key.
- **Matrix** — MW axis 100…300; body is a bounded `SUMIFS` formula
  (`USE_LIVE_SUMIFS` defaults True); `pl_Stage4MatrixValid` passes a good matrix
  and stops an emptied one with the explicit *“0 valid substation rows”* reason.
- **Matrix shape, headroom, and the 165 ceiling.**

Results print to the Immediate window (**Ctrl+G**). `SelfTest` also prints a
reminder to run **Debug → Compile VBAProject** (the one check that can only be
done in the VBA editor).
