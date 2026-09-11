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
   rest = name; bare tab = name only). Consolidated behind the metadata prefix
   `Substation Name | Voltage (kV) | Source File | Source Sheet`, then the
   source columns verbatim — including `Size Overload Occurs (MW)` (trigger) and
   `Proposed Project Allocation ($)` (allocation).
2. **Site files → `Site Data`.** From each `Summary` tab (name = A, state = C,
   voltage = G), de-duplicated on the full **(name, voltage, state) triple**
   (a different voltage or state is a different substation).
3. **Join → `Matrix`.** Intersection only, matched by name (plus voltage when
   the cost tab carried one). Cost-per-MW per (substation, MW).
4. **Analysis + ranking + weighted scoring + headroom → `Cost Curve Analysis`.**

Also produced: a run log **`_Pipeline Log`** (per-file status, dropped
duplicate triples, join alignment errors). Existing data/output sheets prompt
**overwrite / new-timestamped / cancel**; the log is rewritten each run.

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
- **Tab parsing, matrix shape, headroom, and the 165 ceiling.**

Results print to the Immediate window (**Ctrl+G**).
