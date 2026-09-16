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
   `Application.DefaultFilePath`). The data sheet is **located by content** — the
   worksheet whose leading rows carry the known headers — rather than assuming a
   fixed tab index, so files where the data isn't the 3rd tab are still read. The
   substation name is **parsed from the (messy) tab name into clean `(name, voltage)`**, in this
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
   occurrence kept, drops logged. **This dedupe serves consolidation and the
   State attachment only — it never decides which substations are analyzed.**
3. **Matrix (all Cost Data substations) → `Matrix`.** Matrix rows are **every
   distinct substation in `Cost Data` column A** (`Substation Name`, plus
   `Voltage (kV)` where the cost tab carried one) — taken as-is from the cost
   consolidation, which already cleans and de-duplicates them. **This is not an
   intersection with the site data**: a substation present in `Cost Data` (e.g.
   `Cecelia`) always gets a full row, even with no matching site entry.
   `State` is a **LEFT JOIN** attached from `Site Data`: filled where the
   substation matches (by name, plus voltage when carried), **left blank when
   there is no match — never a reason to drop the row**.

   **One shared name + voltage normalization on both sides.** The match key runs
   both the cost-side name and the site-side `Summary!A` name through the same
   two shared routines (via `pl_CostKey`), so the sides can never drift:
   - **`pl_NormSubName`** (name-clean): remove the embedded voltage token
     (`138kV` / `230 kV` / `230. kV`); strip a trailing standalone integer that
     **followed** the voltage — a file counter, so `Cecelia 138kV 1` and
     `Cecelia 138kV 2` both become `cecelia` and collapse together — while
     **keeping** a unit number that preceded the voltage (the `1` in
     `Big Cajun 1 230kV` → `big cajun 1`, and the already-clean cost name
     `Big Cajun 1` → `big cajun 1`, so they still match); collapse nbsp/spaces,
     trim, lower-case.
   - **`pl_VoltFrom` + `pl_VoltStr`** (voltage): extract the 2–4 digit integer
     **immediately adjacent** to `kV` (only spaces / one period between), so
     `Explorer Claremore 138kV` → `138` (never a stray `168`), and render it as a
     **clean integer** — the key reads `…|138`, never `…|138.`.

   So `Cecelia 138kV 1` (site) keys to `cecelia|138` and matches the clean cost
   `Cecelia` @ 138. Before these fixes the site keys were malformed
   (`cecelia 1|138.` — counter kept, voltage with a trailing dot) and ~299
   genuine pairs reported *no match*; **after, unmatched drops below ~100.**
   Stage 3 logs the reconciliation — *site triples, cost substations, matched,
   unmatched* — and every remaining unmatched site row with **both its raw name
   and its normalized key**, so true residuals are diagnosable. Site rows with no
   Cost Data match are informational only. The `Matrix` identity is
   **three separate columns** — `Substation` (clean name), `Voltage (kV)`,
   `State` — followed by the MW × cost body (never a concatenated
   `Chalkley 230. kV (Louisiana)` label). The body is **live `SUMIFS` by
   default** (`USE_LIVE_SUMIFS = True`): each cell is a bounded `SUMIFS` over
   `Cost Data` — resolving `Proposed Project Allocation ($)`, `Substation Name`,
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

   **Column order (identity + headline first).** `Cost Curve Analysis` is one
   contiguous table whose identity is **three separate columns** —
   `Substation` (clean name only), `Voltage (kV)`, `State` — never concatenated.
   The headline results sit next to the name, then the per-threshold ranks, then
   the detailed blocks, left to right:
   `Substation | Voltage (kV) | State | Weighted Score | Weighted Score Pctile |
   Headroom (MW) | Headroom Pctile | Flattening Point | Flattening Pctile |
   [Rank by MW Breadth / Rank by Slope per threshold] | Fitted Function |
   Segments | Step-Change Points | Slope at Knee | Marginal Slowdown |
   [per threshold: MW Range, Slope over Range, Breadth (0-5), Slope (0-5), Band
   Score]`. Every live-formula reference (`Weighted Score = Flattening pctile +
   Σ band scores`, each band score `= $B$2·(breadth pctile + slope pctile)`, and
   every `PERCENTRANK.EXC` population range) is **derived from a single
   column-index map**, so a column move relocates its references automatically —
   the math and the 165 ceiling are unchanged, only positions.

   **No chart.** Stage 4 no longer builds a chart. Excel caps a chart at 256
   series and ~1,691 substations overflowed it (*“A chart can only have up to
   256 series”*), so the chart routine and its helpers, the Source-Data / helper
   copy blocks it read, and its progress phase were removed. Stage 4's outputs
   are the `Cost Curve Analysis` sheet (table + threshold leaderboard) and the
   reconciliation.

Also produced: a run log **`_Pipeline Log`** (per-file status, dropped
duplicate triples, join alignment errors), a **`_Pipeline State`** checkpoint
sheet, and a durable **`_Skipped Files`** record — the log is regenerable from
the sheets via `RebuildAudit` (see below).

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

## Regenerable audit (`RebuildAudit`)

`_Pipeline Log` is a **run transcript** — deleting the tab must not force an
(expensive) re-run to recover the diagnostics. Everything the audit needs is
**persisted outside the log**, so the summary can be recomputed from the output
sheets on demand:

- **Per-row provenance is retained.** `Cost Data` keeps a **`Source File`** and
  **`Source Sheet`** column on every row (columns C/D, ahead of the resolved
  cost columns). With `Source File` present, the mapping of files → substations,
  the collapse counts (which substations came from >1 file, e.g. the
  `Cecelia …1` / `…2` merges), and the distinct total are all **derivable from
  the sheet at any time** — no run history required. (The clean sheet already
  carries this, so no separate raw sheet is needed.)
- **Skips/fails are persisted durably.** A file that was unreadable, lacked its
  data sheet, had no header, or an empty `Summary` leaves **no rows** in
  `Cost Data`, so its outcome can’t be reconstructed from the output sheets.
  These are written to a dedicated **`_Skipped Files`** sheet
  (`Stage | File | Sheet | Outcome | Intended Substation | Loss / Coverage | Reason`)
  during Stages 1–2 — not only to the deletable `_Pipeline Log` — so a log
  deletion never erases them.

### Every selected file is accounted for

The reconciliation is **anchored on the count of files the user selected**, never
on the contributing count alone:

```
selected = contributed + skipped/failed
e.g. 2090 selected = 1836 contributed + 254 not contributed
```

The selected count is captured at pick time (`pl_PickFiles`), carried through, and
persisted on `_Pipeline State` so the anchor survives a log wipe. If the parts
don’t sum to `selected`, that is itself a bug — the code asserts it and logs any
**unaccounted** remainder.

**Every file gets an explicit outcome**, logged per file and (for non-contributors)
itemized in `_Skipped Files`:

- `Contributed` — *n* data rows read.
- `Skipped: sheet not found` — no worksheet at all (content search found none).
- `Skipped: header not found` — no worksheet carried `Size Overload Occurs (MW)`
  / `Proposed Project Allocation ($)`; **the headers actually seen are logged**, so
  a wording mismatch is visible.
- `Skipped: no data rows` — header found but nothing below it.
- `Skipped: empty substation name` — the tab-name parse yielded no name.
- `Skipped: header signature mismatch` — a valid sheet whose column layout differs
  from the first valid file and the user chose to skip it.
- `Failed: open error` — couldn’t open (locked/corrupt/format), with `Err.Description`.

**LOST vs redundant.** A skip is real substation loss only if **no other file**
supplied that substation. For each skipped file the intended substation is derived
from its tab name; if that substation is **absent** from `Cost Data` the row is
flagged **`LOST — only source for this substation`**, otherwise
**`redundant skip (substation covered)`**; an unreadable tab name is
**`unknown`**. The Stage-1 summary counts the split — *of the 254, how many are
LOST vs redundant vs unreadable* — so a dominant root cause is obvious. The likely
culprits are surfaced by design: the content-based **sheet** location recovers
files where the data isn’t the 3rd tab, the content-based **header** match absorbs
wording variants, and the logged “headers seen” exposes any different-template
files as a group.

**`Public Sub RebuildAudit()`** regenerates the diagnostic summary from those
sheets alone, callable any time (including after `_Pipeline Log` is deleted),
**without re-running consolidation**. It recreates `_Pipeline Log` (fresh) and
writes the file-level reconciliation anchored on the persisted selected count
(**selected = contributed + skipped/failed**, flagging any unaccounted remainder),
the distinct-substation total, and the files-per-substation collapse counts (with a
few examples). That is the “why are there 2,090 files but ~1,600 substations”
answer — and now also the “where did the other 254 go” answer — recovered with no
re-run.

**Optional external log copy.** With the constant `WRITE_LOG_TO_FILE = True`,
`pl_WriteLog` also writes the run log to `<workbook base>_PipelineLog.txt` beside
the saved workbook, so a run’s transcript survives independent of any tab
deletion. Best-effort: an unwritable path is ignored, never fatal.

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
end** (no intermediate recalcs).

| Stage | Before (bottleneck) | After |
|-------|---------------------|-------|
| 1 Consolidation | per-cell reads/writes | every workbook opened `UpdateLinks:=0, ReadOnly:=True, AddToMru:=False` and closed immediately; used range read in one `Range.Value`; output written one `Range.Value = array` per file; one `DoEvents` + status line per file; a locked file is logged and skipped, never aborting the batch |
| 2 Dedupe | O(N²) pairwise compare | `Scripting.Dictionary` keyed on `LCase(name)|voltage|LCase(state)` — one O(N) pass |
| 3 Matrix membership + State join | rescans Cost Data per substation | one-pass dictionaries (identity key → id = all distinct Cost Data subs; name → records); `State` attached by O(1) left-join lookup, blank when unmatched |
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
headroom is better, no inversion). **Headroom = 0 (or no positive trigger) scores
percentile 0** — the worst case, excluded from the ranked population exactly like
a threshold non-qualifier: only strictly-positive headrooms are ranked 1–5 among
themselves, and a zero never lands in a 1–5 bucket and is never blank. Both paths
enforce this — the values path uses the masked bucket helper
(`pl_RawBucketArrayMasked`, which returns 0 for the masked-out zeros), and the
`USE_LIVE_FORMULAS` path writes the headroom blank (so `PERCENTRANK.EXC` excludes
it) and falls the percentile to 0 via `IFERROR(…,0)`. Reading it: a first trigger
at/below the smallest sampled MW means effectively **no headroom** in the
practical range; above the largest means headroom **exceeds the studied range**.
Headroom is a display-and-rank lens **only — never added to the Weighted
Score**, and the composite maximum stays 165.

---

## Self-test / verification (`SelfTest`)

- **Parity** — VBA buckets equal Excel `PERCENTRANK.EXC`/`CEILING` on a scratch
  sheet, for raw-is-better and rank-based columns.
- **Ranking** — sort-based competition ranks match the naive definition,
  including the ties (Cunningham/Hobbs, Pleasant Hill/Roosevelt, the 13-way
  \$100MM breadth tie) asserted on the 17-substation fixture.
- **Dedupe** — dictionary results match the triple rule (name / name+voltage /
  bare); the (name, voltage, state) triple collapses correctly.
- **Matrix membership + State left join** — a scratch `Cost Data` (with `Cecelia`
  spread over two rows and a `Marlin` row) and a `Site Data` that has `Cecelia`
  (Kentucky) but **no `Marlin`** assert: every distinct Cost Data substation
  appears exactly once (Cecelia + Marlin), `Cecelia` gets its `State` from the
  left join, `Marlin` still appears with a **blank `State`** (never dropped), the
  identity is three separate columns (not concatenated), the body cells are live
  `SUMIFS`, and Cecelia's curve sums **both** source rows above its second
  trigger.
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
- **Rebuild audit** — a scratch `Cost Data` (with `Cecelia` consolidated from
  two files and a `Marlin` from one) and a `_Skipped Files` record with one
  failed file: with `_Pipeline Log` gone, `RebuildAudit`’s computation
  regenerates the summary from the sheets alone (2 distinct substations, 3
  contributing files, 1 collapsed, 1 skipped/failed) and writes the
  reconciliation *files seen (4) = distinct (2) + collapsed (1) + skipped (1)* —
  with **no consolidation re-run**.
- **File accounting** — a mixed synthetic fixture of 5 files (two good, one
  missing-sheet, one bad-header for a unique substation, one empty duplicate of a
  good one) asserts `selected (5) = contributed (2) + skipped (3)` exactly; that
  `_Skipped Files` itemizes all 3 non-contributors; and that the bad-header file
  is flagged **LOST** (its substation appears nowhere else) while the empty
  duplicate is flagged **redundant** (its substation is covered by another file).
- **Site↔Cost name matching** — the shared routines strip the trailing counter
  (`Cecelia 138kV 1` and `…2` both → `cecelia`, and both `pl_CostKey` to
  `cecelia|138`, matching the clean cost `Cecelia` @ 138) while keeping a unit
  number (`Big Cajun 1 230kV` → `big cajun 1`); the voltage extracts the digits
  adjacent to `kV` (`Explorer Claremore 138kV` → `138`, not `168`) and renders a
  clean integer (`grimes|138`, never a trailing `.`); and every named bug-report
  pair (`Big Cajun 1`, `Ponderosa`, `Cincinnati`, `Mockingbird`, `Grimes`)
  produces the same key on both sides.
- **Headroom zero** — a 5-substation population (`50, 0, 200, 300, 0`) asserts the
  two zeros score percentile **0** (masked out), the three positives rank **1–5**
  among themselves (`300 ≥ 200 ≥ 50`), and only those three enter the ranked
  population.
- **Analysis layout** — traces the `nt = 4` column map: identity at `A/B/C`,
  Weighted Score/pctile/Headroom at `D/E/F`, Flattening pctile at `I`, and the
  Weighted Score formula resolves to `=I4+AA4+AF4+AK4+AP4` with the first band
  score `=$B$2*(Y4+Z4)` — confirming references land on the reordered columns.
- **Matrix shape, headroom, and the 165 ceiling.**

Results print to the Immediate window (**Ctrl+G**). `SelfTest` also prints a
reminder to run **Debug → Compile VBAProject** (the one check that can only be
done in the VBA editor).
