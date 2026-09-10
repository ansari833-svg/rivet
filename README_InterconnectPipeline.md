# Interconnection Site-Scoring Pipeline (`modInterconnectPipeline.bas`)

One importable Excel VBA module with a **single public entry point,
`RunInterconnectPipeline`**, that runs the whole workflow end to end from a
blank workbook on one **Alt+F8** call. A second public procedure, `SelfTest`,
is an in-memory test harness (no file pickers, no sheets touched) and is the
only other entry point.

Import it via the VBA editor (**Alt+F11 → File → Import File…**) and run it
with **Alt+F8 → `RunInterconnectPipeline`**. It is a plain `.bas` — never a
`.bat`, `.vbs`, or runner workbook. `Option Explicit`, private helpers are
`pl_`-prefixed, no external references, no `.Select` / `.Activate` /
`Selection`. The four application state flags (`ScreenUpdating`,
`EnableEvents`, `DisplayAlerts`, `Calculation`) are saved and restored on every
exit path, including errors, and one source file's failure never aborts the run.

---

## The four-step flow (run order)

### Step 1 — Consolidate the cost / results files → `Cost Data`
- You multi-select the cost source workbooks (picker rooted at
  `ThisWorkbook.Path`, falling back to `Application.DefaultFilePath` for an
  unsaved blank workbook).
- Each source workbook holds three sheets; **sheet 3 carries the data**, header
  on row 1.
- The **substation name comes from sheet 3's tab name**, not a cell:
  - A tab that carries a voltage (multi-voltage substation, e.g.
    `Chaves County 345`) is parsed into name + voltage — the **last** plausible
    numeric token is the voltage, the cleaned remainder is the name.
  - A bare tab (single-voltage substation, e.g. `Cunningham`) becomes the name
    with a blank voltage.
- Each sheet's used range is block-read (values only), blank rows skipped, and
  appended to `Cost Data` behind four metadata columns the tool creates:
  **`Substation Name` | `Voltage (kV)` | `Source File` | `Source Sheet`**, then
  the source columns **verbatim**. The two that matter downstream —
  `Size Overload Occurs (MW)` (tier trigger) and
  `Proposed Project Allocation ($)` (allocation) — are carried through
  unchanged so they can be resolved later by header name.
- Header signatures are validated across files (case-insensitive); mismatches
  are surfaced with an include / skip / abort choice, and every file's status
  is logged.

### Step 2 — Consolidate the dedupe / site files → `Site Data`
- You multi-select the site source workbooks (same picker behaviour).
- The relevant sheet is the **`Summary`** tab: **name = column A, state =
  column C, voltage = column G**, header on row 1, data below.
- Those three fields per row are read into `Site Data` under headers
  **`Substation Name` | `State` | `Voltage (kV)`**.
- Rows are **de-duplicated on the full (name, voltage, state) triple**: two rows
  collapse only when name **and** voltage **and** state are all equal. A
  different voltage is a different substation; a different state is a different
  substation. **Dedupe is on the triple, never on the name alone.** The first
  occurrence of each distinct triple is kept; every dropped duplicate is logged
  (file, row, triple) and the distinct-triple count is reported.

### Step 3 — Join and build `Matrix` (live `SUMIFS`)
- One matrix row per **`Site Data` triple**. Each triple's cost curve is
  attached from `Cost Data` by **matching on name**; when the cost tab carried
  a voltage (the multi-voltage case) the match is on **name + voltage** to pick
  the correct one of several same-named cost tabs; when the cost tab was bare,
  name alone matches.
- **Intersection only.** The sets are expected to align exactly, so a `Site
  Data` triple with no cost match, or a cost tab with no site match, is **logged
  as an alignment error** rather than silently dropped; the run then proceeds
  with the intersection.
- Layout: **row 1** is a leading label then the MW axis header
  **100, 110, … 300** (numeric, strictly ascending); **column A** from row 2
  down is the substation identity label (name + voltage + state, so
  multi-voltage and multi-state entries are distinct rows); the **body** is
  cost-per-MW as one **live `SUMIFS`** per (substation, MW):

  ```
  =SUMIFS('Cost Data'!<AllocCol>, 'Cost Data'!<NameCol>, <thisSubstation>,
          'Cost Data'!<TriggerCol>, "<=" & <thisMW>) / <thisMW>
  ```

  `<AllocCol>`, `<NameCol>`, `<TriggerCol>` (and `<VoltCol>` for the
  multi-voltage criterion) are **resolved at run time by header name** on the
  `Cost Data` sheet (full-column references) — **never by fixed letter**,
  because the four metadata columns shift the source columns right. `<thisMW>`
  references the MW header cell, so editing a header re-drives the row. The
  metadata (state, voltage) stays on `Site Data`; the matrix body is purely the
  MW × cost block the analysis consumes. A full recompute is forced before the
  analysis reads the values.

### Step 4 — Analysis, ranking, weighted scoring → `Cost Curve Analysis`
The proven cost-curve analysis, per-threshold ranking and chart/leaderboard
logic (from `modCostCurves`) is reused **unchanged**, driven off the `Matrix`
ranges instead of interactive prompts, and writing the analysis table, chart,
threshold leaderboard and a self-contained data copy to `Cost Curve Analysis`.
Two points apply on top:

- **(a) Column resolution.** Every reference into `Cost Data` (the matrix
  `SUMIFS` and the headroom `MINIFS` below) resolves by header name, never by a
  fixed column letter.
- **(b) Weighted composite score.** See below.

---

## Weighted composite score

A single live weight cell **`$B$2`** (default **4**) multiplies **every**
threshold band. All four bands share this one multiplier — **the threshold axis
as a whole is worth 4× the flattening axis; the weight does not grade
$25 > $50 > $75 > $100.** If graded weights are wanted later, each band needs
its own weight cell.

| Axis | Contribution |
|------|--------------|
| Flattening percentile | max **5** |
| Each threshold band | `$B$2 × (breadth pctile + slope pctile)` = `4 × (5 + 5)` = **40** |
| Four bands | **160** |
| **Composite maximum** | **5 + 160 = 165** |

> Earlier drafts said 105; that was a different sketch. **The correct ceiling
> for these formulas is 165.** The **practical maximum today is 125**, because
> nothing prices under \$25MM on the current data (cheapest total ≈ \$47.6MM),
> so the \$25MM band scores 0 for every substation regardless of the multiplier.

Percentiles (0–5) are computed in VBA; the band-score and composite cells are
**formulas referencing `$B$2`**, so re-weighting in the cell is live.

---

## Standalone Headroom columns (display-and-rank only — **not scored**)

Two columns are appended at the end of the scoring block. **Headroom is a
standalone lens: it is not added to the Weighted Score, and the composite
maximum stays 165.**

**Headroom (MW)** — the project size below which no upgrade is triggered, i.e.
the **minimum** overload-trigger size for that substation (a substation has
several triggers; the smallest is the binding constraint). It is computed from
the `Cost Data` trigger column directly — **not** from the matrix cells — so
triggers below the smallest sampled MW or above the largest are captured
accurately rather than clipped to the sampled range:

```
=MINIFS('Cost Data'!<TriggerCol>, 'Cost Data'!<NameCol>, <thisSubstation>
        [, 'Cost Data'!<VoltageCol>, <thisVoltage>  when multi-voltage])
```

Resolved by header name (`Size Overload Occurs (MW)`, `Substation Name`,
`Voltage (kV)`), matched by the **same key the matrix rows use** (name, plus
voltage when the cost tab carried one). Number format `#,##0`. A substation with
no trigger records (shouldn't occur after the intersection join, but guarded)
yields `MINIFS = 0` and is shown blank via `IFERROR`/`IF`, dropping it out of
the percentile population.

**Headroom Percentile (1–5)** — more headroom is better, so there is **no
inversion** (same shape as the flattening percentile), over a fully absolute
population across all data rows:

```
=CEILING(PERCENTRANK.EXC($<HeadroomCol>$<first>:$<HeadroomCol>$<last>, <thisHeadroomCell>)*5, 1)
```

Centred, format `0`. Reading the value:
- A first trigger **at or below** the smallest sampled MW (100) means
  effectively **no headroom** in the practical range.
- A first trigger **above** the largest sampled MW (300) means headroom
  **exceeds the studied range**.

---

## Sheets produced (all in `ThisWorkbook`, created if absent)

`Cost Data`, `Site Data`, `Matrix`, `Cost Curve Analysis`, and a run log
**`_Pipeline Log`** capturing per-file status, dropped duplicate triples and any
join alignment errors. For the four data/output sheets, if one already exists
you are offered **overwrite / new-timestamped / cancel** (the downstream steps
follow the actual sheet returned, so a timestamped choice does not break the
chain). The log is rewritten each run.

---

## Self-test

`SelfTest` runs the full analysis logic in memory against the 17-substation ×
21-MW reference fixture (or `substation_cost_per_mw.csv` in the workbook folder
if present) and, in the same pass, the new front-half and scoring assertions
(no file pickers):

- **Tab parsing** — `Chaves County 345` → name `Chaves County`, voltage 345;
  `Cunningham` → name `Cunningham`, voltage blank.
- **Triple dedupe** — five rows where two are identical on (name, voltage,
  state) collapse to three; rows differing only in voltage, or only in state,
  stay separate; drops are logged.
- **Join intersection** — a site triple with no cost match and a cost tab with
  no site match are both logged as alignment errors and excluded; a bare cost
  tab matches on name; the matched set builds correctly.
- **Matrix shape** — MW axis is 100…300 step 10 (21 points, strictly
  ascending), column A the identities, body cells `SUMIFS` that resolve
  trigger/allocation by header name and divide by the MW cell (with a voltage
  criterion for multi-voltage tabs).
- **Scoring ceiling** — composite maximum is 165, and all 17 `$25MM` band
  percentiles are 0 on the fixture.
- **Headroom** — triggers 130/200/260 report headroom 130; largest headroom →
  percentile 5, smallest → 1; and, because headroom is **not** summed into the
  composite, the ceiling stays 165.

Results print to the Immediate window (**Ctrl+G**).
