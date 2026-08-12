# Interconnection Cost Curve Analysis (`modCostCurves.bas`)

An Excel VBA module that reads a substation interconnection cost-per-MW matrix
from a worksheet, discovers the piecewise tier structure hidden in the data, and
writes a formatted cost-curve chart plus a per-substation analytical summary onto
a fresh, self-contained output tab.

Public entry point: **`Sub BuildCostCurveAnalysis()`**.

---

## How the fit works (read this first)

The fitted function this tool reports is **exact, not a regression.** For each
substation `i` and MW point `j` it computes the *implied total cost*
`T(i,j) = Cost(i,j) * MW(j)`. In real interconnection data a substation carries a
fixed network-upgrade cost amortized across project size, so cost-per-MW traces
`y = T / x` **exactly** (R² = 1 by construction) until a larger project trips a new
upgrade tier and `T` steps up to a new constant. `T` is therefore
*piecewise-constant*, and the module fits each segment as `y = Tₖ / x` rather than
running any power-law or polynomial regression.

Two consequences shape the output:

- **The threshold test resolves at segment boundaries, never inside a segment.**
  Because `T` is constant across a segment, either the whole segment satisfies
  `T ≤ threshold` or none of it does. Qualifying MW ranges are reported at segment
  boundaries and are never interpolated to a fractional crossing MW.
- **Column E is a geometric *knee*, not a mathematical inflection point.** The
  curve `y = T/x` is convex everywhere (`y'' = 2T/x³ > 0`), so it has no true
  inflection. The module instead reports the Kneedle maximum-vertical-distance
  knee, which for a pure hyperbola on `[x₁, x₂]` has the closed form
  `x* = √(x₁·x₂)`, the geometric mean of the segment endpoints. It is labelled
  honestly as a knee / flattening point, not as an inflection.

(A global power-law fit `y = a·x^b` is used only as a fallback when segmentation
produces more than `MAX_SEGMENTS` segments, i.e. when the data is not behaving as
a tier structure.)

---

## Importing the module

1. Open the target workbook and press **Alt+F11** to open the VBA editor.
2. **File → Import File…** and select `modCostCurves.bas`.
3. Back in Excel, run the macro from **Developer → Macros → `BuildCostCurveAnalysis`**,
   or press **Alt+F8** and choose it.

The module is fully self-contained and imports into any workbook — it does not
depend on specific sheet names.

---

## Using it: the two selection prompts

When you run `BuildCostCurveAnalysis` you are prompted to select two ranges with
the mouse:

1. **Substation name list** — a single column, one substation name per row, no
   header.
2. **Cost matrix** — a rectangular block whose **first row is the MW header** and
   whose remaining rows are one substation's cost-per-MW curve each, in dollars
   per MW. Row *i* of the name list corresponds to data row *i* of the matrix, in
   the same order.

The matrix must have `n + 1` rows (1 header + `n` data) and at least 3 MW columns.
The macro validates shapes and cell contents up front and stops with a clear
message if the MW header is non-numeric/non-positive/non-increasing, if any data
cell is blank/non-numeric/non-positive, or if the row counts don't line up.

The result is written to a **new tab** named *Cost Curve Analysis* (with ` (2)`,
` (3)`… appended if that name is taken — existing sheets are never overwritten),
inserted immediately after the source sheet.

---

## Output columns

**Block A — analysis table** (one row per substation):

| Col | Header | Meaning |
|-----|--------|---------|
| A | Substation | Name |
| B | Fitted Function | Piecewise `y = T / x` formula string, one term per segment, joined by ` \| ` |
| C | Segments | Number of tier segments detected |
| D | Step-Change Points (MW) | MW values where `T` steps to a new tier, or `None` |
| E | Knee / Flattening Point (MW) | Geometric knee `√(x₁·x₂)` on the longest segment, to one decimal |
| F | Slope at Knee ($/MW per MW) | Analytic slope `−T / x*²` at the knee |
| G | Marginal Slowdown Point (MW) | First MW where the slope magnitude drops below `SLOWDOWN_TOL` (`√(T/SLOWDOWN_TOL)`), or `Beyond range` |

Then **one pair of columns per threshold** (H onward, ascending threshold order):

| Col | Header | Meaning |
|-----|--------|---------|
| H | MW Range ≤ $25MM | MW range(s) where `T ≤ threshold` — e.g. `100-280 MW`, `100-170, 210-240 MW`, `None`, or `All (100-300 MW)` |
| I | Slope over Range ($/MW per MW) | OLS slope of cost vs MW over the qualifying points (`n/a` / `n/a (single point)` when too few qualify) |
| … | (repeats for $50MM, $75MM, $100MM) | |

> Note on column E vs G: `x*` depends only on the segment endpoints, not on `T`,
> so column E separates substations by *where their segments break*, not by cost
> level. Column G is the one that reflects cost magnitude. The threshold slope in
> the I/K/M/O columns is a *secant-style average* across the qualifying window;
> because the curve is convex it understates steepness at the low-MW end and
> overstates it at the high-MW end.

**Block B — chart:** an XY scatter-with-lines chart (`xlXYScatterLines`, so MW sits
on a true numeric axis). One line per substation, plus one dashed gray→black
threshold reference curve per threshold drawn on top.

**Block C — Source Data:** a copy of the MW header and full name+cost matrix. The
chart series reference **this copied block**, not the original sheet, so the output
tab is portable on its own.

**Block D — Threshold Curve Helper:** the `threshold / MW` reference curves that
feed the threshold chart series. Shown in light-gray font (not hidden — hidden
cells can drop out of chart series).

---

## Constants you can edit

All tunable values sit in a single block at the top of the module:

| Constant | Default | Purpose |
|----------|---------|---------|
| `OUT_SHEET` | `"Cost Curve Analysis"` | Base name of the output tab |
| `SEG_TOL` | `0.005` | Relative `T` tolerance for treating two points as the same segment (absorbs cent-level rounding) |
| `MAX_SEGMENTS` | `8` | Above this many segments, fall back to a global power-law fit |
| `SLOWDOWN_TOL` | `1500` | Slope magnitude ($/MW per MW) defining the marginal-slowdown point |
| `CHART_W`, `CHART_H` | `1100`, `620` | Chart dimensions in points |
| `TITLE_ROW`, `TABLE_HDR_ROW`, `CHART_GAP_ROWS`, `CHART_ROWS`, `COLB_WIDTH` | — | Layout anchors, adjustable in one place |

**To change the threshold set** (the ≤ $25MM / $50MM / $75MM / $100MM columns and
the corresponding chart curves), edit the single function:

```vba
Private Function GetThresholds() As Variant
    GetThresholds = Array(25000000#, 50000000#, 75000000#, 100000000#)
End Function
```

The threshold column block, the chart threshold series, and the legend all read
from this one function — adding or removing a threshold requires no other change.

---

## Self-test

`Sub SelfTest()` runs the full analysis logic in memory (no Excel ranges or charts
touched) against a 17-substation × 21-MW reference fixture and prints each
pass/fail to the Immediate window (**Ctrl+G** in the VBA editor). If
`substation_cost_per_mw.csv` is present in the workbook folder it is loaded;
otherwise a hardcoded fixture reproducing the reference cases (Cunningham & Hobbs
single-tier `T = 47,633,000`; Chaves County five-tier) is used.
