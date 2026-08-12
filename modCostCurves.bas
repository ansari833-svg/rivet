Attribute VB_Name = "modCostCurves"
Option Explicit

' ============================================================
'  modCostCurves.bas
'  Interconnection Cost Curve Analysis
'
'  Public entry point:  BuildCostCurveAnalysis
'
'  Reads a substation interconnection cost-per-MW matrix from a
'  worksheet, discovers the piecewise tier structure hidden in the
'  data, and writes a formatted cost-curve chart plus a per-substation
'  analytical summary onto a fresh, self-contained output tab.
'
'  --- Why the fit is exact, not a regression --------------------
'  For substation i at MW point j define the IMPLIED TOTAL cost
'
'        T(i,j) = Cost(i,j) * MW(j)
'
'  A substation carries a FIXED network-upgrade cost that is
'  amortised across project size, so cost-per-MW traces  y = T / x
'  EXACTLY (R^2 = 1 by construction) until a larger project trips a
'  new upgrade tier and T steps up to a new constant. T is therefore
'  piecewise-constant. Consequences the code exploits:
'
'    1. Each segment is  y = T_k / x  exactly. No power-law or
'       polynomial regression is used as the primary fit.
'    2. The threshold test is binary per segment: since T is constant
'       across a segment, either the whole segment satisfies
'       T <= L or none of it does. Crossings are reported at segment
'       boundaries, never interpolated inside a segment.
'    3. The curve is convex everywhere ( y'' = 2T / x^3 > 0 ), so it
'       has NO true inflection point. Column E reports a geometric
'       KNEE, honestly labelled, not an inflection.
' ============================================================

' ---------- Tunable constants (edit here only) --------------

Private Const OUT_SHEET       As String = "Cost Curve Analysis"

Private Const SEG_TOL         As Double = 0.005   ' relative T tolerance for "same segment"
Private Const MAX_SEGMENTS    As Long = 8         ' above this, fall back to a power-law fit
Private Const SLOWDOWN_TOL    As Double = 1500#   ' $/MW per MW: marginal-slowdown slope threshold

' Layout anchors (change these to move the whole layout)
Private Const TITLE_ROW       As Long = 1
Private Const TABLE_HDR_ROW   As Long = 3         ' Block A headers; data begins on the next row
Private Const N_FIXED_COLS    As Long = 7         ' cols A..G before the per-threshold pairs
Private Const CHART_GAP_ROWS  As Long = 2         ' blank rows between table and chart
Private Const CHART_ROWS      As Long = 44        ' vertical rows reserved for the chart band
Private Const COLB_WIDTH      As Double = 60      ' fixed width of the Fitted Function column

' Chart dimensions, in points
Private Const CHART_W         As Double = 1100
Private Const CHART_H         As Double = 620

' Number formats
Private Const FMT_DOLLAR      As String = "$#,##0"
Private Const FMT_SLOPE       As String = "$#,##0.0"
Private Const FMT_MW          As String = "#,##0"
Private Const FMT_MW1         As String = "#,##0.0"

' ---------- Segment record ----------------------------------

Private Type TSegment
    FirstJ As Long      ' 1-based index of the segment's first MW point
    LastJ  As Long      ' 1-based index of the segment's last  MW point
    Ttotal As Double    ' representative total = mean of T over the segment, rounded to $1,000
End Type

' Module-level stage tag, surfaced by the error handler
Private mStage As String

' ============================================================
'  PUBLIC ENTRY POINT
' ============================================================

Public Sub BuildCostCurveAnalysis()

    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc   As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ---- Stage 1: capture and validate the two inputs --------
    mStage = "input capture and validation"
    Dim mw()    As Double        ' 1..m  MW header values
    Dim names() As String        ' 1..n  substation names
    Dim costM() As Double        ' 1..n, 1..m  cost-per-MW matrix
    Dim n As Long, m As Long
    Dim srcSheet As Worksheet

    If Not CaptureAndValidate(mw, names, costM, n, m, srcSheet) Then GoTo Cleanup

    ' ---- Stage 2: run the analysis fully in memory -----------
    mStage = "analysis"
    Dim nt As Long: nt = UBound(GetThresholds()) - LBound(GetThresholds()) + 1
    Dim totalCols As Long: totalCols = N_FIXED_COLS + 2 * nt

    Dim aTable() As Variant
    ReDim aTable(1 To n, 1 To totalCols)
    AnalyseAll mw, names, costM, n, m, aTable

    ' ---- Stage 3: create the output sheet --------------------
    mStage = "sheet creation"
    Dim wsOut As Worksheet
    Set wsOut = CreateOutputSheet(srcSheet)

    ' ---- Stage 4: write every block ---------------------------
    mStage = "sheet writing"
    Dim srcMWRow As Long, srcFirstDataRow As Long
    Dim hlpMWCol As Long, hlpFirstThreshCol As Long, hlpFirstRow As Long, hlpRows As Long
    WriteSheet wsOut, mw, names, costM, n, m, aTable, totalCols, _
               srcMWRow, srcFirstDataRow, _
               hlpMWCol, hlpFirstThreshCol, hlpFirstRow, hlpRows

    ' ---- Stage 5: build the chart -----------------------------
    mStage = "chart build"
    BuildChart wsOut, mw, costM, n, m, _
               srcMWRow, srcFirstDataRow, _
               hlpMWCol, hlpFirstThreshCol, hlpFirstRow, hlpRows

    ' ---- Stage 6: freeze panes below the header row ----------
    ' FreezePanes is a view operation that acts on the active window,
    ' so the output sheet must be the active sheet for this one step.
    ' This is the sole Activate in the module; no DATA operation uses
    ' selection anywhere. Landing the user on the results is also the
    ' natural finish.
    mStage = "freeze panes"
    wsOut.Activate
    With Application.ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = TABLE_HDR_ROW
        .FreezePanes = True
    End With

    MsgBox "Cost Curve Analysis complete." & vbCrLf & vbCrLf & _
           "Substations analysed: " & n & vbCrLf & _
           "MW points:            " & m & vbCrLf & _
           "Output tab:           " & wsOut.Name, _
           vbInformation, "Cost Curve Analysis"

Cleanup:
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.Calculation = prevCalc
    Exit Sub

ErrHandler:
    MsgBox "Error during " & mStage & ":" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Cost Curve Analysis"
    Resume Cleanup
End Sub

' ============================================================
'  THRESHOLDS
'  Single source of truth. The threshold column block, the chart
'  threshold series and the chart legend all read from this one
'  function, so adding or removing a threshold requires no other
'  code change.
' ============================================================

Private Function GetThresholds() As Variant
    GetThresholds = Array(25000000#, 50000000#, 75000000#, 100000000#)
End Function

' ============================================================
'  INPUT CAPTURE + VALIDATION
' ============================================================

Private Function CaptureAndValidate(ByRef mw() As Double, _
                                    ByRef names() As String, _
                                    ByRef costM() As Double, _
                                    ByRef n As Long, ByRef m As Long, _
                                    ByRef srcSheet As Worksheet) As Boolean
    CaptureAndValidate = False

    Dim rngNames As Range, rngData As Range

    On Error Resume Next
    Set rngNames = Application.InputBox( _
        Prompt:="Select the SUBSTATION NAME list:" & vbCrLf & _
                "(a single column, one name per row, no header)", _
        Title:="Cost Curve Analysis  -  step 1 of 2", Type:=8)
    On Error GoTo 0
    If rngNames Is Nothing Then Exit Function      ' user cancelled

    On Error Resume Next
    Set rngData = Application.InputBox( _
        Prompt:="Select the COST MATRIX:" & vbCrLf & _
                "(first row = MW header, then one cost-per-MW row per substation)", _
        Title:="Cost Curve Analysis  -  step 2 of 2", Type:=8)
    On Error GoTo 0
    If rngData Is Nothing Then Exit Function        ' user cancelled

    ' --- shape checks straight off the Range objects -----------
    If rngNames.Columns.Count <> 1 Then
        MsgBox "The substation name list must be a single column.", _
               vbExclamation, "Cost Curve Analysis": Exit Function
    End If

    n = rngNames.Rows.Count
    m = rngData.Columns.Count

    If m < 3 Then
        MsgBox "The cost matrix needs at least 3 MW points (found " & m & ").", _
               vbExclamation, "Cost Curve Analysis": Exit Function
    End If

    If rngData.Rows.Count - 1 <> n Then
        MsgBox "Row-count mismatch." & vbCrLf & _
               "Cost matrix has " & rngData.Rows.Count & " rows (1 header + " & _
               (rngData.Rows.Count - 1) & " data), but the name list has " & n & " rows.", _
               vbExclamation, "Cost Curve Analysis": Exit Function
    End If

    ' --- one bulk read each, no cell-by-cell looping -----------
    Dim vNames As Variant: vNames = RangeTo2D(rngNames)
    Dim vData  As Variant: vData = RangeTo2D(rngData)

    ' --- validate the MW header: numeric, positive, increasing -
    Dim j As Long
    Dim prevH As Double
    For j = 1 To m
        If Not IsUsableNumber(vData(1, j)) Then
            MsgBox "MW header cell " & j & " is blank, non-numeric or non-positive.", _
                   vbExclamation, "Cost Curve Analysis": Exit Function
        End If
        If j > 1 Then
            If CDbl(vData(1, j)) <= prevH Then
                MsgBox "MW headers must be strictly increasing." & vbCrLf & _
                       "Value at position " & j & " is not greater than the previous one.", _
                       vbExclamation, "Cost Curve Analysis": Exit Function
            End If
        End If
        prevH = CDbl(vData(1, j))
    Next j

    ' --- validate every data cell ------------------------------
    Dim i As Long
    For i = 2 To n + 1
        For j = 1 To m
            If Not IsUsableNumber(vData(i, j)) Then
                MsgBox "Cost cell at data row " & (i - 1) & ", MW point " & j & _
                       " is blank, non-numeric or non-positive.", _
                       vbExclamation, "Cost Curve Analysis": Exit Function
            End If
        Next j
    Next i

    ' --- materialise typed arrays ------------------------------
    ReDim mw(1 To m)
    For j = 1 To m: mw(j) = CDbl(vData(1, j)): Next j

    ReDim names(1 To n)
    For i = 1 To n: names(i) = CStr(vNames(i, 1)): Next i

    ReDim costM(1 To n, 1 To m)
    For i = 1 To n
        For j = 1 To m
            costM(i, j) = CDbl(vData(i + 1, j))
        Next j
    Next i

    Set srcSheet = rngData.Worksheet
    CaptureAndValidate = True
End Function

' Force any range value to a 1-based 2-D array (single cells arrive scalar)
Private Function RangeTo2D(r As Range) As Variant
    Dim v As Variant
    v = r.Value
    If IsArray(v) Then
        RangeTo2D = v
    Else
        Dim a(1 To 1, 1 To 1) As Variant
        a(1, 1) = v
        RangeTo2D = a
    End If
End Function

Private Function IsUsableNumber(ByVal v As Variant) As Boolean
    IsUsableNumber = False
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CDbl(v) <= 0 Then Exit Function
    IsUsableNumber = True
End Function

' ============================================================
'  ANALYSIS DRIVER
'  Fills aTable(1..n, 1..totalCols) with the Block A row content.
' ============================================================

Private Sub AnalyseAll(ByRef mw() As Double, ByRef names() As String, _
                       ByRef costM() As Double, ByVal n As Long, ByVal m As Long, _
                       ByRef aTable() As Variant)
    Dim thr As Variant: thr = GetThresholds()
    Dim i As Long, j As Long, t As Long

    Dim c() As Double: ReDim c(1 To m)
    Dim T() As Double: ReDim T(1 To m)

    For i = 1 To n
        For j = 1 To m
            c(j) = costM(i, j)
            T(j) = c(j) * mw(j)
        Next j

        Dim segs() As TSegment
        segs = Segmentize(mw, T, m)
        Dim nseg As Long: nseg = UBound(segs)

        ' Column A..G
        aTable(i, 1) = names(i)
        aTable(i, 2) = FittedString(mw, T, segs, nseg, m, c)
        aTable(i, 3) = nseg
        aTable(i, 4) = StepChangeStr(mw, segs, nseg)

        Dim li As Long: li = LongestSegIdx(mw, segs, nseg)
        Dim x1 As Double, x2 As Double, Tlong As Double
        x1 = mw(segs(li).FirstJ): x2 = mw(segs(li).LastJ): Tlong = segs(li).Ttotal

        Dim xstar As Double: xstar = KneeXStar(mw, T, segs(li))
        aTable(i, 5) = Round(xstar, 1)
        aTable(i, 6) = -Tlong / (xstar * xstar)          ' slope at knee, -T / x*^2
        aTable(i, 7) = SlowdownPoint(Tlong, x1, x2)

        ' Threshold pairs, ascending order, starting at column H
        For t = LBound(thr) To UBound(thr)
            Dim baseCol As Long: baseCol = N_FIXED_COLS + 1 + 2 * (t - LBound(thr))
            aTable(i, baseCol) = ThresholdRangeStr(T, mw, m, CDbl(thr(t)))
            aTable(i, baseCol + 1) = ThresholdSlope(c, mw, T, m, CDbl(thr(t)))
        Next t
    Next i
End Sub

' ============================================================
'  SEGMENTATION
'  Walk j from 1..m. Points j-1 and j share a segment when
'      |T(j) - T(j-1)| / T(j-1) <= SEG_TOL.
'  SEG_TOL (0.5%) absorbs cent-level rounding in the source data,
'  where a genuine tier step is orders of magnitude larger. Each
'  segment's representative total is the mean of its T values,
'  rounded to the nearest $1,000.
' ============================================================

Private Function Segmentize(ByRef mw() As Double, ByRef T() As Double, _
                            ByVal m As Long) As TSegment()
    Dim segs() As TSegment
    ReDim segs(1 To m)                   ' at most m segments

    Dim nseg As Long: nseg = 0
    Dim segStart As Long: segStart = 1
    Dim j As Long
    For j = 2 To m
        If Abs(T(j) - T(j - 1)) / T(j - 1) > SEG_TOL Then
            nseg = nseg + 1
            segs(nseg).FirstJ = segStart
            segs(nseg).LastJ = j - 1
            segStart = j
        End If
    Next j
    nseg = nseg + 1
    segs(nseg).FirstJ = segStart
    segs(nseg).LastJ = m

    ReDim Preserve segs(1 To nseg)

    Dim k As Long, jj As Long
    Dim s As Double, cnt As Long
    For k = 1 To nseg
        s = 0: cnt = 0
        For jj = segs(k).FirstJ To segs(k).LastJ
            s = s + T(jj): cnt = cnt + 1
        Next jj
        segs(k).Ttotal = Round((s / cnt) / 1000#, 0) * 1000#
    Next k

    Segmentize = segs
End Function

' ============================================================
'  FITTED FUNCTION STRING (column B)
'  One string per substation, segments joined by " | ".
'  Falls back to a global power-law fit only when segmentation
'  yields more than MAX_SEGMENTS segments.
' ============================================================

Private Function FittedString(ByRef mw() As Double, ByRef T() As Double, _
                              ByRef segs() As TSegment, ByVal nseg As Long, _
                              ByVal m As Long, ByRef c() As Double) As String
    If nseg > MAX_SEGMENTS Then
        FittedString = PowerFitString(mw, c, m)
        Exit Function
    End If

    ' StringBuilder pattern: accumulate parts, Join once (no quadratic
    ' concatenation even for large n).
    Dim parts() As String
    ReDim parts(1 To nseg)
    Dim k As Long
    For k = 1 To nseg
        Dim rngTxt As String
        If segs(k).FirstJ = segs(k).LastJ Then
            rngTxt = Format(mw(segs(k).FirstJ), FMT_MW) & " MW"
        Else
            rngTxt = Format(mw(segs(k).FirstJ), FMT_MW) & "-" & _
                     Format(mw(segs(k).LastJ), FMT_MW) & " MW"
        End If
        parts(k) = "y = " & Format(segs(k).Ttotal, "#,##0") & " / x  [" & rngTxt & "]"
    Next k

    FittedString = Join(parts, " | ")
End Function

' Global power-law y = a * x^b, least squares on ln(y) vs ln(x).
' R^2 computed on the log-transformed values. Fallback path only.
Private Function PowerFitString(ByRef mw() As Double, ByRef c() As Double, _
                                ByVal m As Long) As String
    Dim j As Long
    Dim sx As Double, sy As Double, sxx As Double, sxy As Double
    Dim lx As Double, ly As Double
    For j = 1 To m
        lx = Log(mw(j)): ly = Log(c(j))
        sx = sx + lx: sy = sy + ly
        sxx = sxx + lx * lx: sxy = sxy + lx * ly
    Next j
    Dim b As Double, lnA As Double
    b = (m * sxy - sx * sy) / (m * sxx - sx * sx)
    lnA = (sy - b * sx) / m
    Dim a As Double: a = Exp(lnA)

    ' R^2 on ln(y)
    Dim meanLy As Double: meanLy = sy / m
    Dim ssTot As Double, ssRes As Double, resid As Double
    For j = 1 To m
        ly = Log(c(j))
        ssTot = ssTot + (ly - meanLy) ^ 2
        resid = ly - (lnA + b * Log(mw(j)))
        ssRes = ssRes + resid * resid
    Next j
    Dim r2 As Double
    If ssTot > 0 Then r2 = 1 - ssRes / ssTot Else r2 = 1

    PowerFitString = "y = " & Format(a, "0.00E+00") & " * x^" & Format(b, "0.000") & _
                     "  (power fit, R" & ChrW(178) & " = " & Format(r2, "0.000") & ")"
End Function

' ============================================================
'  STEP-CHANGE POINTS (column D)
'  MW values where T steps up = the first MW of each segment after
'  the first. "None" for a single-segment substation.
' ============================================================

Private Function StepChangeStr(ByRef mw() As Double, ByRef segs() As TSegment, _
                               ByVal nseg As Long) As String
    If nseg <= 1 Then StepChangeStr = "None": Exit Function
    Dim parts() As String
    ReDim parts(1 To nseg - 1)
    Dim k As Long
    For k = 2 To nseg
        parts(k - 1) = Format(mw(segs(k).FirstJ), FMT_MW)
    Next k
    StepChangeStr = Join(parts, ", ")
End Function

' ============================================================
'  LONGEST SEGMENT (by MW span)
' ============================================================

Private Function LongestSegIdx(ByRef mw() As Double, ByRef segs() As TSegment, _
                               ByVal nseg As Long) As Long
    Dim k As Long, best As Long, bestSpan As Double, span As Double
    best = 1: bestSpan = -1
    For k = 1 To nseg
        span = mw(segs(k).LastJ) - mw(segs(k).FirstJ)
        If span > bestSpan Then bestSpan = span: best = k
    Next k
    LongestSegIdx = best
End Function

' ============================================================
'  KNEE POINT (column E)  -  Kneedle max-vertical-distance method
'  applied to the longest segment.
'
'  For a pure hyperbola y = T/x on [x1,x2] the knee has the closed
'  form  x* = Sqr(x1 * x2)  (geometric mean of the endpoints).
'  Because Kneedle normalises both axes to [0,1] with an AFFINE map,
'  and an affine reparametrisation does not move an argmax, the
'  discrete search and the closed form agree. So we:
'    - run the discrete Kneedle search over the actual data points,
'    - Debug.Assert the discrete winner sits within one MW step of
'      Sqr(x1*x2)  (this is the primary unit test of the segmentation
'      walk), and
'    - report the continuous x* = Sqr(x1*x2).
'
'  Note: x* depends only on the segment endpoints, NOT on T. This
'  column therefore separates substations by WHERE their segments
'  break, not by cost level. Column G is the one that reflects cost
'  magnitude.
' ============================================================

Private Function KneeXStar(ByRef mw() As Double, ByRef T() As Double, _
                           ByRef seg As TSegment) As Double
    Dim x1 As Double, x2 As Double
    x1 = mw(seg.FirstJ): x2 = mw(seg.LastJ)

    Dim geo As Double: geo = Sqr(x1 * x2)

    If x1 = x2 Then                       ' degenerate single-point segment
        KneeXStar = x1: Exit Function
    End If

    Dim nPts As Long: nPts = seg.LastJ - seg.FirstJ + 1
    If nPts >= 3 Then
        ' Discrete Kneedle: normalise x and y across the segment to
        ' [0,1]; y falls as x rises, so the chord runs from (0,1) to
        ' (1,0), chord(xn) = 1 - xn. Vertical distance below the chord
        ' is (1 - xn) - yn. Take the argmax over data points.
        Dim yLo As Double, yHi As Double
        yHi = T(seg.FirstJ) / mw(seg.FirstJ)   ' cost is highest at x1
        yLo = T(seg.LastJ) / mw(seg.LastJ)      ' cost is lowest  at x2

        Dim j As Long, xn As Double, yn As Double, dist As Double
        Dim bestDist As Double, bestMW As Double, maxGap As Double, prevX As Double
        bestDist = -1E+308
        prevX = mw(seg.FirstJ): maxGap = 0
        For j = seg.FirstJ To seg.LastJ
            xn = (mw(j) - x1) / (x2 - x1)
            yn = ((T(j) / mw(j)) - yLo) / (yHi - yLo)
            dist = (1 - xn) - yn
            If dist > bestDist Then bestDist = dist: bestMW = mw(j)
            If mw(j) - prevX > maxGap Then maxGap = mw(j) - prevX
            prevX = mw(j)
        Next j

        ' Primary unit test: the discrete knee must land within one MW
        ' step of the closed-form geometric mean.
        Debug.Assert Abs(bestMW - geo) <= maxGap + 0.000001
    End If

    KneeXStar = geo
End Function

' ============================================================
'  MARGINAL SLOWDOWN POINT (column G)
'  First MW at which |discrete slope| drops below SLOWDOWN_TOL,
'  evaluated over the longest segment. The analytic slope magnitude
'  is T / x^2, so T / x^2 = SLOWDOWN_TOL gives the continuous
'      x = Sqr(T / SLOWDOWN_TOL).
'  Report the continuous value clamped to the segment range; write
'  "Beyond range" when the condition is never met inside the segment.
'  This column DOES reflect cost magnitude (through T).
' ============================================================

Private Function SlowdownPoint(ByVal Tlong As Double, ByVal x1 As Double, _
                               ByVal x2 As Double) As Variant
    Dim xc As Double: xc = Sqr(Tlong / SLOWDOWN_TOL)
    If xc > x2 Then
        SlowdownPoint = "Beyond range"          ' slope never falls below tol within segment
    ElseIf xc < x1 Then
        SlowdownPoint = Round(x1, 1)            ' already below tol at the segment start
    Else
        SlowdownPoint = Round(xc, 1)
    End If
End Function

' ============================================================
'  THRESHOLD RANGE (columns H, J, L, N ...)
'  Collect every MW point where T(j) <= L. Per the piecewise-constant
'  structure this is always a whole number of segments; consecutive
'  qualifying points collapse into ranges.
' ============================================================

Private Function ThresholdRangeStr(ByRef T() As Double, ByRef mw() As Double, _
                                   ByVal m As Long, ByVal L As Double) As String
    Dim j As Long, cntQ As Long
    Dim qual() As Boolean: ReDim qual(1 To m)
    For j = 1 To m
        qual(j) = (T(j) <= L)
        If qual(j) Then cntQ = cntQ + 1
    Next j

    If cntQ = 0 Then ThresholdRangeStr = "None": Exit Function
    If cntQ = m Then
        ThresholdRangeStr = "All (" & Format(mw(1), FMT_MW) & "-" & _
                            Format(mw(m), FMT_MW) & " MW)"
        Exit Function
    End If

    ' Collapse consecutive (by index) qualifying points into runs.
    Dim parts As String: parts = ""
    Dim runStart As Long: runStart = 0
    For j = 1 To m
        If qual(j) And runStart = 0 Then
            runStart = j
        ElseIf (Not qual(j)) And runStart > 0 Then
            parts = AppendRun(parts, mw, runStart, j - 1)
            runStart = 0
        End If
    Next j
    If runStart > 0 Then parts = AppendRun(parts, mw, runStart, m)

    ThresholdRangeStr = parts & " MW"
End Function

Private Function AppendRun(ByVal acc As String, ByRef mw() As Double, _
                           ByVal a As Long, ByVal b As Long) As String
    Dim piece As String
    If a = b Then
        piece = Format(mw(a), FMT_MW)
    Else
        piece = Format(mw(a), FMT_MW) & "-" & Format(mw(b), FMT_MW)
    End If
    If Len(acc) = 0 Then
        AppendRun = piece
    Else
        AppendRun = acc & ", " & piece
    End If
End Function

' ============================================================
'  THRESHOLD SLOPE (columns I, K, M, O ...)
'  Ordinary least-squares slope of cost vs MW over the qualifying
'  points only (equivalent to WorksheetFunction.Slope).
'
'  NOTE: this is a secant-style AVERAGE across the qualifying window.
'  Because the underlying curve is convex, it understates the
'  steepness at the low-MW end and overstates it at the high-MW end.
' ============================================================

Private Function ThresholdSlope(ByRef c() As Double, ByRef mw() As Double, _
                                ByRef T() As Double, ByVal m As Long, _
                                ByVal L As Double) As Variant
    Dim j As Long, cnt As Long
    Dim xs() As Double, ys() As Double
    ReDim xs(1 To m): ReDim ys(1 To m)
    For j = 1 To m
        If T(j) <= L Then
            cnt = cnt + 1
            xs(cnt) = mw(j): ys(cnt) = c(j)
        End If
    Next j

    If cnt = 0 Then ThresholdSlope = "n/a": Exit Function
    If cnt = 1 Then ThresholdSlope = "n/a (single point)": Exit Function
    ThresholdSlope = OLSSlope(xs, ys, cnt)
End Function

Private Function OLSSlope(ByRef x() As Double, ByRef y() As Double, _
                          ByVal cnt As Long) As Double
    Dim k As Long, sx As Double, sy As Double
    For k = 1 To cnt: sx = sx + x(k): sy = sy + y(k): Next k
    Dim mx As Double: mx = sx / cnt
    Dim my As Double: my = sy / cnt
    Dim num As Double, den As Double
    For k = 1 To cnt
        num = num + (x(k) - mx) * (y(k) - my)
        den = den + (x(k) - mx) ^ 2
    Next k
    OLSSlope = num / den
End Function

' ============================================================
'  OUTPUT SHEET CREATION  (never overwrites, inserts after source)
' ============================================================

Private Function CreateOutputSheet(ByVal srcSheet As Worksheet) As Worksheet
    Dim wb As Workbook: Set wb = srcSheet.Parent
    Dim nm As String: nm = UniqueSheetName(wb, OUT_SHEET)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=srcSheet)
    ws.Name = nm
    Set CreateOutputSheet = ws
End Function

Private Function UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    If Not SheetExists(wb, baseName) Then UniqueSheetName = baseName: Exit Function
    Dim k As Long: k = 2
    Do While SheetExists(wb, baseName & " (" & k & ")")
        k = k + 1
    Loop
    UniqueSheetName = baseName & " (" & k & ")"
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next ws
End Function

' ============================================================
'  SHEET WRITER
'  Writes Block A (analysis table), Block C (source-data copy) and
'  Block D (threshold helper curves). Returns, via ByRef, the anchor
'  rows/cols the chart builder needs. All writes are bulk range
'  assignments.
' ============================================================

Private Sub WriteSheet(ByVal ws As Worksheet, ByRef mw() As Double, _
                       ByRef names() As String, ByRef costM() As Double, _
                       ByVal n As Long, ByVal m As Long, ByRef aTable() As Variant, _
                       ByVal totalCols As Long, _
                       ByRef srcMWRow As Long, ByRef srcFirstDataRow As Long, _
                       ByRef hlpMWCol As Long, ByRef hlpFirstThreshCol As Long, _
                       ByRef hlpFirstRow As Long, ByRef hlpRows As Long)

    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1
    Dim i As Long, j As Long, t As Long

    ' ---- Title -----------------------------------------------
    ws.Cells(TITLE_ROW, 1).Value = "Interconnection Cost Curve Analysis"
    ws.Cells(TITLE_ROW, 1).Font.Bold = True
    ws.Cells(TITLE_ROW, 1).Font.Size = 14

    ' ---- Block A header row -----------------------------------
    Dim hdr() As Variant: ReDim hdr(1 To 1, 1 To totalCols)
    hdr(1, 1) = "Substation"
    hdr(1, 2) = "Fitted Function"
    hdr(1, 3) = "Segments"
    hdr(1, 4) = "Step-Change Points (MW)"
    hdr(1, 5) = "Knee / Flattening Point (MW)"
    hdr(1, 6) = "Slope at Knee ($/MW per MW)"
    hdr(1, 7) = "Marginal Slowdown Point (MW)"
    For t = LBound(thr) To UBound(thr)
        Dim bc As Long: bc = N_FIXED_COLS + 1 + 2 * (t - LBound(thr))
        hdr(1, bc) = "MW Range <= " & ThreshLabel(CDbl(thr(t)))
        hdr(1, bc + 1) = "Slope over Range ($/MW per MW)"
    Next t
    ws.Cells(TABLE_HDR_ROW, 1).Resize(1, totalCols).Value = hdr

    ' ---- Block A data (bulk) ----------------------------------
    ws.Cells(TABLE_HDR_ROW + 1, 1).Resize(n, totalCols).Value = aTable

    ' ---- Block A formatting -----------------------------------
    Dim dataTop As Long: dataTop = TABLE_HDR_ROW + 1
    ws.Cells(dataTop, 3).Resize(n, 1).NumberFormat = "0"                     ' segments
    ws.Cells(dataTop, 5).Resize(n, 1).NumberFormat = FMT_MW1                 ' knee MW
    ws.Cells(dataTop, 6).Resize(n, 1).NumberFormat = FMT_SLOPE               ' slope at knee
    ws.Cells(dataTop, 7).Resize(n, 1).NumberFormat = FMT_MW1                 ' slowdown MW
    For t = 0 To nt - 1
        Dim scol As Long: scol = N_FIXED_COLS + 2 + 2 * t                    ' slope columns
        ws.Cells(dataTop, scol).Resize(n, 1).NumberFormat = FMT_SLOPE
    Next t

    StyleHeaderRow ws.Cells(TABLE_HDR_ROW, 1).Resize(1, totalCols)

    ' ---- Block C: Source Data copy (chart reads THIS) --------
    Dim tableLastRow As Long: tableLastRow = TABLE_HDR_ROW + n
    Dim chartTopRow As Long: chartTopRow = tableLastRow + CHART_GAP_ROWS
    Dim srcHdrTextRow As Long: srcHdrTextRow = chartTopRow + CHART_ROWS
    srcMWRow = srcHdrTextRow + 1
    srcFirstDataRow = srcMWRow + 1

    ws.Cells(srcHdrTextRow, 1).Value = "Source Data"
    ws.Cells(srcHdrTextRow, 1).Font.Bold = True

    ' MW header row: col A label, cols 2..m+1 = MW points
    Dim srcHdr() As Variant: ReDim srcHdr(1 To 1, 1 To m + 1)
    srcHdr(1, 1) = "Substation"
    For j = 1 To m: srcHdr(1, j + 1) = mw(j): Next j
    ws.Cells(srcMWRow, 1).Resize(1, m + 1).Value = srcHdr

    ' name + cost matrix
    Dim srcBody() As Variant: ReDim srcBody(1 To n, 1 To m + 1)
    For i = 1 To n
        srcBody(i, 1) = names(i)
        For j = 1 To m: srcBody(i, j + 1) = costM(i, j): Next j
    Next i
    ws.Cells(srcFirstDataRow, 1).Resize(n, m + 1).Value = srcBody

    ws.Cells(srcMWRow, 2).Resize(1, m).NumberFormat = FMT_MW
    ws.Cells(srcFirstDataRow, 2).Resize(n, m).NumberFormat = FMT_DOLLAR
    StyleHeaderRow ws.Cells(srcMWRow, 1).Resize(1, m + 1)

    ' ---- Block D: Threshold Curve Helper ----------------------
    ' One MW column MW(1)..MW(m) in 1-MW steps, then threshold/MW per
    ' threshold. Left visible (light-gray font) rather than hidden,
    ' because hidden cells can drop out of chart series.
    hlpMWCol = m + 3                                   ' gap column at m+2
    hlpFirstThreshCol = hlpMWCol + 1
    Dim hlpHdrRow As Long: hlpHdrRow = srcHdrTextRow
    Dim hlpSubRow As Long: hlpSubRow = srcMWRow
    hlpFirstRow = srcFirstDataRow
    hlpRows = CLng(mw(m) - mw(1)) + 1

    ws.Cells(hlpHdrRow, hlpMWCol).Value = "Threshold Curve Helper"
    ws.Cells(hlpHdrRow, hlpMWCol).Font.Bold = True

    ' sub-header row
    ws.Cells(hlpSubRow, hlpMWCol).Value = "MW"
    For t = LBound(thr) To UBound(thr)
        ws.Cells(hlpSubRow, hlpFirstThreshCol + (t - LBound(thr))).Value = _
            ThreshLabel(CDbl(thr(t))) & " total cost"
    Next t

    Dim hlp() As Variant: ReDim hlp(1 To hlpRows, 1 To nt + 1)
    Dim r As Long, xVal As Double
    For r = 1 To hlpRows
        xVal = mw(1) + (r - 1)
        hlp(r, 1) = xVal
        For t = LBound(thr) To UBound(thr)
            hlp(r, 2 + (t - LBound(thr))) = CDbl(thr(t)) / xVal
        Next t
    Next r
    ws.Cells(hlpFirstRow, hlpMWCol).Resize(hlpRows, nt + 1).Value = hlp

    ws.Cells(hlpSubRow, hlpMWCol).Resize(hlpRows + 1, nt + 1).Font.Color = RGB(190, 190, 190)
    ws.Cells(hlpHdrRow, hlpMWCol).Font.Color = RGB(190, 190, 190)
    ws.Cells(hlpFirstRow, hlpMWCol).Resize(hlpRows, 1).NumberFormat = FMT_MW
    ws.Cells(hlpFirstRow, hlpFirstThreshCol).Resize(hlpRows, nt).NumberFormat = FMT_DOLLAR

    ' ---- Global column sizing ---------------------------------
    ws.Columns(1).AutoFit
    ws.Columns(2).ColumnWidth = COLB_WIDTH
    ws.Columns(2).WrapText = True
    Dim cc As Long
    For cc = 3 To totalCols
        ws.Columns(cc).AutoFit
    Next cc
End Sub

Private Function ThreshLabel(ByVal L As Double) As String
    ThreshLabel = "$" & Format(L / 1000000#, "0") & "MM"
End Function

Private Sub StyleHeaderRow(ByVal rng As Range)
    rng.Font.Bold = True
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
End Sub

' ============================================================
'  CHART BUILDER
'  xlXYScatterLines so MW sits on a true numeric X axis. Every series
'  references the copied Block C / Block D data on the output sheet,
'  so the tab is self-contained and portable.
' ============================================================

Private Sub BuildChart(ByVal ws As Worksheet, ByRef mw() As Double, _
                       ByRef costM() As Double, ByVal n As Long, ByVal m As Long, _
                       ByVal srcMWRow As Long, ByVal srcFirstDataRow As Long, _
                       ByVal hlpMWCol As Long, ByVal hlpFirstThreshCol As Long, _
                       ByVal hlpFirstRow As Long, ByVal hlpRows As Long)

    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1

    Dim tableLastRow As Long: tableLastRow = TABLE_HDR_ROW + n
    Dim chartTopRow As Long: chartTopRow = tableLastRow + CHART_GAP_ROWS

    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add( _
        Left:=ws.Columns(1).Left + 4, _
        Top:=ws.Rows(chartTopRow).Top, _
        Width:=CHART_W, Height:=CHART_H)

    Dim ch As Chart: Set ch = chObj.Chart
    ch.ChartType = xlXYScatterLines
    ch.PlotVisibleOnly = False                    ' keep light-gray helper cells in the plot

    ' clear any auto-created series
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ' ---- Substation series ------------------------------------
    Dim i As Long
    For i = 1 To n
        Dim srsRow As Long: srsRow = srcMWRow + i          ' data row for substation i
        Dim s As Series: Set s = ch.SeriesCollection.NewSeries
        s.Name = ws.Cells(srsRow, 1).Value                  ' name from the copied name cell
        s.XValues = ws.Range(ws.Cells(srcMWRow, 2), ws.Cells(srcMWRow, m + 1))
        s.Values = ws.Range(ws.Cells(srsRow, 2), ws.Cells(srsRow, m + 1))
        StyleSubSeries s, i
    Next i

    ' ---- Threshold series (added last, drawn on top) ---------
    Dim t As Long
    For t = 0 To nt - 1
        Dim st As Series: Set st = ch.SeriesCollection.NewSeries
        st.Name = ThreshLabel(CDbl(thr(LBound(thr) + t))) & " total cost"
        st.XValues = ws.Range(ws.Cells(hlpFirstRow, hlpMWCol), _
                              ws.Cells(hlpFirstRow + hlpRows - 1, hlpMWCol))
        st.Values = ws.Range(ws.Cells(hlpFirstRow, hlpFirstThreshCol + t), _
                             ws.Cells(hlpFirstRow + hlpRows - 1, hlpFirstThreshCol + t))
        StyleThreshSeries st, t, nt
    Next t

    ' ---- Axes and chrome --------------------------------------
    Dim gMin As Double, gMax As Double
    GlobalMinMax costM, n, m, gMin, gMax
    Dim incr As Double: incr = NiceIncrement(gMax)

    ch.HasTitle = True
    ch.ChartTitle.Text = "Cost per MW by Substation and Project Size"

    With ch.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Project Size (MW)"
        .MinimumScale = mw(1)
        .MaximumScale = mw(m)
        .MajorUnit = ModalStep(mw, m)
        .HasMajorGridlines = False                 ' no vertical gridlines
        .HasMinorGridlines = False
    End With

    With ch.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Interconnection Cost ($/MW)"
        .TickLabels.NumberFormat = FMT_DOLLAR
        .MinimumScale = NiceRound(0.9 * gMin, incr, False)
        .MaximumScale = NiceRound(1.05 * gMax, incr, True)
        .HasMajorGridlines = True                  ' light horizontal gridlines only
        .HasMinorGridlines = False
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
        .MajorGridlines.Format.Line.Weight = 0.5
    End With

    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionRight
    ch.Legend.Format.Line.Visible = msoFalse

    ch.PlotArea.Format.Fill.Visible = msoTrue
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.ChartArea.Format.Line.Visible = msoFalse    ' no chart border
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
End Sub

' Color cycles a 17-entry palette; dash cycles 4; marker cycles 7.
' The three cycles run on co-prime-ish period lengths deliberately, so
' series stay distinguishable even where the palette repeats (n > 17).
Private Sub StyleSubSeries(ByVal s As Series, ByVal idx0 As Long)
    Dim pal As Variant: pal = GetPalette()
    Dim dsh As Variant: dsh = GetDashes()
    Dim mk As Variant: mk = GetMarkers()

    Dim palN As Long: palN = UBound(pal) - LBound(pal) + 1
    Dim dshN As Long: dshN = UBound(dsh) - LBound(dsh) + 1
    Dim mkN As Long: mkN = UBound(mk) - LBound(mk) + 1

    Dim k As Long: k = idx0 - 1
    Dim col As Long: col = pal(LBound(pal) + (k Mod palN))

    s.Format.Line.ForeColor.RGB = col
    s.Format.Line.Weight = 1.75
    s.Border.LineStyle = dsh(LBound(dsh) + (k Mod dshN))

    s.MarkerStyle = mk(LBound(mk) + (k Mod mkN))
    s.MarkerSize = 5
    s.MarkerBackgroundColor = col
    s.MarkerForegroundColor = col                  ' no contrasting marker border
End Sub

' Threshold helper curves: no markers, dashed, gray-to-black ramp in
' ascending threshold order, drawn heavier so they sit over the data.
Private Sub StyleThreshSeries(ByVal s As Series, ByVal idx0 As Long, ByVal count As Long)
    s.MarkerStyle = xlMarkerStyleNone
    s.Border.LineStyle = xlDash
    s.Format.Line.Weight = 2.5
    s.Format.Line.ForeColor.RGB = ThreshColor(idx0, count)
End Sub

' ============================================================
'  Palette / marker / dash tables and small numeric helpers
' ============================================================

Private Function GetPalette() As Variant
    GetPalette = Array( _
        RGB(13, 63, 74), RGB(120, 185, 15), RGB(0, 140, 70), RGB(74, 176, 196), _
        RGB(193, 39, 45), RGB(232, 163, 61), RGB(106, 76, 147), RGB(31, 119, 180), _
        RGB(140, 86, 75), RGB(214, 44, 168), RGB(44, 160, 44), RGB(255, 127, 14), _
        RGB(23, 190, 207), RGB(127, 127, 127), RGB(188, 189, 34), RGB(148, 103, 189), _
        RGB(0, 0, 0))
End Function

Private Function GetDashes() As Variant
    GetDashes = Array(xlContinuous, xlDash, xlDashDot, xlDot)
End Function

Private Function GetMarkers() As Variant
    GetMarkers = Array(xlMarkerStyleCircle, xlMarkerStyleSquare, xlMarkerStyleTriangle, _
                       xlMarkerStyleDiamond, xlMarkerStyleX, xlMarkerStyleStar, _
                       xlMarkerStylePlus)
End Function

' Gray-to-black ramp for the threshold series (ascending threshold).
Private Function ThreshColor(ByVal idx0 As Long, ByVal count As Long) As Long
    Dim ramp As Variant
    ramp = Array(RGB(176, 176, 176), RGB(122, 122, 122), RGB(69, 69, 69), RGB(0, 0, 0))
    Dim rN As Long: rN = UBound(ramp) - LBound(ramp) + 1
    If idx0 <= rN - 1 Then
        ThreshColor = ramp(idx0)
    Else
        ' more thresholds than table entries: interpolate gray 176 -> 0
        Dim g As Long
        If count > 1 Then
            g = CLng(176 - 176 * idx0 / (count - 1))
        Else
            g = 0
        End If
        If g < 0 Then g = 0
        ThreshColor = RGB(g, g, g)
    End If
End Function

Private Sub GlobalMinMax(ByRef costM() As Double, ByVal n As Long, ByVal m As Long, _
                         ByRef gMin As Double, ByRef gMax As Double)
    Dim i As Long, j As Long
    gMin = costM(1, 1): gMax = costM(1, 1)
    For i = 1 To n
        For j = 1 To m
            If costM(i, j) < gMin Then gMin = costM(i, j)
            If costM(i, j) > gMax Then gMax = costM(i, j)
        Next j
    Next i
End Sub

' Modal (most-frequent) consecutive MW step in the header.
Private Function ModalStep(ByRef mw() As Double, ByVal m As Long) As Double
    Dim j As Long, k As Long
    Dim vals() As Double: ReDim vals(1 To m - 1)
    Dim cnts() As Long: ReDim cnts(1 To m - 1)
    Dim nv As Long: nv = 0
    Dim d As Double, found As Boolean
    For j = 2 To m
        d = mw(j) - mw(j - 1)
        found = False
        For k = 1 To nv
            If Abs(vals(k) - d) < 0.000001 Then cnts(k) = cnts(k) + 1: found = True: Exit For
        Next k
        If Not found Then nv = nv + 1: vals(nv) = d: cnts(nv) = 1
    Next j
    Dim best As Long: best = 1
    For k = 2 To nv
        If cnts(k) > cnts(best) Then best = k
    Next k
    ModalStep = vals(best)
End Function

' Round outward (ceil) or inward (floor) to a clean increment.
Private Function NiceRound(ByVal x As Double, ByVal incr As Double, _
                           ByVal up As Boolean) As Double
    If incr <= 0 Then NiceRound = x: Exit Function
    If up Then
        NiceRound = -Int(-x / incr) * incr
    Else
        NiceRound = Int(x / incr) * incr
    End If
End Function

' A clean axis increment ~ one order of magnitude below the max value.
Private Function NiceIncrement(ByVal gMax As Double) As Double
    If gMax <= 0 Then NiceIncrement = 1: Exit Function
    Dim ord As Double: ord = Int(Log(gMax) / Log(10#))
    Dim incr As Double: incr = 10# ^ (ord - 1)
    If incr < 1 Then incr = 1
    NiceIncrement = incr
End Function

' ============================================================
'  SELF TEST
'  Runs the analysis logic fully in memory (no Excel ranges, no
'  charts) against the 17-substation x 21-MW reference fixture, and
'  reports each assertion to the Immediate window. Reads
'  substation_cost_per_mw.csv from the workbook folder if present;
'  otherwise uses the hardcoded fixture.
' ============================================================

Public Sub SelfTest()
    Dim mw() As Double, names() As String, costM() As Double
    Dim n As Long, m As Long

    LoadFixture mw, names, costM, n, m

    Debug.Print "=== Cost Curve Analysis SelfTest ==="
    Debug.Print "Fixture: " & n & " substations x " & m & " MW points"

    Dim passCount As Long, failCount As Long
    Dim i As Long, j As Long
    Dim c() As Double: ReDim c(1 To m)
    Dim T() As Double: ReDim T(1 To m)

    Dim iCun As Long, iHob As Long, iChv As Long
    iCun = NameIndex(names, n, "Cunningham")
    iHob = NameIndex(names, n, "Hobbs")
    iChv = NameIndex(names, n, "Chaves County")

    ' ---- Cunningham & Hobbs (identical single-tier curves) ----
    Dim who As Variant: who = Array(iCun, iHob)
    Dim w As Long
    For w = 0 To 1
        i = who(w)
        If i = 0 Then
            Assert False, "Cunningham/Hobbs row missing", passCount, failCount
        Else
            RowVectors costM, mw, i, m, c, T
            Dim segs() As TSegment: segs = Segmentize(mw, T, m)
            Assert UBound(segs) = 1, GetName(names, i) & ": 1 segment", passCount, failCount
            Assert segs(1).Ttotal = 47633000#, GetName(names, i) & ": T = 47,633,000", passCount, failCount
            Assert FittedString(mw, T, segs, UBound(segs), m, c) = _
                   "y = 47,633,000 / x  [100-300 MW]", _
                   GetName(names, i) & ": formula string", passCount, failCount
            Dim li As Long: li = LongestSegIdx(mw, segs, UBound(segs))
            Assert Round(KneeXStar(mw, T, segs(li)), 1) = 173.2, _
                   GetName(names, i) & ": knee 173.2 MW", passCount, failCount
            Assert ThresholdRangeStr(T, mw, m, 25000000#) = "None", _
                   GetName(names, i) & ": <=$25MM None", passCount, failCount
            Assert ThresholdRangeStr(T, mw, m, 50000000#) = "All (100-300 MW)", _
                   GetName(names, i) & ": <=$50MM All", passCount, failCount
            Assert ThresholdRangeStr(T, mw, m, 75000000#) = "All (100-300 MW)", _
                   GetName(names, i) & ": <=$75MM All", passCount, failCount
            Assert ThresholdRangeStr(T, mw, m, 100000000#) = "All (100-300 MW)", _
                   GetName(names, i) & ": <=$100MM All", passCount, failCount
        End If
    Next w

    ' ---- Chaves County (multi-tier) ---------------------------
    If iChv = 0 Then
        Assert False, "Chaves County row missing", passCount, failCount
    Else
        RowVectors costM, mw, iChv, m, c, T
        Dim csg() As TSegment: csg = Segmentize(mw, T, m)
        Assert UBound(csg) = 5, "Chaves: 5 segments", passCount, failCount
        Assert csg(1).Ttotal = 49473000#, "Chaves: seg1 T", passCount, failCount
        Assert csg(2).Ttotal = 98003000#, "Chaves: seg2 T", passCount, failCount
        Assert csg(3).Ttotal = 185783000#, "Chaves: seg3 T", passCount, failCount
        Assert csg(4).Ttotal = 205353000#, "Chaves: seg4 T", passCount, failCount
        Assert csg(5).Ttotal = 207843000#, "Chaves: seg5 T", passCount, failCount
        Assert StepChangeStr(mw, csg, 5) = "200, 210, 240, 280", _
               "Chaves: step points 200, 210, 240, 280", passCount, failCount
        Dim cli As Long: cli = LongestSegIdx(mw, csg, 5)
        Assert mw(csg(cli).FirstJ) = 100 And mw(csg(cli).LastJ) = 190, _
               "Chaves: longest segment 100-190", passCount, failCount
        Assert Round(KneeXStar(mw, T, csg(cli)), 1) = 137.8, _
               "Chaves: knee 137.8 MW", passCount, failCount
        Assert ThresholdRangeStr(T, mw, m, 25000000#) = "None", _
               "Chaves: <=$25MM None", passCount, failCount
        Assert ThresholdRangeStr(T, mw, m, 50000000#) = "100-190 MW", _
               "Chaves: <=$50MM 100-190 MW", passCount, failCount
        Assert ThresholdRangeStr(T, mw, m, 75000000#) = "100-190 MW", _
               "Chaves: <=$75MM 100-190 MW", passCount, failCount
        Assert ThresholdRangeStr(T, mw, m, 100000000#) = "100-200 MW", _
               "Chaves: <=$100MM 100-200 MW", passCount, failCount
    End If

    ' ---- Global checks ----------------------------------------
    Dim gMinT As Double: gMinT = 1E+300
    Dim allNone25 As Boolean: allNone25 = True
    Dim segSumOK As Boolean: segSumOK = True
    For i = 1 To n
        RowVectors costM, mw, i, m, c, T
        For j = 1 To m
            If T(j) < gMinT Then gMinT = T(j)
        Next j
        If ThresholdRangeStr(T, mw, m, 25000000#) <> "None" Then allNone25 = False
        Dim sg() As TSegment: sg = Segmentize(mw, T, m)
        Dim k As Long, ptSum As Long: ptSum = 0
        For k = 1 To UBound(sg)
            ptSum = ptSum + (sg(k).LastJ - sg(k).FirstJ + 1)
        Next k
        If ptSum <> m Then segSumOK = False
    Next i

    Assert allNone25, "Global: <=$25MM is None for all 17 rows", passCount, failCount
    Assert Round(gMinT, 0) = 47633000#, "Global: min implied total = 47,633,000", passCount, failCount
    Assert segSumOK, "Global: segment point counts sum to m for every substation", passCount, failCount

    Debug.Print "=== SelfTest done: " & passCount & " passed, " & failCount & " failed ==="
    If failCount = 0 Then
        Debug.Print "ALL TESTS PASSED"
    Else
        Debug.Print "*** " & failCount & " FAILURE(S) ***"
    End If
End Sub

Private Sub Assert(ByVal cond As Boolean, ByVal label As String, _
                   ByRef passCount As Long, ByRef failCount As Long)
    If cond Then
        passCount = passCount + 1
        Debug.Print "  PASS  " & label
    Else
        failCount = failCount + 1
        Debug.Print "  FAIL  " & label
    End If
End Sub

Private Sub RowVectors(ByRef costM() As Double, ByRef mw() As Double, ByVal i As Long, _
                       ByVal m As Long, ByRef c() As Double, ByRef T() As Double)
    Dim j As Long
    For j = 1 To m
        c(j) = costM(i, j)
        T(j) = c(j) * mw(j)
    Next j
End Sub

Private Function NameIndex(ByRef names() As String, ByVal n As Long, ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To n
        If StrComp(names(i), nm, vbTextCompare) = 0 Then NameIndex = i: Exit Function
    Next i
    NameIndex = 0
End Function

Private Function GetName(ByRef names() As String, ByVal i As Long) As String
    If i >= LBound(names) And i <= UBound(names) Then GetName = names(i) Else GetName = "?"
End Function

' ------------------------------------------------------------
'  Fixture loader. Tries substation_cost_per_mw.csv in the workbook
'  folder; falls back to a hardcoded 17 x 21 matrix that reproduces
'  the reference cases from section 9 (Cunningham & Hobbs single-tier,
'  Chaves County five-tier, min implied total 47,633,000, and no point
'  at or below $25MM anywhere).
' ------------------------------------------------------------

Private Sub LoadFixture(ByRef mw() As Double, ByRef names() As String, _
                        ByRef costM() As Double, ByRef n As Long, ByRef m As Long)
    If TryLoadCsv(mw, names, costM, n, m) Then Exit Sub

    m = 21: n = 17
    ReDim mw(1 To m)
    Dim j As Long
    For j = 1 To m: mw(j) = 100 + (j - 1) * 10: Next j     ' 100..300 step 10

    ReDim names(1 To n)
    ReDim costM(1 To n, 1 To m)

    ' Per-substation total-cost tier profile T(j). cost(j) = T(j)/mw(j).
    Dim T() As Double: ReDim T(1 To m)

    ' 1) Cunningham  - single tier 47,633,000
    names(1) = "Cunningham": FillConst T, m, 47633000#: SetRow costM, mw, T, 1, m

    ' 2) Hobbs       - identical single tier
    names(2) = "Hobbs": FillConst T, m, 47633000#: SetRow costM, mw, T, 2, m

    ' 3) Chaves County - five tiers
    '    100-190: 49,473,000 | 200: 98,003,000 | 210-230: 185,783,000
    '    240-270: 205,353,000 | 280-300: 207,843,000
    For j = 1 To m
        Select Case mw(j)
            Case Is <= 190: T(j) = 49473000#
            Case 200: T(j) = 98003000#
            Case 210, 220, 230: T(j) = 185783000#
            Case 240, 250, 260, 270: T(j) = 205353000#
            Case Else: T(j) = 207843000#
        End Select
    Next j
    names(3) = "Chaves County": SetRow costM, mw, T, 3, m

    ' 4..17) Filler substations, all single- or two-tier, every total
    ' strictly above 47,633,000 so the global minimum stays 47,633,000
    ' and no point falls at or below $25MM.
    Dim fillNames As Variant
    fillNames = Array("Artesia", "Roswell", "Carlsbad", "Lovington", "Portales", _
                      "Clovis", "Tucumcari", "Ruidoso", "Alamogordo", "Deming", _
                      "Silver City", "Socorro", "Las Cruces", "Truth or Consequences")
    Dim fillBase As Variant
    fillBase = Array(55000000#, 62000000#, 70000000#, 78000000#, 85000000#, _
                     92000000#, 100000000#, 110000000#, 120000000#, 135000000#, _
                     150000000#, 165000000#, 180000000#, 195000000#)
    Dim f As Long
    For f = 0 To 13
        Dim base As Double: base = fillBase(f)
        For j = 1 To m
            ' single upward tier step at 250 MW for variety (still all > 47.633M)
            If mw(j) >= 250 Then
                T(j) = base * 1.15
            Else
                T(j) = base
            End If
        Next j
        names(4 + f) = fillNames(f)
        SetRow costM, mw, T, 4 + f, m
    Next f
End Sub

Private Sub FillConst(ByRef T() As Double, ByVal m As Long, ByVal v As Double)
    Dim j As Long
    For j = 1 To m: T(j) = v: Next j
End Sub

Private Sub SetRow(ByRef costM() As Double, ByRef mw() As Double, ByRef T() As Double, _
                   ByVal i As Long, ByVal m As Long)
    Dim j As Long
    For j = 1 To m: costM(i, j) = T(j) / mw(j): Next j
End Sub

' Attempt to read substation_cost_per_mw.csv from the workbook folder.
' Layout: first row = header with a leading name cell then MW values;
' each subsequent row = name then cost-per-MW values. Returns False if
' the file is absent or unreadable (SelfTest then uses the hardcoded set).
Private Function TryLoadCsv(ByRef mw() As Double, ByRef names() As String, _
                            ByRef costM() As Double, ByRef n As Long, ByRef m As Long) As Boolean
    TryLoadCsv = False
    On Error GoTo Fail

    Dim path As String
    path = ThisWorkbook.path
    If Len(path) = 0 Then Exit Function
    Dim fn As String: fn = path & Application.PathSeparator & "substation_cost_per_mw.csv"
    If Len(Dir(fn)) = 0 Then Exit Function

    Dim ff As Integer: ff = FreeFile
    Dim lines() As String, cnt As Long
    ReDim lines(1 To 1000)
    Open fn For Input As #ff
    Do While Not EOF(ff)
        Dim ln As String: Line Input #ff, ln
        If Len(Trim(ln)) > 0 Then
            cnt = cnt + 1
            If cnt > UBound(lines) Then ReDim Preserve lines(1 To UBound(lines) + 1000)
            lines(cnt) = ln
        End If
    Loop
    Close #ff
    If cnt < 2 Then Exit Function

    Dim hdr() As String: hdr = Split(lines(1), ",")
    m = UBound(hdr) - LBound(hdr)                 ' minus the leading name column
    n = cnt - 1
    ReDim mw(1 To m): ReDim names(1 To n): ReDim costM(1 To n, 1 To m)

    Dim j As Long
    For j = 1 To m: mw(j) = CDbl(Trim(hdr(LBound(hdr) + j))): Next j

    Dim i As Long
    For i = 1 To n
        Dim parts() As String: parts = Split(lines(i + 1), ",")
        names(i) = Trim(parts(LBound(parts)))
        For j = 1 To m
            costM(i, j) = CDbl(Trim(parts(LBound(parts) + j)))
        Next j
    Next i

    TryLoadCsv = True
    Exit Function
Fail:
    On Error Resume Next
    Close #ff
    TryLoadCsv = False
End Function
