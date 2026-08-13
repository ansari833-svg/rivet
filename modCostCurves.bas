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

' Rank B (slope) direction. Steepest first reads a steep slope as the
' curve still capturing scale economies (incremental MW still buying
' meaningful unit-cost reduction); flattest first reads the opposite
' (curve already flattened, site near efficient size). Both are
' legitimate, so the direction is a constant rather than a hardcoded sort.
' 1 = steepest first (largest magnitude gets rank 1)
' 2 = flattest first (smallest magnitude gets rank 1)
Private Const SLOPE_RANK_MODE As Long = 1

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
    ' Each threshold now occupies FOUR columns: range, slope, Rank A, Rank B.
    ' The column anchors are derived from GetThresholds() so adding or
    ' removing a threshold shifts the whole layout with no other edit.
    Dim totalCols As Long: totalCols = N_FIXED_COLS + 4 * nt

    ' Cross-substation rankings, one Rank A and one Rank B per threshold.
    Dim rkCount() As Long, rkMaxMW() As Double, rkCostMax() As Double
    Dim rkSlope() As Double, rkHasSlope() As Boolean
    Dim rankA() As Long, rankB() As Long
    ComputeRankings mw, names, costM, n, m, _
                    rkCount, rkMaxMW, rkCostMax, rkSlope, rkHasSlope, rankA, rankB

    Dim aTable() As Variant
    ReDim aTable(1 To n, 1 To totalCols)
    AnalyseAll mw, names, costM, n, m, aTable, rankA, rankB

    ' ---- Stage 3: create the output sheet --------------------
    mStage = "sheet creation"
    Dim wsOut As Worksheet
    Set wsOut = CreateOutputSheet(srcSheet)

    ' ---- Stage 4: write every block ---------------------------
    mStage = "sheet writing"
    Dim srcMWRow As Long, srcFirstDataRow As Long
    Dim hlpMWCol As Long, hlpFirstThreshCol As Long, hlpFirstRow As Long, hlpRows As Long
    WriteSheet wsOut, mw, names, costM, n, m, aTable, totalCols, _
               rkCount, rkMaxMW, rkSlope, rankA, rankB, _
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
                       ByRef aTable() As Variant, _
                       ByRef rankA() As Long, ByRef rankB() As Long)
    Dim thr As Variant: thr = GetThresholds()
    Dim i As Long, j As Long, tt As Long

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

        ' Four columns per threshold, ascending order, starting at column H:
        '   +0 range   +1 slope   +2 Rank by MW breadth   +3 Rank by slope
        For tt = 0 To UBound(thr) - LBound(thr)
            Dim baseCol As Long: baseCol = N_FIXED_COLS + 1 + 4 * tt
            Dim L As Double: L = CDbl(thr(LBound(thr) + tt))
            aTable(i, baseCol) = ThresholdRangeStr(T, mw, m, L)
            aTable(i, baseCol + 1) = ThresholdSlope(c, mw, T, m, L)
            ' Rank 0 means "not ranked" (no qualifying points, or a single
            ' point for slope). Write n/a rather than a last-place rank.
            aTable(i, baseCol + 2) = IIf(rankA(i, tt + 1) = 0, "n/a", CLng(rankA(i, tt + 1)))
            aTable(i, baseCol + 3) = IIf(rankB(i, tt + 1) = 0, "n/a", CLng(rankB(i, tt + 1)))
        Next tt
    Next i
End Sub

' ============================================================
'  RANKINGS
'  Two independent ranks per substation per threshold band, never
'  blended and never computed across thresholds:
'    Rank A  by MW breadth   (more qualifying MW is better)
'    Rank B  by slope        (direction set by SLOPE_RANK_MODE)
'
'  Both use COMPETITION ranking on the metric key only: equal metrics
'  share the lower rank and the next distinct value skips ahead
'  (1, 2, 2, 4 - not dense 1, 2, 2, 3). The alphabetical backstop is a
'  deterministic DISPLAY order within a metric tie, not a rank
'  differentiator, so genuinely equal metrics stay tied.
'
'  Rank A tiebreak chain (metric key): higher qualifying count, then
'  higher maximum qualifying MW, then LOWER cost per MW at that maximum
'  qualifying MW. Tiebreak 3 is load bearing - e.g. many substations
'  qualify at every MW point and share the same max MW, and cost at the
'  largest affordable size is what separates them.
'
'  Rank B tiebreak: equal slope to $0.1, then higher qualifying count.
'  Substations with zero qualifying points (Rank A) or fewer than two
'  (Rank B, slope undefined) are excluded from the population entirely
'  and get rank 0 (written as n/a), so ranks run 1..k over the rankable
'  set, not 1..n.
' ============================================================

Private Sub ComputeRankings(ByRef mw() As Double, ByRef names() As String, _
                            ByRef costM() As Double, ByVal n As Long, ByVal m As Long, _
                            ByRef rkCount() As Long, ByRef rkMaxMW() As Double, _
                            ByRef rkCostMax() As Double, ByRef rkSlope() As Double, _
                            ByRef rkHasSlope() As Boolean, _
                            ByRef rankA() As Long, ByRef rankB() As Long)
    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1

    ReDim rkCount(1 To n, 1 To nt)
    ReDim rkMaxMW(1 To n, 1 To nt)
    ReDim rkCostMax(1 To n, 1 To nt)
    ReDim rkSlope(1 To n, 1 To nt)
    ReDim rkHasSlope(1 To n, 1 To nt)
    ReDim rankA(1 To n, 1 To nt)
    ReDim rankB(1 To n, 1 To nt)

    ' ---- Per-substation, per-threshold metrics ---------------
    Dim i As Long, j As Long, tt As Long
    Dim c() As Double: ReDim c(1 To m)
    Dim T() As Double: ReDim T(1 To m)
    For i = 1 To n
        For j = 1 To m: c(j) = costM(i, j): T(j) = c(j) * mw(j): Next j
        For tt = 1 To nt
            Dim L As Double: L = CDbl(thr(LBound(thr) + tt - 1))
            Dim cnt As Long: cnt = 0
            Dim lastJ As Long: lastJ = 0
            Dim sx As Double, sy As Double, sxx As Double, sxy As Double
            sx = 0: sy = 0: sxx = 0: sxy = 0
            For j = 1 To m
                If T(j) <= L Then
                    cnt = cnt + 1: lastJ = j
                    sx = sx + mw(j): sy = sy + c(j)
                    sxx = sxx + mw(j) * mw(j): sxy = sxy + mw(j) * c(j)
                End If
            Next j
            rkCount(i, tt) = cnt
            If cnt >= 1 Then
                rkMaxMW(i, tt) = mw(lastJ)
                rkCostMax(i, tt) = c(lastJ)
            End If
            If cnt >= 2 Then
                rkHasSlope(i, tt) = True
                rkSlope(i, tt) = (cnt * sxy - sx * sy) / (cnt * sxx - sx * sx)
            End If
        Next tt
    Next i

    ' ---- Competition ranking within each threshold -----------
    Dim ord() As Long: ReDim ord(1 To n)
    Dim popN As Long, p As Long
    For tt = 1 To nt
        ' Rank A population: substations with >=1 qualifying point.
        popN = 0
        For i = 1 To n
            If rkCount(i, tt) >= 1 Then popN = popN + 1: ord(popN) = i
        Next i
        If popN > 0 Then
            SortIdx ord, popN, tt, True, rkCount, rkMaxMW, rkCostMax, rkSlope, names
            For p = 1 To popN
                If p = 1 Then
                    rankA(ord(p), tt) = 1
                ElseIf EqRankA(ord(p), ord(p - 1), tt, rkCount, rkMaxMW, rkCostMax) Then
                    rankA(ord(p), tt) = rankA(ord(p - 1), tt)
                Else
                    rankA(ord(p), tt) = p
                End If
            Next p
        End If

        ' Rank B population: substations with a defined slope (>=2 points).
        popN = 0
        For i = 1 To n
            If rkHasSlope(i, tt) Then popN = popN + 1: ord(popN) = i
        Next i
        If popN > 0 Then
            SortIdx ord, popN, tt, False, rkCount, rkMaxMW, rkCostMax, rkSlope, names
            For p = 1 To popN
                If p = 1 Then
                    rankB(ord(p), tt) = 1
                ElseIf EqRankB(ord(p), ord(p - 1), tt, rkCount, rkSlope) Then
                    rankB(ord(p), tt) = rankB(ord(p - 1), tt)
                Else
                    rankB(ord(p), tt) = p
                End If
            Next p
        End If
    Next tt
End Sub

' Insertion sort of the index array ord(1..popN) for one threshold.
' isRankA selects the ordering; the alphabetical name comparison is the
' final DISPLAY tiebreak so the order is deterministic without making
' otherwise-equal metrics count as distinct for ranking.
Private Sub SortIdx(ByRef ord() As Long, ByVal popN As Long, ByVal tt As Long, _
                    ByVal isRankA As Boolean, _
                    ByRef rkCount() As Long, ByRef rkMaxMW() As Double, _
                    ByRef rkCostMax() As Double, ByRef rkSlope() As Double, _
                    ByRef names() As String)
    Dim a As Long, b As Long, keyIdx As Long
    For a = 2 To popN
        keyIdx = ord(a): b = a - 1
        Do While b >= 1
            If LessThan(keyIdx, ord(b), tt, isRankA, rkCount, rkMaxMW, rkCostMax, rkSlope, names) Then
                ord(b + 1) = ord(b): b = b - 1
            Else
                Exit Do
            End If
        Loop
        ord(b + 1) = keyIdx
    Next a
End Sub

' True when substation i1 sorts before i2 for the given rank type.
Private Function LessThan(ByVal i1 As Long, ByVal i2 As Long, ByVal tt As Long, _
                          ByVal isRankA As Boolean, _
                          ByRef rkCount() As Long, ByRef rkMaxMW() As Double, _
                          ByRef rkCostMax() As Double, ByRef rkSlope() As Double, _
                          ByRef names() As String) As Boolean
    If isRankA Then
        If rkCount(i1, tt) <> rkCount(i2, tt) Then
            LessThan = rkCount(i1, tt) > rkCount(i2, tt): Exit Function      ' more MW is better
        End If
        If rkMaxMW(i1, tt) <> rkMaxMW(i2, tt) Then
            LessThan = rkMaxMW(i1, tt) > rkMaxMW(i2, tt): Exit Function       ' higher max MW
        End If
        If Round(rkCostMax(i1, tt), 2) <> Round(rkCostMax(i2, tt), 2) Then
            LessThan = rkCostMax(i1, tt) < rkCostMax(i2, tt): Exit Function    ' lower cost at max MW
        End If
    Else
        Dim s1 As Double, s2 As Double
        s1 = Round(rkSlope(i1, tt), 1): s2 = Round(rkSlope(i2, tt), 1)
        If s1 <> s2 Then
            If SLOPE_RANK_MODE = 2 Then
                LessThan = Abs(s1) < Abs(s2)                                  ' flattest first
            Else
                LessThan = Abs(s1) > Abs(s2)                                  ' steepest first
            End If
            Exit Function
        End If
        If rkCount(i1, tt) <> rkCount(i2, tt) Then
            LessThan = rkCount(i1, tt) > rkCount(i2, tt): Exit Function
        End If
    End If
    LessThan = (StrComp(names(i1), names(i2), vbTextCompare) < 0)             ' display backstop
End Function

Private Function EqRankA(ByVal i1 As Long, ByVal i2 As Long, ByVal tt As Long, _
                         ByRef rkCount() As Long, ByRef rkMaxMW() As Double, _
                         ByRef rkCostMax() As Double) As Boolean
    EqRankA = (rkCount(i1, tt) = rkCount(i2, tt)) And _
              (rkMaxMW(i1, tt) = rkMaxMW(i2, tt)) And _
              (Round(rkCostMax(i1, tt), 2) = Round(rkCostMax(i2, tt), 2))
End Function

Private Function EqRankB(ByVal i1 As Long, ByVal i2 As Long, ByVal tt As Long, _
                         ByRef rkCount() As Long, ByRef rkSlope() As Double) As Boolean
    EqRankB = (Round(rkSlope(i1, tt), 1) = Round(rkSlope(i2, tt), 1)) And _
              (rkCount(i1, tt) = rkCount(i2, tt))
End Function

Private Function SlopeRankHeader() As String
    If SLOPE_RANK_MODE = 2 Then
        SlopeRankHeader = "Rank by Slope (flattest first)"
    Else
        SlopeRankHeader = "Rank by Slope (steepest first)"
    End If
End Function

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
                       ByRef rkCount() As Long, ByRef rkMaxMW() As Double, _
                       ByRef rkSlope() As Double, ByRef rankA() As Long, ByRef rankB() As Long, _
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
    For t = 0 To nt - 1
        Dim bc As Long: bc = N_FIXED_COLS + 1 + 4 * t
        hdr(1, bc) = "MW Range <= " & ThreshLabel(CDbl(thr(LBound(thr) + t)))
        hdr(1, bc + 1) = "Slope over Range ($/MW per MW)"
        hdr(1, bc + 2) = "Rank by MW Breadth"
        hdr(1, bc + 3) = SlopeRankHeader()
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
        Dim scol As Long: scol = N_FIXED_COLS + 2 + 4 * t                    ' slope over range
        ws.Cells(dataTop, scol).Resize(n, 1).NumberFormat = FMT_SLOPE
        ' Rank columns: centred, plain integer, no colour scale (rank is
        ' ordinal, so a gradient would imply a magnitude it does not have).
        ' n/a cells shown as centred grey text.
        Dim rc As Long
        For rc = scol + 1 To scol + 2
            With ws.Cells(dataTop, rc).Resize(n, 1)
                .NumberFormat = "0"
                .HorizontalAlignment = xlCenter
            End With
            For i = 1 To n
                If ws.Cells(dataTop + i - 1, rc).Value = "n/a" Then
                    ws.Cells(dataTop + i - 1, rc).Font.Color = RGB(150, 150, 150)
                End If
            Next i
        Next rc
    Next t

    StyleHeaderRow ws.Cells(TABLE_HDR_ROW, 1).Resize(1, totalCols)

    ' ---- Block E: Threshold Rankings leaderboard -------------
    ' Two rows below the chart. Pushes the source-data / helper blocks
    ' further down so nothing overlaps; chart series still reference the
    ' source-data copy wherever it lands.
    Dim tableLastRow As Long: tableLastRow = TABLE_HDR_ROW + n
    Dim chartTopRow As Long: chartTopRow = tableLastRow + CHART_GAP_ROWS
    Dim lbTopRow As Long: lbTopRow = chartTopRow + CHART_ROWS + 2
    Dim lbLastRow As Long
    BuildLeaderboard ws, lbTopRow, names, n, rkCount, rkMaxMW, rkSlope, rankA, rankB, lbLastRow

    ' ---- Block C: Source Data copy (chart reads THIS) --------
    Dim srcHdrTextRow As Long: srcHdrTextRow = lbLastRow + 2
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
'  LEADERBOARD  (Block E, "Threshold Rankings")
'  Two side-by-side sorted lists per threshold - by MW breadth and by
'  slope - for reading, complementing the inline rank lookup columns.
'  Sorted in memory (never write-then-Range.Sort). A monospace font
'  keeps the two columns aligned. Rank-1 rows are bold; a zero-qualifier
'  threshold collapses to a single explanatory line.
' ============================================================

Private Sub BuildLeaderboard(ByVal ws As Worksheet, ByVal topRow As Long, _
                             ByRef names() As String, ByVal n As Long, _
                             ByRef rkCount() As Long, ByRef rkMaxMW() As Double, _
                             ByRef rkSlope() As Double, _
                             ByRef rankA() As Long, ByRef rankB() As Long, _
                             ByRef lastRow As Long)
    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1

    Dim r As Long: r = topRow
    ws.Cells(r, 1).Value = "Threshold Rankings"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    r = r + 2

    Dim tt As Long, i As Long
    For tt = 1 To nt
        Dim L As Double: L = CDbl(thr(LBound(thr) + tt - 1))
        Dim qcnt As Long: qcnt = 0
        For i = 1 To n
            If rkCount(i, tt) >= 1 Then qcnt = qcnt + 1
        Next i

        ws.Cells(r, 1).Value = "<= " & ThreshLabel(L) & "   (" & qcnt & " of " & n & " substations qualify)"
        ws.Cells(r, 1).Font.Bold = True
        r = r + 1

        If qcnt = 0 Then
            ws.Cells(r, 2).Value = "No substations qualify at this threshold"
            r = r + 2
        Else
            ws.Cells(r, 2).Value = "By MW Breadth"
            ws.Cells(r, 6).Value = "By Slope, " & IIf(SLOPE_RANK_MODE = 2, "flattest", "steepest") & " first"
            ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 6).Font.Bold = True
            r = r + 1

            Dim ordA() As Long, kA As Long: LeaderOrder rankA, names, n, tt, ordA, kA
            Dim ordB() As Long, kB As Long: LeaderOrder rankB, names, n, tt, ordB, kB
            Dim rows As Long: rows = kA: If kB > rows Then rows = kB

            Dim rr As Long
            For rr = 1 To rows
                If rr <= kA Then
                    i = ordA(rr)
                    ws.Cells(r, 1).Value = rankA(i, tt)
                    ws.Cells(r, 2).Value = names(i)
                    ws.Cells(r, 3).Value = rkCount(i, tt) & " pts, max " & Format(rkMaxMW(i, tt), FMT_MW)
                    If rankA(i, tt) = 1 Then ws.Cells(r, 1).Resize(1, 3).Font.Bold = True
                End If
                If rr <= kB Then
                    i = ordB(rr)
                    ws.Cells(r, 5).Value = rankB(i, tt)
                    ws.Cells(r, 6).Value = names(i)
                    ws.Cells(r, 7).Value = Format(rkSlope(i, tt), "#,##0.0")
                    If rankB(i, tt) = 1 Then ws.Cells(r, 5).Resize(1, 3).Font.Bold = True
                End If
                r = r + 1
            Next rr
            r = r + 1                           ' spacer between sub-blocks
        End If
    Next tt

    lastRow = r
    ' Monospace so the two side-by-side lists line up (font name only,
    ' preserves the bold already applied to rank-1 and header rows).
    ws.Range(ws.Cells(topRow, 1), ws.Cells(lastRow, 7)).Font.Name = "Consolas"
End Sub

' Population indices (rank > 0) for one threshold, sorted for display by
' (rank asc, name asc). The rank already encodes the full metric order
' with competition ties, so sorting on it reproduces the ranked order.
Private Sub LeaderOrder(ByRef rank() As Long, ByRef names() As String, ByVal n As Long, _
                        ByVal tt As Long, ByRef ord() As Long, ByRef k As Long)
    ReDim ord(1 To n)
    k = 0
    Dim i As Long
    For i = 1 To n
        If rank(i, tt) > 0 Then k = k + 1: ord(k) = i
    Next i
    If k = 0 Then Exit Sub
    Dim a As Long, b As Long, keyIdx As Long
    For a = 2 To k
        keyIdx = ord(a): b = a - 1
        Do While b >= 1
            If LeaderBefore(keyIdx, ord(b), tt, rank, names) Then
                ord(b + 1) = ord(b): b = b - 1
            Else
                Exit Do
            End If
        Loop
        ord(b + 1) = keyIdx
    Next a
End Sub

Private Function LeaderBefore(ByVal i1 As Long, ByVal i2 As Long, ByVal tt As Long, _
                              ByRef rank() As Long, ByRef names() As String) As Boolean
    If rank(i1, tt) <> rank(i2, tt) Then
        LeaderBefore = rank(i1, tt) < rank(i2, tt): Exit Function
    End If
    LeaderBefore = (StrComp(names(i1), names(i2), vbTextCompare) < 0)
End Function

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

    ' ---- Ranking assertions (spec section 8) ------------------
    Dim rkCount() As Long, rkMaxMW() As Double, rkCostMax() As Double
    Dim rkSlope() As Double, rkHasSlope() As Boolean
    Dim rankA() As Long, rankB() As Long
    ComputeRankings mw, names, costM, n, m, _
                    rkCount, rkMaxMW, rkCostMax, rkSlope, rkHasSlope, rankA, rankB

    Dim xCun As Long: xCun = NameIndex(names, n, "Cunningham")
    Dim xHob As Long: xHob = NameIndex(names, n, "Hobbs")
    Dim xCha As Long: xCha = NameIndex(names, n, "Chaves County")
    Dim xEC As Long:  xEC = NameIndex(names, n, "Eddy County")
    Dim xEN As Long:  xEN = NameIndex(names, n, "Eddy North")
    Dim xKio As Long: xKio = NameIndex(names, n, "Kiowa")
    Dim xOa As Long:  xOa = NameIndex(names, n, "Oasis")
    Dim xPH As Long:  xPH = NameIndex(names, n, "Pleasant Hill")
    Dim xRoo As Long: xRoo = NameIndex(names, n, "Roosevelt")
    Dim xCD As Long:  xCD = NameIndex(names, n, "China Draw")
    Dim xRR As Long:  xRR = NameIndex(names, n, "Roadrunner")

    ' $25MM: zero qualify, all ranks n/a
    Assert CountRanked(rankA, n, 1) = 0 And CountRanked(rankB, n, 1) = 0, _
           "$25MM: no substation ranked (all n/a)", passCount, failCount

    ' $50MM: 8 qualify
    Assert CountRanked(rankA, n, 2) = 8, "$50MM: 8 qualify", passCount, failCount
    Assert rankA(xCun, 2) = 1 And rankA(xHob, 2) = 1, "$50MM A: Cun/Hob rank 1", passCount, failCount
    Assert rankA(xEC, 2) = 3, "$50MM A: Eddy County rank 3", passCount, failCount
    Assert rankA(xEN, 2) = 4, "$50MM A: Eddy North rank 4", passCount, failCount
    Assert rankA(xCha, 2) = 5, "$50MM A: Chaves rank 5", passCount, failCount
    Assert rankA(xOa, 2) = 6 And rankA(xPH, 2) = 6 And rankA(xRoo, 2) = 6, _
           "$50MM A: Oasis/PH/Roosevelt tie rank 6", passCount, failCount
    Assert rkCount(xEC, 2) = 19 And rkCount(xEN, 2) = 18 And rkCount(xCha, 2) = 10, _
           "$50MM: EddyC 19 / EddyN 18 / Chaves 10 pts", passCount, failCount
    Assert rankB(xOa, 2) = 1 And rankB(xPH, 2) = 1 And rankB(xRoo, 2) = 1, _
           "$50MM B: Oasis/PH/Roosevelt tie rank 1", passCount, failCount
    Assert rankB(xCha, 2) = 4 And rankB(xEN, 2) = 5 And rankB(xEC, 2) = 6, _
           "$50MM B: Chaves 4 / EddyN 5 / EddyC 6", passCount, failCount
    Assert rankB(xCun, 2) = 7 And rankB(xHob, 2) = 7, "$50MM B: Cun/Hob tie rank 7", passCount, failCount
    Assert Round(rkSlope(xOa, 2), 1) = -2754.6 And Round(rkSlope(xCha, 2), 1) = -2531.7 _
       And Round(rkSlope(xEN, 2), 1) = -1691# And Round(rkSlope(xEC, 2), 1) = -1559.2 _
       And Round(rkSlope(xCun, 2), 1) = -1435.4, "$50MM B: slope values", passCount, failCount

    ' $75MM: 15 qualify
    Assert CountRanked(rankA, n, 3) = 15, "$75MM: 15 qualify", passCount, failCount
    Assert rkCount(xCun, 3) = 21 And rkCount(xEC, 3) = 21 And rkCount(xEN, 3) = 21 _
       And rkCount(xHob, 3) = 21 And rkCount(xKio, 3) = 21, "$75MM: five substations at 21 pts", passCount, failCount
    Assert rankA(xCun, 3) = 1 And rankA(xHob, 3) = 1, "$75MM A: Cun/Hob rank 1", passCount, failCount
    Assert rankA(xEC, 3) = 3 And rankA(xKio, 3) = 3, "$75MM A: EddyC/Kiowa tie rank 3 (equal cost@300)", passCount, failCount
    Assert rankA(xEN, 3) = 5, "$75MM A: Eddy North rank 5", passCount, failCount
    Assert rkCount(xCD, 3) = 8 And rankA(xCD, 3) = MaxRankVal(rankA, n, 3), _
           "$75MM A: China Draw last on breadth at 8 pts", passCount, failCount
    Assert rankB(xCD, 3) = 1 And Round(rkSlope(xCD, 3), 1) = -4227.6, _
           "$75MM B: China Draw first on slope -4227.6 (inversion)", passCount, failCount
    Assert rkCount(xCha, 3) = 10 And Round(rkSlope(xCha, 3), 1) = -2531.7, _
           "$75MM: Chaves 10 pts slope -2531.7", passCount, failCount

    ' $100MM: 17 qualify, 13 tie at 21 pts
    Assert CountRanked(rankA, n, 4) = 17, "$100MM: 17 qualify", passCount, failCount
    Dim c21 As Long: c21 = 0
    For i = 1 To n
        If rkCount(i, 4) = 21 Then c21 = c21 + 1
    Next i
    Assert c21 = 13, "$100MM: 13 tie at 21 qualifying points", passCount, failCount
    Assert rkCount(xRR, 4) = 12 And rankA(xRR, 4) = 16 And rankB(xRR, 4) = 1, _
           "$100MM: Roadrunner 16th breadth / 1st slope (inversion)", passCount, failCount
    Assert Round(rkSlope(xRR, 4), 1) = -4428.9, "$100MM B: Roadrunner -4428.9", passCount, failCount
    Assert rkCount(xCha, 4) = 11 And rankA(xCha, 4) = 17 And rankA(xCha, 4) = MaxRankVal(rankA, n, 4), _
           "$100MM A: Chaves last on breadth at 11 pts", passCount, failCount
    Assert rankB(xCha, 4) = 14 And Round(rkSlope(xCha, 4), 1) = -1287.8, _
           "$100MM B: Chaves 14th slope -1287.8", passCount, failCount
    Assert Round(rkSlope(xOa, 4), 1) = -730.3 And Round(rkSlope(xPH, 4), 1) = -687.7 _
       And Round(rkSlope(xRoo, 4), 1) = -687.7, "$100MM B: Oasis/PH/Roosevelt flattest", passCount, failCount

    ' Structural: non-na Rank A count == qualifier count; max rank <= count
    Dim tt As Long
    For tt = 1 To 4
        Dim qc As Long: qc = 0
        For i = 1 To n
            If rkCount(i, tt) >= 1 Then qc = qc + 1
        Next i
        Assert CountRanked(rankA, n, tt) = qc, _
               "Struct t" & tt & ": non-na Rank A count == qualifier count", passCount, failCount
        Assert MaxRankVal(rankA, n, tt) <= qc, _
               "Struct t" & tt & ": max Rank A <= qualifier count", passCount, failCount
    Next tt

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

Private Function CountRanked(ByRef rank() As Long, ByVal n As Long, ByVal tt As Long) As Long
    Dim i As Long, c As Long
    For i = 1 To n
        If rank(i, tt) > 0 Then c = c + 1
    Next i
    CountRanked = c
End Function

Private Function MaxRankVal(ByRef rank() As Long, ByVal n As Long, ByVal tt As Long) As Long
    Dim i As Long, mx As Long
    For i = 1 To n
        If rank(i, tt) > mx Then mx = rank(i, tt)
    Next i
    MaxRankVal = mx
End Function

' ------------------------------------------------------------
'  Fixture loader. Tries substation_cost_per_mw.csv in the workbook
'  folder; otherwise builds the hardcoded 17 x 21 reference matrix.
'
'  Each substation is a piecewise-constant total-cost profile
'  T(MW); cost per MW = T / MW. The tier totals are chosen so the
'  matrix reproduces every reference case in the ranking spec:
'    * Cunningham & Hobbs and Pleasant Hill & Roosevelt are identical
'      pairs (the only genuine ties left after the full tiebreak chain).
'    * Chaves County is the five-tier case.
'    * The $25MM band has no qualifier anywhere (global min = 47,633,000);
'      8 qualify at $50MM, 15 at $75MM, all 17 at $100MM with 13 tied at
'      21 qualifying points; China Draw and Roadrunner are the
'      breadth-vs-slope inversions.
'  Cielo..House are the fillers that complete those aggregate counts.
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

    ' name;  tier breakpoints (upper MW of each tier);  tier totals
    names(1) = "Cunningham":    SetProfile costM, mw, 1, m, Array(300), Array(47633000#)
    names(2) = "Hobbs":         SetProfile costM, mw, 2, m, Array(300), Array(47633000#)
    names(3) = "Chaves County": SetProfile costM, mw, 3, m, _
                    Array(190, 200, 230, 270, 300), _
                    Array(49473000#, 98003000#, 185783000#, 205353000#, 207843000#)
    names(4) = "Eddy County":   SetProfile costM, mw, 4, m, Array(280, 300), Array(47633000#, 64033000#)
    names(5) = "Eddy North":    SetProfile costM, mw, 5, m, Array(270, 300), Array(49473000#, 65873000#)
    names(6) = "Kiowa":         SetProfile costM, mw, 6, m, Array(300), Array(64033000#)
    names(7) = "Oasis":         SetProfile costM, mw, 7, m, Array(170, 240, 300), Array(47633000#, 60000000#, 79083142#)
    names(8) = "Pleasant Hill": SetProfile costM, mw, 8, m, Array(170, 240, 300), Array(47633000#, 60000000#, 81108681#)
    names(9) = "Roosevelt":     SetProfile costM, mw, 9, m, Array(170, 240, 300), Array(47633000#, 60000000#, 81108681#)
    names(10) = "China Draw":   SetProfile costM, mw, 10, m, Array(170, 230, 300), Array(73103436#, 92000000#, 150000000#)
    names(11) = "Roadrunner":   SetProfile costM, mw, 11, m, Array(210, 300), Array(96832165#, 160000000#)
    names(12) = "Cielo":        SetProfile costM, mw, 12, m, Array(250, 300), Array(71000000#, 86000000#)
    names(13) = "Datil":        SetProfile costM, mw, 13, m, Array(250, 300), Array(72000000#, 89000000#)
    names(14) = "Encino":       SetProfile costM, mw, 14, m, Array(250, 300), Array(73000000#, 91000000#)
    names(15) = "Fence Lake":   SetProfile costM, mw, 15, m, Array(250, 300), Array(74000000#, 95000000#)
    names(16) = "Grama":        SetProfile costM, mw, 16, m, Array(250, 300), Array(75000000#, 97000000#)
    names(17) = "House":        SetProfile costM, mw, 17, m, Array(220, 300), Array(90000000#, 140000000#)
End Sub

' Fill row i from a piecewise-constant tier profile: for each MW point,
' T is the total of the first tier whose upper-MW bound it falls within.
Private Sub SetProfile(ByRef costM() As Double, ByRef mw() As Double, ByVal i As Long, _
                       ByVal m As Long, ByVal ups As Variant, ByVal ts As Variant)
    Dim T() As Double: ReDim T(1 To m)
    Dim j As Long, k As Long
    For j = 1 To m
        For k = LBound(ups) To UBound(ups)
            If mw(j) <= CDbl(ups(k)) Then T(j) = CDbl(ts(k)): Exit For
        Next k
    Next j
    SetRow costM, mw, T, i, m
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
