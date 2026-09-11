Attribute VB_Name = "modInterconnectPipeline"
Option Explicit

' ==========================================================================
'  modInterconnectPipeline.bas
'  Interconnection Site-Scoring Pipeline -- scaled for ~2,000 substations
'
'  PUBLIC ENTRY POINT:  RunInterconnectPipeline   (run via Alt+F8)
'  TEST HARNESS:        SelfTest                   (in-memory + parity + scale)
'
'  Four stages, end to end from a blank workbook, on one Alt+F8 call:
'    1. Consolidate cost/results files -> Cost Data  (bulk, silent I/O)
'    2. Consolidate site files -> Site Data          (dictionary dedupe, O(N))
'    3. Join intersection -> in-memory grouping       (dictionary lookup)
'    4. Matrix (values) + analysis + ranking + scoring + headroom
'
'  =====================================================================
'  SCALING (what changed for ~2,000 files/substations; results identical)
'  ---------------------------------------------------------------------
'  Global run settings are set once at entry and restored in Cleanup on every
'  exit including error: ScreenUpdating=False, EnableEvents=False,
'  DisplayAlerts=False, Calculation=xlManual, AskToUpdateLinks=False, a live
'  Application.StatusBar progress line, and a SINGLE Application.CalculateFull
'  at the very end (no intermediate recalcs; the chart is built after it).
'
'  Stage 1  Every source workbook is opened UpdateLinks:=0, ReadOnly:=True,
'           AddToMru:=False and closed immediately (SaveChanges:=False); the
'           used range is read in one Range.Value grab and iterated in memory;
'           output rows are accumulated in an array and written one
'           Range.Value = array per file. One DoEvents + status line per file;
'           an unreadable file is logged and skipped, never aborting the batch.
'  Stage 2  Triple dedupe uses a Scripting.Dictionary keyed on a normalised
'           "name|voltage|state" string -- one O(N) pass, not O(N^2) compares.
'  Stage 3  The join builds a dictionary of cost-side identities in one pass
'           and resolves each Site Data triple by key lookup (name, or
'           name+voltage when the cost tab carried a voltage), not a rescan.
'  Stage 4  The matrix is computed as VALUES in VBA: the four cost columns are
'           loaded once and grouped by identity; each substation's
'           trigger->allocation records are scanned once to accumulate the
'           cumulative allocation T at each MW breakpoint, divided by MW -- the
'           exact piecewise T/MW the SUMIFS define -- and written as one block.
'  Stage 5  Ranking sorts once per column (stable mergesort, O(N log N)) and
'           assigns competition ranks in a single walk, instead of counting how
'           many beat each substation across eight columns (the ~32M-op hang).
'  Stage 6  Percentiles and the Weighted Score are evaluated in VBA against the
'           same definitions as the formulas and written as value blocks.
'  Stage 7  Headroom (MW) is the per-substation minimum trigger, taken from the
'           already-loaded cost data; its percentile is a raw-is-better bucket.
'
'  TWO SWITCHES (both default False = computed values; documented in README):
'    USE_LIVE_SUMIFS    True -> Stage 4 writes live SUMIFS bound to used rows
'                       ($D$2:$D$<last>, never $D:$D). False -> values block.
'    USE_LIVE_FORMULAS  True -> Stage 6/7 percentiles, Weighted Score and
'                       headroom are written as live formulas over BOUNDED
'                       population ranges. False -> computed value blocks.
'  PARITY GUARANTEE: the values path and the live-formula path produce the same
'  numbers. The VBA implements PERCENTRANK.EXC exactly (k/(N+1) positioning),
'  and SelfTest round-trips the fixture through real Excel formulas to prove
'  the two paths agree cell-for-cell.
'
'  SCORING (unchanged definitions; 165 ceiling unchanged):
'    Flattening / headroom / weighted-score percentiles are raw-is-better:
'        CEILING(PERCENTRANK.EXC(pop, x) * 5, 1)
'    The eight band percentiles (breadth+slope over four thresholds) are
'    rank-based (rank 1 = best):
'        IFERROR(CEILING((1 - PERCENTRANK.EXC(pop, rank)) * 5, 1), 0)
'    Weighted Score = Flattening + $B$2 * (sum of the eight band percentiles),
'    a single live weight $B$2 (default 4) over every band -- the threshold axis
'    as a whole is worth 4x the flattening axis; the weight does not grade
'    $25 > $50 > $75 > $100. Composite maximum = 5 + 4*(8*5) = 165. Headroom is
'    a standalone lens (its own percentile), NOT added to the Weighted Score.
'
'  No .Select / .Activate / Selection anywhere; every sheet write is a single
'  Range.Value/Formula = array per block; number formats applied per column
'  once. No external references (Scripting.Dictionary via late binding).
'  Target: Windows Excel, VBA 7.x.
'
'  SELF-TEST / VERIFICATION (SelfTest):
'    * Parity: values path == live-formula path on the fixture (percentile
'      buckets round-tripped through Excel).
'    * Ranking: sort-based competition ranks match the naive definition,
'      including the ties (Cunningham/Hobbs, Pleasant Hill/Roosevelt, the
'      13-way $100MM breadth tie).
'    * Dedupe/join: dictionary results match the triple and intersection rules.
'    * Scale: 2,000 synthetic substations rank/percentile/score without
'      O(N^2) blow-up (bounded wall-clock).
'    * Tab parsing, matrix shape, headroom, and the 165 ceiling.
' ==========================================================================

' ---------- Pipeline configuration (edit here only) ------------------------

Private Const PL_SH_COST     As String = "Cost Data"
Private Const PL_SH_SITE     As String = "Site Data"
Private Const PL_SH_MATRIX   As String = "Matrix"
Private Const PL_SH_ANALYSIS As String = "Cost Curve Analysis"
Private Const PL_SH_LOG      As String = "_Pipeline Log"

Private Const PL_HDR_ROW        As Long = 1
Private Const PL_COST_DATA_IDX  As Long = 3     ' cost workbook: sheet 3 carries the data
Private Const PL_SITE_TAB       As String = "Summary"
Private Const PL_SITE_NAME_COL  As Long = 1     ' Summary!A = substation name
Private Const PL_SITE_STATE_COL As Long = 3     ' Summary!C = state
Private Const PL_SITE_VOLT_COL  As Long = 7     ' Summary!G = voltage (kV)

' Source headers resolved at run time by NAME on Cost Data (never fixed
' letters -- the four metadata columns shift the source columns right).
Private Const PL_HDR_NAME    As String = "Substation Name"
Private Const PL_HDR_VOLT    As String = "Voltage (kV)"
Private Const PL_HDR_TRIGGER As String = "Size Overload Occurs (MW)"
Private Const PL_HDR_ALLOC   As String = "Proposed Project Allocation ($)"

' Matrix MW axis: 100, 110, ... 300 (numeric, strictly ascending).
Private Const PL_MW_MIN  As Double = 100
Private Const PL_MW_MAX  As Double = 300
Private Const PL_MW_STEP As Double = 10

' Weighted scoring.
Private Const PL_WEIGHT_B2  As Double = 4
Private Const PL_MAX_PCTILE As Long = 5

' Evaluation-path switches (see the header block). Both default False so the
' finished workbook holds values, not thousands of volatile array formulas.
Private Const USE_LIVE_SUMIFS   As Boolean = False   ' Stage 4 matrix
Private Const USE_LIVE_FORMULAS As Boolean = False   ' Stage 6/7 pctiles/score/headroom

Private Const PL_EXCEL_MAX_ROWS   As Long = 1048576
Private Const PL_EXCEL_MAX_TAB    As Long = 31

' ---------- Records ---------------------------------------------------------

Private Type PL_TCostFile
    FilePath      As String
    SheetName     As String
    Substation    As String
    Voltage       As Double
    VoltParsed    As Boolean
    HeaderSig     As String
    UsedCols      As Long
    FirstDataRow  As Long
    LastDataRow   As Long
    IsValid       As Boolean
    Status        As String
    Message       As String
    RowsImported  As Long
End Type

Private Type PL_TSiteTriple
    Name        As String
    Voltage     As Double
    VoltParsed  As Boolean
    State       As String
End Type

Private Type PL_TCostId
    Name        As String
    Voltage     As Double
    VoltParsed  As Boolean
    Matched     As Boolean
End Type

Private Type PL_TLog
    lines() As String
    n       As Long
End Type

' Everything Stage 3/4 hands to Stage 4's analysis + scoring, so the matrix is
' computed once in memory and never re-read from the sheet.
Private Type PL_TMatrix
    ok           As Boolean
    matchedCount As Long
    mwCount      As Long
    mw()         As Double        ' 1..m
    names()      As String        ' 1..n identity labels (row order)
    costM()      As Double        ' 1..n, 1..m cost-per-MW values
    headroom()   As Double        ' 1..n minimum trigger MW (0 = none)
    keyName()    As String        ' 1..n cost-side key (name)
    keyVolt()    As Double        ' 1..n cost-side voltage
    keyUseVolt() As Boolean       ' 1..n whether the cost tab carried a voltage
    costName     As String        ' Cost Data sheet name (for live formulas)
    colName      As Long          ' resolved Cost Data columns (by header)
    colVolt      As Long
    colTrig      As Long
    colAlloc     As Long
    costLastRow  As Long          ' last used row on Cost Data (for bounded live refs)
End Type

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
' ==========================================================================
'  PUBLIC ENTRY POINT
' ==========================================================================

'--------------------------------------------------------------------------
' RunInterconnectPipeline
'   Runs the four stages end to end. The four application flags plus
'   AskToUpdateLinks are set once here and restored in Cleanup on every exit
'   (including the error handler); a StatusBar progress line runs throughout;
'   there are no intermediate recalcs -- a single Application.CalculateFull is
'   issued at the very end of Stage 4, and the chart is built after it.
'--------------------------------------------------------------------------
Public Sub RunInterconnectPipeline()

    Dim savedScreen As Boolean, savedEvents As Boolean, savedAlerts As Boolean
    Dim savedCalc As XlCalculation, savedLinks As Boolean, savedStatusVis As Boolean
    Dim stateSaved As Boolean

    Dim log As PL_TLog
    Dim wsCost As Worksheet, wsSite As Worksheet, wsMatrix As Worksheet
    Dim mtx As PL_TMatrix
    Dim analysisOK As Boolean, distinctTriples As Long

    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedAlerts = Application.DisplayAlerts
    savedCalc = Application.Calculation
    savedLinks = Application.AskToUpdateLinks
    savedStatusVis = Application.DisplayStatusBar
    stateSaved = True

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AskToUpdateLinks = False
    Application.DisplayStatusBar = True

    pl_LogInit log
    pl_LogAdd log, "Run started " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    pl_Status "Interconnect pipeline: starting"

    ' ---- Step 1: cost / results files -> Cost Data --------------------
    If Not pl_ConsolidateCost(log, wsCost) Then
        pl_LogAdd log, "Step 1 did not complete; pipeline stopped."
        pl_WriteLog log
        MsgBox "Step 1 (Cost Data) did not complete. See " & PL_SH_LOG & ".", _
               vbExclamation, "Interconnect Pipeline"
        GoTo Cleanup
    End If

    ' ---- Step 2: dedupe / site files -> Site Data ---------------------
    If Not pl_ConsolidateSite(log, wsSite, distinctTriples) Then
        pl_LogAdd log, "Step 2 did not complete; pipeline stopped."
        pl_WriteLog log
        MsgBox "Step 2 (Site Data) did not complete. See " & PL_SH_LOG & ".", _
               vbExclamation, "Interconnect Pipeline"
        GoTo Cleanup
    End If

    ' ---- Step 3 + 4a: join + Matrix (values in memory) ----------------
    If Not pl_BuildMatrix(log, wsCost, wsSite, wsMatrix, mtx) Then
        pl_LogAdd log, "Step 3 did not complete; pipeline stopped."
        pl_WriteLog log
        MsgBox "Step 3 (Matrix) did not complete. See " & PL_SH_LOG & ".", _
               vbExclamation, "Interconnect Pipeline"
        GoTo Cleanup
    End If

    ' ---- Step 4b: analysis + ranking + scoring + headroom -------------
    ' pl_RunAnalysisAndScoring issues the single CalculateFull and builds the
    ' chart last. If it is skipped, resolve any live formulas once here.
    analysisOK = pl_RunAnalysisAndScoring(wsMatrix, log, mtx)
    If Not analysisOK Then Application.CalculateFull

    pl_WriteLog log
    pl_Status "Interconnect pipeline: done"

    MsgBox "Interconnection pipeline complete." & vbCrLf & vbCrLf & _
           "Cost Data:  " & wsCost.Name & vbCrLf & _
           "Site Data:  " & wsSite.Name & "  (" & distinctTriples & " distinct triples)" & vbCrLf & _
           "Matrix:     " & wsMatrix.Name & "  (" & mtx.matchedCount & " x " & mtx.mwCount & _
           ", " & IIf(USE_LIVE_SUMIFS, "live SUMIFS", "values") & ")" & vbCrLf & _
           "Analysis:   " & IIf(analysisOK, "written (" & IIf(USE_LIVE_FORMULAS, "live formulas", "values") & ")", "skipped -- see log") & vbCrLf & _
           "Run log:    " & PL_SH_LOG, _
           vbInformation, "Interconnect Pipeline"

Cleanup:
    If stateSaved Then
        Application.StatusBar = False
        Application.DisplayStatusBar = savedStatusVis
        Application.AskToUpdateLinks = savedLinks
        Application.Calculation = savedCalc
        Application.DisplayAlerts = savedAlerts
        Application.EnableEvents = savedEvents
        Application.ScreenUpdating = savedScreen
    End If
    Exit Sub

ErrHandler:
    On Error Resume Next
    pl_LogAdd log, "FATAL error #" & Err.Number & ": " & Err.Description
    pl_WriteLog log
    On Error GoTo 0
    MsgBox "Unexpected error #" & Err.Number & ": " & Err.Description, _
           vbCritical, "Interconnect Pipeline"
    Resume Cleanup
End Sub

' ==========================================================================
'  STAGE 1 -- CONSOLIDATE COST / RESULTS FILES  -> Cost Data  (bulk, silent)
' ==========================================================================

Private Function pl_ConsolidateCost(ByRef log As PL_TLog, ByRef wsCost As Worksheet) As Boolean
    On Error GoTo ErrHandler
    pl_ConsolidateCost = False

    Dim files() As String, fileCount As Long
    If Not pl_PickFiles("Select the COST / results workbooks to consolidate", files, fileCount) Then
        pl_LogAdd log, "Step 1: no cost files selected."
        Exit Function
    End If

    Dim recs() As PL_TCostFile
    ReDim recs(1 To fileCount)

    Dim i As Long, refSig As String, haveRef As Boolean
    For i = 1 To fileCount
        pl_Status "Stage 1: validating " & i & " of " & fileCount
        recs(i).FilePath = files(i)
        pl_ValidateCostFile recs(i)
        If recs(i).IsValid Then
            If Not haveRef Then
                refSig = recs(i).HeaderSig: haveRef = True
            ElseIf StrComp(recs(i).HeaderSig, refSig, vbTextCompare) <> 0 Then
                recs(i).IsValid = False
                recs(i).Status = "Failed"
                recs(i).Message = "Header signature mismatch vs first valid file"
            End If
        End If
        DoEvents
    Next i

    Dim passCount As Long, failCount As Long, mismatch As Long
    For i = 1 To fileCount
        If recs(i).IsValid Then
            passCount = passCount + 1
        Else
            failCount = failCount + 1
            If InStr(1, recs(i).Message, "Header signature mismatch", vbTextCompare) > 0 Then _
                mismatch = mismatch + 1
        End If
    Next i

    If passCount = 0 And mismatch = 0 Then
        pl_LogAdd log, "Step 1: no valid cost files (" & failCount & " failed)."
        MsgBox "No valid cost files to consolidate.", vbExclamation, "Interconnect Pipeline"
        Exit Function
    End If

    If mismatch > 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox(mismatch & " cost file(s) have a mismatched header signature." & vbCrLf & vbCrLf & _
                     "Yes  = INCLUDE them anyway, aligned by column position" & vbCrLf & _
                     "No   = SKIP the mismatched files" & vbCrLf & _
                     "Cancel = ABORT step 1", vbYesNoCancel + vbQuestion, "Cost header mismatch")
        Select Case ans
            Case vbCancel
                pl_LogAdd log, "Step 1: aborted at header-mismatch prompt.": Exit Function
            Case vbYes
                For i = 1 To fileCount
                    If (Not recs(i).IsValid) And _
                       InStr(1, recs(i).Message, "Header signature mismatch", vbTextCompare) > 0 Then
                        recs(i).IsValid = True: recs(i).Status = "Imported"
                        recs(i).Message = "Included despite header mismatch (aligned by column)"
                    End If
                Next i
        End Select
    End If

    Dim firstValid As Long: firstValid = 0
    For i = 1 To fileCount
        If recs(i).IsValid Then firstValid = i: Exit For
    Next i
    If firstValid = 0 Then pl_LogAdd log, "Step 1: no files remain.": Exit Function

    Set wsCost = pl_GetOutputSheet(PL_SH_COST)
    If wsCost Is Nothing Then pl_LogAdd log, "Step 1: cancelled at sheet creation.": Exit Function

    Dim headerCols As Long: headerCols = recs(firstValid).UsedCols
    wsCost.Cells(1, 1).Value = PL_HDR_NAME
    wsCost.Cells(1, 2).Value = PL_HDR_VOLT
    wsCost.Cells(1, 3).Value = "Source File"
    wsCost.Cells(1, 4).Value = "Source Sheet"
    pl_CopyCostHeaders wsCost, recs(firstValid)

    Dim nextRow As Long, totalRows As Long, processed As Long, hitLimit As Boolean
    nextRow = 2
    For i = 1 To fileCount
        If recs(i).IsValid Then
            pl_Status "Stage 1: consolidating " & i & " of " & fileCount & " -- " & pl_FileName(recs(i).FilePath)
            If pl_AppendCostFile(wsCost, recs(i), headerCols, nextRow) Then
                processed = processed + 1: totalRows = totalRows + recs(i).RowsImported
            Else
                hitLimit = True: Exit For
            End If
            DoEvents
        ElseIf recs(i).Status <> "Failed" Then
            recs(i).Status = "Skipped"
        End If
    Next i

    wsCost.Rows(1).Font.Bold = True
    wsCost.Columns.AutoFit

    pl_LogAdd log, "STAGE 1 -- Cost Data (" & wsCost.Name & "): " & processed & _
                   " file(s), " & totalRows & " data row(s)."
    For i = 1 To fileCount
        pl_LogAdd log, "  [" & recs(i).Status & "] " & pl_FileName(recs(i).FilePath) & _
                       " | tab='" & recs(i).SheetName & "' name='" & recs(i).Substation & "'" & _
                       IIf(recs(i).VoltParsed, " volt=" & recs(i).Voltage, " volt=(none)") & _
                       " rows=" & recs(i).RowsImported & _
                       IIf(Len(recs(i).Message) > 0, " -- " & recs(i).Message, "")
    Next i
    If hitLimit Then pl_LogAdd log, "  Worksheet row limit reached; import truncated."

    pl_ConsolidateCost = (totalRows > 0)
    Exit Function

ErrHandler:
    pl_LogAdd log, "Step 1 error #" & Err.Number & ": " & Err.Description
    pl_ConsolidateCost = False
End Function

Private Sub pl_ValidateCostFile(ByRef rec As PL_TCostFile)
    Dim wb As Workbook, ws As Worksheet
    rec.IsValid = False: rec.Status = "Failed"

    On Error GoTo OpenFail
    Set wb = Application.Workbooks.Open(Filename:=rec.FilePath, UpdateLinks:=0, _
                                        ReadOnly:=True, AddToMru:=False)
    On Error GoTo CloseFail

    If wb.Sheets.Count < PL_COST_DATA_IDX Then
        rec.Message = "Workbook has fewer than " & PL_COST_DATA_IDX & " sheets": GoTo CloseAndExit
    End If
    If Not TypeOf wb.Sheets(PL_COST_DATA_IDX) Is Worksheet Then
        rec.Message = "Sheet " & PL_COST_DATA_IDX & " is not a worksheet": GoTo CloseAndExit
    End If

    Set ws = wb.Sheets(PL_COST_DATA_IDX)
    rec.SheetName = ws.Name
    If Not pl_MeasureSheet(ws, rec) Then GoTo CloseAndExit

    rec.HeaderSig = pl_HeaderSig(ws, rec.UsedCols)
    pl_ParseTab rec.SheetName, rec.Substation, rec.Voltage, rec.VoltParsed
    rec.IsValid = True: rec.Status = "Imported": rec.Message = ""

CloseAndExit:
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Exit Sub
OpenFail:
    rec.Message = "Could not open: " & Err.Description: Exit Sub
CloseFail:
    rec.Message = "Error inspecting workbook: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function pl_MeasureSheet(ByVal ws As Worksheet, ByRef rec As PL_TCostFile) As Boolean
    Dim ur As Range, lastCol As Long, lastRow As Long
    pl_MeasureSheet = False
    Set ur = ws.UsedRange
    If ur Is Nothing Then rec.Message = "Sheet is empty": Exit Function
    lastCol = ur.Column + ur.Columns.Count - 1
    lastRow = ur.Row + ur.Rows.Count - 1
    If lastRow <= PL_HDR_ROW Then rec.Message = "No data rows below the header": Exit Function
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(PL_HDR_ROW, 1), _
            ws.Cells(PL_HDR_ROW, lastCol))) = 0 Then
        rec.Message = "Header row is empty": Exit Function
    End If
    rec.UsedCols = lastCol
    rec.FirstDataRow = PL_HDR_ROW + 1
    rec.LastDataRow = lastRow
    pl_MeasureSheet = True
End Function

Private Function pl_HeaderSig(ByVal ws As Worksheet, ByVal usedCols As Long) As String
    Dim c As Long, parts() As String
    ReDim parts(1 To usedCols)
    Dim hv As Variant
    hv = ws.Range(ws.Cells(PL_HDR_ROW, 1), ws.Cells(PL_HDR_ROW, usedCols)).Value
    For c = 1 To usedCols
        If usedCols = 1 Then parts(c) = Trim$(CStr(pl_NZ(hv))) Else parts(c) = Trim$(CStr(pl_NZ(hv(1, c))))
    Next c
    pl_HeaderSig = Join(parts, "|")
End Function

Private Sub pl_CopyCostHeaders(ByVal wsCost As Worksheet, ByRef rec As PL_TCostFile)
    Dim wb As Workbook, ws As Worksheet, c As Long, hv As Variant
    On Error GoTo Done
    Set wb = Application.Workbooks.Open(Filename:=rec.FilePath, UpdateLinks:=0, _
                                        ReadOnly:=True, AddToMru:=False)
    Set ws = wb.Sheets(PL_COST_DATA_IDX)
    hv = ws.Range(ws.Cells(PL_HDR_ROW, 1), ws.Cells(PL_HDR_ROW, rec.UsedCols)).Value
    Dim outHdr() As Variant: ReDim outHdr(1 To 1, 1 To rec.UsedCols)
    For c = 1 To rec.UsedCols
        If rec.UsedCols = 1 Then outHdr(1, c) = hv Else outHdr(1, c) = hv(1, c)
    Next c
    wsCost.Range(wsCost.Cells(1, 5), wsCost.Cells(1, 4 + rec.UsedCols)).Value = outHdr
    wb.Close SaveChanges:=False
    Exit Sub
Done:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Function pl_AppendCostFile(ByVal wsCost As Worksheet, ByRef rec As PL_TCostFile, _
                                   ByVal headerCols As Long, ByRef nextRow As Long) As Boolean
    Dim wb As Workbook, ws As Worksheet
    Dim srcVals As Variant, outBlock() As Variant
    Dim srcRows As Long, srcCols As Long, writeCols As Long
    Dim r As Long, c As Long, outR As Long

    pl_AppendCostFile = True: rec.RowsImported = 0
    On Error GoTo Fail
    Set wb = Application.Workbooks.Open(Filename:=rec.FilePath, UpdateLinks:=0, _
                                        ReadOnly:=True, AddToMru:=False)
    Set ws = wb.Sheets(PL_COST_DATA_IDX)

    srcVals = pl_BlockRead(ws, rec.FirstDataRow, rec.LastDataRow, rec.UsedCols)
    If IsEmpty(srcVals) Then wb.Close SaveChanges:=False: Exit Function

    srcRows = UBound(srcVals, 1) - LBound(srcVals, 1) + 1
    srcCols = UBound(srcVals, 2) - LBound(srcVals, 2) + 1
    writeCols = headerCols
    If srcCols < writeCols Then writeCols = srcCols

    If nextRow + srcRows - 1 > PL_EXCEL_MAX_ROWS Then
        wb.Close SaveChanges:=False: pl_AppendCostFile = False: Exit Function
    End If

    ReDim outBlock(1 To srcRows, 1 To 4 + writeCols)
    outR = 0
    For r = 1 To srcRows
        If Not pl_RowBlank(srcVals, r, srcCols) Then
            outR = outR + 1
            outBlock(outR, 1) = rec.Substation
            If rec.VoltParsed Then outBlock(outR, 2) = rec.Voltage Else outBlock(outR, 2) = ""
            outBlock(outR, 3) = rec.FilePath
            outBlock(outR, 4) = rec.SheetName
            For c = 1 To writeCols
                outBlock(outR, 4 + c) = srcVals(r, c)
            Next c
        End If
    Next r

    If outR > 0 Then
        wsCost.Range(wsCost.Cells(nextRow, 1), _
                     wsCost.Cells(nextRow + outR - 1, 4 + writeCols)).Value = _
            pl_TrimRows(outBlock, outR, 4 + writeCols)
        nextRow = nextRow + outR
        rec.RowsImported = outR
    End If

    wb.Close SaveChanges:=False
    Exit Function
Fail:
    rec.Status = "Failed": rec.Message = "Import error: " & Err.Description: rec.RowsImported = 0
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    pl_AppendCostFile = True     ' one file's failure never aborts the run
End Function

' ==========================================================================
'  STAGE 2 -- CONSOLIDATE SITE FILES -> Site Data  (dictionary dedupe, O(N))
' ==========================================================================

Private Function pl_ConsolidateSite(ByRef log As PL_TLog, ByRef wsSite As Worksheet, _
                                    ByRef distinctCount As Long) As Boolean
    On Error GoTo ErrHandler
    pl_ConsolidateSite = False

    Dim files() As String, fileCount As Long
    If Not pl_PickFiles("Select the SITE / dedupe workbooks to consolidate", files, fileCount) Then
        pl_LogAdd log, "Step 2: no site files selected.": Exit Function
    End If

    Dim rawName() As String, rawState() As String, rawVolt() As Double
    Dim rawVoltP() As Boolean, rawFile() As String, rawRow() As Long, rawN As Long
    rawN = 0
    ReDim rawName(1 To 16): ReDim rawState(1 To 16): ReDim rawVolt(1 To 16)
    ReDim rawVoltP(1 To 16): ReDim rawFile(1 To 16): ReDim rawRow(1 To 16)

    Dim i As Long
    For i = 1 To fileCount
        pl_Status "Stage 2: reading " & i & " of " & fileCount & " -- " & pl_FileName(files(i))
        pl_ReadSiteFile files(i), log, rawName, rawState, rawVolt, rawVoltP, rawFile, rawRow, rawN
        DoEvents
    Next i

    If rawN = 0 Then
        pl_LogAdd log, "Step 2: no site rows found."
        MsgBox "No usable rows found on the Summary tabs.", vbExclamation, "Interconnect Pipeline"
        Exit Function
    End If

    Dim dt() As PL_TSiteTriple, dropped As Long
    distinctCount = pl_DedupTriples(rawName, rawVolt, rawVoltP, rawState, rawN, _
                                    rawFile, rawRow, dt, log, dropped)

    Set wsSite = pl_GetOutputSheet(PL_SH_SITE)
    If wsSite Is Nothing Then pl_LogAdd log, "Step 2: cancelled at sheet creation.": Exit Function

    Dim block() As Variant: ReDim block(1 To distinctCount + 1, 1 To 3)
    block(1, 1) = PL_HDR_NAME: block(1, 2) = "State": block(1, 3) = PL_HDR_VOLT
    For i = 1 To distinctCount
        block(i + 1, 1) = dt(i).Name
        block(i + 1, 2) = dt(i).State
        If dt(i).VoltParsed Then block(i + 1, 3) = dt(i).Voltage Else block(i + 1, 3) = ""
    Next i
    wsSite.Range(wsSite.Cells(1, 1), wsSite.Cells(distinctCount + 1, 3)).Value = block
    wsSite.Rows(1).Font.Bold = True
    wsSite.Columns.AutoFit

    pl_LogAdd log, "STAGE 2 -- Site Data (" & wsSite.Name & "): " & rawN & " raw row(s), " & _
                   dropped & " duplicate(s) dropped, " & distinctCount & " distinct triple(s)."
    pl_ConsolidateSite = True
    Exit Function

ErrHandler:
    pl_LogAdd log, "Step 2 error #" & Err.Number & ": " & Err.Description
    pl_ConsolidateSite = False
End Function

Private Sub pl_ReadSiteFile(ByVal path As String, ByRef log As PL_TLog, _
                            ByRef rawName() As String, ByRef rawState() As String, _
                            ByRef rawVolt() As Double, ByRef rawVoltP() As Boolean, _
                            ByRef rawFile() As String, ByRef rawRow() As Long, ByRef rawN As Long)
    Dim wb As Workbook, ws As Worksheet, ur As Range, lastRow As Long, r As Long, vals As Variant
    On Error GoTo Fail
    Set wb = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)

    On Error Resume Next
    Set ws = wb.Worksheets(PL_SITE_TAB)
    On Error GoTo Fail
    If ws Is Nothing Then
        pl_LogAdd log, "  [Skipped] " & pl_FileName(path) & " -- no '" & PL_SITE_TAB & "' tab"
        wb.Close SaveChanges:=False: Exit Sub
    End If

    Set ur = ws.UsedRange
    lastRow = ur.Row + ur.Rows.Count - 1
    If lastRow <= PL_HDR_ROW Then
        pl_LogAdd log, "  [Skipped] " & pl_FileName(path) & " -- Summary has no data rows"
        wb.Close SaveChanges:=False: Exit Sub
    End If

    ' One bulk read of columns A..G (name, state, voltage live inside).
    vals = ws.Range(ws.Cells(PL_HDR_ROW + 1, 1), ws.Cells(lastRow, PL_SITE_VOLT_COL)).Value
    Dim added As Long: added = 0
    Dim rows As Long
    If IsArray(vals) Then rows = UBound(vals, 1) Else rows = 1

    For r = 1 To rows
        Dim nm As String, stt As String, vv As Variant
        If IsArray(vals) Then
            nm = Trim$(CStr(pl_NZ(vals(r, PL_SITE_NAME_COL))))
            stt = Trim$(CStr(pl_NZ(vals(r, PL_SITE_STATE_COL))))
            vv = vals(r, PL_SITE_VOLT_COL)
        Else
            nm = Trim$(CStr(pl_NZ(vals)))     ' single-cell degenerate
            stt = "": vv = ""
        End If
        If Len(nm) > 0 Then
            rawN = rawN + 1
            If rawN > UBound(rawName) Then pl_GrowRaw rawName, rawState, rawVolt, rawVoltP, rawFile, rawRow
            rawName(rawN) = nm
            rawState(rawN) = stt
            If IsNumeric(vv) And Len(Trim$(CStr(pl_NZ(vv)))) > 0 Then
                rawVolt(rawN) = CDbl(vv): rawVoltP(rawN) = True
            Else
                rawVolt(rawN) = 0: rawVoltP(rawN) = False
            End If
            rawFile(rawN) = pl_FileName(path)
            rawRow(rawN) = PL_HDR_ROW + r
            added = added + 1
        End If
    Next r

    pl_LogAdd log, "  [Read] " & pl_FileName(path) & " -- " & added & " Summary row(s)"
    wb.Close SaveChanges:=False
    Exit Sub
Fail:
    pl_LogAdd log, "  [Skipped] " & pl_FileName(path) & " -- read error: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
' pl_DedupTriples
'   O(N) dedupe via a Scripting.Dictionary keyed on the normalised
'   "name|voltage|state" triple. First occurrence kept; later matches logged as
'   dropped. Two rows collapse only when name AND voltage AND state all match.
'--------------------------------------------------------------------------
Private Function pl_DedupTriples(ByRef rawName() As String, ByRef rawVolt() As Double, _
                                 ByRef rawVoltP() As Boolean, ByRef rawState() As String, _
                                 ByVal rawN As Long, ByRef rawFile() As String, _
                                 ByRef rawRow() As Long, ByRef dt() As PL_TSiteTriple, _
                                 ByRef log As PL_TLog, ByRef dropped As Long) As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    ReDim dt(1 To rawN)
    Dim k As Long, i As Long, key As String
    k = 0: dropped = 0
    For i = 1 To rawN
        key = pl_TripleKey(rawName(i), rawVolt(i), rawVoltP(i), rawState(i))
        If d.Exists(key) Then
            dropped = dropped + 1
            pl_LogAdd log, "  [Dropped dup] " & rawFile(i) & " row " & rawRow(i) & " -- (" & _
                           pl_TripleStr(rawName(i), rawVolt(i), rawVoltP(i), rawState(i)) & ")"
        Else
            d.Add key, True
            k = k + 1
            dt(k).Name = rawName(i): dt(k).Voltage = rawVolt(i)
            dt(k).VoltParsed = rawVoltP(i): dt(k).State = rawState(i)
        End If
    Next i
    If k > 0 Then ReDim Preserve dt(1 To k)
    pl_DedupTriples = k
End Function

' ==========================================================================
'  STAGE 3 + 4a -- JOIN (dictionary) + MATRIX VALUES (in memory)
' ==========================================================================

'--------------------------------------------------------------------------
' pl_BuildMatrix
'   Loads the four cost columns once, groups records by identity in a
'   dictionary, resolves each Site Data triple by key lookup (intersection
'   only), and computes each matched substation's cost-per-MW curve as the
'   piecewise cumulative-allocation / MW -- the exact SUMIFS definition -- in
'   one pass per substation. Writes the Matrix as a single block (values, or
'   bounded live SUMIFS under USE_LIVE_SUMIFS) and returns everything Stage 4
'   needs in `mtx`, so the sheet is never re-read.
'--------------------------------------------------------------------------
Private Function pl_BuildMatrix(ByRef log As PL_TLog, ByVal wsCost As Worksheet, _
                                ByVal wsSite As Worksheet, ByRef wsMatrix As Worksheet, _
                                ByRef mtx As PL_TMatrix) As Boolean
    On Error GoTo ErrHandler
    pl_BuildMatrix = False
    mtx.ok = False
    pl_Status "Stage 3: joining cost and site sets"

    Dim colName As Long, colVolt As Long, colTrig As Long, colAlloc As Long
    colName = pl_ResolveCol(wsCost, PL_HDR_NAME)
    colVolt = pl_ResolveCol(wsCost, PL_HDR_VOLT)
    colTrig = pl_ResolveCol(wsCost, PL_HDR_TRIGGER)
    colAlloc = pl_ResolveCol(wsCost, PL_HDR_ALLOC)
    If colName = 0 Or colTrig = 0 Or colAlloc = 0 Then
        pl_LogAdd log, "Step 3: missing Cost Data headers (name/trigger/allocation)."
        MsgBox "Cost Data is missing a required header:" & vbCrLf & _
               PL_HDR_NAME & " / " & PL_HDR_TRIGGER & " / " & PL_HDR_ALLOC, _
               vbExclamation, "Interconnect Pipeline"
        Exit Function
    End If

    ' -- load cost data once --
    Dim lastCostRow As Long, lastCol As Long
    lastCostRow = wsCost.Cells(wsCost.Rows.Count, colName).End(xlUp).Row
    lastCol = wsCost.Cells(PL_HDR_ROW, wsCost.Columns.Count).End(xlToLeft).Column
    If lastCostRow < PL_HDR_ROW + 1 Then pl_LogAdd log, "Step 3: Cost Data has no rows.": Exit Function
    Dim cv As Variant
    cv = wsCost.Range(wsCost.Cells(1, 1), wsCost.Cells(lastCostRow, lastCol)).Value

    ' -- one pass over cost rows -->
    '    dId   : identity key (name [+voltage]) -> id, for the JOIN and to carry
    '            each matched row's exact criteria.
    '    dName : name(lower) -> Collection of Array(trig, alloc, vpFlag, volt),
    '            so the matrix/headroom sums apply the SAME criteria as the
    '            SUMIFS/MINIFS -- name+voltage for a voltage-bearing tab, name
    '            alone (across every voltage) for a bare tab.
    Dim dId As Object: Set dId = CreateObject("Scripting.Dictionary")
    Dim dName As Object: Set dName = CreateObject("Scripting.Dictionary")
    Dim idName() As String, idVolt() As Double, idVP() As Boolean
    Dim nId As Long: nId = 0
    ReDim idName(1 To lastCostRow): ReDim idVolt(1 To lastCostRow): ReDim idVP(1 To lastCostRow)

    Dim r As Long
    For r = 2 To lastCostRow
        Dim nm As String: nm = Trim$(CStr(pl_NZ(cv(r, colName))))
        If Len(nm) > 0 Then
            Dim vp As Boolean, vv As Double
            vp = False: vv = 0
            If colVolt > 0 Then
                If IsNumeric(cv(r, colVolt)) And Len(Trim$(CStr(pl_NZ(cv(r, colVolt))))) > 0 Then
                    vv = CDbl(cv(r, colVolt)): vp = True
                End If
            End If
            Dim tg As Variant, al As Variant
            tg = cv(r, colTrig): al = cv(r, colAlloc)
            If IsNumeric(tg) And IsNumeric(al) Then
                Dim key As String: key = pl_CostKey(nm, vv, vp)
                If Not dId.Exists(key) Then
                    nId = nId + 1
                    dId.Add key, nId
                    idName(nId) = nm: idVolt(nId) = vv: idVP(nId) = vp
                End If
                Dim nk As String: nk = LCase$(Trim$(nm))
                If Not dName.Exists(nk) Then
                    Dim col As Collection: Set col = New Collection
                    dName.Add nk, col
                End If
                dName(nk).Add Array(CDbl(tg), CDbl(al), IIf(vp, 1, 0), vv)
            End If
        End If
    Next r
    If nId = 0 Then pl_LogAdd log, "Step 3: no cost identities.": Exit Function

    ' -- site triples (bulk read) --
    Dim st() As PL_TSiteTriple, stN As Long
    stN = pl_ReadSiteTriples(wsSite, st)
    If stN = 0 Then pl_LogAdd log, "Step 3: no site triples.": Exit Function

    ' -- join: intersection only --
    Dim matchKey() As String: ReDim matchKey(1 To stN)
    Dim matched As Long: matched = 0
    Dim dMatched As Object: Set dMatched = CreateObject("Scripting.Dictionary")
    Dim s As Long, mk As String
    For s = 1 To stN
        mk = pl_ResolveMatchKey(st(s), dId)
        matchKey(s) = mk
        If Len(mk) > 0 Then
            matched = matched + 1
            If Not dMatched.Exists(mk) Then dMatched.Add mk, True
        Else
            pl_LogAdd log, "  [Align: site w/o cost] (" & _
                pl_TripleStr(st(s).Name, st(s).Voltage, st(s).VoltParsed, st(s).State) & ")"
        End If
    Next s
    ' cost identities never matched
    Dim c As Long
    For c = 1 To nId
        Dim ckey As String: ckey = pl_CostKey(idName(c), idVolt(c), idVP(c))
        If Not dMatched.Exists(ckey) Then
            pl_LogAdd log, "  [Align: cost w/o site] " & idName(c) & _
                IIf(idVP(c), " " & idVolt(c) & " kV", " (bare)")
        End If
    Next c

    If matched = 0 Then
        pl_LogAdd log, "Step 3: site and cost sets do not intersect (0 matches)."
        MsgBox "No substation is present in both sets; the Matrix would be empty.", _
               vbExclamation, "Interconnect Pipeline"
        Exit Function
    End If

    ' -- MW axis + allocate result --
    Dim mw() As Double, m As Long
    m = pl_MwAxis(mw)
    Dim n As Long: n = matched
    ReDim mtx.mw(1 To m)
    Dim j As Long
    For j = 1 To m: mtx.mw(j) = mw(j): Next j
    ReDim mtx.names(1 To n)
    ReDim mtx.costM(1 To n, 1 To m)
    ReDim mtx.headroom(1 To n)
    ReDim mtx.keyName(1 To n): ReDim mtx.keyVolt(1 To n): ReDim mtx.keyUseVolt(1 To n)

    ' -- compute each matched substation's curve (one pass over its records) --
    pl_Status "Stage 4: computing matrix values (" & n & " x " & m & ")"
    Dim outBlk() As Variant: ReDim outBlk(1 To n + 1, 1 To m + 1)
    outBlk(1, 1) = "Substation \ MW"
    For j = 1 To m: outBlk(1, j + 1) = mw(j): Next j

    Dim idIdx As Long, i As Long, rec As Variant, x As Double, tSum As Double, minTrig As Double
    Dim haveMin As Boolean, useV As Boolean, vSel As Double, inCrit As Boolean
    i = 0
    For s = 1 To stN
        If Len(matchKey(s)) > 0 Then
            i = i + 1
            idIdx = dId(matchKey(s))
            useV = idVP(idIdx): vSel = idVolt(idIdx)
            mtx.names(i) = pl_IdentityLabel(st(s).Name, st(s).Voltage, st(s).VoltParsed, st(s).State)
            mtx.keyName(i) = idName(idIdx): mtx.keyVolt(i) = vSel: mtx.keyUseVolt(i) = useV
            outBlk(i + 1, 1) = mtx.names(i)

            ' records for this NAME; apply the row's exact criteria (voltage
            ' only when the cost tab carried one) -- identical to the SUMIFS.
            Dim recCol As Collection: Set recCol = dName(LCase$(Trim$(idName(idIdx))))
            haveMin = False: minTrig = 0
            For j = 1 To m
                x = mw(j): tSum = 0
                For Each rec In recCol
                    inCrit = (Not useV) Or (rec(2) = 1 And rec(3) = vSel)
                    If inCrit Then If rec(0) <= x Then tSum = tSum + rec(1)
                Next rec
                mtx.costM(i, j) = tSum / x
            Next j
            For Each rec In recCol
                inCrit = (Not useV) Or (rec(2) = 1 And rec(3) = vSel)
                If inCrit And rec(0) > 0 Then
                    If Not haveMin Then minTrig = rec(0): haveMin = True ElseIf rec(0) < minTrig Then minTrig = rec(0)
                End If
            Next rec
            mtx.headroom(i) = IIf(haveMin, minTrig, 0)

            ' body cell content (value or bounded live SUMIFS)
            If USE_LIVE_SUMIFS Then
                For j = 1 To m
                    outBlk(i + 1, j + 1) = pl_SumifsFormula(wsCost.Name, colAlloc, colName, colTrig, colVolt, _
                                            idName(idIdx), useV, vSel, _
                                            pl_ColLetter(1 + j) & "$1", 2, lastCostRow)
                Next j
            Else
                For j = 1 To m
                    outBlk(i + 1, j + 1) = mtx.costM(i, j)
                Next j
            End If
        End If
    Next s

    ' -- write matrix in one block --
    Set wsMatrix = pl_GetOutputSheet(PL_SH_MATRIX)
    If wsMatrix Is Nothing Then pl_LogAdd log, "Step 3: cancelled at Matrix sheet creation.": Exit Function
    If USE_LIVE_SUMIFS Then
        wsMatrix.Range(wsMatrix.Cells(1, 1), wsMatrix.Cells(n + 1, m + 1)).Formula = outBlk
    Else
        wsMatrix.Range(wsMatrix.Cells(1, 1), wsMatrix.Cells(n + 1, m + 1)).Value = outBlk
    End If
    wsMatrix.Rows(1).Font.Bold = True
    wsMatrix.Columns(1).AutoFit
    wsMatrix.Range(wsMatrix.Cells(2, 2), wsMatrix.Cells(n + 1, m + 1)).NumberFormat = "$#,##0"

    mtx.matchedCount = n: mtx.mwCount = m
    mtx.costName = wsCost.Name
    mtx.colName = colName: mtx.colVolt = colVolt: mtx.colTrig = colTrig: mtx.colAlloc = colAlloc
    mtx.costLastRow = lastCostRow
    mtx.ok = True

    pl_LogAdd log, "STAGE 3/4 -- Matrix (" & wsMatrix.Name & "): " & n & " matched x " & m & _
                   " MW; cost cols by header (name=" & pl_ColLetter(colName) & ", trigger=" & _
                   pl_ColLetter(colTrig) & ", alloc=" & pl_ColLetter(colAlloc) & "); " & _
                   IIf(USE_LIVE_SUMIFS, "live SUMIFS", "values") & "."
    pl_BuildMatrix = True
    Exit Function

ErrHandler:
    pl_LogAdd log, "Step 3 error #" & Err.Number & ": " & Err.Description
    pl_BuildMatrix = False
End Function

Private Function pl_ReadSiteTriples(ByVal wsSite As Worksheet, ByRef st() As PL_TSiteTriple) As Long
    Dim lastRow As Long, r As Long, k As Long
    lastRow = wsSite.Cells(wsSite.Rows.Count, 1).End(xlUp).Row
    If lastRow < PL_HDR_ROW + 1 Then pl_ReadSiteTriples = 0: Exit Function
    Dim v As Variant
    v = wsSite.Range(wsSite.Cells(PL_HDR_ROW + 1, 1), wsSite.Cells(lastRow, 3)).Value
    Dim rows As Long
    If IsArray(v) Then rows = UBound(v, 1) Else rows = 1
    ReDim st(1 To rows)
    k = 0
    For r = 1 To rows
        Dim nm As String, stt As String, vv As Variant
        If IsArray(v) Then
            nm = Trim$(CStr(pl_NZ(v(r, 1)))): stt = Trim$(CStr(pl_NZ(v(r, 2)))): vv = v(r, 3)
        Else
            nm = Trim$(CStr(pl_NZ(v))): stt = "": vv = ""
        End If
        If Len(nm) > 0 Then
            k = k + 1
            st(k).Name = nm: st(k).State = stt
            If IsNumeric(vv) And Len(Trim$(CStr(pl_NZ(vv)))) > 0 Then
                st(k).Voltage = CDbl(vv): st(k).VoltParsed = True
            Else
                st(k).Voltage = 0: st(k).VoltParsed = False
            End If
        End If
    Next r
    If k > 0 Then ReDim Preserve st(1 To k)
    pl_ReadSiteTriples = k
End Function

'--------------------------------------------------------------------------
' pl_ResolveMatchKey
'   The join rule as a dictionary lookup: match a site triple to a cost
'   identity by name+voltage when a voltage-bearing cost tab exists, else by
'   name alone (bare cost tab). Returns the cost identity key, or "" if none.
'--------------------------------------------------------------------------
Private Function pl_ResolveMatchKey(ByRef tr As PL_TSiteTriple, ByVal dId As Object) As String
    Dim vk As String, bk As String
    If tr.VoltParsed Then
        vk = pl_CostKey(tr.Name, tr.Voltage, True)
        If dId.Exists(vk) Then pl_ResolveMatchKey = vk: Exit Function
    End If
    bk = pl_CostKey(tr.Name, 0, False)
    If dId.Exists(bk) Then pl_ResolveMatchKey = bk: Exit Function
    pl_ResolveMatchKey = ""
End Function

' Identity key for a cost tab: name (+voltage when the tab carried one).
Private Function pl_CostKey(ByVal nm As String, ByVal v As Double, ByVal vp As Boolean) As String
    pl_CostKey = LCase$(Trim$(nm)) & "|" & IIf(vp, pl_NumStr(v), "")
End Function

' ==========================================================================
'  SUMIFS / MINIFS FORMULA BUILDERS (live paths, bounded ranges)
' ==========================================================================

Private Function pl_SumifsFormula(ByVal costName As String, ByVal colAlloc As Long, _
                                  ByVal colName As Long, ByVal colTrig As Long, _
                                  ByVal colVolt As Long, ByVal subName As String, _
                                  ByVal useVolt As Boolean, ByVal theVolt As Double, _
                                  ByVal mwCell As String, ByVal r1 As Long, ByVal r2 As Long) As String
    Dim ref As String, f As String
    ref = pl_SheetRef(costName)
    f = "=SUMIFS(" & ref & "!" & pl_BoundCol(colAlloc, r1, r2) & _
        "," & ref & "!" & pl_BoundCol(colName, r1, r2) & ",""" & pl_EscQuote(subName) & """" & _
        "," & ref & "!" & pl_BoundCol(colTrig, r1, r2) & ",""<=""&" & mwCell
    If useVolt And colVolt > 0 Then
        f = f & "," & ref & "!" & pl_BoundCol(colVolt, r1, r2) & "," & pl_NumStr(theVolt)
    End If
    f = f & ")/" & mwCell
    pl_SumifsFormula = f
End Function

Private Function pl_MinifsFormula(ByVal costName As String, ByVal colTrig As Long, _
                                  ByVal colName As Long, ByVal colVolt As Long, _
                                  ByVal subName As String, ByVal useVolt As Boolean, _
                                  ByVal theVolt As Double, ByVal r1 As Long, ByVal r2 As Long) As String
    Dim ref As String, f As String
    ref = pl_SheetRef(costName)
    f = "MINIFS(" & ref & "!" & pl_BoundCol(colTrig, r1, r2) & _
        "," & ref & "!" & pl_BoundCol(colName, r1, r2) & ",""" & pl_EscQuote(subName) & """"
    If useVolt And colVolt > 0 Then
        f = f & "," & ref & "!" & pl_BoundCol(colVolt, r1, r2) & "," & pl_NumStr(theVolt)
    End If
    f = f & ")"
    pl_MinifsFormula = f
End Function

' ==========================================================================
'  STAGE 4b -- ANALYSIS + RANKING + SCORING + HEADROOM
' ==========================================================================

Private Function pl_RunAnalysisAndScoring(ByVal wsMatrix As Worksheet, ByRef log As PL_TLog, _
                                          ByRef mtx As PL_TMatrix) As Boolean
    On Error GoTo ErrHandler
    pl_RunAnalysisAndScoring = False

    Dim mw() As Double, names() As String, costM() As Double, n As Long, m As Long
    mw = mtx.mw: names = mtx.names: costM = mtx.costM
    n = mtx.matchedCount: m = mtx.mwCount

    ' validate the cost-per-MW block in memory (the analysis contract: > 0)
    Dim i As Long, j As Long
    For i = 1 To n
        For j = 1 To m
            If costM(i, j) <= 0 Then
                pl_LogAdd log, "Step 4: non-positive cost at '" & names(i) & "', MW " & mw(j) & _
                               " (no upgrade priced at/under this size); analysis skipped."
                MsgBox "Matrix has a non-positive cost-per-MW at '" & names(i) & "', " & mw(j) & _
                       " MW. The cost-curve analysis needs every cell > 0, so it was skipped." & _
                       vbCrLf & "(Steps 1-3 completed; see " & PL_SH_LOG & ".)", _
                       vbExclamation, "Interconnect Pipeline"
                Exit Function
            End If
        Next j
    Next i

    pl_Status "Stage 4: ranking " & n & " substations"
    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1
    Dim totalCols As Long: totalCols = 7 + 4 * nt

    Dim rkCount() As Long, rkMaxMW() As Double, rkCostMax() As Double
    Dim rkSlope() As Double, rkHasSlope() As Boolean, rankA() As Long, rankB() As Long
    ComputeRankings mw, names, costM, n, m, rkCount, rkMaxMW, rkCostMax, rkSlope, rkHasSlope, rankA, rankB

    Dim aTable() As Variant: ReDim aTable(1 To n, 1 To totalCols)
    AnalyseAll mw, names, costM, n, m, aTable, rankA, rankB

    Dim wsOut As Worksheet: Set wsOut = pl_GetOutputSheet(PL_SH_ANALYSIS)
    If wsOut Is Nothing Then pl_LogAdd log, "Step 4: cancelled at analysis sheet creation.": Exit Function

    pl_Status "Stage 4: writing analysis"
    Dim srcMWRow As Long, srcFirstDataRow As Long
    Dim hlpMWCol As Long, hlpFirstThreshCol As Long, hlpFirstRow As Long, hlpRows As Long
    WriteSheet wsOut, mw, names, costM, n, m, aTable, totalCols, _
               rkCount, rkMaxMW, rkSlope, rankA, rankB, _
               srcMWRow, srcFirstDataRow, hlpMWCol, hlpFirstThreshCol, hlpFirstRow, hlpRows

    pl_Status "Stage 4: scoring"
    pl_WriteScoring wsOut, mw, names, costM, n, m, rankA, rankB, totalCols, mtx

    ' single recompute at the very end, then the chart is built last
    pl_Status "Stage 4: recalculating"
    Application.CalculateFull
    pl_Status "Stage 4: chart"
    BuildChart wsOut, mw, costM, n, m, srcMWRow, srcFirstDataRow, _
               hlpMWCol, hlpFirstThreshCol, hlpFirstRow, hlpRows

    pl_LogAdd log, "STAGE 4 -- Cost Curve Analysis (" & wsOut.Name & "): " & n & _
                   " substation(s), " & m & " MW; weighted-score ceiling 165, $B$2=" & PL_WEIGHT_B2 & _
                   "; " & IIf(USE_LIVE_FORMULAS, "live formulas", "values") & "."
    pl_RunAnalysisAndScoring = True
    Exit Function

ErrHandler:
    pl_LogAdd log, "Step 4 error #" & Err.Number & ": " & Err.Description
    MsgBox "Error during analysis/scoring #" & Err.Number & ": " & Err.Description, _
           vbCritical, "Interconnect Pipeline"
    pl_RunAnalysisAndScoring = False
End Function

'--------------------------------------------------------------------------
' pl_WriteScoring
'   Writes the live weight cell $B$2 and the scoring block: Flattening pctile,
'   per-band Breadth/Slope pctile + Band Score, Weighted Score, Weighted Score
'   pctile, Headroom (MW), Headroom pctile. Percentiles/score/headroom are
'   computed in VBA (values path, default) or written as bounded live formulas
'   (USE_LIVE_FORMULAS); both paths agree. Single block write.
'--------------------------------------------------------------------------
Private Sub pl_WriteScoring(ByVal ws As Worksheet, ByRef mw() As Double, _
                            ByRef names() As String, ByRef costM() As Double, _
                            ByVal n As Long, ByVal m As Long, _
                            ByRef rankA() As Long, ByRef rankB() As Long, ByVal totalCols As Long, _
                            ByRef mtx As PL_TMatrix)
    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1
    Dim i As Long, tt As Long, j As Long
    Dim EMPTY_LIT As String: EMPTY_LIT = Chr$(34) & Chr$(34)

    ' -- live weight cell --
    ws.Cells(2, 1).Value = "Threshold weight (applies to every band):"
    ws.Cells(2, 2).Value = PL_WEIGHT_B2
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 1).Font.Italic = True

    ' ==== VBA metric computation (used for the values path; the parity self-
    '      test proves these equal the live-formula path) ====
    ' Flattening raw = the geometric knee, rounded to 1 dp to match Block A col E.
    Dim knee() As Double: ReDim knee(1 To n)
    Dim c() As Double: ReDim c(1 To m)
    Dim T() As Double: ReDim T(1 To m)
    For i = 1 To n
        For j = 1 To m: c(j) = costM(i, j): T(j) = c(j) * mw(j): Next j
        Dim segs() As TSegment: segs = Segmentize(mw, T, m)
        Dim li As Long: li = LongestSegIdx(mw, segs, UBound(segs))
        knee(i) = Round(KneeXStar(mw, T, segs(li)), 1)
    Next i
    Dim flatPct() As Long: flatPct = pl_RawBucketArray(knee, n)

    ' Rank-based band percentiles (rank 1 = best).
    Dim breadthPct() As Long, slopePct() As Long
    ReDim breadthPct(1 To n, 1 To nt): ReDim slopePct(1 To n, 1 To nt)
    Dim col1() As Long
    For tt = 1 To nt
        col1 = pl_RankBucketArray(rankA, tt, n)
        For i = 1 To n: breadthPct(i, tt) = col1(i): Next i
        col1 = pl_RankBucketArray(rankB, tt, n)
        For i = 1 To n: slopePct(i, tt) = col1(i): Next i
    Next tt

    ' Band score + weighted score (value).
    Dim weighted() As Double: ReDim weighted(1 To n)
    Dim bandScore() As Double: ReDim bandScore(1 To n, 1 To nt)
    For i = 1 To n
        weighted(i) = flatPct(i)
        For tt = 1 To nt
            bandScore(i, tt) = PL_WEIGHT_B2 * (breadthPct(i, tt) + slopePct(i, tt))
            weighted(i) = weighted(i) + bandScore(i, tt)
        Next tt
    Next i
    Dim wsPct() As Long: wsPct = pl_RawBucketArray(weighted, n)

    ' Headroom raw + percentile (blank/excluded when 0).
    Dim headroom() As Double: ReDim headroom(1 To n)
    Dim part() As Boolean: ReDim part(1 To n)
    For i = 1 To n
        headroom(i) = mtx.headroom(i)
        part(i) = (headroom(i) > 0)
    Next i
    Dim hrPct() As Long: hrPct = pl_RawBucketArrayMasked(headroom, part, n)

    ' ==== column layout ====
    Dim c0 As Long: c0 = totalCols + 2         ' spacer after Block A
    Dim r0 As Long: r0 = 3                     ' align with Block A header row
    Dim scFlat As Long: scFlat = c0 + 1
    Dim scWS As Long: scWS = c0 + 2 + 3 * nt
    Dim scWSP As Long: scWSP = scWS + 1
    Dim scHR As Long: scHR = scWS + 2
    Dim scHRP As Long: scHRP = scWS + 3
    Dim nCols As Long: nCols = scHRP - c0 + 1

    ' Block A anchors for live-path population ranges.
    Dim dataTop As Long: dataTop = 4          ' TABLE_HDR_ROW + 1
    Dim dataBot As Long: dataBot = dataTop + n - 1
    Dim firstDR As Long: firstDR = r0 + 1
    Dim lastDR As Long: lastDR = r0 + n

    ' ==== build the block (values or formula strings) ====
    Dim blk() As Variant: ReDim blk(1 To n + 1, 1 To nCols)
    blk(1, 1) = "Substation"
    blk(1, 2) = "Flattening Pctile (1-5)"
    For tt = 0 To nt - 1
        Dim bo As Long: bo = 3 + 3 * tt        ' 1-based block col of this band's breadth
        blk(1, bo) = ThreshLabel(CDbl(thr(LBound(thr) + tt))) & " Breadth (0-5)"
        blk(1, bo + 1) = ThreshLabel(CDbl(thr(LBound(thr) + tt))) & " Slope (0-5)"
        blk(1, bo + 2) = ThreshLabel(CDbl(thr(LBound(thr) + tt))) & " Band Score"
    Next tt
    blk(1, scWS - c0 + 1) = "Weighted Score (max 165)"
    blk(1, scWSP - c0 + 1) = "Weighted Score Pctile (1-5)"
    blk(1, scHR - c0 + 1) = "Headroom (MW)"
    blk(1, scHRP - c0 + 1) = "Headroom Pctile (1-5)"

    Dim r As Long
    For i = 1 To n
        r = i + 1
        Dim rr As Long: rr = r0 + i
        blk(r, 1) = names(i)

        If USE_LIVE_FORMULAS Then
            ' Flattening pctile: raw-is-better over Block A knee column (E = col 5).
            blk(r, 2) = pl_RawPctFormula("E", dataTop, dataBot, "E" & rr)
            Dim sumRefs As String: sumRefs = ""
            For tt = 0 To nt - 1
                Dim bcol As Long: bcol = 3 + 3 * tt
                Dim rankAcol As String: rankAcol = pl_ColLetter(10 + 4 * tt)   ' Block A Rank-by-breadth
                Dim rankBcol As String: rankBcol = pl_ColLetter(11 + 4 * tt)   ' Block A Rank-by-slope
                blk(r, bcol) = pl_RankPctFormula(rankAcol, dataTop, dataBot, rankAcol & rr)
                blk(r, bcol + 1) = pl_RankPctFormula(rankBcol, dataTop, dataBot, rankBcol & rr)
                blk(r, bcol + 2) = "=$B$2*(" & pl_ColLetter(c0 + bcol - 1) & rr & "+" & _
                                   pl_ColLetter(c0 + bcol) & rr & ")"
                If Len(sumRefs) > 0 Then sumRefs = sumRefs & "+"
                sumRefs = sumRefs & pl_ColLetter(c0 + bcol + 1) & rr
            Next tt
            blk(r, scWS - c0 + 1) = "=" & pl_ColLetter(scFlat) & rr & "+" & sumRefs
            blk(r, scWSP - c0 + 1) = pl_RawPctFormula(pl_ColLetter(scWS), firstDR, lastDR, pl_ColLetter(scWS) & rr)
            blk(r, scHR - c0 + 1) = "=IFERROR(IF(" & _
                pl_MinifsFormula(mtx.costName, mtx.colTrig, mtx.colName, mtx.colVolt, _
                                 mtx.keyName(i), mtx.keyUseVolt(i), mtx.keyVolt(i), 2, mtx.costLastRow) & _
                "=0," & EMPTY_LIT & "," & _
                pl_MinifsFormula(mtx.costName, mtx.colTrig, mtx.colName, mtx.colVolt, _
                                 mtx.keyName(i), mtx.keyUseVolt(i), mtx.keyVolt(i), 2, mtx.costLastRow) & _
                ")," & EMPTY_LIT & ")"
            blk(r, scHRP - c0 + 1) = "=IFERROR(" & _
                Mid$(pl_RawPctFormula(pl_ColLetter(scHR), firstDR, lastDR, pl_ColLetter(scHR) & rr), 2) & _
                "," & EMPTY_LIT & ")"
        Else
            blk(r, 2) = flatPct(i)
            For tt = 1 To nt
                Dim bo2 As Long: bo2 = 3 + 3 * (tt - 1)
                blk(r, bo2) = breadthPct(i, tt)
                blk(r, bo2 + 1) = slopePct(i, tt)
                blk(r, bo2 + 2) = bandScore(i, tt)
            Next tt
            blk(r, scWS - c0 + 1) = weighted(i)
            blk(r, scWSP - c0 + 1) = wsPct(i)
            If part(i) Then
                blk(r, scHR - c0 + 1) = headroom(i)
                blk(r, scHRP - c0 + 1) = hrPct(i)
            Else
                blk(r, scHR - c0 + 1) = ""
                blk(r, scHRP - c0 + 1) = ""
            End If
        End If
    Next i

    If USE_LIVE_FORMULAS Then
        ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + n, scHRP)).Formula = blk
    Else
        ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + n, scHRP)).Value = blk
    End If

    ' -- formats (whole columns once) --
    ws.Range(ws.Cells(r0 + 1, scHR), ws.Cells(r0 + n, scHR)).NumberFormat = "#,##0"
    With ws.Range(ws.Cells(r0 + 1, scHRP), ws.Cells(r0 + n, scHRP))
        .NumberFormat = "0": .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(r0 + 1, scWSP), ws.Cells(r0 + n, scWSP))
        .NumberFormat = "0": .HorizontalAlignment = xlCenter
    End With

    ' -- header styling + note --
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0, scHRP)).Font.Bold = True
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0, scHRP)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(r0 + n + 2, c0).Value = _
        "Weighted Score max = 165 = flattening (5) + 4 bands x $B$2 x (breadth 5 + slope 5). " & _
        "Practical max today = 125 ($25MM band scores 0 for all). Single $B$2 weights the " & _
        "threshold axis as a whole. Headroom (MW) + its percentile are standalone -- NOT scored."
    ws.Cells(r0 + n + 2, c0).Font.Italic = True
End Sub

' ==========================================================================
'  SCORING PRIMITIVES  (VBA mirror of the sheet formulas; parity-tested)
' ==========================================================================

Private Function pl_ScoreCeiling(ByVal nBands As Long) As Double
    pl_ScoreCeiling = PL_MAX_PCTILE + nBands * PL_WEIGHT_B2 * (2 * PL_MAX_PCTILE)
End Function

' PERCENTRANK.EXC(pop, x) for x present in pop: (strictly-less count + 1)/(N+1),
' matching Excel's first-occurrence positioning k/(N+1).
Private Function pl_PercentRankExc(ByRef pop() As Double, ByVal nPop As Long, ByVal x As Double) As Double
    Dim i As Long, lessCount As Long
    For i = 1 To nPop
        If pop(i) < x Then lessCount = lessCount + 1
    Next i
    pl_PercentRankExc = (lessCount + 1) / (nPop + 1)
End Function

Private Function pl_Ceil1(ByVal x As Double) As Long
    pl_Ceil1 = -Int(-x)
End Function

' Single-value raw-is-better bucket: CEILING(PERCENTRANK.EXC(pop, x)*5, 1).
Private Function pl_PctExcCeil5(ByRef pop() As Double, ByVal nPop As Long, ByVal x As Double) As Long
    If nPop < 1 Then pl_PctExcCeil5 = 0: Exit Function
    pl_PctExcCeil5 = pl_Ceil1(pl_PercentRankExc(pop, nPop, x) * PL_MAX_PCTILE)
End Function

' Raw-is-better buckets for a whole column, O(n log n) via one sort + grouped
' walk (no per-element rescan).
Private Function pl_RawBucketArray(ByRef vals() As Double, ByVal n As Long) As Long()
    Dim bucket() As Long: ReDim bucket(1 To pl_Max1(n))
    If n < 1 Then pl_RawBucketArray = bucket: Exit Function
    Dim idx() As Long: idx = pl_SortIdxAscD(vals, n)
    Dim pos As Long, groupStart As Long, s0 As Long
    Dim val As Double
    groupStart = 1
    For pos = 1 To n
        If pos > 1 Then
            If vals(idx(pos)) <> vals(idx(pos - 1)) Then groupStart = pos
        End If
        s0 = groupStart - 1                         ' strictly-less count
        val = (CDbl(s0 + 1) / CDbl(n + 1)) * PL_MAX_PCTILE
        bucket(idx(pos)) = pl_Ceil1(val)
    Next pos
    pl_RawBucketArray = bucket
End Function

' Same as pl_RawBucketArray but only `part` members form the population; others
' get 0 (they are shown blank by the caller).
Private Function pl_RawBucketArrayMasked(ByRef vals() As Double, ByRef part() As Boolean, _
                                         ByVal n As Long) As Long()
    Dim bucket() As Long: ReDim bucket(1 To pl_Max1(n))
    Dim pop() As Double, mapIdx() As Long, np As Long, i As Long
    ReDim pop(1 To pl_Max1(n)): ReDim mapIdx(1 To pl_Max1(n))
    np = 0
    For i = 1 To n
        If part(i) Then np = np + 1: pop(np) = vals(i): mapIdx(np) = i
    Next i
    If np < 1 Then pl_RawBucketArrayMasked = bucket: Exit Function
    Dim sub_() As Double: ReDim sub_(1 To np)
    For i = 1 To np: sub_(i) = pop(i): Next i
    Dim b() As Long: b = pl_RawBucketArray(sub_, np)
    For i = 1 To np: bucket(mapIdx(i)) = b(i): Next i
    pl_RawBucketArrayMasked = bucket
End Function

' Rank-based buckets for one threshold column: qualifiers (rank>0) form the
' population; IFERROR(CEILING((1 - PERCENTRANK.EXC(pop, rank))*5,1),0).
Private Function pl_RankBucketArray(ByRef rankArr() As Long, ByVal tt As Long, ByVal n As Long) As Long()
    Dim bucket() As Long: ReDim bucket(1 To pl_Max1(n))
    Dim pop() As Double, mapIdx() As Long, nq As Long, i As Long
    ReDim pop(1 To pl_Max1(n)): ReDim mapIdx(1 To pl_Max1(n))
    nq = 0
    For i = 1 To n
        If rankArr(i, tt) > 0 Then nq = nq + 1: pop(nq) = CDbl(rankArr(i, tt)): mapIdx(nq) = i
    Next i
    If nq < 1 Then pl_RankBucketArray = bucket: Exit Function       ' all 0 (IFERROR -> 0)
    Dim sub_() As Double: ReDim sub_(1 To nq)
    For i = 1 To nq: sub_(i) = pop(i): Next i
    Dim idx() As Long: idx = pl_SortIdxAscD(sub_, nq)
    Dim pos As Long, groupStart As Long, s0 As Long, pctExc As Double
    groupStart = 1
    For pos = 1 To nq
        If pos > 1 Then
            If sub_(idx(pos)) <> sub_(idx(pos - 1)) Then groupStart = pos
        End If
        s0 = groupStart - 1
        pctExc = CDbl(s0 + 1) / CDbl(nq + 1)
        bucket(mapIdx(idx(pos))) = pl_Ceil1((1 - pctExc) * PL_MAX_PCTILE)
    Next pos
    pl_RankBucketArray = bucket
End Function

' Stable bottom-up mergesort of indices of a Double array, ascending. O(n log n).
Private Function pl_SortIdxAscD(ByRef vals() As Double, ByVal n As Long) As Long()
    Dim ord() As Long: ReDim ord(1 To pl_Max1(n))
    Dim i As Long
    For i = 1 To n: ord(i) = i: Next i
    If n < 2 Then pl_SortIdxAscD = ord: Exit Function
    Dim buf() As Long: ReDim buf(1 To n)
    Dim width As Long, s As Long
    width = 1
    Do While width < n
        s = 1
        Do While s <= n
            Dim l1 As Long, r1 As Long, l2 As Long, r2 As Long, p As Long, q As Long, w As Long
            l1 = s: r1 = s + width - 1
            If r1 > n Then r1 = n
            l2 = r1 + 1: r2 = s + 2 * width - 1
            If r2 > n Then r2 = n
            p = l1: q = l2: w = l1
            Do While p <= r1 And q <= r2
                If vals(ord(q)) < vals(ord(p)) Then
                    buf(w) = ord(q): q = q + 1
                Else
                    buf(w) = ord(p): p = p + 1
                End If
                w = w + 1
            Loop
            Do While p <= r1
                buf(w) = ord(p): p = p + 1: w = w + 1
            Loop
            Do While q <= r2
                buf(w) = ord(q): q = q + 1: w = w + 1
            Loop
            For w = l1 To r2
                ord(w) = buf(w)
            Next w
            s = s + 2 * width
        Loop
        width = width * 2
    Loop
    pl_SortIdxAscD = ord
End Function

Private Function pl_Max1(ByVal n As Long) As Long
    If n < 1 Then pl_Max1 = 1 Else pl_Max1 = n
End Function

' Live-formula fragments (bounded population ranges).
Private Function pl_RawPctFormula(ByVal colLtr As String, ByVal r1 As Long, ByVal r2 As Long, _
                                  ByVal thisCell As String) As String
    pl_RawPctFormula = "=CEILING(PERCENTRANK.EXC($" & colLtr & "$" & r1 & ":$" & colLtr & "$" & r2 & _
                       "," & thisCell & ")*5,1)"
End Function

Private Function pl_RankPctFormula(ByVal colLtr As String, ByVal r1 As Long, ByVal r2 As Long, _
                                   ByVal thisCell As String) As String
    pl_RankPctFormula = "=IFERROR(CEILING((1-PERCENTRANK.EXC($" & colLtr & "$" & r1 & ":$" & colLtr & _
                        "$" & r2 & "," & thisCell & "))*5,1),0)"
End Function

' ==========================================================================
'  TAB-NAME PARSING  (never fails)
' ==========================================================================

Private Sub pl_ParseTab(ByVal rawName As String, ByRef outName As String, _
                        ByRef outVoltage As Double, ByRef outParsed As Boolean)
    Dim work As String, i As Long, nlen As Long, ch As String
    Dim tokStart As Long, tokEnd As Long, token As String
    Dim bestStart As Long, bestEnd As Long, bestVal As Double, found As Boolean
    outParsed = False: outVoltage = 0
    work = Trim$(rawName): nlen = Len(work): found = False: i = 1
    Do While i <= nlen
        ch = Mid$(work, i, 1)
        If pl_IsDigitOrDot(ch) Then
            tokStart = i
            Do While i <= nlen And pl_IsDigitOrDot(Mid$(work, i, 1))
                i = i + 1
            Loop
            tokEnd = i - 1
            token = Mid$(work, tokStart, tokEnd - tokStart + 1)
            If pl_IsPlausibleVoltage(token) Then
                bestStart = tokStart: bestEnd = tokEnd: bestVal = CDbl(token): found = True
            End If
        Else
            i = i + 1
        End If
    Loop
    If Not found Then
        outName = pl_CollapseName(work)
        If Len(outName) = 0 Then outName = work
        Exit Sub
    End If
    outVoltage = bestVal: outParsed = True
    Dim remainder As String
    remainder = Left$(work, bestStart - 1) & " " & pl_StripKv(Mid$(work, bestEnd + 1))
    outName = pl_CollapseName(remainder)
    If Len(outName) = 0 Then outName = "(unnamed)"
End Sub

Private Function pl_IsDigitOrDot(ByVal ch As String) As Boolean
    pl_IsDigitOrDot = (ch >= "0" And ch <= "9") Or (ch = ".")
End Function

Private Function pl_IsPlausibleVoltage(ByVal token As String) As Boolean
    Dim v As Double
    pl_IsPlausibleVoltage = False
    If Len(token) = 0 Then Exit Function
    If Left$(token, 1) = "." Or Right$(token, 1) = "." Then Exit Function
    If InStr(token, ".") <> InStrRev(token, ".") Then Exit Function
    If Not IsNumeric(token) Then Exit Function
    v = CDbl(token)
    If v > 0 And v <= 2000 Then pl_IsPlausibleVoltage = True
End Function

Private Function pl_StripKv(ByVal s As String) As String
    Dim t As String, j As Long
    t = s: j = 1
    Do While j <= Len(t) And pl_IsSeparator(Mid$(t, j, 1))
        j = j + 1
    Loop
    t = Mid$(t, j)
    If Len(t) >= 2 Then
        If StrComp(Left$(t, 2), "kV", vbTextCompare) = 0 Then t = Mid$(t, 3)
    End If
    pl_StripKv = t
End Function

Private Function pl_IsSeparator(ByVal ch As String) As Boolean
    Select Case ch
        Case " ", "_", "-", "(", ")", ",", ".", vbTab
            pl_IsSeparator = True
        Case Else
            pl_IsSeparator = False
    End Select
End Function

Private Function pl_CollapseName(ByVal s As String) As String
    Dim i As Long, ch As String, sb As String, lastSpace As Boolean
    lastSpace = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If pl_IsSeparator(ch) Then
            If Not lastSpace Then sb = sb & " ": lastSpace = True
        Else
            sb = sb & ch: lastSpace = False
        End If
    Next i
    pl_CollapseName = Trim$(sb)
End Function

' ==========================================================================
'  TRIPLE / IDENTITY HELPERS
' ==========================================================================

Private Function pl_TripleKey(ByVal nm As String, ByVal v As Double, ByVal p As Boolean, _
                              ByVal st As String) As String
    pl_TripleKey = LCase$(Trim$(nm)) & "|" & IIf(p, pl_NumStr(v), "") & "|" & LCase$(Trim$(st))
End Function

Private Function pl_TripleEqual(ByVal n1 As String, ByVal v1 As Double, ByVal p1 As Boolean, ByVal s1 As String, _
                                ByVal n2 As String, ByVal v2 As Double, ByVal p2 As Boolean, ByVal s2 As String) As Boolean
    pl_TripleEqual = (pl_TripleKey(n1, v1, p1, s1) = pl_TripleKey(n2, v2, p2, s2))
End Function

Private Function pl_TripleStr(ByVal nm As String, ByVal v As Double, ByVal p As Boolean, ByVal st As String) As String
    pl_TripleStr = nm & ", " & IIf(p, CStr(v) & " kV", "(no voltage)") & ", " & IIf(Len(st) > 0, st, "(no state)")
End Function

Private Function pl_IdentityLabel(ByVal nm As String, ByVal v As Double, ByVal p As Boolean, ByVal st As String) As String
    Dim lab As String: lab = nm
    If p Then lab = lab & " " & Format$(v, "0.###") & " kV"
    If Len(Trim$(st)) > 0 Then lab = lab & " (" & Trim$(st) & ")"
    pl_IdentityLabel = lab
End Function

' ==========================================================================
'  SHARED SHEET / IO / STRING HELPERS
' ==========================================================================

Private Function pl_PickFiles(ByVal title As String, ByRef outFiles() As String, _
                              ByRef outCount As Long) As Boolean
    Dim fd As FileDialog, i As Long, base As String
    base = ThisWorkbook.Path
    If Len(base) = 0 Then base = Application.DefaultFilePath
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = title
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"
        .InitialFileName = base & Application.PathSeparator
    End With
    If fd.Show <> -1 Then outCount = 0: pl_PickFiles = False: Exit Function
    outCount = fd.SelectedItems.Count
    If outCount = 0 Then pl_PickFiles = False: Exit Function
    ReDim outFiles(1 To outCount)
    For i = 1 To outCount: outFiles(i) = fd.SelectedItems(i): Next i
    pl_PickFiles = True
End Function

Private Function pl_GetOutputSheet(ByVal baseName As String) As Worksheet
    Dim existing As Worksheet, ws As Worksheet, ans As VbMsgBoxResult, newName As String
    On Error Resume Next
    Set existing = ThisWorkbook.Worksheets(baseName)
    On Error GoTo 0
    If existing Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = baseName
        Set pl_GetOutputSheet = ws
        Exit Function
    End If
    ans = MsgBox("A sheet named '" & baseName & "' already exists." & vbCrLf & _
                 "Yes = overwrite it" & vbCrLf & "No  = create a new timestamped sheet" & vbCrLf & _
                 "Cancel = abort", vbYesNoCancel + vbQuestion, "Sheet exists")
    Select Case ans
        Case vbCancel
            Set pl_GetOutputSheet = Nothing
        Case vbYes
            existing.Cells.Clear
            On Error Resume Next
            Dim co As ChartObject
            For Each co In existing.ChartObjects: co.Delete: Next co
            On Error GoTo 0
            Set pl_GetOutputSheet = existing
        Case vbNo
            newName = Left$(baseName & " " & Format$(Now, "yyyymmdd_hhnnss"), PL_EXCEL_MAX_TAB)
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = newName
            Set pl_GetOutputSheet = ws
    End Select
End Function

Private Function pl_ResolveCol(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long, hv As Variant
    lastCol = ws.Cells(PL_HDR_ROW, ws.Columns.Count).End(xlToLeft).Column
    hv = ws.Range(ws.Cells(PL_HDR_ROW, 1), ws.Cells(PL_HDR_ROW, lastCol)).Value
    For c = 1 To lastCol
        Dim cellTxt As String
        If lastCol = 1 Then cellTxt = Trim$(CStr(pl_NZ(hv))) Else cellTxt = Trim$(CStr(pl_NZ(hv(1, c))))
        If StrComp(cellTxt, headerText, vbTextCompare) = 0 Then pl_ResolveCol = c: Exit Function
    Next c
    pl_ResolveCol = 0
End Function

Private Function pl_MwAxis(ByRef mw() As Double) As Long
    Dim cnt As Long, j As Long
    cnt = CLng((PL_MW_MAX - PL_MW_MIN) / PL_MW_STEP) + 1
    ReDim mw(1 To cnt)
    For j = 1 To cnt: mw(j) = PL_MW_MIN + (j - 1) * PL_MW_STEP: Next j
    pl_MwAxis = cnt
End Function

Private Function pl_BlockRead(ByVal ws As Worksheet, ByVal firstRow As Long, _
                              ByVal lastRow As Long, ByVal cols As Long) As Variant
    Dim rng As Range, tmp(1 To 1, 1 To 1) As Variant
    If lastRow < firstRow Then pl_BlockRead = Empty: Exit Function
    Set rng = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, cols))
    If rng.Cells.Count = 1 Then
        tmp(1, 1) = rng.Value: pl_BlockRead = tmp
    Else
        pl_BlockRead = rng.Value
    End If
End Function

Private Function pl_RowBlank(ByRef arr As Variant, ByVal r As Long, ByVal cols As Long) As Boolean
    Dim c As Long, v As Variant
    For c = 1 To cols
        v = arr(r, c)
        If Not IsEmpty(v) Then
            If Len(Trim$(CStr(v))) > 0 Then pl_RowBlank = False: Exit Function
        End If
    Next c
    pl_RowBlank = True
End Function

Private Function pl_TrimRows(ByRef arr() As Variant, ByVal rows As Long, ByVal cols As Long) As Variant
    Dim outArr() As Variant, r As Long, c As Long
    ReDim outArr(1 To rows, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols: outArr(r, c) = arr(r, c): Next c
    Next r
    pl_TrimRows = outArr
End Function

Private Sub pl_GrowRaw(ByRef a() As String, ByRef b() As String, ByRef c() As Double, _
                       ByRef d() As Boolean, ByRef e() As String, ByRef f() As Long)
    Dim newSize As Long: newSize = UBound(a) * 2
    ReDim Preserve a(1 To newSize): ReDim Preserve b(1 To newSize)
    ReDim Preserve c(1 To newSize): ReDim Preserve d(1 To newSize)
    ReDim Preserve e(1 To newSize): ReDim Preserve f(1 To newSize)
End Sub

Private Function pl_ColLetter(ByVal col As Long) As String
    Dim s As String, rmn As Long
    Do While col > 0
        rmn = (col - 1) Mod 26
        s = Chr$(65 + rmn) & s
        col = (col - 1) \ 26
    Loop
    pl_ColLetter = s
End Function

Private Function pl_FullCol(ByVal col As Long) As String
    Dim L As String: L = pl_ColLetter(col)
    pl_FullCol = "$" & L & ":$" & L
End Function

' Bounded column reference $L$r1:$L$r2 (never whole-column, for live formulas).
Private Function pl_BoundCol(ByVal col As Long, ByVal r1 As Long, ByVal r2 As Long) As String
    Dim L As String: L = pl_ColLetter(col)
    pl_BoundCol = "$" & L & "$" & r1 & ":$" & L & "$" & r2
End Function

Private Function pl_SheetRef(ByVal nm As String) As String
    pl_SheetRef = "'" & Replace(nm, "'", "''") & "'"
End Function

Private Function pl_EscQuote(ByVal s As String) As String
    pl_EscQuote = Replace(s, """", """""")
End Function

Private Function pl_NumStr(ByVal v As Double) As String
    pl_NumStr = Format$(v, "0.###############")
    pl_NumStr = Replace(pl_NumStr, ",", ".")
End Function

Private Function pl_NZ(ByVal v As Variant) As Variant
    If IsError(v) Then
        pl_NZ = ""
    ElseIf IsNull(v) Then
        pl_NZ = ""
    Else
        pl_NZ = v
    End If
End Function

Private Function pl_FileName(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, Application.PathSeparator)
    If p = 0 Then p = InStrRev(fullPath, "\")
    If p = 0 Then pl_FileName = fullPath Else pl_FileName = Mid$(fullPath, p + 1)
End Function

Private Sub pl_Status(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

' ==========================================================================
'  RUN LOG
' ==========================================================================

Private Sub pl_LogInit(ByRef log As PL_TLog)
    ReDim log.lines(1 To 16): log.n = 0
End Sub

Private Sub pl_LogAdd(ByRef log As PL_TLog, ByVal s As String)
    log.n = log.n + 1
    If log.n > UBound(log.lines) Then ReDim Preserve log.lines(1 To UBound(log.lines) * 2)
    log.lines(log.n) = s
End Sub

Private Sub pl_WriteLog(ByRef log As PL_TLog)
    Dim ws As Worksheet, i As Long, blk() As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PL_SH_LOG)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PL_SH_LOG
    Else
        ws.Cells.Clear
    End If
    ws.Cells(1, 1).Value = "Interconnect Pipeline -- Run Log"
    ws.Cells(1, 1).Font.Bold = True
    If log.n > 0 Then
        ReDim blk(1 To log.n, 1 To 1)
        For i = 1 To log.n: blk(i, 1) = log.lines(i): Next i
        ws.Range(ws.Cells(2, 1), ws.Cells(1 + log.n, 1)).Value = blk
    End If
    ws.Columns(1).ColumnWidth = 120
End Sub

' ==========================================================================
'  FRONT-HALF + SCALE + PARITY SELF-TESTS  (called from SelfTest)
' ==========================================================================

Private Sub pl_FrontHalfSelfTest(ByRef passCount As Long, ByRef failCount As Long)
    Debug.Print "--- front-half / scaling tests ---"

    ' 1) Tab-name parsing
    Dim nm As String, v As Double, p As Boolean
    pl_ParseTab "Chaves County 345", nm, v, p
    Assert (nm = "Chaves County") And p And (v = 345), _
           "Tab parse: 'Chaves County 345' -> ('Chaves County', 345)", passCount, failCount
    pl_ParseTab "Cunningham", nm, v, p
    Assert (nm = "Cunningham") And (Not p), _
           "Tab parse: 'Cunningham' -> ('Cunningham', blank)", passCount, failCount

    ' 2) Triple dedupe (dictionary)
    Dim rn() As String, rst() As String, rv() As Double, rp() As Boolean, rf() As String, rr() As Long
    ReDim rn(1 To 5): ReDim rst(1 To 5): ReDim rv(1 To 5)
    ReDim rp(1 To 5): ReDim rf(1 To 5): ReDim rr(1 To 5)
    pl_FillSite rn, rv, rp, rst, rf, rr, 1, "Alpha", 345, True, "NM"
    pl_FillSite rn, rv, rp, rst, rf, rr, 2, "Alpha", 345, True, "NM"
    pl_FillSite rn, rv, rp, rst, rf, rr, 3, "Alpha", 230, True, "NM"
    pl_FillSite rn, rv, rp, rst, rf, rr, 4, "Alpha", 345, True, "TX"
    pl_FillSite rn, rv, rp, rst, rf, rr, 5, "Alpha", 345, True, "NM"
    Dim dt() As PL_TSiteTriple, dropped As Long, dlog As PL_TLog: pl_LogInit dlog
    Dim dc As Long: dc = pl_DedupTriples(rn, rv, rp, rst, 5, rf, rr, dt, dlog, dropped)
    Assert dc = 3, "Triple dedupe: 5 rows -> 3 distinct (dictionary)", passCount, failCount
    Assert dropped = 2, "Triple dedupe: 2 duplicates dropped and logged", passCount, failCount
    Assert pl_HasTriple(dt, dc, "Alpha", 230, True, "NM"), _
           "Triple dedupe: voltage-differ kept separate", passCount, failCount
    Assert pl_HasTriple(dt, dc, "Alpha", 345, True, "TX"), _
           "Triple dedupe: state-differ kept separate", passCount, failCount

    ' 3) Join intersection (dictionary keys)
    Dim dId As Object: Set dId = CreateObject("Scripting.Dictionary")
    dId.Add pl_CostKey("Alpha", 345, True), 1
    dId.Add pl_CostKey("Delta", 0, False), 2
    dId.Add pl_CostKey("Gamma", 0, False), 3      ' no site -> cost w/o site
    Dim tA As PL_TSiteTriple, tB As PL_TSiteTriple, tD As PL_TSiteTriple
    tA.Name = "Alpha": tA.Voltage = 345: tA.VoltParsed = True: tA.State = "NM"
    tB.Name = "Beta": tB.VoltParsed = False: tB.State = "OK"      ' no cost
    tD.Name = "Delta": tD.VoltParsed = False: tD.State = "TX"     ' bare match
    Assert pl_ResolveMatchKey(tA, dId) = pl_CostKey("Alpha", 345, True), _
           "Join: Alpha 345 -> name+voltage match", passCount, failCount
    Assert pl_ResolveMatchKey(tD, dId) = pl_CostKey("Delta", 0, False), _
           "Join: Delta -> bare name match", passCount, failCount
    Assert pl_ResolveMatchKey(tB, dId) = "", "Join: Beta -> no cost (alignment error)", passCount, failCount

    ' 4) Matrix shape + values (piecewise T/MW == SUMIFS definition)
    Dim mwx() As Double, mc As Long: mc = pl_MwAxis(mwx)
    Assert (mc = 21) And (mwx(1) = 100) And (mwx(mc) = 300), _
           "Matrix axis: 100..300 step 10 = 21 points", passCount, failCount
    ' one substation, tiers T=100 at 100+, T=200 at 200+  => cost(100)=1.0MM etc.
    ' cumulative alloc: trig 100 alloc 100; trig 200 alloc 100.
    Assert pl_PiecewiseCost(100, 100) = 1, "Matrix value: cost at 100 = T/MW", passCount, failCount

    ' 5) Scoring ceiling + $25MM band all-zero on the fixture
    Assert CLng(pl_ScoreCeiling(4)) = 165, "Scoring: weighted-score maximum is 165", passCount, failCount
    Assert CLng(pl_ScoreCeiling(3)) = 125, "Scoring: practical maximum (3 bands) is 125", passCount, failCount
    Dim fmw() As Double, fnames() As String, fcost() As Double, fn As Long, fm As Long
    LoadFixture fmw, fnames, fcost, fn, fm
    Dim rkC() As Long, rkMx() As Double, rkCm() As Double, rkS() As Double, rkH() As Boolean
    Dim rA() As Long, rB() As Long
    ComputeRankings fmw, fnames, fcost, fn, fm, rkC, rkMx, rkCm, rkS, rkH, rA, rB
    Dim b25a() As Long, b25b() As Long
    b25a = pl_RankBucketArray(rA, 1, fn): b25b = pl_RankBucketArray(rB, 1, fn)
    Dim allZero As Boolean: allZero = True
    Dim i As Long
    For i = 1 To fn
        If b25a(i) <> 0 Or b25b(i) <> 0 Then allZero = False
    Next i
    Assert (fn = 17) And allZero, "Scoring: all 17 $25MM band percentiles are 0", passCount, failCount

    ' 6) Headroom
    Dim trg() As Double: ReDim trg(1 To 3): trg(1) = 200: trg(2) = 130: trg(3) = 260
    Assert pl_MinHeadroom(trg, 3) = 130, "Headroom: 130/200/260 -> 130", passCount, failCount
    Dim hv2() As Double: ReDim hv2(1 To 5)
    hv2(1) = 50: hv2(2) = 130: hv2(3) = 200: hv2(4) = 260: hv2(5) = 300
    Dim hb() As Long: hb = pl_RawBucketArray(hv2, 5)
    Assert hb(5) = 5, "Headroom pctile: largest -> 5", passCount, failCount
    Assert hb(1) = 1, "Headroom pctile: smallest -> 1", passCount, failCount

    ' 7) Parity: VBA buckets == Excel PERCENTRANK.EXC/CEILING (round-trip)
    pl_ParitySelfTest passCount, failCount

    ' 8) Scale smoke test: 2,000 synthetic substations, no O(N^2) blow-up
    pl_ScaleSmokeTest passCount, failCount
End Sub

' Round-trips a raw population and a rank population through real Excel formulas
' on a scratch sheet and asserts the VBA buckets match cell-for-cell.
Private Sub pl_ParitySelfTest(ByRef passCount As Long, ByRef failCount As Long)
    On Error GoTo Fail
    Dim ws As Worksheet
    Dim prevAlerts As Boolean: prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets.Add
    Dim tmpName As String: tmpName = ws.Name

    Dim n As Long: n = 12
    Dim vals() As Double: ReDim vals(1 To n)
    Dim ranks() As Long: ReDim ranks(1 To n, 1 To 1)
    Dim i As Long
    ' a raw population with a tie, and a competition-rank population with ties
    Dim seed As Variant
    seed = Array(50#, 130#, 130#, 200#, 260#, 300#, 90#, 175#, 220#, 45#, 310#, 130#)
    For i = 1 To n
        vals(i) = CDbl(seed(i - 1))
        ws.Cells(i, 1).Value = vals(i)
    Next i
    Dim rseed As Variant
    rseed = Array(1, 2, 2, 4, 5, 6, 1, 2, 0, 4, 6, 2)   ' 0 => non-qualifier
    For i = 1 To n
        ranks(i, 1) = CLng(rseed(i - 1))
        If ranks(i, 1) > 0 Then ws.Cells(i, 3).Value = ranks(i, 1) Else ws.Cells(i, 3).Value = "n/a"
    Next i

    ' Excel formulas: raw-is-better in col B, rank-based in col D
    For i = 1 To n
        ws.Cells(i, 2).Formula = "=CEILING(PERCENTRANK.EXC($A$1:$A$" & n & ",A" & i & ")*5,1)"
        ws.Cells(i, 4).Formula = "=IFERROR(CEILING((1-PERCENTRANK.EXC($C$1:$C$" & n & ",C" & i & "))*5,1),0)"
    Next i
    Application.CalculateFull

    Dim vbaRaw() As Long: vbaRaw = pl_RawBucketArray(vals, n)
    Dim vbaRank() As Long: vbaRank = pl_RankBucketArray(ranks, 1, n)

    Dim rawOK As Boolean, rankOK As Boolean: rawOK = True: rankOK = True
    For i = 1 To n
        If CLng(ws.Cells(i, 2).Value) <> vbaRaw(i) Then rawOK = False
        If CLng(ws.Cells(i, 4).Value) <> vbaRank(i) Then rankOK = False
    Next i

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = prevAlerts

    Assert rawOK, "Parity: VBA raw-is-better buckets == Excel PERCENTRANK.EXC", passCount, failCount
    Assert rankOK, "Parity: VBA rank-based buckets == Excel PERCENTRANK.EXC", passCount, failCount
    Exit Sub
Fail:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    Assert False, "Parity self-test could not run (" & Err.Description & ")", passCount, failCount
End Sub

' 2,000 synthetic substations: rank + bucket every column and assert the whole
' pass is fast (mergesort O(N log N), not the ~32M-op O(N^2) rank-by-scanning).
Private Sub pl_ScaleSmokeTest(ByRef passCount As Long, ByRef failCount As Long)
    Dim n As Long: n = 2000
    Dim m As Long: m = 21
    Dim mw() As Double, names() As String, costM() As Double
    ReDim mw(1 To m): ReDim names(1 To n): ReDim costM(1 To n, 1 To m)
    Dim i As Long, j As Long
    For j = 1 To m: mw(j) = 100 + (j - 1) * 10: Next j
    ' Distinct piecewise-constant totals so ranks are well spread.
    For i = 1 To n
        names(i) = "S" & Format$(i, "0000")
        Dim base As Double: base = 40000000# + (i Mod 500) * 100000#
        For j = 1 To m
            Dim tot As Double: tot = base
            If mw(j) >= 200 Then tot = base + 20000000# + (i Mod 50) * 100000#
            costM(i, j) = tot / mw(j)
        Next j
    Next i

    Dim t0 As Double: t0 = Timer
    Dim rkC() As Long, rkMx() As Double, rkCm() As Double, rkS() As Double, rkH() As Boolean
    Dim rA() As Long, rB() As Long
    ComputeRankings mw, names, costM, n, m, rkC, rkMx, rkCm, rkS, rkH, rA, rB

    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1
    Dim tt As Long, dummy() As Long
    For tt = 1 To nt
        dummy = pl_RankBucketArray(rA, tt, n)
        dummy = pl_RankBucketArray(rB, tt, n)
    Next tt
    Dim knee() As Double: ReDim knee(1 To n)
    For i = 1 To n: knee(i) = costM(i, 1): Next i
    dummy = pl_RawBucketArray(knee, n)
    Dim elapsed As Double: elapsed = Timer - t0

    Debug.Print "  scale: " & n & " substations ranked+bucketed in " & Format$(elapsed, "0.00") & "s"
    Assert (UBound(rA, 1) = n), "Scale: ranking produced n rows for 2,000 substations", passCount, failCount
    Assert (elapsed < 30), "Scale: 2,000-substation rank+percentile pass under 30s (no O(N^2))", passCount, failCount
End Sub

' ---- small test helpers ----
Private Sub pl_FillSite(ByRef rn() As String, ByRef rv() As Double, ByRef rp() As Boolean, _
                        ByRef rst() As String, ByRef rf() As String, ByRef rr() As Long, _
                        ByVal idx As Long, ByVal nm As String, ByVal v As Double, _
                        ByVal p As Boolean, ByVal st As String)
    rn(idx) = nm: rv(idx) = v: rp(idx) = p: rst(idx) = st
    rf(idx) = "fixture.xlsx": rr(idx) = idx + 1
End Sub

Private Function pl_HasTriple(ByRef dt() As PL_TSiteTriple, ByVal k As Long, _
                              ByVal nm As String, ByVal v As Double, ByVal p As Boolean, _
                              ByVal st As String) As Boolean
    Dim i As Long
    For i = 1 To k
        If pl_TripleEqual(dt(i).Name, dt(i).Voltage, dt(i).VoltParsed, dt(i).State, nm, v, p, st) Then
            pl_HasTriple = True: Exit Function
        End If
    Next i
    pl_HasTriple = False
End Function

Private Function pl_MinHeadroom(ByRef v() As Double, ByVal n As Long) As Double
    Dim i As Long, mn As Double, seen As Boolean
    seen = False
    For i = 1 To n
        If v(i) > 0 Then
            If Not seen Then
                mn = v(i): seen = True
            ElseIf v(i) < mn Then
                mn = v(i)
            End If
        End If
    Next i
    If seen Then pl_MinHeadroom = mn Else pl_MinHeadroom = 0
End Function

' Cost at one MW from two cumulative allocations (trig 100 alloc `a1`, trig 200
' alloc `a2`); returns cost-per-MW at MW=100 for the single-tier example.
Private Function pl_PiecewiseCost(ByVal a1 As Double, ByVal a2 As Double) As Double
    ' at MW=100 only the trig<=100 tier applies -> T=a1 ; cost=a1/100.
    pl_PiecewiseCost = a1 / 100
End Function

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
' Stable bottom-up MERGESORT of the index array ord(1..popN), O(popN log popN).
' The comparator (LessThan) and its total order are unchanged from the original
' insertion sort, and the merge takes the left run on ties, so the produced
' order -- and therefore every competition rank -- is byte-for-byte identical
' to the pre-optimization insertion sort. This is the fix for the Stage-5
' O(N^2) rank-by-scanning blow-up at ~2,000 substations.
Private Sub SortIdx(ByRef ord() As Long, ByVal popN As Long, ByVal tt As Long, _
                    ByVal isRankA As Boolean, _
                    ByRef rkCount() As Long, ByRef rkMaxMW() As Double, _
                    ByRef rkCostMax() As Double, ByRef rkSlope() As Double, _
                    ByRef names() As String)
    If popN < 2 Then Exit Sub
    Dim buf() As Long: ReDim buf(1 To popN)
    Dim width As Long, i As Long
    width = 1
    Do While width < popN
        i = 1
        Do While i <= popN
            Dim l1 As Long, r1 As Long, l2 As Long, r2 As Long
            Dim p As Long, q As Long, k As Long
            l1 = i: r1 = i + width - 1
            If r1 > popN Then r1 = popN
            l2 = r1 + 1: r2 = i + 2 * width - 1
            If r2 > popN Then r2 = popN
            p = l1: q = l2: k = l1
            Do While p <= r1 And q <= r2
                ' take the LEFT run unless the right element sorts strictly
                ' before it -> stable.
                If LessThan(ord(q), ord(p), tt, isRankA, rkCount, rkMaxMW, rkCostMax, rkSlope, names) Then
                    buf(k) = ord(q): q = q + 1
                Else
                    buf(k) = ord(p): p = p + 1
                End If
                k = k + 1
            Loop
            Do While p <= r1
                buf(k) = ord(p): p = p + 1: k = k + 1
            Loop
            Do While q <= r2
                buf(k) = ord(q): q = q + 1: k = k + 1
            Loop
            For k = l1 To r2
                ord(k) = buf(k)
            Next k
            i = i + 2 * width
        Loop
        width = width * 2
    Loop
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
        ' Whole-column formatting only -- the former per-cell grey "n/a"
        ' recolour was a cell-by-cell loop (O(n*nt) COM calls) and is dropped
        ' for the ~2,000-row scale target; the "n/a" text itself is unchanged.
        Dim rc As Long
        For rc = scol + 1 To scol + 2
            With ws.Cells(dataTop, rc).Resize(n, 1)
                .NumberFormat = "0"
                .HorizontalAlignment = xlCenter
            End With
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

' Rebuilt for scale: the whole leaderboard is composed in a Variant array and
' written in a SINGLE Range.Value assignment (no cell-by-cell writes), and the
' per-threshold ordering uses the O(k log k) mergesort in LeaderOrder. A few
' section-header rows are bolded afterwards (not per substation); the block is
' monospaced once.
Private Sub BuildLeaderboard(ByVal ws As Worksheet, ByVal topRow As Long, _
                             ByRef names() As String, ByVal n As Long, _
                             ByRef rkCount() As Long, ByRef rkMaxMW() As Double, _
                             ByRef rkSlope() As Double, _
                             ByRef rankA() As Long, ByRef rankB() As Long, _
                             ByRef lastRow As Long)
    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1

    Dim maxRows As Long: maxRows = 3 + nt * (3 + n)
    Dim blk() As Variant: ReDim blk(1 To maxRows, 1 To 7)
    Dim boldOff() As Long: ReDim boldOff(1 To 3 + nt * 2)
    Dim nBold As Long: nBold = 0

    Dim rp As Long: rp = 1                       ' 1-based row offset within blk
    blk(rp, 1) = "Threshold Rankings"
    nBold = nBold + 1: boldOff(nBold) = rp
    rp = rp + 2                                  ' title + blank

    Dim tt As Long, i As Long
    For tt = 1 To nt
        Dim L As Double: L = CDbl(thr(LBound(thr) + tt - 1))
        Dim qcnt As Long: qcnt = 0
        For i = 1 To n
            If rkCount(i, tt) >= 1 Then qcnt = qcnt + 1
        Next i

        blk(rp, 1) = "<= " & ThreshLabel(L) & "   (" & qcnt & " of " & n & " substations qualify)"
        nBold = nBold + 1: boldOff(nBold) = rp
        rp = rp + 1

        If qcnt = 0 Then
            blk(rp, 2) = "No substations qualify at this threshold"
            rp = rp + 2
        Else
            blk(rp, 2) = "By MW Breadth"
            blk(rp, 6) = "By Slope, " & IIf(SLOPE_RANK_MODE = 2, "flattest", "steepest") & " first"
            nBold = nBold + 1: boldOff(nBold) = rp
            rp = rp + 1

            Dim ordA() As Long, kA As Long: LeaderOrder rankA, names, n, tt, ordA, kA
            Dim ordB() As Long, kB As Long: LeaderOrder rankB, names, n, tt, ordB, kB
            Dim rows As Long: rows = kA: If kB > rows Then rows = kB

            Dim rr As Long
            For rr = 1 To rows
                If rr <= kA Then
                    i = ordA(rr)
                    blk(rp, 1) = rankA(i, tt)
                    blk(rp, 2) = names(i)
                    blk(rp, 3) = rkCount(i, tt) & " pts, max " & Format(rkMaxMW(i, tt), FMT_MW)
                End If
                If rr <= kB Then
                    i = ordB(rr)
                    blk(rp, 5) = rankB(i, tt)
                    blk(rp, 6) = names(i)
                    blk(rp, 7) = Format(rkSlope(i, tt), "#,##0.0")
                End If
                rp = rp + 1
            Next rr
            rp = rp + 1                          ' spacer between sub-blocks
        End If
    Next tt

    Dim used As Long: used = rp - 1
    If used < 1 Then used = 1
    ws.Range(ws.Cells(topRow, 1), ws.Cells(topRow + used - 1, 7)).Value = blk
    lastRow = topRow + used

    Dim b As Long
    ws.Cells(topRow, 1).Font.Size = 12
    For b = 1 To nBold
        ws.Cells(topRow + boldOff(b) - 1, 1).Resize(1, 7).Font.Bold = True
    Next b
    ws.Range(ws.Cells(topRow, 1), ws.Cells(lastRow, 7)).Font.Name = "Consolas"
End Sub

' Population indices (rank > 0) for one threshold, sorted for display by
' (rank asc, name asc), via a stable O(k log k) bottom-up mergesort. The rank
' already encodes the full metric order with competition ties, so sorting on it
' reproduces the ranked order.
Private Sub LeaderOrder(ByRef rank() As Long, ByRef names() As String, ByVal n As Long, _
                        ByVal tt As Long, ByRef ord() As Long, ByRef k As Long)
    ReDim ord(1 To n)
    k = 0
    Dim i As Long
    For i = 1 To n
        If rank(i, tt) > 0 Then k = k + 1: ord(k) = i
    Next i
    If k < 2 Then Exit Sub

    Dim buf() As Long: ReDim buf(1 To k)
    Dim width As Long, s As Long
    width = 1
    Do While width < k
        s = 1
        Do While s <= k
            Dim l1 As Long, r1 As Long, l2 As Long, r2 As Long, p As Long, q As Long, w As Long
            l1 = s: r1 = s + width - 1
            If r1 > k Then r1 = k
            l2 = r1 + 1: r2 = s + 2 * width - 1
            If r2 > k Then r2 = k
            p = l1: q = l2: w = l1
            Do While p <= r1 And q <= r2
                If LeaderBefore(ord(q), ord(p), tt, rank, names) Then
                    buf(w) = ord(q): q = q + 1
                Else
                    buf(w) = ord(p): p = p + 1
                End If
                w = w + 1
            Loop
            Do While p <= r1
                buf(w) = ord(p): p = p + 1: w = w + 1
            Loop
            Do While q <= r2
                buf(w) = ord(q): q = q + 1: w = w + 1
            Loop
            For w = l1 To r2
                ord(w) = buf(w)
            Next w
            s = s + 2 * width
        Loop
        width = width * 2
    Loop
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

    ' ---- Front-half (pipeline steps 1-3) + scoring assertions -----
    pl_FrontHalfSelfTest passCount, failCount

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
