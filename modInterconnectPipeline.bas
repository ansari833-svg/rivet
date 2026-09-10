Attribute VB_Name = "modInterconnectPipeline"
Option Explicit

' ==========================================================================
'  modInterconnectPipeline.bas
'  Interconnection Site-Scoring Pipeline  (single module, one entry point)
'
'  PUBLIC ENTRY POINT:  RunInterconnectPipeline   (run via Alt+F8)
'  TEST HARNESS:        SelfTest                   (in-memory, no file pickers)
'
'  Runs the whole workflow end to end from a blank workbook, in order:
'
'    1. Consolidate the cost / results files (per-tier upgrade records)
'       into a "Cost Data" sheet. Substation name comes from sheet 3's TAB
'       NAME (parsed into name + optional voltage), never from a cell.
'    2. Consolidate the dedupe / site files (Summary tab: name = col A,
'       state = col C, voltage = col G) into a "Site Data" sheet,
'       de-duplicated on the full (name, voltage, state) TRIPLE.
'    3. Join the two on substation identity (intersection only, matched by
'       name, plus voltage when the cost tab carried one) and build a
'       "Matrix" sheet of cost-per-MW as live SUMIFS formulas.
'    4. Run the cost-curve analysis, per-threshold ranking and weighted
'       scoring (folded in verbatim from the proven modCostCurves logic)
'       onto a "Cost Curve Analysis" sheet.
'
'  A run log ("_Pipeline Log") captures per-file status, dropped duplicate
'  triples and any join alignment errors.
'
'  DESIGN NOTES
'    * One importable module. Everything below the declarations is a
'      procedure; module-level state is limited to configuration constants
'      plus the analysis module's original constants/Type/stage tag. Data
'      moves between procedures as typed arrays or Private Type records.
'    * The step-4 analysis/ranking/chart/leaderboard code is reused
'      UNCHANGED from modCostCurves; only its input capture (now the Matrix
'      ranges instead of InputBoxes) and its output-sheet creation (now the
'      shared overwrite/new/cancel prompt) are adapted, and the sole
'      Activate it used (for FreezePanes) is dropped so this module contains
'      no .Select / .Activate / Selection anywhere.
'    * Column references into "Cost Data" are resolved at run time BY HEADER
'      NAME ("Substation Name", "Voltage (kV)", "Size Overload Occurs (MW)",
'      "Proposed Project Allocation ($)"), never by fixed letter, because the
'      four metadata columns the tool prepends shift the source columns right.
'
'  SCORING (weighted composite) -- see README for the full rationale
'    A single weight cell $B$2 (default 4) multiplies EVERY threshold band;
'    the threshold axis as a whole is worth 4x the flattening axis. The
'    weight does NOT grade $25 > $50 > $75 > $100 -- if graded weights are
'    wanted later, each band needs its own weight cell.
'      Flattening percentile:                 max 5
'      Each band: $B$2 x (breadth + slope) =   4 x (5 + 5) = 40 ; x4 bands = 160
'      Composite maximum:                      5 + 160 = 165
'      Practical maximum today:                125 -- nothing prices under
'        $25MM (cheapest total ~ $47.6MM), so that band scores 0 for every
'        substation regardless of the multiplier.
'    A standalone Headroom (MW) column -- minimum overload-trigger size via
'    MINIFS on the Cost Data trigger column -- and its 1-5 percentile are added
'    to the scoring sheet as a display-and-rank lens ONLY. Headroom is NOT
'    summed into the composite; the ceiling stays 165.
'
'  ACCEPTANCE / SELF-TEST CASES (asserted in SelfTest, in memory):
'    * Tab parsing: "Chaves County 345" -> name "Chaves County", voltage 345;
'      "Cunningham" -> name "Cunningham", voltage blank.
'    * Triple dedupe: identical (name, voltage, state) rows collapse; rows
'      differing only in voltage, or only in state, stay separate; drops
'      are logged.
'    * Join intersection: a site triple with no cost match and a cost tab
'      with no site match are both logged as alignment errors and excluded;
'      the matched set builds correctly; a bare cost tab matches on name.
'    * Matrix shape: header row is the MW axis (100..300, strictly
'      ascending), column A the identities, body cells SUMIFS formulas that
'      resolve trigger/allocation by header name.
'    * Scoring ceiling: composite maximum is 165; all 17 $25MM band
'      percentiles are 0 on the reference fixture.
'    * Headroom: triggers 130/200/260 report 130; largest headroom -> pctile
'      5, smallest -> 1; headroom is not scored, so the ceiling stays 165.
'
'  Target platform: Windows Excel, VBA 7.x. No external references.
' ==========================================================================

' ---------- Pipeline configuration (edit here only) ------------------------

Private Const PL_SH_COST     As String = "Cost Data"
Private Const PL_SH_SITE     As String = "Site Data"
Private Const PL_SH_MATRIX   As String = "Matrix"
Private Const PL_SH_ANALYSIS As String = "Cost Curve Analysis"
Private Const PL_SH_LOG      As String = "_Pipeline Log"

Private Const PL_HDR_ROW        As Long = 1     ' header row on every source sheet
Private Const PL_COST_DATA_IDX  As Long = 3     ' cost workbook: sheet 3 carries the data
Private Const PL_SITE_TAB       As String = "Summary"
Private Const PL_SITE_NAME_COL  As Long = 1     ' Summary!A = substation name
Private Const PL_SITE_STATE_COL As Long = 3     ' Summary!C = state
Private Const PL_SITE_VOLT_COL  As Long = 7     ' Summary!G = voltage (kV)

' The two source headers that must survive verbatim so the Matrix SUMIFS can
' resolve them by name on Cost Data (their column LETTERS shift because of the
' four metadata columns the consolidator prepends -- never hardcode L/O).
Private Const PL_HDR_NAME    As String = "Substation Name"
Private Const PL_HDR_VOLT    As String = "Voltage (kV)"
Private Const PL_HDR_TRIGGER As String = "Size Overload Occurs (MW)"
Private Const PL_HDR_ALLOC   As String = "Proposed Project Allocation ($)"

' Matrix MW axis: 100, 110, ... 300 (numeric, strictly ascending).
Private Const PL_MW_MIN  As Double = 100
Private Const PL_MW_MAX  As Double = 300
Private Const PL_MW_STEP As Double = 10

' Weighted scoring. PL_WEIGHT_B2 seeds the live $B$2 cell; PL_MAX_PCTILE is the
' per-axis percentile ceiling. Composite max = PL_MAX_PCTILE + nBands x
' PL_WEIGHT_B2 x (2 x PL_MAX_PCTILE) = 5 + 4 x 4 x 10 = 165.
Private Const PL_WEIGHT_B2  As Double = 4
Private Const PL_MAX_PCTILE As Long = 5

Private Const PL_EXCEL_MAX_ROWS   As Long = 1048576
Private Const PL_EXCEL_MAX_TAB    As Long = 31

' ---------- Per-cost-file working record -----------------------------------

Private Type PL_TCostFile
    FilePath      As String
    SheetName     As String    ' raw sheet-3 tab name, verbatim
    Substation    As String    ' parsed from the tab name
    Voltage       As Double     ' parsed voltage (0 when none)
    VoltParsed    As Boolean    ' True when the tab carried a voltage
    HeaderSig     As String
    UsedCols      As Long
    FirstDataRow  As Long
    LastDataRow   As Long
    IsValid       As Boolean
    Status        As String     ' Imported | Skipped | Failed
    Message       As String
    RowsImported  As Long
End Type

' ---------- Distinct-identity records --------------------------------------

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
    Matched     As Boolean      ' set during the join (for cost-no-site logging)
End Type

' ---------- Run log (threaded by reference, never a global) ----------------

Private Type PL_TLog
    lines() As String
    n       As Long
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
'   Orchestrates the four stages end to end. Saves and restores the four
'   application state flags on every exit path, including errors. One file's
'   failure never aborts the run; a whole stage that cannot proceed (no files
'   picked, no intersection, user cancel) stops the pipeline cleanly and still
'   writes the run log.
'--------------------------------------------------------------------------
Public Sub RunInterconnectPipeline()

    Dim savedScreen As Boolean, savedEvents As Boolean
    Dim savedAlerts As Boolean, savedCalc As XlCalculation
    Dim stateSaved As Boolean

    Dim log As PL_TLog
    Dim wsCost As Worksheet, wsSite As Worksheet, wsMatrix As Worksheet
    Dim analysisOK As Boolean

    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedAlerts = Application.DisplayAlerts
    savedCalc = Application.Calculation
    stateSaved = True

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    pl_LogInit log
    pl_LogAdd log, "Run started " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    ' ---- Step 1: cost / results files -> Cost Data --------------------
    If Not pl_ConsolidateCost(log, wsCost) Then
        pl_LogAdd log, "Step 1 did not complete; pipeline stopped."
        pl_WriteLog log
        MsgBox "Step 1 (Cost Data consolidation) did not complete. See the " & _
               PL_SH_LOG & " sheet.", vbExclamation, "Interconnect Pipeline"
        GoTo Cleanup
    End If

    ' ---- Step 2: dedupe / site files -> Site Data ---------------------
    Dim distinctTriples As Long
    If Not pl_ConsolidateSite(log, wsSite, distinctTriples) Then
        pl_LogAdd log, "Step 2 did not complete; pipeline stopped."
        pl_WriteLog log
        MsgBox "Step 2 (Site Data consolidation) did not complete. See the " & _
               PL_SH_LOG & " sheet.", vbExclamation, "Interconnect Pipeline"
        GoTo Cleanup
    End If

    ' ---- Step 3: join + Matrix (live SUMIFS) --------------------------
    Dim matchedCount As Long, mwCount As Long
    Dim keyName() As String, keyVolt() As Double, keyUseVolt() As Boolean
    Dim costName As String, colTrig As Long, colName As Long, colVolt As Long
    If Not pl_BuildMatrix(log, wsCost, wsSite, wsMatrix, matchedCount, mwCount, _
                          keyName, keyVolt, keyUseVolt, costName, colTrig, colName, colVolt) Then
        pl_LogAdd log, "Step 3 did not complete; pipeline stopped."
        pl_WriteLog log
        MsgBox "Step 3 (Matrix build) did not complete. See the " & _
               PL_SH_LOG & " sheet.", vbExclamation, "Interconnect Pipeline"
        GoTo Cleanup
    End If

    ' Force a recompute so the analysis reads computed SUMIFS values, not 0s.
    Application.CalculateFull

    ' ---- Step 4: analysis + ranking + scoring -> Cost Curve Analysis --
    analysisOK = pl_RunAnalysisAndScoring(wsMatrix, log, keyName, keyVolt, keyUseVolt, _
                                          costName, colTrig, colName, colVolt)

    pl_WriteLog log

    MsgBox "Interconnection pipeline complete." & vbCrLf & vbCrLf & _
           "Cost Data sheet:   " & wsCost.Name & vbCrLf & _
           "Site Data sheet:   " & wsSite.Name & "  (" & distinctTriples & " distinct triples)" & vbCrLf & _
           "Matrix sheet:      " & wsMatrix.Name & "  (" & matchedCount & " substations x " & mwCount & " MW)" & vbCrLf & _
           "Analysis + scoring: " & IIf(analysisOK, "written", "skipped -- see log") & vbCrLf & _
           "Run log:           " & PL_SH_LOG, _
           vbInformation, "Interconnect Pipeline"

Cleanup:
    If stateSaved Then
        Application.ScreenUpdating = savedScreen
        Application.EnableEvents = savedEvents
        Application.DisplayAlerts = savedAlerts
        Application.Calculation = savedCalc
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
'  STEP 1 -- CONSOLIDATE COST / RESULTS FILES  -> Cost Data
' ==========================================================================

'--------------------------------------------------------------------------
' pl_ConsolidateCost
'   Picks the cost source workbooks, validates each (sheet 3 = data, header
'   on row 1), consolidates the sheet-3 used ranges into Cost Data with the
'   metadata prefix  Substation Name | Voltage (kV) | Source File |
'   Source Sheet  then the source columns verbatim. Header signatures are
'   compared case-insensitively; mismatches are surfaced for include/exclude.
'   Returns True and sets wsCost when at least one file was written.
'--------------------------------------------------------------------------
Private Function pl_ConsolidateCost(ByRef log As PL_TLog, _
                                    ByRef wsCost As Worksheet) As Boolean
    On Error GoTo ErrHandler
    pl_ConsolidateCost = False

    Dim files() As String, fileCount As Long
    If Not pl_PickFiles("Select the COST / results workbooks to consolidate", files, fileCount) Then
        pl_LogAdd log, "Step 1: no cost files selected."
        Exit Function
    End If

    Dim recs() As PL_TCostFile
    ReDim recs(1 To fileCount)

    ' -- validation pass (nothing written yet) --
    Dim i As Long, refSig As String, haveRef As Boolean
    For i = 1 To fileCount
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
    Next i

    ' -- count passes / mismatches --
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

    ' -- header-mismatch decision --
    If mismatch > 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox(mismatch & " cost file(s) have a mismatched header signature." & vbCrLf & vbCrLf & _
                     "Yes  = INCLUDE them anyway, aligned by column position" & vbCrLf & _
                     "No   = SKIP the mismatched files and consolidate the rest" & vbCrLf & _
                     "Cancel = ABORT step 1", _
                     vbYesNoCancel + vbQuestion, "Cost header mismatch")
        Select Case ans
            Case vbCancel
                pl_LogAdd log, "Step 1: aborted at header-mismatch prompt."
                Exit Function
            Case vbYes
                For i = 1 To fileCount
                    If (Not recs(i).IsValid) And _
                       InStr(1, recs(i).Message, "Header signature mismatch", vbTextCompare) > 0 Then
                        recs(i).IsValid = True
                        recs(i).Status = "Imported"
                        recs(i).Message = "Included despite header mismatch (aligned by column)"
                    End If
                Next i
        End Select
    End If

    Dim firstValid As Long
    firstValid = 0
    For i = 1 To fileCount
        If recs(i).IsValid Then firstValid = i: Exit For
    Next i
    If firstValid = 0 Then
        pl_LogAdd log, "Step 1: no files remain after mismatch decision."
        Exit Function
    End If

    ' -- create Cost Data --
    Set wsCost = pl_GetOutputSheet(PL_SH_COST)
    If wsCost Is Nothing Then
        pl_LogAdd log, "Step 1: user cancelled Cost Data sheet creation."
        Exit Function
    End If

    ' -- header row: metadata prefix + source headers from first valid file --
    Dim headerCols As Long
    headerCols = recs(firstValid).UsedCols
    wsCost.Cells(1, 1).Value = PL_HDR_NAME
    wsCost.Cells(1, 2).Value = PL_HDR_VOLT
    wsCost.Cells(1, 3).Value = "Source File"
    wsCost.Cells(1, 4).Value = "Source Sheet"
    pl_CopyCostHeaders wsCost, recs(firstValid)

    ' -- consolidation pass --
    Dim nextRow As Long, totalRows As Long, processed As Long, hitLimit As Boolean
    nextRow = 2
    For i = 1 To fileCount
        If recs(i).IsValid Then
            If pl_AppendCostFile(wsCost, recs(i), headerCols, nextRow) Then
                processed = processed + 1
                totalRows = totalRows + recs(i).RowsImported
            Else
                hitLimit = True
                Exit For
            End If
        ElseIf recs(i).Status <> "Failed" Then
            recs(i).Status = "Skipped"
        End If
    Next i

    wsCost.Rows(1).Font.Bold = True
    wsCost.Columns.AutoFit

    ' -- log every file --
    pl_LogAdd log, "STEP 1 -- Cost Data (" & wsCost.Name & "): " & processed & _
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

'--------------------------------------------------------------------------
' pl_ValidateCostFile
'   Opens one cost workbook read-only and checks it has a usable sheet 3.
'   Fills IsValid / Status / Message, and on success the parsed name/voltage,
'   header signature and data-row extents. Never leaves the workbook open.
'--------------------------------------------------------------------------
Private Sub pl_ValidateCostFile(ByRef rec As PL_TCostFile)
    Dim wb As Workbook, ws As Worksheet
    rec.IsValid = False
    rec.Status = "Failed"

    On Error GoTo OpenFail
    Set wb = Application.Workbooks.Open(Filename:=rec.FilePath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo CloseFail

    If wb.Sheets.Count < PL_COST_DATA_IDX Then
        rec.Message = "Workbook has fewer than " & PL_COST_DATA_IDX & " sheets"
        GoTo CloseAndExit
    End If
    If Not TypeOf wb.Sheets(PL_COST_DATA_IDX) Is Worksheet Then
        rec.Message = "Sheet " & PL_COST_DATA_IDX & " is not a worksheet"
        GoTo CloseAndExit
    End If

    Set ws = wb.Sheets(PL_COST_DATA_IDX)
    rec.SheetName = ws.Name
    If Not pl_MeasureSheet(ws, rec) Then GoTo CloseAndExit

    rec.HeaderSig = pl_HeaderSig(ws, rec.UsedCols)
    pl_ParseTab rec.SheetName, rec.Substation, rec.Voltage, rec.VoltParsed
    rec.IsValid = True
    rec.Status = "Imported"
    rec.Message = ""

CloseAndExit:
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Exit Sub

OpenFail:
    rec.Message = "Could not open: " & Err.Description
    Exit Sub

CloseFail:
    rec.Message = "Error inspecting workbook: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
' pl_MeasureSheet
'   Determines the used-column count and data-row range on a source sheet and
'   confirms a non-empty header row with at least one data row below it.
'--------------------------------------------------------------------------
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

'--------------------------------------------------------------------------
' pl_HeaderSig
'   Case-preserving delimited signature of the trimmed header cells (compared
'   case-insensitively by the caller).
'--------------------------------------------------------------------------
Private Function pl_HeaderSig(ByVal ws As Worksheet, ByVal usedCols As Long) As String
    Dim c As Long, parts() As String
    ReDim parts(1 To usedCols)
    For c = 1 To usedCols
        parts(c) = Trim$(CStr(ws.Cells(PL_HDR_ROW, c).Value))
    Next c
    pl_HeaderSig = Join(parts, "|")
End Function

'--------------------------------------------------------------------------
' pl_CopyCostHeaders
'   Copies the source header row of the first valid file into Cost Data,
'   verbatim, after the four metadata columns.
'--------------------------------------------------------------------------
Private Sub pl_CopyCostHeaders(ByVal wsCost As Worksheet, ByRef rec As PL_TCostFile)
    Dim wb As Workbook, ws As Worksheet, c As Long
    On Error GoTo Done
    Set wb = Application.Workbooks.Open(Filename:=rec.FilePath, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Sheets(PL_COST_DATA_IDX)
    For c = 1 To rec.UsedCols
        wsCost.Cells(1, 4 + c).Value = ws.Cells(PL_HDR_ROW, c).Value
    Next c
    wb.Close SaveChanges:=False
    Exit Sub
Done:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'--------------------------------------------------------------------------
' pl_AppendCostFile
'   Block-reads one cost file's data range (values only), skips blank rows,
'   and writes the surviving rows with the metadata prefix. Returns False only
'   when the worksheet row limit would be exceeded; a per-file read error is
'   logged on the record and returns True so the run continues.
'--------------------------------------------------------------------------
Private Function pl_AppendCostFile(ByVal wsCost As Worksheet, ByRef rec As PL_TCostFile, _
                                   ByVal headerCols As Long, ByRef nextRow As Long) As Boolean
    Dim wb As Workbook, ws As Worksheet
    Dim srcVals As Variant, outBlock() As Variant
    Dim srcRows As Long, srcCols As Long, writeCols As Long
    Dim r As Long, c As Long, outR As Long

    pl_AppendCostFile = True
    rec.RowsImported = 0

    On Error GoTo Fail
    Set wb = Application.Workbooks.Open(Filename:=rec.FilePath, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Sheets(PL_COST_DATA_IDX)

    srcVals = pl_BlockRead(ws, rec.FirstDataRow, rec.LastDataRow, rec.UsedCols)
    If IsEmpty(srcVals) Then wb.Close SaveChanges:=False: Exit Function

    srcRows = UBound(srcVals, 1) - LBound(srcVals, 1) + 1
    srcCols = UBound(srcVals, 2) - LBound(srcVals, 2) + 1
    writeCols = headerCols
    If srcCols < writeCols Then writeCols = srcCols

    If nextRow + srcRows - 1 > PL_EXCEL_MAX_ROWS Then
        wb.Close SaveChanges:=False
        pl_AppendCostFile = False
        Exit Function
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
    rec.Status = "Failed"
    rec.Message = "Import error: " & Err.Description
    rec.RowsImported = 0
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    pl_AppendCostFile = True     ' one file's failure never aborts the run
End Function

' ==========================================================================
'  STEP 2 -- CONSOLIDATE DEDUPE / SITE FILES  -> Site Data
' ==========================================================================

'--------------------------------------------------------------------------
' pl_ConsolidateSite
'   Picks the site source workbooks, reads name/state/voltage from each
'   Summary tab, de-duplicates on the full (name, voltage, state) triple
'   (keeping the first occurrence, logging each drop), and writes Site Data
'   with headers Substation Name | State | Voltage (kV). Returns True and sets
'   wsSite / distinctCount on success.
'--------------------------------------------------------------------------
Private Function pl_ConsolidateSite(ByRef log As PL_TLog, ByRef wsSite As Worksheet, _
                                    ByRef distinctCount As Long) As Boolean
    On Error GoTo ErrHandler
    pl_ConsolidateSite = False

    Dim files() As String, fileCount As Long
    If Not pl_PickFiles("Select the SITE / dedupe workbooks to consolidate", files, fileCount) Then
        pl_LogAdd log, "Step 2: no site files selected."
        Exit Function
    End If

    ' -- accumulate raw rows across all files --
    Dim rawName() As String, rawState() As String
    Dim rawVolt() As Double, rawVoltP() As Boolean
    Dim rawFile() As String, rawRow() As Long
    Dim rawN As Long
    rawN = 0
    ReDim rawName(1 To 16): ReDim rawState(1 To 16)
    ReDim rawVolt(1 To 16): ReDim rawVoltP(1 To 16)
    ReDim rawFile(1 To 16): ReDim rawRow(1 To 16)

    Dim i As Long
    For i = 1 To fileCount
        pl_ReadSiteFile files(i), log, rawName, rawState, rawVolt, rawVoltP, rawFile, rawRow, rawN
    Next i

    If rawN = 0 Then
        pl_LogAdd log, "Step 2: no site rows found on any Summary tab."
        MsgBox "No usable rows found on the Summary tab of the selected site files.", _
               vbExclamation, "Interconnect Pipeline"
        Exit Function
    End If

    ' -- de-duplicate on the triple --
    Dim dt() As PL_TSiteTriple, dropped As Long
    distinctCount = pl_DedupTriples(rawName, rawVolt, rawVoltP, rawState, rawN, _
                                    rawFile, rawRow, dt, log, dropped)

    ' -- write Site Data --
    Set wsSite = pl_GetOutputSheet(PL_SH_SITE)
    If wsSite Is Nothing Then
        pl_LogAdd log, "Step 2: user cancelled Site Data sheet creation."
        Exit Function
    End If

    Dim block() As Variant
    ReDim block(1 To distinctCount + 1, 1 To 3)
    block(1, 1) = PL_HDR_NAME
    block(1, 2) = "State"
    block(1, 3) = PL_HDR_VOLT
    For i = 1 To distinctCount
        block(i + 1, 1) = dt(i).Name
        block(i + 1, 2) = dt(i).State
        If dt(i).VoltParsed Then block(i + 1, 3) = dt(i).Voltage Else block(i + 1, 3) = ""
    Next i
    wsSite.Range(wsSite.Cells(1, 1), wsSite.Cells(distinctCount + 1, 3)).Value = block
    wsSite.Rows(1).Font.Bold = True
    wsSite.Columns.AutoFit

    pl_LogAdd log, "STEP 2 -- Site Data (" & wsSite.Name & "): " & rawN & " raw row(s), " & _
                   dropped & " duplicate(s) dropped, " & distinctCount & " distinct triple(s)."
    pl_ConsolidateSite = True
    Exit Function

ErrHandler:
    pl_LogAdd log, "Step 2 error #" & Err.Number & ": " & Err.Description
    pl_ConsolidateSite = False
End Function

'--------------------------------------------------------------------------
' pl_ReadSiteFile
'   Opens one site workbook read-only, locates its Summary tab, and appends
'   each non-blank row's (name, state, voltage) to the raw accumulators. A
'   missing Summary tab or read error is logged and skipped, never fatal.
'--------------------------------------------------------------------------
Private Sub pl_ReadSiteFile(ByVal path As String, ByRef log As PL_TLog, _
                            ByRef rawName() As String, ByRef rawState() As String, _
                            ByRef rawVolt() As Double, ByRef rawVoltP() As Boolean, _
                            ByRef rawFile() As String, ByRef rawRow() As Long, ByRef rawN As Long)
    Dim wb As Workbook, ws As Worksheet
    Dim ur As Range, lastRow As Long, r As Long
    Dim vals As Variant

    On Error GoTo Fail
    Set wb = Application.Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)

    On Error Resume Next
    Set ws = wb.Worksheets(PL_SITE_TAB)
    On Error GoTo Fail
    If ws Is Nothing Then
        pl_LogAdd log, "  [Skipped] " & pl_FileName(path) & " -- no '" & PL_SITE_TAB & "' tab"
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    Set ur = ws.UsedRange
    lastRow = ur.Row + ur.Rows.Count - 1
    If lastRow <= PL_HDR_ROW Then
        pl_LogAdd log, "  [Skipped] " & pl_FileName(path) & " -- Summary has no data rows"
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' One bulk read spanning columns A..G (name, state, voltage live inside).
    vals = ws.Range(ws.Cells(PL_HDR_ROW + 1, 1), ws.Cells(lastRow, PL_SITE_VOLT_COL)).Value
    If Not IsArray(vals) Then
        Dim tmp(1 To 1, 1 To PL_SITE_VOLT_COL) As Variant
        Dim cc As Long
        For cc = 1 To PL_SITE_VOLT_COL
            tmp(1, cc) = ws.Cells(PL_HDR_ROW + 1, cc).Value
        Next cc
        vals = tmp
    End If

    Dim added As Long
    For r = 1 To UBound(vals, 1)
        Dim nm As String
        nm = Trim$(CStr(pl_NZ(vals(r, PL_SITE_NAME_COL))))
        If Len(nm) > 0 Then
            rawN = rawN + 1
            If rawN > UBound(rawName) Then pl_GrowRaw rawName, rawState, rawVolt, rawVoltP, rawFile, rawRow
            rawName(rawN) = nm
            rawState(rawN) = Trim$(CStr(pl_NZ(vals(r, PL_SITE_STATE_COL))))
            Dim vv As Variant
            vv = vals(r, PL_SITE_VOLT_COL)
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
'   Collapses raw site rows to distinct (name, voltage, state) triples,
'   keeping the first occurrence. Two rows collapse only when name AND voltage
'   AND state are all equal -- a different voltage, or a different state, is a
'   different substation. Each dropped duplicate is logged with its file/row
'   and the triple. Returns the distinct count and fills dt().
'--------------------------------------------------------------------------
Private Function pl_DedupTriples(ByRef rawName() As String, ByRef rawVolt() As Double, _
                                 ByRef rawVoltP() As Boolean, ByRef rawState() As String, _
                                 ByVal rawN As Long, ByRef rawFile() As String, _
                                 ByRef rawRow() As Long, ByRef dt() As PL_TSiteTriple, _
                                 ByRef log As PL_TLog, ByRef dropped As Long) As Long
    Dim k As Long, i As Long, j As Long, isDup As Boolean
    ReDim dt(1 To rawN)
    k = 0
    dropped = 0

    For i = 1 To rawN
        isDup = False
        For j = 1 To k
            If pl_TripleEqual(rawName(i), rawVolt(i), rawVoltP(i), rawState(i), _
                              dt(j).Name, dt(j).Voltage, dt(j).VoltParsed, dt(j).State) Then
                isDup = True: Exit For
            End If
        Next j
        If isDup Then
            dropped = dropped + 1
            pl_LogAdd log, "  [Dropped dup] " & rawFile(i) & " row " & rawRow(i) & " -- (" & _
                           pl_TripleStr(rawName(i), rawVolt(i), rawVoltP(i), rawState(i)) & ")"
        Else
            k = k + 1
            dt(k).Name = rawName(i)
            dt(k).Voltage = rawVolt(i)
            dt(k).VoltParsed = rawVoltP(i)
            dt(k).State = rawState(i)
        End If
    Next i

    If k > 0 Then ReDim Preserve dt(1 To k)
    pl_DedupTriples = k
End Function

' ==========================================================================
'  STEP 3 -- JOIN + MATRIX  (live SUMIFS)
' ==========================================================================

'--------------------------------------------------------------------------
' pl_BuildMatrix
'   Joins Site Data triples to Cost Data identities (intersection only) and
'   writes the Matrix sheet: row 1 = leading label + MW axis; column A = one
'   identity per matched triple; body = cost-per-MW live SUMIFS formulas that
'   resolve Cost Data columns by header name (plus a voltage criterion when
'   the cost tab carried one). Alignment errors (site-without-cost and
'   cost-without-site) are logged, not silently dropped. Returns True with
'   wsMatrix / counts when at least one substation matched.
'--------------------------------------------------------------------------
Private Function pl_BuildMatrix(ByRef log As PL_TLog, ByVal wsCost As Worksheet, _
                                ByVal wsSite As Worksheet, ByRef wsMatrix As Worksheet, _
                                ByRef matchedCount As Long, ByRef mwCount As Long, _
                                ByRef keyName() As String, ByRef keyVolt() As Double, _
                                ByRef keyUseVolt() As Boolean, ByRef costNameOut As String, _
                                ByRef colTrigOut As Long, ByRef colNameOut As Long, _
                                ByRef colVoltOut As Long) As Boolean
    On Error GoTo ErrHandler
    pl_BuildMatrix = False

    ' -- resolve Cost Data columns by header name --
    Dim colName As Long, colVolt As Long, colTrig As Long, colAlloc As Long
    colName = pl_ResolveCol(wsCost, PL_HDR_NAME)
    colVolt = pl_ResolveCol(wsCost, PL_HDR_VOLT)
    colTrig = pl_ResolveCol(wsCost, PL_HDR_TRIGGER)
    colAlloc = pl_ResolveCol(wsCost, PL_HDR_ALLOC)
    If colName = 0 Or colTrig = 0 Or colAlloc = 0 Then
        pl_LogAdd log, "Step 3: could not resolve Cost Data headers (name/trigger/allocation)."
        MsgBox "Cost Data is missing one of the required headers:" & vbCrLf & _
               PL_HDR_NAME & " / " & PL_HDR_TRIGGER & " / " & PL_HDR_ALLOC, _
               vbExclamation, "Interconnect Pipeline"
        Exit Function
    End If

    ' -- distinct cost identities present in Cost Data --
    Dim cid() As PL_TCostId, cidN As Long
    cidN = pl_ReadCostIds(wsCost, colName, colVolt, cid)
    If cidN = 0 Then pl_LogAdd log, "Step 3: no cost identities found.": Exit Function

    ' -- site triples --
    Dim st() As PL_TSiteTriple, stN As Long
    stN = pl_ReadSiteTriples(wsSite, st)
    If stN = 0 Then pl_LogAdd log, "Step 3: no site triples found.": Exit Function

    ' -- join, intersection only --
    Dim matchIdx() As Long          ' matchIdx(s) = cost identity index or 0
    ReDim matchIdx(1 To stN)
    Dim s As Long, c As Long, chosen As Long
    matchedCount = 0
    For s = 1 To stN
        chosen = pl_MatchCost(st(s), cid, cidN)
        matchIdx(s) = chosen
        If chosen > 0 Then
            matchedCount = matchedCount + 1
            cid(chosen).Matched = True
        Else
            pl_LogAdd log, "  [Align: site w/o cost] (" & _
                pl_TripleStr(st(s).Name, st(s).Voltage, st(s).VoltParsed, st(s).State) & ")"
        End If
    Next s
    For c = 1 To cidN
        If Not cid(c).Matched Then
            pl_LogAdd log, "  [Align: cost w/o site] " & cid(c).Name & _
                IIf(cid(c).VoltParsed, " " & cid(c).Voltage & " kV", " (bare)")
        End If
    Next c

    If matchedCount = 0 Then
        pl_LogAdd log, "Step 3: the site and cost sets do not intersect (0 matches)."
        MsgBox "No substation is present in both the site and cost sets; " & _
               "the Matrix would be empty.", vbExclamation, "Interconnect Pipeline"
        Exit Function
    End If

    ' -- MW axis --
    Dim mw() As Double
    mwCount = pl_MwAxis(mw)

    ' -- create Matrix --
    Set wsMatrix = pl_GetOutputSheet(PL_SH_MATRIX)
    If wsMatrix Is Nothing Then
        pl_LogAdd log, "Step 3: user cancelled Matrix sheet creation."
        Exit Function
    End If

    ' -- header row: leading label + MW values --
    wsMatrix.Cells(1, 1).Value = "Substation \ MW"
    Dim j As Long
    For j = 1 To mwCount
        wsMatrix.Cells(1, 1 + j).Value = mw(j)
    Next j

    ' -- one matched row at a time: identity in col A, SUMIFS across the MW block --
    ' Per-row cost keys are captured in matrix-row order so step 4's headroom
    ' MINIFS can reuse the exact same (name [, voltage]) key the SUMIFS use.
    Dim costName As String
    costName = wsCost.Name
    ReDim keyName(1 To matchedCount)
    ReDim keyVolt(1 To matchedCount)
    ReDim keyUseVolt(1 To matchedCount)
    Dim kk As Long: kk = 0
    Dim outRow As Long
    outRow = 1
    For s = 1 To stN
        If matchIdx(s) > 0 Then
            outRow = outRow + 1
            wsMatrix.Cells(outRow, 1).Value = _
                pl_IdentityLabel(st(s).Name, st(s).Voltage, st(s).VoltParsed, st(s).State)

            Dim useVolt As Boolean, theVolt As Double
            useVolt = cid(matchIdx(s)).VoltParsed          ' add voltage criterion for multi-voltage tabs
            theVolt = cid(matchIdx(s)).Voltage
            kk = kk + 1
            keyName(kk) = cid(matchIdx(s)).Name
            keyVolt(kk) = theVolt
            keyUseVolt(kk) = useVolt

            Dim rowF() As Variant
            ReDim rowF(1 To 1, 1 To mwCount)
            For j = 1 To mwCount
                rowF(1, j) = pl_SumifsFormula(costName, colAlloc, colName, colTrig, colVolt, _
                                              cid(matchIdx(s)).Name, useVolt, theVolt, _
                                              pl_ColLetter(1 + j) & "$1")
            Next j
            wsMatrix.Range(wsMatrix.Cells(outRow, 2), wsMatrix.Cells(outRow, 1 + mwCount)).Formula = rowF
        End If
    Next s

    wsMatrix.Rows(1).Font.Bold = True
    wsMatrix.Columns(1).AutoFit
    wsMatrix.Range(wsMatrix.Cells(2, 2), wsMatrix.Cells(outRow, 1 + mwCount)).NumberFormat = "$#,##0"

    ' Hand the resolved columns and cost sheet name to step 4 for the headroom
    ' MINIFS (resolved by header name, same as the matrix SUMIFS).
    costNameOut = costName
    colTrigOut = colTrig
    colNameOut = colName
    colVoltOut = colVolt

    pl_LogAdd log, "STEP 3 -- Matrix (" & wsMatrix.Name & "): " & matchedCount & _
                   " matched substation(s) x " & mwCount & " MW; SUMIFS resolve Cost Data " & _
                   "cols by header (name=" & pl_ColLetter(colName) & ", trigger=" & _
                   pl_ColLetter(colTrig) & ", alloc=" & pl_ColLetter(colAlloc) & ")."
    pl_BuildMatrix = True
    Exit Function

ErrHandler:
    pl_LogAdd log, "Step 3 error #" & Err.Number & ": " & Err.Description
    pl_BuildMatrix = False
End Function

'--------------------------------------------------------------------------
' pl_ReadCostIds
'   Distinct (name, voltage) identities present in Cost Data. VoltParsed is
'   True when the Voltage (kV) cell is non-blank (i.e. the cost tab carried a
'   voltage), which drives whether the join matches on name+voltage or name.
'--------------------------------------------------------------------------
Private Function pl_ReadCostIds(ByVal wsCost As Worksheet, ByVal colName As Long, _
                                ByVal colVolt As Long, ByRef cid() As PL_TCostId) As Long
    Dim lastRow As Long, r As Long, k As Long, j As Long, dup As Boolean
    Dim nm As String, vp As Boolean, vv As Double
    lastRow = wsCost.Cells(wsCost.Rows.Count, colName).End(xlUp).Row
    ReDim cid(1 To Application.WorksheetFunction.Max(1, lastRow))
    k = 0
    For r = PL_HDR_ROW + 1 To lastRow
        nm = Trim$(CStr(pl_NZ(wsCost.Cells(r, colName).Value)))
        If Len(nm) > 0 Then
            vp = False: vv = 0
            If colVolt > 0 Then
                Dim raw As Variant
                raw = wsCost.Cells(r, colVolt).Value
                If IsNumeric(raw) And Len(Trim$(CStr(pl_NZ(raw)))) > 0 Then vv = CDbl(raw): vp = True
            End If
            dup = False
            For j = 1 To k
                If StrComp(cid(j).Name, nm, vbTextCompare) = 0 And _
                   cid(j).VoltParsed = vp And cid(j).Voltage = vv Then dup = True: Exit For
            Next j
            If Not dup Then
                k = k + 1
                cid(k).Name = nm: cid(k).Voltage = vv
                cid(k).VoltParsed = vp: cid(k).Matched = False
            End If
        End If
    Next r
    If k > 0 Then ReDim Preserve cid(1 To k)
    pl_ReadCostIds = k
End Function

'--------------------------------------------------------------------------
' pl_ReadSiteTriples
'   Reads the distinct triples straight back off the Site Data sheet.
'--------------------------------------------------------------------------
Private Function pl_ReadSiteTriples(ByVal wsSite As Worksheet, ByRef st() As PL_TSiteTriple) As Long
    Dim lastRow As Long, r As Long, k As Long
    lastRow = wsSite.Cells(wsSite.Rows.Count, 1).End(xlUp).Row
    If lastRow < PL_HDR_ROW + 1 Then pl_ReadSiteTriples = 0: Exit Function
    ReDim st(1 To lastRow)
    k = 0
    For r = PL_HDR_ROW + 1 To lastRow
        Dim nm As String
        nm = Trim$(CStr(pl_NZ(wsSite.Cells(r, 1).Value)))
        If Len(nm) > 0 Then
            k = k + 1
            st(k).Name = nm
            st(k).State = Trim$(CStr(pl_NZ(wsSite.Cells(r, 2).Value)))
            Dim vv As Variant
            vv = wsSite.Cells(r, 3).Value
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
' pl_MatchCost
'   Finds the cost identity for a site triple. Match on name; when a matching
'   cost identity carries a voltage, the site voltage must match it too (this
'   picks the correct one of several same-named multi-voltage cost tabs); when
'   the cost identity is bare, name alone matches. Returns the cost index or 0.
'--------------------------------------------------------------------------
Private Function pl_MatchCost(ByRef tr As PL_TSiteTriple, ByRef cid() As PL_TCostId, _
                              ByVal cidN As Long) As Long
    Dim c As Long
    ' First pass: honour a voltage-bearing cost tab (name + voltage).
    For c = 1 To cidN
        If StrComp(cid(c).Name, tr.Name, vbTextCompare) = 0 Then
            If cid(c).VoltParsed Then
                If tr.VoltParsed And cid(c).Voltage = tr.Voltage Then pl_MatchCost = c: Exit Function
            End If
        End If
    Next c
    ' Second pass: a bare cost tab matches on name alone.
    For c = 1 To cidN
        If StrComp(cid(c).Name, tr.Name, vbTextCompare) = 0 Then
            If Not cid(c).VoltParsed Then pl_MatchCost = c: Exit Function
        End If
    Next c
    pl_MatchCost = 0
End Function

'--------------------------------------------------------------------------
' pl_SumifsFormula
'   Builds the live cost-per-MW SUMIFS for one (substation, MW) cell:
'     =SUMIFS('Cost'!alloc, 'Cost'!name, "sub", 'Cost'!trig, "<="&MW [,
'             'Cost'!volt, v]) / MW
'   Cost Data columns are full-column references resolved by header name; the
'   MW criterion and the divisor both reference the MW header cell so editing
'   the header re-drives the row.
'--------------------------------------------------------------------------
Private Function pl_SumifsFormula(ByVal costName As String, ByVal colAlloc As Long, _
                                  ByVal colName As Long, ByVal colTrig As Long, _
                                  ByVal colVolt As Long, ByVal subName As String, _
                                  ByVal useVolt As Boolean, ByVal theVolt As Double, _
                                  ByVal mwCell As String) As String
    Dim ref As String, f As String
    ref = pl_SheetRef(costName)
    f = "=SUMIFS(" & ref & "!" & pl_FullCol(colAlloc) & _
        "," & ref & "!" & pl_FullCol(colName) & ",""" & pl_EscQuote(subName) & """" & _
        "," & ref & "!" & pl_FullCol(colTrig) & ",""<=""&" & mwCell
    If useVolt And colVolt > 0 Then
        f = f & "," & ref & "!" & pl_FullCol(colVolt) & "," & pl_NumStr(theVolt)
    End If
    f = f & ")/" & mwCell
    pl_SumifsFormula = f
End Function

'--------------------------------------------------------------------------
' pl_MinifsFormula
'   Builds the MINIFS expression (no leading "=") for a substation's headroom:
'     MINIFS('Cost'!trig, 'Cost'!name, "sub" [, 'Cost'!volt, v])
'   Columns are full-column references resolved by header name; the voltage
'   criterion is added only when the cost tab carried a voltage -- the same key
'   the matrix SUMIFS use.
'--------------------------------------------------------------------------
Private Function pl_MinifsFormula(ByVal costName As String, ByVal colTrig As Long, _
                                  ByVal colName As Long, ByVal colVolt As Long, _
                                  ByVal subName As String, ByVal useVolt As Boolean, _
                                  ByVal theVolt As Double) As String
    Dim ref As String, f As String
    ref = pl_SheetRef(costName)
    f = "MINIFS(" & ref & "!" & pl_FullCol(colTrig) & _
        "," & ref & "!" & pl_FullCol(colName) & ",""" & pl_EscQuote(subName) & """"
    If useVolt And colVolt > 0 Then
        f = f & "," & ref & "!" & pl_FullCol(colVolt) & "," & pl_NumStr(theVolt)
    End If
    f = f & ")"
    pl_MinifsFormula = f
End Function

' ==========================================================================
'  STEP 4 -- ANALYSIS + RANKING + SCORING  -> Cost Curve Analysis
'  Drives the reused modCostCurves logic on the Matrix ranges, then adds the
'  weighted composite score. Contains no InputBox, no Activate.
' ==========================================================================

'--------------------------------------------------------------------------
' pl_RunAnalysisAndScoring
'   Reads the recomputed Matrix, runs ComputeRankings + AnalyseAll, writes the
'   analysis table / chart / leaderboard (reused WriteSheet + BuildChart), then
'   the weighted composite score block and the live $B$2 weight cell.
'--------------------------------------------------------------------------
Private Function pl_RunAnalysisAndScoring(ByVal wsMatrix As Worksheet, ByRef log As PL_TLog, _
                                          ByRef keyName() As String, ByRef keyVolt() As Double, _
                                          ByRef keyUseVolt() As Boolean, ByVal costName As String, _
                                          ByVal colTrig As Long, ByVal colName As Long, _
                                          ByVal colVolt As Long) As Boolean
    On Error GoTo ErrHandler
    pl_RunAnalysisAndScoring = False

    Dim mw() As Double, names() As String, costM() As Double, n As Long, m As Long
    If Not pl_ReadMatrix(wsMatrix, mw, names, costM, n, m) Then
        pl_LogAdd log, "Step 4: Matrix did not validate as a cost-per-MW block; analysis skipped."
        MsgBox "The Matrix did not validate as a numeric cost-per-MW block " & _
               "(every cell must be > 0 after recompute). Analysis skipped.", _
               vbExclamation, "Interconnect Pipeline"
        Exit Function
    End If

    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1
    Dim totalCols As Long: totalCols = 7 + 4 * nt      ' N_FIXED_COLS + 4 per threshold

    Dim rkCount() As Long, rkMaxMW() As Double, rkCostMax() As Double
    Dim rkSlope() As Double, rkHasSlope() As Boolean
    Dim rankA() As Long, rankB() As Long
    ComputeRankings mw, names, costM, n, m, rkCount, rkMaxMW, rkCostMax, rkSlope, rkHasSlope, rankA, rankB

    Dim aTable() As Variant
    ReDim aTable(1 To n, 1 To totalCols)
    AnalyseAll mw, names, costM, n, m, aTable, rankA, rankB

    Dim wsOut As Worksheet
    Set wsOut = pl_GetOutputSheet(PL_SH_ANALYSIS)
    If wsOut Is Nothing Then
        pl_LogAdd log, "Step 4: user cancelled Cost Curve Analysis sheet creation."
        Exit Function
    End If

    Dim srcMWRow As Long, srcFirstDataRow As Long
    Dim hlpMWCol As Long, hlpFirstThreshCol As Long, hlpFirstRow As Long, hlpRows As Long
    WriteSheet wsOut, mw, names, costM, n, m, aTable, totalCols, _
               rkCount, rkMaxMW, rkSlope, rankA, rankB, _
               srcMWRow, srcFirstDataRow, hlpMWCol, hlpFirstThreshCol, hlpFirstRow, hlpRows

    BuildChart wsOut, mw, costM, n, m, srcMWRow, srcFirstDataRow, _
               hlpMWCol, hlpFirstThreshCol, hlpFirstRow, hlpRows

    pl_WriteScoring wsOut, mw, names, costM, n, m, rankA, rankB, totalCols, _
                    keyName, keyVolt, keyUseVolt, costName, colTrig, colName, colVolt

    ' Resolve the live band-score / composite / headroom formulas now, so the
    ' sheet shows values even if the workbook's calc mode was left on manual.
    Application.CalculateFull

    pl_LogAdd log, "STEP 4 -- Cost Curve Analysis (" & wsOut.Name & "): " & n & _
                   " substation(s), " & m & " MW points; composite ceiling 165, weight $B$2=" & PL_WEIGHT_B2 & "."
    pl_RunAnalysisAndScoring = True
    Exit Function

ErrHandler:
    pl_LogAdd log, "Step 4 error #" & Err.Number & ": " & Err.Description
    MsgBox "Error during analysis/scoring #" & Err.Number & ": " & Err.Description, _
           vbCritical, "Interconnect Pipeline"
    pl_RunAnalysisAndScoring = False
End Function

'--------------------------------------------------------------------------
' pl_ReadMatrix
'   Reads the recomputed Matrix into typed arrays: column A (row 2 down) =
'   identity labels; row 1 (col 2 right) = MW axis; body = cost-per-MW values.
'   Validates m >= 3, a positive strictly-increasing MW header, and every body
'   cell usable (> 0) -- the same contract the analysis input expects.
'--------------------------------------------------------------------------
Private Function pl_ReadMatrix(ByVal ws As Worksheet, ByRef mw() As Double, _
                               ByRef names() As String, ByRef costM() As Double, _
                               ByRef n As Long, ByRef m As Long) As Boolean
    pl_ReadMatrix = False

    ' width: MW header cells across row 1 from col 2 until blank
    m = 0
    Do While Len(Trim$(CStr(pl_NZ(ws.Cells(1, 2 + m).Value)))) > 0
        m = m + 1
    Loop
    ' height: identity labels down col A from row 2 until blank
    n = 0
    Do While Len(Trim$(CStr(pl_NZ(ws.Cells(2 + n, 1).Value)))) > 0
        n = n + 1
    Loop
    If m < 3 Or n < 1 Then Exit Function

    ReDim mw(1 To m): ReDim names(1 To n): ReDim costM(1 To n, 1 To m)

    Dim j As Long, prev As Double
    For j = 1 To m
        Dim hv As Variant: hv = ws.Cells(1, 1 + j).Value
        If Not IsNumeric(hv) Then Exit Function
        mw(j) = CDbl(hv)
        If mw(j) <= 0 Then Exit Function
        If j > 1 Then If mw(j) <= prev Then Exit Function
        prev = mw(j)
    Next j

    Dim i As Long
    For i = 1 To n
        names(i) = CStr(ws.Cells(1 + i, 1).Value)
        For j = 1 To m
            Dim v As Variant: v = ws.Cells(1 + i, 1 + j).Value
            If Not IsUsableNumber(v) Then Exit Function     ' reused: blank/non-numeric/<=0 fails
            costM(i, j) = CDbl(v)
        Next j
    Next i

    pl_ReadMatrix = True
End Function

'--------------------------------------------------------------------------
' pl_WriteScoring
'   Writes the live weight cell ($B$2) and a weighted composite score block to
'   the right of the analysis table. Each band contributes
'   $B$2 x (breadth pctile + slope pctile); the composite adds the flattening
'   pctile. Percentiles (0..5) are computed in VBA; band-score and composite
'   cells are FORMULAS referencing $B$2 so re-weighting is live.
'--------------------------------------------------------------------------
Private Sub pl_WriteScoring(ByVal ws As Worksheet, ByRef mw() As Double, _
                            ByRef names() As String, ByRef costM() As Double, _
                            ByVal n As Long, ByVal m As Long, _
                            ByRef rankA() As Long, ByRef rankB() As Long, ByVal totalCols As Long, _
                            ByRef keyName() As String, ByRef keyVolt() As Double, _
                            ByRef keyUseVolt() As Boolean, ByVal costName As String, _
                            ByVal colTrig As Long, ByVal colName As Long, ByVal colVolt As Long)
    Dim thr As Variant: thr = GetThresholds()
    Dim nt As Long: nt = UBound(thr) - LBound(thr) + 1
    Dim i As Long, tt As Long, j As Long
    Dim EMPTY_LIT As String: EMPTY_LIT = Chr$(34) & Chr$(34)   ' Excel "" literal inside a formula

    ' -- live weight cell --
    ws.Cells(2, 1).Value = "Threshold weight (applies to every band):"
    ws.Cells(2, 2).Value = PL_WEIGHT_B2
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 1).Font.Italic = True

    ' -- flattening pctile: rank by geometric knee (earlier flattening = better) --
    Dim knee() As Double: ReDim knee(1 To n)
    Dim c() As Double: ReDim c(1 To m)
    Dim T() As Double: ReDim T(1 To m)
    For i = 1 To n
        For j = 1 To m: c(j) = costM(i, j): T(j) = c(j) * mw(j): Next j
        Dim segs() As TSegment: segs = Segmentize(mw, T, m)
        Dim li As Long: li = LongestSegIdx(mw, segs, UBound(segs))
        knee(i) = KneeXStar(mw, T, segs(li))
    Next i
    Dim rankFlat() As Long: rankFlat = pl_CompRankAsc(knee, n)
    Dim pFlat() As Long: ReDim pFlat(1 To n)
    For i = 1 To n: pFlat(i) = pl_RankToPctile(rankFlat(i), n): Next i

    ' -- per-band populations and percentiles --
    Dim popA() As Long, popB() As Long: ReDim popA(1 To nt): ReDim popB(1 To nt)
    For tt = 1 To nt
        For i = 1 To n
            If rankA(i, tt) > 0 Then popA(tt) = popA(tt) + 1
            If rankB(i, tt) > 0 Then popB(tt) = popB(tt) + 1
        Next i
    Next tt

    ' -- header --
    Dim c0 As Long: c0 = totalCols + 2      ' one blank spacer column after the table
    Dim r0 As Long: r0 = 3                  ' align with the analysis table header row
    ws.Cells(r0, c0).Value = "Substation"
    ws.Cells(r0, c0 + 1).Value = "Flattening Pctile (0-5)"
    Dim base As Long
    For tt = 0 To nt - 1
        base = c0 + 2 + 3 * tt
        ws.Cells(r0, base).Value = ThreshLabel(CDbl(thr(LBound(thr) + tt))) & " Breadth (0-5)"
        ws.Cells(r0, base + 1).Value = ThreshLabel(CDbl(thr(LBound(thr) + tt))) & " Slope (0-5)"
        ws.Cells(r0, base + 2).Value = ThreshLabel(CDbl(thr(LBound(thr) + tt))) & " Band Score"
    Next tt
    Dim compCol As Long: compCol = c0 + 2 + 3 * nt
    ws.Cells(r0, compCol).Value = "Composite Score (max 165)"

    ' -- standalone headroom lens (NOT part of the composite; ranked on its own) --
    Dim haveKeys As Boolean
    haveKeys = False
    On Error Resume Next
    haveKeys = (UBound(keyName) >= n)
    On Error GoTo 0
    Dim hrCol As Long, hrpCol As Long, hrLetter As String, firstDR As Long, lastDR As Long
    If haveKeys Then
        hrCol = compCol + 1
        hrpCol = compCol + 2
        hrLetter = pl_ColLetter(hrCol)
        firstDR = r0 + 1
        lastDR = r0 + n
        ws.Cells(r0, hrCol).Value = "Headroom (MW)"
        ws.Cells(r0, hrpCol).Value = "Headroom Pctile (1-5)"
    End If

    ' -- rows (in analysis-table order) --
    Dim rr As Long
    For i = 1 To n
        rr = r0 + i
        ws.Cells(rr, c0).Value = names(i)
        ws.Cells(rr, c0 + 1).Value = pFlat(i)

        Dim bandRefs As String: bandRefs = ""
        For tt = 1 To nt
            base = c0 + 2 + 3 * (tt - 1)
            Dim pb As Long, ps As Long
            pb = pl_RankToPctile(rankA(i, tt), popA(tt))
            ps = pl_RankToPctile(rankB(i, tt), popB(tt))
            ws.Cells(rr, base).Value = pb
            ws.Cells(rr, base + 1).Value = ps
            ' Band score = $B$2 * (breadth + slope), live on the weight cell.
            ws.Cells(rr, base + 2).Formula = "=$B$2*(" & pl_ColLetter(base) & rr & "+" & _
                                             pl_ColLetter(base + 1) & rr & ")"
            If Len(bandRefs) > 0 Then bandRefs = bandRefs & "+"
            bandRefs = bandRefs & pl_ColLetter(base + 2) & rr
        Next tt
        ' Composite = flattening pctile + sum of band scores. Headroom is
        ' deliberately absent here: it is a standalone lens, not a score input.
        ws.Cells(rr, compCol).Formula = "=" & pl_ColLetter(c0 + 1) & rr & "+" & bandRefs

        If haveKeys Then
            ' Headroom (MW) = minimum overload-trigger size, straight from the
            ' Cost Data trigger column (not the clipped matrix cells), keyed by
            ' the same name (+voltage when the cost tab carried one) as the row.
            ' MINIFS returns 0 when a substation has no trigger records; show
            ' blank so it drops out of the percentile population.
            Dim minifs As String
            minifs = pl_MinifsFormula(costName, colTrig, colName, colVolt, _
                                      keyName(i), keyUseVolt(i), keyVolt(i))
            ws.Cells(rr, hrCol).Formula = "=IFERROR(IF(" & minifs & "=0," & EMPTY_LIT & _
                                          "," & minifs & ")," & EMPTY_LIT & ")"
            ' Headroom percentile, 1-5, more headroom is better (no inversion),
            ' absolute population over every data row.
            ws.Cells(rr, hrpCol).Formula = "=IFERROR(CEILING(PERCENTRANK.EXC($" & hrLetter & "$" & _
                firstDR & ":$" & hrLetter & "$" & lastDR & "," & hrLetter & rr & ")*5,1)," & EMPTY_LIT & ")"
        End If
    Next i

    ' -- number formats for the headroom columns --
    If haveKeys Then
        ws.Range(ws.Cells(r0 + 1, hrCol), ws.Cells(r0 + n, hrCol)).NumberFormat = "#,##0"
        With ws.Range(ws.Cells(r0 + 1, hrpCol), ws.Cells(r0 + n, hrpCol))
            .NumberFormat = "0"
            .HorizontalAlignment = xlCenter
        End With
    End If

    ' -- styling + explanatory note --
    Dim lastHdrCol As Long
    lastHdrCol = compCol
    If haveKeys Then lastHdrCol = hrpCol
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0, lastHdrCol)).Font.Bold = True
    ws.Range(ws.Cells(r0, c0), ws.Cells(r0, lastHdrCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(r0 + n + 2, c0).Value = _
        "Composite max = 165 = flattening (5) + 4 bands x $B$2 x (breadth 5 + slope 5). " & _
        "Practical max today = 125: the $25MM band scores 0 for all (cheapest total > $25MM). " & _
        "The single $B$2 weights the threshold axis as a whole, not $25 > $50 > $75 > $100. " & _
        "Headroom (MW) and its percentile are a standalone lens -- NOT part of the composite."
    ws.Cells(r0 + n + 2, c0).Font.Italic = True
End Sub

' ==========================================================================
'  SCORING PRIMITIVES  (pure; unit-tested in SelfTest)
' ==========================================================================

'--------------------------------------------------------------------------
' pl_ScoreCeiling
'   Composite ceiling for nBands active bands:
'     PL_MAX_PCTILE + nBands * PL_WEIGHT_B2 * (2 * PL_MAX_PCTILE)
'   4 bands -> 5 + 4*4*10 = 165 ; 3 bands -> 125.
'--------------------------------------------------------------------------
Private Function pl_ScoreCeiling(ByVal nBands As Long) As Double
    pl_ScoreCeiling = PL_MAX_PCTILE + nBands * PL_WEIGHT_B2 * (2 * PL_MAX_PCTILE)
End Function

'--------------------------------------------------------------------------
' pl_RankToPctile
'   Maps a competition rank (1 = best) within a population of size pop to a
'   0..PL_MAX_PCTILE score. Rank 0 (not ranked / no qualifier) scores 0; a
'   sole ranked member scores the maximum.
'--------------------------------------------------------------------------
Private Function pl_RankToPctile(ByVal rank As Long, ByVal pop As Long) As Long
    If rank <= 0 Or pop <= 0 Then
        pl_RankToPctile = 0
    ElseIf pop = 1 Then
        pl_RankToPctile = PL_MAX_PCTILE
    Else
        Dim v As Long
        v = CLng(Round((CDbl(pop - rank) / CDbl(pop - 1)) * PL_MAX_PCTILE, 0))
        If v < 0 Then v = 0
        If v > PL_MAX_PCTILE Then v = PL_MAX_PCTILE
        pl_RankToPctile = v
    End If
End Function

'--------------------------------------------------------------------------
' pl_CompRankAsc
'   Competition rank of each value, smallest value = rank 1, equal values
'   share the lower rank and the next distinct value skips ahead (1,2,2,4).
'--------------------------------------------------------------------------
Private Function pl_CompRankAsc(ByRef vals() As Double, ByVal n As Long) As Long()
    Dim ord() As Long, i As Long, a As Long, b As Long, keyi As Long
    ReDim ord(1 To n)
    For i = 1 To n: ord(i) = i: Next i
    ' insertion sort ascending by value
    For a = 2 To n
        keyi = ord(a): b = a - 1
        Do While b >= 1
            If vals(ord(b)) > vals(keyi) Then ord(b + 1) = ord(b): b = b - 1 Else Exit Do
        Loop
        ord(b + 1) = keyi
    Next a
    Dim rank() As Long: ReDim rank(1 To n)
    Dim p As Long
    For p = 1 To n
        If p = 1 Then
            rank(ord(p)) = 1
        ElseIf vals(ord(p)) = vals(ord(p - 1)) Then
            rank(ord(p)) = rank(ord(p - 1))
        Else
            rank(ord(p)) = p
        End If
    Next p
    pl_CompRankAsc = rank
End Function

'--------------------------------------------------------------------------
' pl_MinHeadroom
'   Minimum positive value (the binding overload-trigger size). Returns 0 when
'   there is no positive trigger. Pure mirror of the sheet's MINIFS, for tests.
'--------------------------------------------------------------------------
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

'--------------------------------------------------------------------------
' pl_PctExcCeil5
'   Pure mirror of the sheet formula CEILING(PERCENTRANK.EXC(range, x) * 5, 1)
'   for a value x within the population -- more headroom scores higher (no
'   inversion), so the largest maps to 5 and the smallest to 1.
'--------------------------------------------------------------------------
Private Function pl_PctExcCeil5(ByRef v() As Double, ByVal n As Long, ByVal x As Double) As Long
    Dim i As Long, lessCount As Long
    For i = 1 To n
        If v(i) < x Then lessCount = lessCount + 1
    Next i
    Dim val As Double
    val = ((lessCount + 1) / (n + 1)) * 5     ' PERCENTRANK.EXC = k/(N+1), k=1..N
    pl_PctExcCeil5 = -Int(-val)               ' CEILING(., 1)
End Function

' ==========================================================================
'  TAB-NAME PARSING  (ported from the consolidator; never fails)
' ==========================================================================

'--------------------------------------------------------------------------
' pl_ParseTab
'   Extracts substation name + optional voltage from a sheet-3 tab name. The
'   LAST plausible numeric token is the voltage; the rest, cleaned, is the
'   name. When no voltage is found the whole cleaned tab is the name and
'   parsed = False. Examples: "Chaves County 345" -> ("Chaves County", 345);
'   "Cunningham" -> ("Cunningham", blank).
'--------------------------------------------------------------------------
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
'  TRIPLE-IDENTITY HELPERS
' ==========================================================================

'--------------------------------------------------------------------------
' pl_TripleEqual
'   True only when name AND voltage AND state all match. Names/states compare
'   case-insensitively after trimming; a parsed voltage never equals a blank
'   voltage; two blank voltages are equal.
'--------------------------------------------------------------------------
Private Function pl_TripleEqual(ByVal n1 As String, ByVal v1 As Double, ByVal p1 As Boolean, ByVal s1 As String, _
                                ByVal n2 As String, ByVal v2 As Double, ByVal p2 As Boolean, ByVal s2 As String) As Boolean
    pl_TripleEqual = False
    If StrComp(Trim$(n1), Trim$(n2), vbTextCompare) <> 0 Then Exit Function
    If StrComp(Trim$(s1), Trim$(s2), vbTextCompare) <> 0 Then Exit Function
    If p1 <> p2 Then Exit Function
    If p1 Then If v1 <> v2 Then Exit Function
    pl_TripleEqual = True
End Function

Private Function pl_TripleStr(ByVal nm As String, ByVal v As Double, ByVal p As Boolean, ByVal st As String) As String
    pl_TripleStr = nm & ", " & IIf(p, CStr(v) & " kV", "(no voltage)") & ", " & IIf(Len(st) > 0, st, "(no state)")
End Function

Private Function pl_IdentityLabel(ByVal nm As String, ByVal v As Double, ByVal p As Boolean, ByVal st As String) As String
    Dim lab As String
    lab = nm
    If p Then lab = lab & " " & Format$(v, "0.###") & " kV"
    If Len(Trim$(st)) > 0 Then lab = lab & " (" & Trim$(st) & ")"
    pl_IdentityLabel = lab
End Function

' ==========================================================================
'  SHARED SHEET / IO / STRING HELPERS
' ==========================================================================

'--------------------------------------------------------------------------
' pl_PickFiles
'   Multi-select Excel file picker rooted at ThisWorkbook's folder, falling
'   back to Application.DefaultFilePath for an unsaved blank workbook.
'--------------------------------------------------------------------------
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

'--------------------------------------------------------------------------
' pl_GetOutputSheet
'   Returns an empty output sheet of the given base name. When one exists,
'   offers overwrite / new-timestamped / cancel (as the existing consolidator
'   does). Returns Nothing only when the user cancels.
'--------------------------------------------------------------------------
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
                 "Yes = overwrite it" & vbCrLf & _
                 "No  = create a new timestamped sheet" & vbCrLf & _
                 "Cancel = abort", _
                 vbYesNoCancel + vbQuestion, "Sheet exists")
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

'--------------------------------------------------------------------------
' pl_ResolveCol
'   1-based column index of the first header cell (row 1) whose trimmed text
'   equals headerText (case-insensitive), or 0 when absent.
'--------------------------------------------------------------------------
Private Function pl_ResolveCol(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(PL_HDR_ROW, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(pl_NZ(ws.Cells(PL_HDR_ROW, c).Value))), headerText, vbTextCompare) = 0 Then
            pl_ResolveCol = c: Exit Function
        End If
    Next c
    pl_ResolveCol = 0
End Function

'--------------------------------------------------------------------------
' pl_MwAxis
'   Fills mw() with PL_MW_MIN..PL_MW_MAX step PL_MW_STEP and returns the count.
'--------------------------------------------------------------------------
Private Function pl_MwAxis(ByRef mw() As Double) As Long
    Dim cnt As Long, x As Double
    cnt = CLng((PL_MW_MAX - PL_MW_MIN) / PL_MW_STEP) + 1
    ReDim mw(1 To cnt)
    Dim j As Long
    For j = 1 To cnt
        mw(j) = PL_MW_MIN + (j - 1) * PL_MW_STEP
    Next j
    pl_MwAxis = cnt
End Function

'--------------------------------------------------------------------------
' pl_BlockRead
'   Reads a rectangular value range into a 1-based 2-D array; a single cell is
'   normalised to a 1x1 array. Empty when the range has no rows.
'--------------------------------------------------------------------------
Private Function pl_BlockRead(ByVal ws As Worksheet, ByVal firstRow As Long, _
                              ByVal lastRow As Long, ByVal cols As Long) As Variant
    Dim rng As Range, tmp(1 To 1, 1 To 1) As Variant
    If lastRow < firstRow Then pl_BlockRead = Empty: Exit Function
    Set rng = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, cols))
    If rng.Cells.Count = 1 Then
        tmp(1, 1) = rng.Value
        pl_BlockRead = tmp
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
    Dim newSize As Long
    newSize = UBound(a) * 2
    ReDim Preserve a(1 To newSize): ReDim Preserve b(1 To newSize)
    ReDim Preserve c(1 To newSize): ReDim Preserve d(1 To newSize)
    ReDim Preserve e(1 To newSize): ReDim Preserve f(1 To newSize)
End Sub

'--------------------------------------------------------------------------
' pl_ColLetter / pl_FullCol / pl_SheetRef / pl_EscQuote / pl_NumStr / pl_NZ /
' pl_FileName  -- small formatting helpers.
'--------------------------------------------------------------------------
Private Function pl_ColLetter(ByVal col As Long) As String
    Dim s As String, r As Long
    Do While col > 0
        r = (col - 1) Mod 26
        s = Chr$(65 + r) & s
        col = (col - 1) \ 26
    Loop
    pl_ColLetter = s
End Function

Private Function pl_FullCol(ByVal col As Long) As String
    Dim L As String: L = pl_ColLetter(col)
    pl_FullCol = "$" & L & ":$" & L
End Function

Private Function pl_SheetRef(ByVal nm As String) As String
    pl_SheetRef = "'" & Replace(nm, "'", "''") & "'"
End Function

Private Function pl_EscQuote(ByVal s As String) As String
    pl_EscQuote = Replace(s, """", """""")
End Function

Private Function pl_NumStr(ByVal v As Double) As String
    ' Locale-independent number literal for a formula (always a dot decimal).
    pl_NumStr = Format$(v, "0.###############")
    pl_NumStr = Replace(pl_NumStr, ",", ".")
End Function

Private Function pl_NZ(ByVal v As Variant) As Variant
    If IsError(v) Then pl_NZ = "" Else If IsNull(v) Then pl_NZ = "" Else pl_NZ = v
End Function

Private Function pl_FileName(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, Application.PathSeparator)
    If p = 0 Then p = InStrRev(fullPath, "\")
    If p = 0 Then pl_FileName = fullPath Else pl_FileName = Mid$(fullPath, p + 1)
End Function

' ==========================================================================
'  RUN LOG
' ==========================================================================

Private Sub pl_LogInit(ByRef log As PL_TLog)
    ReDim log.lines(1 To 16)
    log.n = 0
End Sub

Private Sub pl_LogAdd(ByRef log As PL_TLog, ByVal s As String)
    log.n = log.n + 1
    If log.n > UBound(log.lines) Then ReDim Preserve log.lines(1 To UBound(log.lines) * 2)
    log.lines(log.n) = s
End Sub

'--------------------------------------------------------------------------
' pl_WriteLog
'   Rewrites the _Pipeline Log sheet with the accumulated lines (per-file
'   status, dropped duplicates, alignment errors). Recreated each run.
'--------------------------------------------------------------------------
Private Sub pl_WriteLog(ByRef log As PL_TLog)
    Dim ws As Worksheet, i As Long
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
    For i = 1 To log.n
        ws.Cells(1 + i, 1).Value = log.lines(i)
    Next i
    ws.Columns(1).ColumnWidth = 120
End Sub

' ==========================================================================
'  FRONT-HALF SELF-TEST  (called from SelfTest; shares Assert / counts)
' ==========================================================================

'--------------------------------------------------------------------------
' pl_FrontHalfSelfTest
'   In-memory assertions for steps 1-3 and the scoring layer -- no file
'   pickers, no sheets touched. Uses the reused Assert / LoadFixture /
'   ComputeRankings helpers.
'--------------------------------------------------------------------------
Private Sub pl_FrontHalfSelfTest(ByRef passCount As Long, ByRef failCount As Long)
    Debug.Print "--- front-half (pipeline) tests ---"

    ' 1) Tab-name parsing
    Dim nm As String, v As Double, p As Boolean
    pl_ParseTab "Chaves County 345", nm, v, p
    Assert (nm = "Chaves County") And p And (v = 345), _
           "Tab parse: 'Chaves County 345' -> ('Chaves County', 345)", passCount, failCount
    pl_ParseTab "Cunningham", nm, v, p
    Assert (nm = "Cunningham") And (Not p), _
           "Tab parse: 'Cunningham' -> ('Cunningham', blank)", passCount, failCount

    ' 2) Triple dedupe
    Dim rn() As String, rst() As String, rv() As Double, rp() As Boolean, rf() As String, rr() As Long
    ReDim rn(1 To 5): ReDim rst(1 To 5): ReDim rv(1 To 5)
    ReDim rp(1 To 5): ReDim rf(1 To 5): ReDim rr(1 To 5)
    pl_FillSite rn, rv, rp, rst, rf, rr, 1, "Alpha", 345, True, "NM"
    pl_FillSite rn, rv, rp, rst, rf, rr, 2, "Alpha", 345, True, "NM"   ' dup of 1
    pl_FillSite rn, rv, rp, rst, rf, rr, 3, "Alpha", 230, True, "NM"   ' diff voltage
    pl_FillSite rn, rv, rp, rst, rf, rr, 4, "Alpha", 345, True, "TX"   ' diff state
    pl_FillSite rn, rv, rp, rst, rf, rr, 5, "Alpha", 345, True, "NM"   ' dup of 1
    Dim dt() As PL_TSiteTriple, dropped As Long, dlog As PL_TLog
    pl_LogInit dlog
    Dim dc As Long
    dc = pl_DedupTriples(rn, rv, rp, rst, 5, rf, rr, dt, dlog, dropped)
    Assert dc = 3, "Triple dedupe: 5 rows -> 3 distinct", passCount, failCount
    Assert dropped = 2, "Triple dedupe: 2 duplicates dropped and logged", passCount, failCount
    Assert pl_HasTriple(dt, dc, "Alpha", 230, True, "NM"), _
           "Triple dedupe: voltage-differ row kept separate", passCount, failCount
    Assert pl_HasTriple(dt, dc, "Alpha", 345, True, "TX"), _
           "Triple dedupe: state-differ row kept separate", passCount, failCount

    ' 3) Join intersection
    Dim sTr() As PL_TSiteTriple: ReDim sTr(1 To 3)
    sTr(1).Name = "Alpha": sTr(1).Voltage = 345: sTr(1).VoltParsed = True: sTr(1).State = "NM"
    sTr(2).Name = "Beta":  sTr(2).VoltParsed = False: sTr(2).State = "OK"   ' no cost match
    sTr(3).Name = "Delta": sTr(3).VoltParsed = False: sTr(3).State = "TX"   ' bare cost match
    Dim cids() As PL_TCostId: ReDim cids(1 To 3)
    cids(1).Name = "Alpha": cids(1).Voltage = 345: cids(1).VoltParsed = True
    cids(2).Name = "Delta": cids(2).VoltParsed = False
    cids(3).Name = "Gamma": cids(3).VoltParsed = False                       ' no site match
    Dim matched As Long, siteNoCost As Long, cc As Long
    matched = 0: siteNoCost = 0
    Dim si As Long, mi As Long
    For si = 1 To 3
        mi = pl_MatchCost(sTr(si), cids, 3)
        If mi > 0 Then matched = matched + 1: cids(mi).Matched = True Else siteNoCost = siteNoCost + 1
    Next si
    Dim costNoSite As Long: costNoSite = 0
    For cc = 1 To 3
        If Not cids(cc).Matched Then costNoSite = costNoSite + 1
    Next cc
    Assert matched = 2, "Join: Alpha(345) and Delta(bare) match -> 2", passCount, failCount
    Assert siteNoCost = 1, "Join: Beta logged as site-without-cost", passCount, failCount
    Assert costNoSite = 1, "Join: Gamma logged as cost-without-site", passCount, failCount

    ' 4) Matrix shape
    Dim mw() As Double, mc As Long
    mc = pl_MwAxis(mw)
    Assert (mc = 21) And (mw(1) = 100) And (mw(mc) = 300), _
           "Matrix axis: 100..300 step 10 = 21 points", passCount, failCount
    Dim ascOK As Boolean: ascOK = True
    Dim k As Long
    For k = 2 To mc
        If mw(k) <= mw(k - 1) Then ascOK = False
    Next k
    Assert ascOK, "Matrix axis: strictly ascending", passCount, failCount
    Dim f As String
    f = pl_SumifsFormula("Cost Data", 8, 1, 5, 2, "Cunningham", False, 0, "C$1")
    Assert (InStr(f, "SUMIFS(") > 0) And (InStr(f, "'Cost Data'!") > 0) _
           And (InStr(f, "Cunningham") > 0) And (InStr(f, "<=") > 0) _
           And (Right$(f, 4) = "/C$1"), _
           "Matrix body: SUMIFS resolves cols and divides by the MW cell", passCount, failCount
    Dim fv As String
    fv = pl_SumifsFormula("Cost Data", 8, 1, 5, 2, "Chaves County", True, 345, "C$1")
    Assert InStr(fv, "$B:$B,345") > 0, _
           "Matrix body: multi-voltage tab adds the voltage criterion", passCount, failCount
    Assert pl_IdentityLabel("Chaves County", 345, True, "NM") = "Chaves County 345 kV (NM)", _
           "Matrix identity: name + voltage + state label", passCount, failCount

    ' 5) Scoring ceiling + $25MM band all-zero on the fixture
    Assert CLng(pl_ScoreCeiling(4)) = 165, "Scoring: composite maximum is 165", passCount, failCount
    Assert CLng(pl_ScoreCeiling(3)) = 125, "Scoring: practical maximum (3 active bands) is 125", passCount, failCount

    Dim fmw() As Double, fnames() As String, fcost() As Double, fn As Long, fm As Long
    LoadFixture fmw, fnames, fcost, fn, fm
    Dim rkCount() As Long, rkMaxMW() As Double, rkCostMax() As Double
    Dim rkSlope() As Double, rkHasSlope() As Boolean, rA() As Long, rB() As Long
    ComputeRankings fmw, fnames, fcost, fn, fm, rkCount, rkMaxMW, rkCostMax, rkSlope, rkHasSlope, rA, rB
    Dim allZero As Boolean: allZero = True
    Dim i As Long
    For i = 1 To fn
        If pl_RankToPctile(rA(i, 1), 0) <> 0 Then allZero = False
        If pl_RankToPctile(rB(i, 1), 0) <> 0 Then allZero = False
    Next i
    Assert (fn = 17) And allZero, _
           "Scoring: all 17 $25MM band percentiles are 0 on the fixture", passCount, failCount

    ' 6) Headroom -- standalone lens, not part of the composite
    Dim trg() As Double: ReDim trg(1 To 3)
    trg(1) = 200: trg(2) = 130: trg(3) = 260
    Assert pl_MinHeadroom(trg, 3) = 130, _
           "Headroom: triggers 130/200/260 report 130 (minimum)", passCount, failCount
    Dim hv() As Double: ReDim hv(1 To 5)
    hv(1) = 50: hv(2) = 130: hv(3) = 200: hv(4) = 260: hv(5) = 300
    Assert pl_PctExcCeil5(hv, 5, 300) = 5, "Headroom pctile: largest headroom -> 5", passCount, failCount
    Assert pl_PctExcCeil5(hv, 5, 50) = 1, "Headroom pctile: smallest headroom -> 1", passCount, failCount
    ' Headroom is not summed into the composite: the ceiling is unchanged at 165
    ' whether or not the headroom columns are present.
    Assert CLng(pl_ScoreCeiling(4)) = 165, _
           "Headroom not scored: composite maximum stays 165", passCount, failCount
End Sub

' Test helper: set one raw site row.
Private Sub pl_FillSite(ByRef rn() As String, ByRef rv() As Double, ByRef rp() As Boolean, _
                        ByRef rst() As String, ByRef rf() As String, ByRef rr() As Long, _
                        ByVal idx As Long, ByVal nm As String, ByVal v As Double, _
                        ByVal p As Boolean, ByVal st As String)
    rn(idx) = nm: rv(idx) = v: rp(idx) = p: rst(idx) = st
    rf(idx) = "fixture.xlsx": rr(idx) = idx + 1
End Sub

' Test helper: does dt() contain this triple?
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
