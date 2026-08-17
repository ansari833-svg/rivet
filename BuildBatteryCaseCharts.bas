Attribute VB_Name = "BuildBatteryCaseCharts"
Option Explicit

' ============================================================
'  BuildBatteryCaseCharts
' ------------------------------------------------------------
'  Interactively collects a date row and an arbitrary number
'  of "battery case" result blocks, then builds a set of
'  comparison line charts (one chart per charted metric, one
'  series per collected case) on a dedicated worksheet.
'
'  Entry points:
'    BuildBatteryCaseCharts  - the interactive macro
'    CreateTestData          - writes a dummy sheet for testing
'
'  Assumptions (flagged in the task):
'    * The date row and each case block are aligned
'      POSITIONALLY. The k-th data column of a case block maps
'      to column k of the date row. No header-value matching.
'    * The metric label column (first column of each block) is
'      used ONLY for the validation warning. Row order alone
'      determines which metric a row is.
'    * "Energy Purchase Price ($)" in the chart list is taken
'      to mean the "Avg Energy Purchase Price ($)" row.
' ============================================================

' ── Fixed metric row order ──────────────────────────────────
' Rows within every case block always appear in this order.
' Edit here in one place if the model layout changes.
' NOTE: this is a 1-based lookup; index 0 is left unused so the
' array subscript equals the 1-based row number in the block.
Private Const N_METRICS As Long = 8

' ── Layout constants for a case block ───────────────────────
Private Const BLOCK_ROWS As Long = N_METRICS   ' each block is 8 rows tall

' Chart size / spacing (points)
Private Const CHART_WIDTH   As Double = 800
Private Const CHART_HEIGHT  As Double = 400
Private Const CHART_LEFT    As Double = 20
Private Const CHART_TOP0    As Double = 20
Private Const CHART_VGAP    As Double = 30

Private Const OUTPUT_SHEET As String = "Case Charts"

' ── User-defined type: one collected case ───────────────────
' Store the parent worksheet explicitly so cases are NOT assumed
' to live on the same sheet, and every range stays fully qualified.
Private Type CaseInfo
    Name  As String        ' series name entered by the user
    Block As Range         ' the full 8 x (nCols+1) block, incl. label col
    Ws    As Worksheet     ' the worksheet the block lives on
End Type

' ── Module-level state shared between the small subs ────────
Private mDateRange As Range     ' the captured date row
Private mNCols     As Long      ' number of date columns (the contract width)
Private mDateFmt   As String    ' number format detected for the axis

' ============================================================
'  Metric names & the set of charted metrics
' ============================================================

' Returns the fixed metric order as a 1-based array (1..N_METRICS).
Private Function MetricNames() As Variant
    Dim a(1 To N_METRICS) As String
    a(1) = "Total Battery Gross Margin ($)"
    a(2) = "Value Added per Installed kW ($)"
    a(3) = "Value Added per Installed kWh ($)"
    a(4) = "Cycles"
    a(5) = "Net Carbon Abatement (lbs CO2)"
    a(6) = "Avg Energy Purchase Price ($)"
    a(7) = "Avg Value of Renewable Charging Energy ($)"
    a(8) = "Avg Battery Energy Sale Price ($)"
    MetricNames = a
End Function

' The metrics that actually get charted, driven off the row
' indices only. Add/remove entries here to change the chart set
' without touching any plotting logic below.
'   3 -> Value Added per Installed kWh ($)
'   6 -> Avg Energy Purchase Price ($)
'   7 -> Avg Value of Renewable Charging Energy ($)
'   8 -> Avg Battery Energy Sale Price ($)
Private Function ChartedMetricRows() As Variant
    ChartedMetricRows = Array(3, 6, 7, 8)
End Function

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Sub BuildBatteryCaseCharts()

    Dim cases As Collection

    On Error GoTo Cleanup

    ' Reset module-level state for a clean run.
    Set mDateRange = Nothing
    mNCols = 0
    mDateFmt = vbNullString

    ' ── Step 1: capture the date row ─────────────────────────
    If Not CaptureDateRow() Then
        ' User cancelled during date capture; exit cleanly, no side effects.
        GoTo Cleanup
    End If

    ' ── Step 2: collect an arbitrary number of cases ─────────
    Set cases = CollectCases()
    If cases Is Nothing Then GoTo Cleanup      ' cancelled
    If cases.Count = 0 Then
        MsgBox "No cases were collected, so no charts were built.", _
               vbInformation, "Build Battery Case Charts"
        GoTo Cleanup
    End If

    ' ── Step 3: build the charts ─────────────────────────────
    Application.ScreenUpdating = False
    BuildCharts cases

    Application.ScreenUpdating = True
    MsgBox "Built charts for " & cases.Count & " case(s) on sheet '" & _
           OUTPUT_SHEET & "'.", vbInformation, "Build Battery Case Charts"

Cleanup:
    ' Restore application state on every exit path.
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then
        MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, _
               vbExclamation, "Build Battery Case Charts"
    End If
End Sub

' ============================================================
'  STEP 1 — Capture the date row
' ------------------------------------------------------------
'  Returns True if a valid date row was captured, False if the
'  user cancelled. On success sets mDateRange, mNCols, mDateFmt.
' ============================================================
Private Function CaptureDateRow() As Boolean
    Dim rng As Range

    Do
        On Error Resume Next
        Set rng = Nothing
        ' Type:=8 => range selection. RefEdit lets the user click cells.
        Set rng = Application.InputBox( _
            Prompt:="Select the ROW of dates (one row only).", _
            Title:="Step 1 of 3 — Date row", _
            Type:=8)
        On Error GoTo 0

        ' Cancel returns Nothing (or an error swallowed above).
        If rng Is Nothing Then
            CaptureDateRow = False
            Exit Function
        End If

        ' Validate: exactly one row.
        If rng.Rows.Count <> 1 Then
            MsgBox "The date selection must be exactly ONE row." & vbCrLf & _
                   "You selected " & rng.Rows.Count & " rows. Please try again.", _
                   vbExclamation, "Step 1 — Date row"
        Else
            Exit Do
        End If
    Loop

    Set mDateRange = rng
    mNCols = rng.Columns.Count        ' the contract width every case must match

    ' Detect time resolution from the first two date cells.
    mDateFmt = DetectDateFormat(rng)

    CaptureDateRow = True
End Function

' Determine the axis number format by comparing the first two
' date cells. Delta < 1 day => sub-daily; show time in the label.
Private Function DetectDateFormat(ByVal dateRange As Range) As String
    Dim v1 As Variant, v2 As Variant
    Dim d1 As Double, d2 As Double

    DetectDateFormat = "mm/dd/yyyy"      ' default: daily or coarser

    If dateRange.Columns.Count < 2 Then Exit Function

    v1 = dateRange.Cells(1, 1).Value
    v2 = dateRange.Cells(1, 2).Value

    ' Only compare when both cells look like dates or numbers. Note a real
    ' Date value is NOT IsNumeric in VBA, so test IsDate too; CDbl coerces
    ' both a Date and a numeric serial to the underlying day-based serial.
    If (IsNumeric(v1) Or IsDate(v1)) And (IsNumeric(v2) Or IsDate(v2)) Then
        d1 = CDbl(v1)
        d2 = CDbl(v2)
        ' Serial-date deltas are in days; < 1 means sub-daily resolution.
        If Abs(d2 - d1) < 1 Then
            DetectDateFormat = "mm/dd hh:mm"
        End If
    End If
End Function

' ============================================================
'  STEP 2 — Collect cases in a loop
' ------------------------------------------------------------
'  Returns a Collection of CaseInfo. Returns Nothing if the user
'  cancelled outright before choosing to stop.
' ============================================================
Private Function CollectCases() As Collection
    Dim result As Collection
    Dim blk As Range
    Dim caseName As String
    Dim ci As CaseInfo
    Dim ans As VbMsgBoxResult

    Set result = New Collection

    Do
        ' 1) Prompt for the case block (label col + nCols data cols).
        Set blk = PromptForCaseBlock()
        If blk Is Nothing Then
            ' User cancelled the block prompt.
            If result.Count = 0 Then
                ' Cancelled before collecting anything -> abort the macro.
                Set CollectCases = Nothing
                Exit Function
            Else
                ' Already have cases; treat cancel as "stop collecting".
                Exit Do
            End If
        End If

        ' 3) Validate labels in the first column (warn, allow override).
        If Not ValidateLabels(blk) Then
            ' User chose to re-select rather than proceed.
            GoTo ContinueLoop
        End If

        ' 4) Prompt for the case name cell.
        If Not PromptForCaseName(result, caseName) Then
            ' Cancelled the name prompt -> re-select this case.
            GoTo ContinueLoop
        End If

        ' Store the collected case.
        ci.Name = caseName
        Set ci.Block = blk
        Set ci.Ws = blk.Worksheet
        AddCase result, ci

        ' 5) Ask whether to add another case.
        ans = MsgBox("Add another case?", vbYesNo + vbQuestion, _
                     "Step 2 — Cases (" & result.Count & " collected)")
        If ans = vbNo Then Exit Do

ContinueLoop:
    Loop

    Set CollectCases = result
End Function

' Prompt for a single case block and validate its shape.
' Returns the range, or Nothing if the user cancelled.
Private Function PromptForCaseBlock() As Range
    Dim rng As Range
    Dim wantCols As Long
    wantCols = mNCols + 1              ' 1 label column + nCols data columns

    Do
        On Error Resume Next
        Set rng = Nothing
        Set rng = Application.InputBox( _
            Prompt:="Select the case block: " & BLOCK_ROWS & " rows tall, " & _
                    wantCols & " columns wide." & vbCrLf & _
                    "(1 metric-label column followed by " & mNCols & _
                    " data columns.)", _
            Title:="Step 2 — Case block", _
            Type:=8)
        On Error GoTo 0

        If rng Is Nothing Then
            Set PromptForCaseBlock = Nothing
            Exit Function
        End If

        ' Shape validation. The block must be exactly BLOCK_ROWS tall and
        ' (nCols + 1) wide: one label column plus one data column per date
        ' column, which is why the expected width is one greater than the
        ' date row's width.
        If rng.Rows.Count <> BLOCK_ROWS Or rng.Columns.Count <> wantCols Then
            MsgBox "Wrong block dimensions." & vbCrLf & vbCrLf & _
                   "Expected: " & BLOCK_ROWS & " rows x " & wantCols & _
                   " columns" & vbCrLf & _
                   "  (one label column plus " & mNCols & " data columns)." & _
                   vbCrLf & _
                   "Received: " & rng.Rows.Count & " rows x " & _
                   rng.Columns.Count & " columns." & vbCrLf & vbCrLf & _
                   "Please re-select.", _
                   vbExclamation, "Step 2 — Case block"
        Else
            Set PromptForCaseBlock = rng
            Exit Function
        End If
    Loop
End Function

' ============================================================
'  Block validation — metric labels
' ------------------------------------------------------------
'  Compare the first column of the block against the expected
'  metric order (trimmed, case-insensitive). Warn per mismatch
'  and let the user proceed or re-select. If the first column is
'  entirely empty, skip the check and trust row order.
'  Returns True to proceed with this block, False to re-select.
' ============================================================
Private Function ValidateLabels(ByVal blk As Range) As Boolean
    Dim names As Variant
    Dim r As Long
    Dim found As String, expected As String
    Dim anyLabel As Boolean
    Dim ans As VbMsgBoxResult

    names = MetricNames()

    ' First pass: is the label column entirely empty?
    anyLabel = False
    For r = 1 To BLOCK_ROWS
        If Len(Trim$(CStr(blk.Cells(r, 1).Value))) > 0 Then
            anyLabel = True
            Exit For
        End If
    Next r

    If Not anyLabel Then
        ' No labels at all -> trust row order, skip the check.
        ValidateLabels = True
        Exit Function
    End If

    ' Second pass: compare each present label to the expected one.
    For r = 1 To BLOCK_ROWS
        found = Trim$(CStr(blk.Cells(r, 1).Value))
        expected = names(r)
        ' Compare trimmed & case-insensitive; skip blank cells in a
        ' partially-labelled column.
        If Len(found) > 0 Then
            If StrComp(found, expected, vbTextCompare) <> 0 Then
                ans = MsgBox( _
                    "Label mismatch in block row " & r & ":" & vbCrLf & _
                    "  Expected: " & expected & vbCrLf & _
                    "  Found:    " & found & vbCrLf & vbCrLf & _
                    "Proceed anyway (Yes) or re-select the block (No)?", _
                    vbYesNo + vbExclamation, "Step 2 — Label check")
                If ans = vbNo Then
                    ValidateLabels = False
                    Exit Function
                End If
                ' Yes: stop nagging for the rest of this block.
                ValidateLabels = True
                Exit Function
            End If
        End If
    Next r

    ValidateLabels = True
End Function

' ============================================================
'  Case-name prompt
' ------------------------------------------------------------
'  Prompt for the single cell holding the case name. Validate a
'  single cell, non-empty, and not a duplicate of an existing
'  case. Returns True with caseName set, or False if cancelled.
' ============================================================
Private Function PromptForCaseName(ByVal existing As Collection, _
                                   ByRef caseName As String) As Boolean
    Dim rng As Range
    Dim nm As String

    Do
        On Error Resume Next
        Set rng = Nothing
        Set rng = Application.InputBox( _
            Prompt:="Select the single cell holding this case's NAME.", _
            Title:="Step 2 — Case name", _
            Type:=8)
        On Error GoTo 0

        If rng Is Nothing Then
            PromptForCaseName = False
            Exit Function
        End If

        ' Validate exactly one cell.
        If rng.Cells.Count <> 1 Then
            MsgBox "Please select exactly ONE cell for the case name." & _
                   vbCrLf & "You selected " & rng.Cells.Count & " cells.", _
                   vbExclamation, "Step 2 — Case name"
        Else
            ' Coerce to text so a numeric/date-valued cell still yields a
            ' usable series name.
            nm = Trim$(CStr(rng.Cells(1, 1).Value))
            If Len(nm) = 0 Then
                MsgBox "The name cell is empty. Please select a cell with a value.", _
                       vbExclamation, "Step 2 — Case name"
            ElseIf CaseNameExists(existing, nm) Then
                MsgBox "A case named '" & nm & "' has already been collected." & _
                       vbCrLf & "Please choose a different name cell.", _
                       vbExclamation, "Step 2 — Case name"
            Else
                caseName = nm
                PromptForCaseName = True
                Exit Function
            End If
        End If
    Loop
End Function

' True if a case with the given name (case-insensitive) is already collected.
Private Function CaseNameExists(ByVal existing As Collection, _
                                ByVal nm As String) As Boolean
    Dim i As Long
    Dim item As Variant
    For i = 1 To existing.Count
        item = existing(i)
        ' item(1) is the stored case name (see AddCase).
        If StrComp(CStr(item(1)), nm, vbTextCompare) = 0 Then
            CaseNameExists = True
            Exit Function
        End If
    Next i
End Function

' Add a CaseInfo to the collection. A UDT cannot be stored directly in a
' Collection, so wrap it in a Variant array of [name, block, worksheet].
Private Sub AddCase(ByVal col As Collection, ByRef ci As CaseInfo)
    Dim item(1 To 3) As Variant
    item(1) = ci.Name
    Set item(2) = ci.Block
    Set item(3) = ci.Ws
    col.Add item
End Sub

' Read a stored case back out of the collection into a CaseInfo.
Private Function GetCase(ByVal col As Collection, ByVal idx As Long) As CaseInfo
    Dim item As Variant
    Dim ci As CaseInfo
    item = col(idx)
    ci.Name = item(1)
    Set ci.Block = item(2)
    Set ci.Ws = item(3)
    GetCase = ci
End Function

' ============================================================
'  STEP 3 — Build the charts
' ============================================================
Private Sub BuildCharts(ByVal cases As Collection)
    Dim ws As Worksheet
    Dim rows As Variant
    Dim names As Variant
    Dim i As Long
    Dim topPos As Double

    names = MetricNames()
    rows = ChartedMetricRows()

    Set ws = FreshOutputSheet()

    topPos = CHART_TOP0
    For i = LBound(rows) To UBound(rows)
        BuildOneChart ws, cases, CLng(rows(i)), CStr(names(rows(i))), topPos
        topPos = topPos + CHART_HEIGHT + CHART_VGAP
    Next i
End Sub

' Delete any existing output sheet (alerts suppressed) and add a fresh one.
Private Function FreshOutputSheet() As Worksheet
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(OUTPUT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = OUTPUT_SHEET
    Set FreshOutputSheet = ws
End Function

' Build a single metric chart: one line series per collected case.
Private Sub BuildOneChart(ByVal ws As Worksheet, ByVal cases As Collection, _
                          ByVal metricRow As Long, ByVal metricName As String, _
                          ByVal topPos As Double)
    Dim chObj As ChartObject
    Dim ch As Chart
    Dim ser As Series
    Dim ci As CaseInfo
    Dim dataRng As Range
    Dim i As Long

    Set chObj = ws.ChartObjects.Add(CHART_LEFT, topPos, CHART_WIDTH, CHART_HEIGHT)
    Set ch = chObj.Chart

    ch.ChartType = xlLine

    ' One series per case.
    For i = 1 To cases.Count
        ci = GetCase(cases, i)
        Set dataRng = MetricDataRange(ci, metricRow)

        Set ser = ch.SeriesCollection.NewSeries
        ' Keep charts linked to the live source data by handing the Range
        ' object itself (not a copied array) to XValues / Values.
        ser.XValues = mDateRange
        ser.Values = dataRng
        ser.Name = ci.Name
        ser.MarkerStyle = xlMarkerStyleNone      ' line only, no markers
    Next i

    ' ── Category axis fix ────────────────────────────────────
    ' Excel auto-detects a "date" axis for date-like X values and re-bins
    ' the points onto a uniform DAILY interval, which destroys sub-daily
    ' resolution. Forcing xlCategoryScale makes it plot exactly one point
    ' per column of the date row, in order, regardless of the timestamps.
    With ch.Axes(xlCategory)
        .CategoryType = xlCategoryScale
        .TickLabels.NumberFormat = mDateFmt
        ' Thin the tick labels so a wide axis stays readable.
        .TickLabelSpacing = TickSpacing(mNCols)
        .HasTitle = False
    End With

    ' Titles, legend, blank handling.
    ch.HasTitle = True
    ch.ChartTitle.Text = metricName

    With ch.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = UnitPart(metricName)
    End With

    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom

    ' Do not draw gaps down to zero.
    ch.DisplayBlanksAs = xlNotPlotted
End Sub

' The data-only portion of a metric row: offset one column to the right of
' the block (skipping the label cell) and resized to exactly nCols. This is
' the live Range that aligns column-for-column with the date row.
Private Function MetricDataRange(ByRef ci As CaseInfo, _
                                 ByVal metricRow As Long) As Range
    ' Top-left cell of the block, move down to the metric row and right by
    ' one column to skip the label, then take nCols cells wide.
    Set MetricDataRange = ci.Block.Cells(metricRow, 2).Resize(1, mNCols)
End Function

' Choose a tick-label spacing so a wide axis does not overcrowd. Aim for
' roughly 20 visible labels.
Private Function TickSpacing(ByVal nCols As Long) As Long
    Dim s As Long
    s = nCols \ 20
    If s < 1 Then s = 1
    TickSpacing = s
End Function

' Extract the unit portion of a metric name, i.e. the text inside the last
' pair of parentheses (e.g. "$", "lbs CO2"). Falls back to the whole name.
Private Function UnitPart(ByVal metricName As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStrRev(metricName, "(")
    p2 = InStrRev(metricName, ")")
    If p1 > 0 And p2 > p1 Then
        UnitPart = Mid$(metricName, p1 + 1, p2 - p1 - 1)
    Else
        UnitPart = metricName
    End If
End Function

' ============================================================
'  CreateTestData
' ------------------------------------------------------------
'  Writes a dummy worksheet with an HOURLY date row and three
'  case blocks in the correct layout so BuildBatteryCaseCharts
'  can be exercised end to end. Each block has its metric labels
'  in the first column and its case name in a single cell above
'  the block, so both the label check and the name-cell
'  selection can be tested.
' ============================================================
Public Sub CreateTestData()
    Const SHEET_NAME As String = "Battery Test Data"
    Const N As Long = 24                 ' number of hourly date columns
    Dim ws As Worksheet
    Dim names As Variant
    Dim caseNames As Variant
    Dim baseDate As Double
    Dim dateRow As Long, firstDataCol As Long
    Dim c As Long, r As Long, k As Long
    Dim blockTopRow As Long
    Dim nameCell As Range

    names = MetricNames()
    caseNames = Array("Base Case", "High Cycling", "Aggressive Arbitrage")

    ' Fresh sheet.
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_NAME).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = SHEET_NAME

    ' ── Date row (row 1). Column 1 is a spacer label column so the
    '    date columns align with each block's data columns. ─────────
    dateRow = 1
    firstDataCol = 2                     ' data starts in column B
    ws.Cells(dateRow, 1).Value = "Timestamp"
    baseDate = CDbl(DateSerial(2026, 1, 1)) + CDbl(TimeSerial(0, 0, 0))
    For c = 0 To N - 1
        ' Hourly steps so the sub-daily category-axis behaviour is exercised.
        ws.Cells(dateRow, firstDataCol + c).Value = baseDate + c * (1# / 24#)
        ws.Cells(dateRow, firstDataCol + c).NumberFormat = "mm/dd hh:mm"
    Next c

    ' ── Three case blocks stacked below the date row. ──────────────
    ' Layout per case:
    '   name-cell row, then the 8-row block starting on the next row.
    blockTopRow = dateRow + 2            ' leave a gap under the date row

    For k = 0 To UBound(caseNames)
        ' Case name in a single cell above the block (column A).
        Set nameCell = ws.Cells(blockTopRow - 1, 1)
        nameCell.Value = caseNames(k)
        nameCell.Font.Bold = True

        For r = 1 To N_METRICS
            ' Metric label in the first column of the block.
            ws.Cells(blockTopRow + r - 1, 1).Value = names(r)
            ' Plausible data values, varied per case index k.
            For c = 0 To N - 1
                ws.Cells(blockTopRow + r - 1, firstDataCol + c).Value = _
                    SampleValue(r, k, c, N)
            Next c
        Next r

        blockTopRow = blockTopRow + N_METRICS + 2   ' gap before next block
    Next k

    ws.Columns(1).AutoFit
    ws.Activate
    ws.Cells(1, 1).Select

    MsgBox "Test data written to '" & SHEET_NAME & "'." & vbCrLf & vbCrLf & _
           "Date row: row " & dateRow & ", columns B:" & _
           ColLetter(firstDataCol + N - 1) & "." & vbCrLf & _
           "Three case blocks with name cells in column A.", _
           vbInformation, "CreateTestData"
End Sub

' Produce a plausible, metric-appropriate value for the test data. Uses a
' gentle sinusoid plus per-case scaling so the three series are distinct.
Private Function SampleValue(ByVal metricRow As Long, ByVal caseIdx As Long, _
                             ByVal col As Long, ByVal n As Long) As Double
    Dim base As Double, amp As Double, phase As Double, scale As Double
    Dim wave As Double

    ' Per-metric baselines and amplitudes.
    Select Case metricRow
        Case 1: base = 500000:  amp = 80000     ' Total Battery Gross Margin ($)
        Case 2: base = 350:     amp = 40        ' Value Added / installed kW ($)
        Case 3: base = 90:      amp = 12        ' Value Added / installed kWh ($)
        Case 4: base = 1.2:     amp = 0.4       ' Cycles
        Case 5: base = 1500:    amp = 300       ' Net Carbon Abatement (lbs CO2)
        Case 6: base = 35:      amp = 15        ' Avg Energy Purchase Price ($)
        Case 7: base = 28:      amp = 10        ' Avg Value of Renewable Charging ($)
        Case 8: base = 65:      amp = 20        ' Avg Battery Energy Sale Price ($)
        Case Else: base = 100:  amp = 10
    End Select

    ' Case-specific scaling and phase shift so series don't overlap.
    scale = 1# + 0.15 * caseIdx
    phase = caseIdx * 0.7

    wave = Sin(2 * 3.14159265358979 * col / n + phase)
    SampleValue = (base + amp * wave) * scale
End Function

' Convert a 1-based column number to its letter(s), for user messages.
Private Function ColLetter(ByVal colNum As Long) As String
    Dim s As String, n As Long
    n = colNum
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function
