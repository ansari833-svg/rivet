Attribute VB_Name = "modNPVBridge"
Option Explicit

' ============================================================
'  modNPVBridge.bas
'  NPV Bridge (Waterfall) Chart Builder
'
'  Public entry point:  BuildNPVBridge
'
'  Interactively prompts the user for every component of an NPV
'  bridge -- a base case, an arbitrary number of positive/negative
'  bridge steps, and a revised (final) base case -- then lays the
'  data out on a dedicated worksheet and draws a waterfall chart.
'
'  --- Why a stacked column, not the native Waterfall -----------
'  Excel's built-in Waterfall chart type does not expose per-point
'  fill control in a portable way. Instead this macro fakes the
'  waterfall with a STACKED COLUMN chart of two series:
'
'      Spacer  (bottom of the stack, made completely invisible)
'      Value   (top of the stack, the visible floating bar)
'
'  A running cumulative total decides how tall each Spacer must be so
'  that its Value bar floats to the correct level. Because Value is a
'  normal column series, each point's colour can be set independently
'  via .Points(i).Format.Fill.ForeColor.RGB.
'
'  --- Waterfall math (cumulative = running total) --------------
'      Base Case            Spacer = 0                   Value = base
'      Increase (d >= 0)    Spacer = cum_before          Value = d
'      Decrease (d <  0)    Spacer = cum_before + d      Value = |d|
'      Revised NPV          Spacer = 0                   Value = final total
'
'  The Revised NPV value is computed automatically as
'  base + sum(steps) so the bridge always foots.
' ============================================================

' ---------- Tunable constants (edit here only) --------------

Private Const DATA_SHEET   As String = "NPV Bridge Data"
Private Const CHART_TITLE  As String = "NPV Bridge"

Private Const HDR_ROW      As Long = 1          ' header row of the helper table
Private Const FIRST_DATA   As Long = 2          ' first data row

' ---------- Small container for one bridge component ---------

Private Type Component
    Label   As String
    Spacer  As Double
    Value   As Double
    HexCode As String        ' 6 hex digits, no leading "#"
End Type

' ============================================================
'  PUBLIC ENTRY POINT
' ============================================================
'  wbTarget: the workbook that receives the data table and chart. When
'  called with no argument (e.g. run from the VBA editor) it defaults to
'  the ActiveWorkbook; the runner-workbook launcher passes the dropped
'  workbook explicitly so output lands there, not in the runner.
Public Sub BuildNPVBridge(Optional ByVal wbTarget As Workbook = Nothing)

    Dim comps() As Component
    Dim nSteps As Long
    Dim baseVal As Double, delta As Double
    Dim cumBefore As Double, cumAfter As Double
    Dim finalTotal As Double
    Dim firstStepHex As String
    Dim i As Long, idx As Long
    Dim s As String

    ' Default the target to the active workbook when none was passed.
    If wbTarget Is Nothing Then Set wbTarget = ActiveWorkbook

    ' ---- 1. How many bridge steps? (integer, minimum 1) ----
    Do
        s = InputBox("How many bridge steps are there? (whole number, at least 1)", _
                     "NPV Bridge")
        If StrPtr(s) = 0 Then Exit Sub          ' Cancel pressed
        s = Trim$(s)
        If IsNumeric(s) Then
            If s = CStr(CLng(Val(s))) And CLng(Val(s)) >= 1 Then
                nSteps = CLng(Val(s))
                Exit Do
            End If
        End If
        MsgBox "Please enter a whole number greater than or equal to 1.", _
               vbExclamation, "NPV Bridge"
    Loop

    ' comps(0) = base, comps(1..nSteps) = steps, comps(nSteps+1) = revised
    ReDim comps(0 To nSteps + 1)

    ' ---- 2. BASE CASE ----
    If Not PromptLabel(comps(0).Label, "Component: [Base Case]", _
                       "Enter the label for the base case:") Then Exit Sub
    If Not PromptValue(baseVal, "Component: [Base Case]", _
                       "Enter the base case value:") Then Exit Sub
    If Not PromptHex(comps(0).HexCode, "Component: [Base Case]", _
                     "Enter the bar colour as a 6-digit hex code (e.g. 1F4E79):", _
                     vbNullString) Then Exit Sub

    ' Base case bar: floats from zero.
    comps(0).Spacer = 0
    comps(0).Value = baseVal

    cumAfter = baseVal
    firstStepHex = vbNullString

    ' ---- 3. BRIDGE STEPS ----
    For i = 1 To nSteps
        idx = i
        Dim ttl As String
        ttl = "Component: [Bridge Value] " & Chr$(151) & " Step " & i & " of " & nSteps

        If Not PromptLabel(comps(idx).Label, ttl, _
                           "Enter the label for step " & i & ":") Then Exit Sub
        If Not PromptValue(delta, ttl, _
                           "Enter the value for step " & i & _
                           " (positive = increase, negative = decrease):") Then Exit Sub

        ' For step 1 there is no default; steps 2..N pre-fill step 1's hex.
        Dim defHex As String
        If i = 1 Then
            defHex = vbNullString
        Else
            defHex = firstStepHex
        End If
        If Not PromptHex(comps(idx).HexCode, ttl, _
                         "Enter the bar colour as a 6-digit hex code (e.g. 1F4E79):", _
                         defHex) Then Exit Sub
        If i = 1 Then firstStepHex = comps(idx).HexCode

        ' Waterfall math for this step.
        cumBefore = cumAfter
        If delta >= 0 Then
            comps(idx).Spacer = cumBefore
            comps(idx).Value = delta
        Else
            comps(idx).Spacer = cumBefore + delta      ' delta is negative
            comps(idx).Value = -delta                  ' |delta|
        End If
        cumAfter = cumBefore + delta
    Next i

    finalTotal = cumAfter                              ' base + sum(steps)

    ' ---- 4. REVISED NPV (value computed automatically) ----
    idx = nSteps + 1
    If Not PromptLabel(comps(idx).Label, "Component: [Revised Base Case]", _
                       "Enter the label for the revised base case:") Then Exit Sub
    If Not PromptHex(comps(idx).HexCode, "Component: [Revised Base Case]", _
                     "Enter the bar colour as a 6-digit hex code (e.g. 1F4E79):", _
                     vbNullString) Then Exit Sub
    comps(idx).Spacer = 0
    comps(idx).Value = finalTotal

    ' ---- 5. Build the sheet and chart ----
    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    WriteTableAndChart wbTarget, comps

    Application.ScreenUpdating = True
    MsgBox "NPV Bridge built successfully.", vbInformation, "NPV Bridge"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Could not build the NPV Bridge:" & vbCrLf & Err.Description, _
           vbExclamation, "NPV Bridge"
End Sub

' ============================================================
'  LAYOUT + CHART
' ============================================================
Private Sub WriteTableAndChart(ByVal wb As Workbook, comps() As Component)

    Dim ws As Worksheet
    Dim nRows As Long, r As Long, i As Long
    Dim lastRow As Long

    Set ws = GetCleanSheet(wb, DATA_SHEET)

    nRows = UBound(comps) - LBound(comps) + 1     ' base + steps + revised

    ' ---- Header ----
    ws.Cells(HDR_ROW, 1).Value = "Label"
    ws.Cells(HDR_ROW, 2).Value = "Spacer"
    ws.Cells(HDR_ROW, 3).Value = "Value"
    ws.Cells(HDR_ROW, 4).Value = "HexColor"
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW, 4)).Font.Bold = True

    ' ---- Data rows ----
    r = FIRST_DATA
    For i = LBound(comps) To UBound(comps)
        ws.Cells(r, 1).Value = comps(i).Label
        ws.Cells(r, 2).Value = comps(i).Spacer
        ws.Cells(r, 3).Value = comps(i).Value
        ws.Cells(r, 4).Value = "'" & comps(i).HexCode     ' keep leading zeros as text
        r = r + 1
    Next i
    lastRow = FIRST_DATA + nRows - 1
    ws.Columns("A:D").AutoFit

    ' ---- Build the stacked column chart ----
    Dim co As ChartObject
    Dim ch As Chart

    ' Remove any charts already on the sheet (idempotent rebuild).
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop

    Set co = ws.ChartObjects.Add(Left:=ws.Columns("F").Left, Top:=ws.Rows(1).Top, _
                                 Width:=520, Height:=320)
    Set ch = co.Chart
    ch.ChartType = xlColumnStacked

    ' Categories = labels (column A); explicit series so order is guaranteed.
    Dim catRng As Range, spcRng As Range, valRng As Range
    Set catRng = ws.Range(ws.Cells(FIRST_DATA, 1), ws.Cells(lastRow, 1))
    Set spcRng = ws.Range(ws.Cells(FIRST_DATA, 2), ws.Cells(lastRow, 2))
    Set valRng = ws.Range(ws.Cells(FIRST_DATA, 3), ws.Cells(lastRow, 3))

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    Dim serSpacer As Series, serValue As Series

    Set serSpacer = ch.SeriesCollection.NewSeries
    serSpacer.Name = "Spacer"
    serSpacer.Values = spcRng
    serSpacer.XValues = catRng

    Set serValue = ch.SeriesCollection.NewSeries
    serValue.Name = "Value"
    serValue.Values = valRng
    serValue.XValues = catRng

    ' ---- Make the Spacer series completely invisible ----
    serSpacer.Format.Fill.Visible = msoFalse
    serSpacer.Format.Line.Visible = msoFalse

    ' ---- Colour each Value point from its component's hex ----
    For i = 1 To nRows
        serValue.Points(i).Format.Fill.Visible = msoTrue
        serValue.Points(i).Format.Fill.ForeColor.RGB = HexToColor(comps(LBound(comps) + i - 1).HexCode)
        serValue.Points(i).Format.Line.Visible = msoFalse
    Next i

    ' ---- Data labels on the Value series ----
    serValue.HasDataLabels = True
    serValue.DataLabels.ShowValue = True
    serValue.DataLabels.NumberFormat = "#,##0"

    ' ---- Title, legend ----
    ch.HasTitle = True
    ch.ChartTitle.Text = CHART_TITLE
    ch.HasLegend = False        ' the only "real" series is Value; hide the legend entirely

    ' ---- Optional: thin connector lines between consecutive bars ----
    On Error Resume Next        ' connectors are cosmetic; never fail the build over them
    AddConnectors ch, comps
    On Error GoTo 0

    ws.Cells(1, 1).Select
End Sub

' ------------------------------------------------------------
'  Draw thin horizontal connector lines linking the top of each
'  bar to the base of the next, so the eye follows the bridge.
'  Cosmetic only -- wrapped in On Error Resume Next by the caller.
' ------------------------------------------------------------
Private Sub AddConnectors(ch As Chart, comps() As Component)

    Dim pa As PlotArea
    Dim i As Long, n As Long
    Dim topLevel As Double
    Dim x1 As Double, x2 As Double, yPix As Double
    Dim vAxis As Axis, cAxis As Axis
    Dim plotL As Double, plotW As Double
    Dim ln As Shape
    Dim topOfBar As Double
    Dim minY As Double, maxY As Double

    n = UBound(comps) - LBound(comps) + 1
    If n < 2 Then Exit Sub

    Set vAxis = ch.Axes(xlValue)
    minY = vAxis.MinimumScale
    maxY = vAxis.MaximumScale
    If maxY <= minY Then Exit Sub

    Set pa = ch.PlotArea
    plotL = pa.InsideLeft
    plotW = pa.InsideWidth

    ' Each category occupies an equal slice of the plot width.
    For i = LBound(comps) To UBound(comps) - 1
        ' Top level = cumulative running total after this component.
        topLevel = comps(i).Spacer + comps(i).Value

        ' Convert the value level to a vertical pixel inside the plot area.
        yPix = pa.InsideTop + pa.InsideHeight * (1 - (topLevel - minY) / (maxY - minY))

        ' Horizontal centres of bar i and bar i+1 (1-based slot index).
        Dim slot As Long
        slot = i - LBound(comps) + 1
        x1 = plotL + plotW * (slot - 0.5) / n
        x2 = plotL + plotW * (slot + 0.5) / n

        Set ln = ch.Shapes.AddLine(x1, yPix, x2, yPix)
        ln.Line.DashStyle = msoLineDash
        ln.Line.Weight = 0.75
        ln.Line.ForeColor.RGB = RGB(128, 128, 128)
    Next i
End Sub

' ============================================================
'  INPUT HELPERS  (each returns False if the user cancels)
' ============================================================

' Free-text label. Returns False on Cancel.
Private Function PromptLabel(ByRef outLabel As String, _
                             ByVal title As String, _
                             ByVal prompt As String) As Boolean
    Dim s As String
    s = InputBox(prompt, title)
    If StrPtr(s) = 0 Then
        PromptLabel = False
    Else
        outLabel = s
        PromptLabel = True
    End If
End Function

' Numeric value. Re-prompts on non-numeric. Returns False on Cancel.
Private Function PromptValue(ByRef outVal As Double, _
                             ByVal title As String, _
                             ByVal prompt As String) As Boolean
    Dim s As String
    Do
        s = InputBox(prompt, title)
        If StrPtr(s) = 0 Then
            PromptValue = False
            Exit Function
        End If
        s = Trim$(s)
        If IsNumeric(s) Then
            outVal = CDbl(s)
            PromptValue = True
            Exit Function
        End If
        MsgBox "Please enter a valid number.", vbExclamation, title
    Loop
End Function

' Hex colour. Re-prompts on invalid. Returns False on Cancel.
' defHex pre-fills the InputBox default (pass vbNullString for none).
Private Function PromptHex(ByRef outHex As String, _
                           ByVal title As String, _
                           ByVal prompt As String, _
                           ByVal defHex As String) As Boolean
    Dim s As String, clean As String
    Do
        s = InputBox(prompt, title, defHex)
        If StrPtr(s) = 0 Then
            PromptHex = False
            Exit Function
        End If
        If IsValidHex(s, clean) Then
            outHex = clean
            PromptHex = True
            Exit Function
        End If
        MsgBox "Please enter a valid 6-digit hex code, e.g. 1F4E79 " & _
               "(a leading # is allowed).", vbExclamation, title
    Loop
End Function

' ============================================================
'  HEX / COLOUR HELPERS
' ============================================================

' Validate a hex string. Accepts an optional leading "#" and
' surrounding whitespace. On success, cleanHex holds the 6 upper-case
' hex digits (no "#").
Private Function IsValidHex(ByVal raw As String, ByRef cleanHex As String) As Boolean
    Dim s As String, i As Long, c As String
    s = Trim$(raw)
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    s = Trim$(s)
    If Len(s) <> 6 Then Exit Function
    For i = 1 To 6
        c = UCase$(Mid$(s, i, 1))
        If InStr("0123456789ABCDEF", c) = 0 Then Exit Function
    Next i
    cleanHex = UCase$(s)
    IsValidHex = True
End Function

' Convert a 6-digit hex code (with or without a leading "#") to a VBA
' colour long. VBA's colour longs are stored BGR, so the byte order is
' RGB(red, green, blue) = red + green*256 + blue*65536 -- built here
' explicitly from the RR, GG, BB pairs.
Public Function HexToColor(ByVal hex As String) As Long
    Dim clean As String
    Dim rr As Long, gg As Long, bb As Long

    If Not IsValidHex(hex, clean) Then
        Err.Raise vbObjectError + 513, "HexToColor", _
                  "Invalid hex colour: '" & hex & "'"
    End If

    rr = CLng("&H" & Mid$(clean, 1, 2))
    gg = CLng("&H" & Mid$(clean, 3, 2))
    bb = CLng("&H" & Mid$(clean, 5, 2))

    HexToColor = RGB(rr, gg, bb)        ' RGB() stores the long as BGR internally
End Function

' ============================================================
'  SHEET HELPER
' ============================================================

' Return a worksheet with the given name, cleared. Creates it if it
' does not exist, otherwise deletes its charts and clears its cells.
Private Function GetCleanSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add( _
                    After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = name
    Else
        Dim co As ChartObject
        For Each co In ws.ChartObjects
            co.Delete
        Next co
        ws.Cells.Clear
    End If

    Set GetCleanSheet = ws
End Function
