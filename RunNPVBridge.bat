@echo off
setlocal DisableDelayedExpansion
rem ============================================================
rem  RunNPVBridge.bat  --  drag-and-drop NPV Bridge builder
rem
rem  Drag an Excel workbook (.xlsx / .xlsm) onto this .bat icon.
rem  It writes a temporary .vbs to %TEMP%, which opens the workbook
rem  in a visible Excel instance, injects the modNPVBridge macro into
rem  the workbook via the VBIDE extensibility model, and runs
rem  BuildNPVBridge. The workbook is left OPEN and unsaved so you can
rem  review the chart and save it yourself. The temp .vbs is deleted.
rem
rem  DEPENDENCY (one-time, per machine): Excel Trust Center >
rem  Macro Settings > "Trust access to the VBA project object model"
rem  must be enabled, otherwise the injection step is blocked.
rem ============================================================

if "%~1"=="" (
    echo Usage: drag an Excel workbook ^(.xlsx or .xlsm^) onto this .bat,
    echo        or run:  "%~nx0" "C:\path\to\workbook.xlsx"
    echo(
    pause
    exit /b 1
)

set "WB=%~1"
set "VBS=%TEMP%\npvbridge_%RANDOM%%RANDOM%.vbs"

rem --- Extract the embedded VBS (lines tagged ::@@) to %TEMP% ---
if exist "%VBS%" del "%VBS%"
(for /f "usebackq delims=" %%L in (`findstr /b /c:"::@@" "%~f0"`) do (
    set "ln=%%L"
    setlocal EnableDelayedExpansion
    echo(!ln:~4!
    endlocal
)) > "%VBS%"

rem --- Run it, passing the dropped workbook path as arg 0 ---
cscript //nologo "%VBS%" "%WB%"
set "RC=%ERRORLEVEL%"

rem --- Clean up the temp .vbs ---
del "%VBS%" >nul 2>&1

exit /b %RC%

rem ==== Embedded VBS payload below (extracted at runtime) ======
::@@Option Explicit
::@@Dim xl, wb, vbProj, comp, c, code, wbPath
::@@wbPath = WScript.Arguments(0)
::@@code = ""
::@@code = code & "Option Explicit" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "'  modNPVBridge.bas" & vbCrLf
::@@code = code & "'  NPV Bridge (Waterfall) Chart Builder" & vbCrLf
::@@code = code & "'" & vbCrLf
::@@code = code & "'  Public entry point:  BuildNPVBridge" & vbCrLf
::@@code = code & "'" & vbCrLf
::@@code = code & "'  Interactively prompts the user for every component of an NPV" & vbCrLf
::@@code = code & "'  bridge -- a base case, an arbitrary number of positive/negative" & vbCrLf
::@@code = code & "'  bridge steps, and a revised (final) base case -- then lays the" & vbCrLf
::@@code = code & "'  data out on a dedicated worksheet and draws a waterfall chart." & vbCrLf
::@@code = code & "'" & vbCrLf
::@@code = code & "'  --- Why a stacked column, not the native Waterfall -----------" & vbCrLf
::@@code = code & "'  Excel's built-in Waterfall chart type does not expose per-point" & vbCrLf
::@@code = code & "'  fill control in a portable way. Instead this macro fakes the" & vbCrLf
::@@code = code & "'  waterfall with a STACKED COLUMN chart of two series:" & vbCrLf
::@@code = code & "'" & vbCrLf
::@@code = code & "'      Spacer  (bottom of the stack, made completely invisible)" & vbCrLf
::@@code = code & "'      Value   (top of the stack, the visible floating bar)" & vbCrLf
::@@code = code & "'" & vbCrLf
::@@code = code & "'  A running cumulative total decides how tall each Spacer must be so" & vbCrLf
::@@code = code & "'  that its Value bar floats to the correct level. Because Value is a" & vbCrLf
::@@code = code & "'  normal column series, each point's colour can be set independently" & vbCrLf
::@@code = code & "'  via .Points(i).Format.Fill.ForeColor.RGB." & vbCrLf
::@@code = code & "'" & vbCrLf
::@@code = code & "'  --- Waterfall math (cumulative = running total) --------------" & vbCrLf
::@@code = code & "'      Base Case            Spacer = 0                   Value = base" & vbCrLf
::@@code = code & "'      Increase (d >= 0)    Spacer = cum_before          Value = d" & vbCrLf
::@@code = code & "'      Decrease (d <  0)    Spacer = cum_before + d      Value = |d|" & vbCrLf
::@@code = code & "'      Revised NPV          Spacer = 0                   Value = final total" & vbCrLf
::@@code = code & "'" & vbCrLf
::@@code = code & "'  The Revised NPV value is computed automatically as" & vbCrLf
::@@code = code & "'  base + sum(steps) so the bridge always foots." & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' ---------- Tunable constants (edit here only) --------------" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "Private Const DATA_SHEET   As String = ""NPV Bridge Data""" & vbCrLf
::@@code = code & "Private Const CHART_TITLE  As String = ""NPV Bridge""" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "Private Const HDR_ROW      As Long = 1          ' header row of the helper table" & vbCrLf
::@@code = code & "Private Const FIRST_DATA   As Long = 2          ' first data row" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' ---------- Small container for one bridge component ---------" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "Private Type Component" & vbCrLf
::@@code = code & "    Label   As String" & vbCrLf
::@@code = code & "    Spacer  As Double" & vbCrLf
::@@code = code & "    Value   As Double" & vbCrLf
::@@code = code & "    HexCode As String        ' 6 hex digits, no leading ""#""" & vbCrLf
::@@code = code & "End Type" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "'  PUBLIC ENTRY POINT" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "Public Sub BuildNPVBridge()" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Dim comps() As Component" & vbCrLf
::@@code = code & "    Dim nSteps As Long" & vbCrLf
::@@code = code & "    Dim baseVal As Double, delta As Double" & vbCrLf
::@@code = code & "    Dim cumBefore As Double, cumAfter As Double" & vbCrLf
::@@code = code & "    Dim finalTotal As Double" & vbCrLf
::@@code = code & "    Dim firstStepHex As String" & vbCrLf
::@@code = code & "    Dim i As Long, idx As Long" & vbCrLf
::@@code = code & "    Dim s As String" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- 1. How many bridge steps? (integer, minimum 1) ----" & vbCrLf
::@@code = code & "    Do" & vbCrLf
::@@code = code & "        s = InputBox(""How many bridge steps are there? (whole number, at least 1)"", _" & vbCrLf
::@@code = code & "                     ""NPV Bridge"")" & vbCrLf
::@@code = code & "        If StrPtr(s) = 0 Then Exit Sub          ' Cancel pressed" & vbCrLf
::@@code = code & "        s = Trim$(s)" & vbCrLf
::@@code = code & "        If IsNumeric(s) Then" & vbCrLf
::@@code = code & "            If s = CStr(CLng(Val(s))) And CLng(Val(s)) >= 1 Then" & vbCrLf
::@@code = code & "                nSteps = CLng(Val(s))" & vbCrLf
::@@code = code & "                Exit Do" & vbCrLf
::@@code = code & "            End If" & vbCrLf
::@@code = code & "        End If" & vbCrLf
::@@code = code & "        MsgBox ""Please enter a whole number greater than or equal to 1."", _" & vbCrLf
::@@code = code & "               vbExclamation, ""NPV Bridge""" & vbCrLf
::@@code = code & "    Loop" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' comps(0) = base, comps(1..nSteps) = steps, comps(nSteps+1) = revised" & vbCrLf
::@@code = code & "    ReDim comps(0 To nSteps + 1)" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- 2. BASE CASE ----" & vbCrLf
::@@code = code & "    If Not PromptLabel(comps(0).Label, ""Component: [Base Case]"", _" & vbCrLf
::@@code = code & "                       ""Enter the label for the base case:"") Then Exit Sub" & vbCrLf
::@@code = code & "    If Not PromptValue(baseVal, ""Component: [Base Case]"", _" & vbCrLf
::@@code = code & "                       ""Enter the base case value:"") Then Exit Sub" & vbCrLf
::@@code = code & "    If Not PromptHex(comps(0).HexCode, ""Component: [Base Case]"", _" & vbCrLf
::@@code = code & "                     ""Enter the bar colour as a 6-digit hex code (e.g. 1F4E79):"", _" & vbCrLf
::@@code = code & "                     vbNullString) Then Exit Sub" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' Base case bar: floats from zero." & vbCrLf
::@@code = code & "    comps(0).Spacer = 0" & vbCrLf
::@@code = code & "    comps(0).Value = baseVal" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    cumAfter = baseVal" & vbCrLf
::@@code = code & "    firstStepHex = vbNullString" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- 3. BRIDGE STEPS ----" & vbCrLf
::@@code = code & "    For i = 1 To nSteps" & vbCrLf
::@@code = code & "        idx = i" & vbCrLf
::@@code = code & "        Dim ttl As String" & vbCrLf
::@@code = code & "        ttl = ""Component: [Bridge Value] "" & Chr$(151) & "" Step "" & i & "" of "" & nSteps" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "        If Not PromptLabel(comps(idx).Label, ttl, _" & vbCrLf
::@@code = code & "                           ""Enter the label for step "" & i & "":"") Then Exit Sub" & vbCrLf
::@@code = code & "        If Not PromptValue(delta, ttl, _" & vbCrLf
::@@code = code & "                           ""Enter the value for step "" & i & _" & vbCrLf
::@@code = code & "                           "" (positive = increase, negative = decrease):"") Then Exit Sub" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "        ' For step 1 there is no default; steps 2..N pre-fill step 1's hex." & vbCrLf
::@@code = code & "        Dim defHex As String" & vbCrLf
::@@code = code & "        If i = 1 Then" & vbCrLf
::@@code = code & "            defHex = vbNullString" & vbCrLf
::@@code = code & "        Else" & vbCrLf
::@@code = code & "            defHex = firstStepHex" & vbCrLf
::@@code = code & "        End If" & vbCrLf
::@@code = code & "        If Not PromptHex(comps(idx).HexCode, ttl, _" & vbCrLf
::@@code = code & "                         ""Enter the bar colour as a 6-digit hex code (e.g. 1F4E79):"", _" & vbCrLf
::@@code = code & "                         defHex) Then Exit Sub" & vbCrLf
::@@code = code & "        If i = 1 Then firstStepHex = comps(idx).HexCode" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "        ' Waterfall math for this step." & vbCrLf
::@@code = code & "        cumBefore = cumAfter" & vbCrLf
::@@code = code & "        If delta >= 0 Then" & vbCrLf
::@@code = code & "            comps(idx).Spacer = cumBefore" & vbCrLf
::@@code = code & "            comps(idx).Value = delta" & vbCrLf
::@@code = code & "        Else" & vbCrLf
::@@code = code & "            comps(idx).Spacer = cumBefore + delta      ' delta is negative" & vbCrLf
::@@code = code & "            comps(idx).Value = -delta                  ' |delta|" & vbCrLf
::@@code = code & "        End If" & vbCrLf
::@@code = code & "        cumAfter = cumBefore + delta" & vbCrLf
::@@code = code & "    Next i" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    finalTotal = cumAfter                              ' base + sum(steps)" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- 4. REVISED NPV (value computed automatically) ----" & vbCrLf
::@@code = code & "    idx = nSteps + 1" & vbCrLf
::@@code = code & "    If Not PromptLabel(comps(idx).Label, ""Component: [Revised Base Case]"", _" & vbCrLf
::@@code = code & "                       ""Enter the label for the revised base case:"") Then Exit Sub" & vbCrLf
::@@code = code & "    If Not PromptHex(comps(idx).HexCode, ""Component: [Revised Base Case]"", _" & vbCrLf
::@@code = code & "                     ""Enter the bar colour as a 6-digit hex code (e.g. 1F4E79):"", _" & vbCrLf
::@@code = code & "                     vbNullString) Then Exit Sub" & vbCrLf
::@@code = code & "    comps(idx).Spacer = 0" & vbCrLf
::@@code = code & "    comps(idx).Value = finalTotal" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- 5. Build the sheet and chart ----" & vbCrLf
::@@code = code & "    Application.ScreenUpdating = False" & vbCrLf
::@@code = code & "    On Error GoTo CleanFail" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    WriteTableAndChart comps" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Application.ScreenUpdating = True" & vbCrLf
::@@code = code & "    MsgBox ""NPV Bridge built successfully."", vbInformation, ""NPV Bridge""" & vbCrLf
::@@code = code & "    Exit Sub" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "CleanFail:" & vbCrLf
::@@code = code & "    Application.ScreenUpdating = True" & vbCrLf
::@@code = code & "    MsgBox ""Could not build the NPV Bridge:"" & vbCrLf & Err.Description, _" & vbCrLf
::@@code = code & "           vbExclamation, ""NPV Bridge""" & vbCrLf
::@@code = code & "End Sub" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "'  LAYOUT + CHART" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "Private Sub WriteTableAndChart(comps() As Component)" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Dim ws As Worksheet" & vbCrLf
::@@code = code & "    Dim nRows As Long, r As Long, i As Long" & vbCrLf
::@@code = code & "    Dim lastRow As Long" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Set ws = GetCleanSheet(DATA_SHEET)" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    nRows = UBound(comps) - LBound(comps) + 1     ' base + steps + revised" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- Header ----" & vbCrLf
::@@code = code & "    ws.Cells(HDR_ROW, 1).Value = ""Label""" & vbCrLf
::@@code = code & "    ws.Cells(HDR_ROW, 2).Value = ""Spacer""" & vbCrLf
::@@code = code & "    ws.Cells(HDR_ROW, 3).Value = ""Value""" & vbCrLf
::@@code = code & "    ws.Cells(HDR_ROW, 4).Value = ""HexColor""" & vbCrLf
::@@code = code & "    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW, 4)).Font.Bold = True" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- Data rows ----" & vbCrLf
::@@code = code & "    r = FIRST_DATA" & vbCrLf
::@@code = code & "    For i = LBound(comps) To UBound(comps)" & vbCrLf
::@@code = code & "        ws.Cells(r, 1).Value = comps(i).Label" & vbCrLf
::@@code = code & "        ws.Cells(r, 2).Value = comps(i).Spacer" & vbCrLf
::@@code = code & "        ws.Cells(r, 3).Value = comps(i).Value" & vbCrLf
::@@code = code & "        ws.Cells(r, 4).Value = ""'"" & comps(i).HexCode     ' keep leading zeros as text" & vbCrLf
::@@code = code & "        r = r + 1" & vbCrLf
::@@code = code & "    Next i" & vbCrLf
::@@code = code & "    lastRow = FIRST_DATA + nRows - 1" & vbCrLf
::@@code = code & "    ws.Columns(""A:D"").AutoFit" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- Build the stacked column chart ----" & vbCrLf
::@@code = code & "    Dim co As ChartObject" & vbCrLf
::@@code = code & "    Dim ch As Chart" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' Remove any charts already on the sheet (idempotent rebuild)." & vbCrLf
::@@code = code & "    Do While ws.ChartObjects.Count > 0" & vbCrLf
::@@code = code & "        ws.ChartObjects(1).Delete" & vbCrLf
::@@code = code & "    Loop" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Set co = ws.ChartObjects.Add(Left:=ws.Columns(""F"").Left, Top:=ws.Rows(1).Top, _" & vbCrLf
::@@code = code & "                                 Width:=520, Height:=320)" & vbCrLf
::@@code = code & "    Set ch = co.Chart" & vbCrLf
::@@code = code & "    ch.ChartType = xlColumnStacked" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' Categories = labels (column A); explicit series so order is guaranteed." & vbCrLf
::@@code = code & "    Dim catRng As Range, spcRng As Range, valRng As Range" & vbCrLf
::@@code = code & "    Set catRng = ws.Range(ws.Cells(FIRST_DATA, 1), ws.Cells(lastRow, 1))" & vbCrLf
::@@code = code & "    Set spcRng = ws.Range(ws.Cells(FIRST_DATA, 2), ws.Cells(lastRow, 2))" & vbCrLf
::@@code = code & "    Set valRng = ws.Range(ws.Cells(FIRST_DATA, 3), ws.Cells(lastRow, 3))" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Do While ch.SeriesCollection.Count > 0" & vbCrLf
::@@code = code & "        ch.SeriesCollection(1).Delete" & vbCrLf
::@@code = code & "    Loop" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Dim serSpacer As Series, serValue As Series" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Set serSpacer = ch.SeriesCollection.NewSeries" & vbCrLf
::@@code = code & "    serSpacer.Name = ""Spacer""" & vbCrLf
::@@code = code & "    serSpacer.Values = spcRng" & vbCrLf
::@@code = code & "    serSpacer.XValues = catRng" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Set serValue = ch.SeriesCollection.NewSeries" & vbCrLf
::@@code = code & "    serValue.Name = ""Value""" & vbCrLf
::@@code = code & "    serValue.Values = valRng" & vbCrLf
::@@code = code & "    serValue.XValues = catRng" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- Make the Spacer series completely invisible ----" & vbCrLf
::@@code = code & "    serSpacer.Format.Fill.Visible = msoFalse" & vbCrLf
::@@code = code & "    serSpacer.Format.Line.Visible = msoFalse" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- Colour each Value point from its component's hex ----" & vbCrLf
::@@code = code & "    For i = 1 To nRows" & vbCrLf
::@@code = code & "        serValue.Points(i).Format.Fill.Visible = msoTrue" & vbCrLf
::@@code = code & "        serValue.Points(i).Format.Fill.ForeColor.RGB = HexToColor(comps(LBound(comps) + i - 1).HexCode)" & vbCrLf
::@@code = code & "        serValue.Points(i).Format.Line.Visible = msoFalse" & vbCrLf
::@@code = code & "    Next i" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- Data labels on the Value series ----" & vbCrLf
::@@code = code & "    serValue.HasDataLabels = True" & vbCrLf
::@@code = code & "    serValue.DataLabels.ShowValue = True" & vbCrLf
::@@code = code & "    serValue.DataLabels.NumberFormat = ""#,##0""" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- Title, legend ----" & vbCrLf
::@@code = code & "    ch.HasTitle = True" & vbCrLf
::@@code = code & "    ch.ChartTitle.Text = CHART_TITLE" & vbCrLf
::@@code = code & "    ch.HasLegend = False        ' the only ""real"" series is Value; hide the legend entirely" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' ---- Optional: thin connector lines between consecutive bars ----" & vbCrLf
::@@code = code & "    On Error Resume Next        ' connectors are cosmetic; never fail the build over them" & vbCrLf
::@@code = code & "    AddConnectors ch, comps" & vbCrLf
::@@code = code & "    On Error GoTo 0" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ws.Cells(1, 1).Select" & vbCrLf
::@@code = code & "End Sub" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' ------------------------------------------------------------" & vbCrLf
::@@code = code & "'  Draw thin horizontal connector lines linking the top of each" & vbCrLf
::@@code = code & "'  bar to the base of the next, so the eye follows the bridge." & vbCrLf
::@@code = code & "'  Cosmetic only -- wrapped in On Error Resume Next by the caller." & vbCrLf
::@@code = code & "' ------------------------------------------------------------" & vbCrLf
::@@code = code & "Private Sub AddConnectors(ch As Chart, comps() As Component)" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Dim pa As PlotArea" & vbCrLf
::@@code = code & "    Dim i As Long, n As Long" & vbCrLf
::@@code = code & "    Dim topLevel As Double" & vbCrLf
::@@code = code & "    Dim x1 As Double, x2 As Double, yPix As Double" & vbCrLf
::@@code = code & "    Dim vAxis As Axis, cAxis As Axis" & vbCrLf
::@@code = code & "    Dim plotL As Double, plotW As Double" & vbCrLf
::@@code = code & "    Dim ln As Shape" & vbCrLf
::@@code = code & "    Dim topOfBar As Double" & vbCrLf
::@@code = code & "    Dim minY As Double, maxY As Double" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    n = UBound(comps) - LBound(comps) + 1" & vbCrLf
::@@code = code & "    If n < 2 Then Exit Sub" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Set vAxis = ch.Axes(xlValue)" & vbCrLf
::@@code = code & "    minY = vAxis.MinimumScale" & vbCrLf
::@@code = code & "    maxY = vAxis.MaximumScale" & vbCrLf
::@@code = code & "    If maxY <= minY Then Exit Sub" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Set pa = ch.PlotArea" & vbCrLf
::@@code = code & "    plotL = pa.InsideLeft" & vbCrLf
::@@code = code & "    plotW = pa.InsideWidth" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    ' Each category occupies an equal slice of the plot width." & vbCrLf
::@@code = code & "    For i = LBound(comps) To UBound(comps) - 1" & vbCrLf
::@@code = code & "        ' Top level = cumulative running total after this component." & vbCrLf
::@@code = code & "        topLevel = comps(i).Spacer + comps(i).Value" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "        ' Convert the value level to a vertical pixel inside the plot area." & vbCrLf
::@@code = code & "        yPix = pa.InsideTop + pa.InsideHeight * (1 - (topLevel - minY) / (maxY - minY))" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "        ' Horizontal centres of bar i and bar i+1 (1-based slot index)." & vbCrLf
::@@code = code & "        Dim slot As Long" & vbCrLf
::@@code = code & "        slot = i - LBound(comps) + 1" & vbCrLf
::@@code = code & "        x1 = plotL + plotW * (slot - 0.5) / n" & vbCrLf
::@@code = code & "        x2 = plotL + plotW * (slot + 0.5) / n" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "        Set ln = ch.Shapes.AddLine(x1, yPix, x2, yPix)" & vbCrLf
::@@code = code & "        ln.Line.DashStyle = msoLineDash" & vbCrLf
::@@code = code & "        ln.Line.Weight = 0.75" & vbCrLf
::@@code = code & "        ln.Line.ForeColor.RGB = RGB(128, 128, 128)" & vbCrLf
::@@code = code & "    Next i" & vbCrLf
::@@code = code & "End Sub" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "'  INPUT HELPERS  (each returns False if the user cancels)" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' Free-text label. Returns False on Cancel." & vbCrLf
::@@code = code & "Private Function PromptLabel(ByRef outLabel As String, _" & vbCrLf
::@@code = code & "                             ByVal title As String, _" & vbCrLf
::@@code = code & "                             ByVal prompt As String) As Boolean" & vbCrLf
::@@code = code & "    Dim s As String" & vbCrLf
::@@code = code & "    s = InputBox(prompt, title)" & vbCrLf
::@@code = code & "    If StrPtr(s) = 0 Then" & vbCrLf
::@@code = code & "        PromptLabel = False" & vbCrLf
::@@code = code & "    Else" & vbCrLf
::@@code = code & "        outLabel = s" & vbCrLf
::@@code = code & "        PromptLabel = True" & vbCrLf
::@@code = code & "    End If" & vbCrLf
::@@code = code & "End Function" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' Numeric value. Re-prompts on non-numeric. Returns False on Cancel." & vbCrLf
::@@code = code & "Private Function PromptValue(ByRef outVal As Double, _" & vbCrLf
::@@code = code & "                             ByVal title As String, _" & vbCrLf
::@@code = code & "                             ByVal prompt As String) As Boolean" & vbCrLf
::@@code = code & "    Dim s As String" & vbCrLf
::@@code = code & "    Do" & vbCrLf
::@@code = code & "        s = InputBox(prompt, title)" & vbCrLf
::@@code = code & "        If StrPtr(s) = 0 Then" & vbCrLf
::@@code = code & "            PromptValue = False" & vbCrLf
::@@code = code & "            Exit Function" & vbCrLf
::@@code = code & "        End If" & vbCrLf
::@@code = code & "        s = Trim$(s)" & vbCrLf
::@@code = code & "        If IsNumeric(s) Then" & vbCrLf
::@@code = code & "            outVal = CDbl(s)" & vbCrLf
::@@code = code & "            PromptValue = True" & vbCrLf
::@@code = code & "            Exit Function" & vbCrLf
::@@code = code & "        End If" & vbCrLf
::@@code = code & "        MsgBox ""Please enter a valid number."", vbExclamation, title" & vbCrLf
::@@code = code & "    Loop" & vbCrLf
::@@code = code & "End Function" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' Hex colour. Re-prompts on invalid. Returns False on Cancel." & vbCrLf
::@@code = code & "' defHex pre-fills the InputBox default (pass vbNullString for none)." & vbCrLf
::@@code = code & "Private Function PromptHex(ByRef outHex As String, _" & vbCrLf
::@@code = code & "                           ByVal title As String, _" & vbCrLf
::@@code = code & "                           ByVal prompt As String, _" & vbCrLf
::@@code = code & "                           ByVal defHex As String) As Boolean" & vbCrLf
::@@code = code & "    Dim s As String, clean As String" & vbCrLf
::@@code = code & "    Do" & vbCrLf
::@@code = code & "        s = InputBox(prompt, title, defHex)" & vbCrLf
::@@code = code & "        If StrPtr(s) = 0 Then" & vbCrLf
::@@code = code & "            PromptHex = False" & vbCrLf
::@@code = code & "            Exit Function" & vbCrLf
::@@code = code & "        End If" & vbCrLf
::@@code = code & "        If IsValidHex(s, clean) Then" & vbCrLf
::@@code = code & "            outHex = clean" & vbCrLf
::@@code = code & "            PromptHex = True" & vbCrLf
::@@code = code & "            Exit Function" & vbCrLf
::@@code = code & "        End If" & vbCrLf
::@@code = code & "        MsgBox ""Please enter a valid 6-digit hex code, e.g. 1F4E79 "" & _" & vbCrLf
::@@code = code & "               ""(a leading # is allowed)."", vbExclamation, title" & vbCrLf
::@@code = code & "    Loop" & vbCrLf
::@@code = code & "End Function" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "'  HEX / COLOUR HELPERS" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' Validate a hex string. Accepts an optional leading ""#"" and" & vbCrLf
::@@code = code & "' surrounding whitespace. On success, cleanHex holds the 6 upper-case" & vbCrLf
::@@code = code & "' hex digits (no ""#"")." & vbCrLf
::@@code = code & "Private Function IsValidHex(ByVal raw As String, ByRef cleanHex As String) As Boolean" & vbCrLf
::@@code = code & "    Dim s As String, i As Long, c As String" & vbCrLf
::@@code = code & "    s = Trim$(raw)" & vbCrLf
::@@code = code & "    If Left$(s, 1) = ""#"" Then s = Mid$(s, 2)" & vbCrLf
::@@code = code & "    s = Trim$(s)" & vbCrLf
::@@code = code & "    If Len(s) <> 6 Then Exit Function" & vbCrLf
::@@code = code & "    For i = 1 To 6" & vbCrLf
::@@code = code & "        c = UCase$(Mid$(s, i, 1))" & vbCrLf
::@@code = code & "        If InStr(""0123456789ABCDEF"", c) = 0 Then Exit Function" & vbCrLf
::@@code = code & "    Next i" & vbCrLf
::@@code = code & "    cleanHex = UCase$(s)" & vbCrLf
::@@code = code & "    IsValidHex = True" & vbCrLf
::@@code = code & "End Function" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' Convert a 6-digit hex code (with or without a leading ""#"") to a VBA" & vbCrLf
::@@code = code & "' colour long. VBA's colour longs are stored BGR, so the byte order is" & vbCrLf
::@@code = code & "' RGB(red, green, blue) = red + green*256 + blue*65536 -- built here" & vbCrLf
::@@code = code & "' explicitly from the RR, GG, BB pairs." & vbCrLf
::@@code = code & "Public Function HexToColor(ByVal hex As String) As Long" & vbCrLf
::@@code = code & "    Dim clean As String" & vbCrLf
::@@code = code & "    Dim rr As Long, gg As Long, bb As Long" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    If Not IsValidHex(hex, clean) Then" & vbCrLf
::@@code = code & "        Err.Raise vbObjectError + 513, ""HexToColor"", _" & vbCrLf
::@@code = code & "                  ""Invalid hex colour: '"" & hex & ""'""" & vbCrLf
::@@code = code & "    End If" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    rr = CLng(""&H"" & Mid$(clean, 1, 2))" & vbCrLf
::@@code = code & "    gg = CLng(""&H"" & Mid$(clean, 3, 2))" & vbCrLf
::@@code = code & "    bb = CLng(""&H"" & Mid$(clean, 5, 2))" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    HexToColor = RGB(rr, gg, bb)        ' RGB() stores the long as BGR internally" & vbCrLf
::@@code = code & "End Function" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "'  SHEET HELPER" & vbCrLf
::@@code = code & "' ============================================================" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "' Return a worksheet with the given name, cleared. Creates it if it" & vbCrLf
::@@code = code & "' does not exist, otherwise deletes its charts and clears its cells." & vbCrLf
::@@code = code & "Private Function GetCleanSheet(ByVal name As String) As Worksheet" & vbCrLf
::@@code = code & "    Dim ws As Worksheet" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    On Error Resume Next" & vbCrLf
::@@code = code & "    Set ws = ActiveWorkbook.Worksheets(name)" & vbCrLf
::@@code = code & "    On Error GoTo 0" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    If ws Is Nothing Then" & vbCrLf
::@@code = code & "        Set ws = ActiveWorkbook.Worksheets.Add( _" & vbCrLf
::@@code = code & "                    After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))" & vbCrLf
::@@code = code & "        ws.Name = name" & vbCrLf
::@@code = code & "    Else" & vbCrLf
::@@code = code & "        Dim co As ChartObject" & vbCrLf
::@@code = code & "        For Each co In ws.ChartObjects" & vbCrLf
::@@code = code & "            co.Delete" & vbCrLf
::@@code = code & "        Next co" & vbCrLf
::@@code = code & "        ws.Cells.Clear" & vbCrLf
::@@code = code & "    End If" & vbCrLf
::@@code = code & "" & vbCrLf
::@@code = code & "    Set GetCleanSheet = ws" & vbCrLf
::@@code = code & "End Function" & vbCrLf
::@@Set xl = CreateObject("Excel.Application")
::@@xl.Visible = True
::@@Set wb = xl.Workbooks.Open(wbPath)
::@@On Error Resume Next
::@@Set vbProj = wb.VBProject
::@@If Err.Number <> 0 Then
::@@  MsgBox "Cannot access the VBA project object model." & vbCrLf & "Enable Excel Trust Center > Macro Settings > ""Trust access to the VBA project object model"", then drag the workbook on again.", vbExclamation, "NPV Bridge"
::@@  WScript.Quit 1
::@@End If
::@@On Error GoTo 0
::@@' Remove any existing copy so re-runs do not duplicate the module
::@@For Each c In vbProj.VBComponents
::@@  If c.Name = "modNPVBridge" Then vbProj.VBComponents.Remove c
::@@Next
::@@Set comp = vbProj.VBComponents.Add(1)   ' 1 = vbext_ct_StdModule
::@@comp.Name = "modNPVBridge"
::@@comp.CodeModule.AddFromString code
::@@xl.Run "BuildNPVBridge"
::@@' Workbook is intentionally left OPEN and unsaved for review.
