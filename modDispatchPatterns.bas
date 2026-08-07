Attribute VB_Name = "modDispatchPatterns"
Option Explicit

'==============================================================================
'  modDispatchPatterns
'------------------------------------------------------------------------------
'  Profiles battery charge/discharge behaviour by hour-ending from a dispatch
'  optimisation timeseries sitting in a worksheet, identifies the top recurring
'  daily dispatch patterns, and writes colour-coded summary tables plus native
'  Excel charts to fresh sheets in the same workbook.
'
'  Public entry point : RunDispatchAnalysis   (wire the Config button to this)
'  Everything else     : Private helpers, prefixed dp_ where practical so other
'                        macros can share the workbook namespace without a clash.
'
'  Design notes (see the Task spec):
'    * The whole used range is read once via .Value2 into a Variant array.
'    * Every output is assembled in a Variant array and written in ONE range
'      assignment. No cell-by-cell loops, no .Select / .Activate anywhere.
'    * Columns are resolved by header text into an index map - never by letter.
'    * Application state is disarmed on entry and always restored in the handler.
'==============================================================================


'==== SECTION: MODULE CONSTANTS =================================================

' --- Output / working sheet names -------------------------------------------
Private Const SH_CONFIG        As String = "Config"
Private Const SH_HOURLY        As String = "HourlyProfile"
Private Const SH_PATIDX        As String = "PatternIndex"
Private Const SH_PATHOURLY     As String = "PatternHourly"
Private Const SH_DAILY         As String = "DailyAssignment"
Private Const SH_RUNLOG        As String = "RunLog"
Private Const SH_CHARTS        As String = "Charts"
Private Const SH_CHARTDATA     As String = "ChartData"

Private Const HDR_ROW          As Long = 1

' --- Source column header text (resolved by name, any order, extras ignored) -
Private Const H_TS             As String = "timestamp"
Private Const H_SIM            As String = "simulation_id"
Private Const H_RT             As String = "energy_rt_price"
Private Const H_DA             As String = "energy_da_price"
Private Const H_DIS            As String = "battery_energy_discharge_mwh"
Private Const H_GRIDCHG        As String = "battery_energy_gridcharge_mwh"
Private Const H_RENCHG         As String = "battery_charge_from_renewable_mwh"
Private Const H_SOC            As String = "soc"
Private Const H_DISREV         As String = "battery_energy_discharge_revenue"
Private Const H_GRIDEXP        As String = "battery_energy_gridcharge_expenditure"

' --- Number formats ----------------------------------------------------------
Private Const FMT_PRICE        As String = "$ #,##0.00"
Private Const FMT_MWH          As String = "#,##0.0"
Private Const FMT_PCT          As String = "0.0%"
Private Const FMT_CYC          As String = "#,##0.00"
Private Const FMT_DATE         As String = "yyyy-mm-dd"

' --- Pattern colours (colour-blind safe Okabe-Ito) ---------------------------
'     P1 blue, P2 vermillion, P3 bluish-green, Other grey.
Private Const CLR_OTHER        As Long = 9868950   ' RGB(150,150,150)


'==== SECTION: MODULE STATE ====================================================
'  App-state cache (restored in the single top-level handler) and run bookkeeping.

Private mScreen   As Boolean
Private mEvents   As Boolean
Private mAlerts   As Boolean
Private mCalc     As XlCalculation
Private mProc     As String            ' name of the step currently executing
Private mLog      As Object            ' Scripting.Dictionary of runlog label -> value
Private mLogOrder As Collection        ' preserves insertion order for the RunLog sheet


'==============================================================================
'  RunDispatchAnalysis - PUBLIC ENTRY POINT
'==============================================================================
Public Sub RunDispatchAnalysis()

    ' Every local is declared here, at the top, before any executable statement -
    ' nothing is Dim'd inside a branch or after the Fail label. The current step
    ' name lives in the module-level mProc, so the handler reports the failing
    ' procedure without needing a local copy.
    Dim t0 As Double
    Dim cfg As Object                  ' validated config carrier
    Dim rc As Object                   ' row-column carrier (typed arrays + counts)
    Dim lErrNum As Long                ' l-prefixed locals; failing-proc name is module-level mProc
    Dim lErrDesc As String

    t0 = Timer

    ' Fail fast on workbook structure protection (adding/clearing sheets would
    ' error deep in the run otherwise) with an actionable message.
    If ThisWorkbook.ProtectStructure Then
        MsgBox "Workbook structure is protected. Unprotect via " & _
               "Review > Protect Workbook, then rerun.", vbExclamation, "modDispatchPatterns"
        Exit Sub
    End If

    On Error GoTo Fail
    dp_InitLog

    mProc = "dp_DisarmApp":            dp_DisarmApp
    mProc = "dp_EnsureConfig":         dp_EnsureConfig
    mProc = "dp_ReadConfig":           Set cfg = dp_ReadConfig()   ' raises on bad input
    mProc = "dp_LoadData":             Set rc = dp_LoadData(cfg)
    mProc = "dp_BuildDays":            dp_BuildDays rc, cfg
    mProc = "dp_BuildPatterns":        dp_BuildPatterns rc, cfg
    mProc = "dp_WriteHourlyProfile":   dp_WriteHourlyProfile rc, cfg
    mProc = "dp_WritePatternIndex":    dp_WritePatternIndex rc, cfg
    mProc = "dp_WritePatternHourly":   dp_WritePatternHourly rc, cfg
    mProc = "dp_WriteDailyAssignment": dp_WriteDailyAssignment rc, cfg

    ' dp_BuildCharts isolates each chart internally (per-chart error logged to
    ' RunLog, failure counted) - so a single bad chart no longer costs them all,
    ' while a genuine setup failure still propagates to the Fail handler.
    mProc = "dp_BuildCharts":          dp_BuildCharts rc, cfg

    mProc = "dp_Finish"
    dp_LogSet "runtime_seconds", Format$(Timer - t0, "0.00")
    dp_WriteRunLog cfg
    dp_RearmApp
    dp_ShowSummary rc, cfg
    Exit Sub

Fail:
    lErrNum = Err.Number
    lErrDesc = Err.Description
    dp_RearmApp
    On Error Resume Next
    dp_LogSet "ERROR", "in " & mProc & ": [" & lErrNum & "] " & lErrDesc
    dp_LogSet "runtime_seconds", Format$(Timer - t0, "0.00")
    dp_WriteRunLog Nothing
    On Error GoTo 0
    MsgBox "RunDispatchAnalysis failed." & vbCrLf & vbCrLf & _
           "Procedure : " & mProc & vbCrLf & _
           "Error     : [" & lErrNum & "] " & lErrDesc, _
           vbCritical, "modDispatchPatterns"
End Sub


'==== SECTION: APP STATE =======================================================

Private Sub dp_DisarmApp()
    With Application
        mScreen = .ScreenUpdating
        mEvents = .EnableEvents
        mAlerts = .DisplayAlerts
        mCalc = .Calculation
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub dp_RearmApp()
    On Error Resume Next
    With Application
        .Calculation = mCalc
        .DisplayAlerts = mAlerts
        .EnableEvents = mEvents
        .ScreenUpdating = mScreen
    End With
    On Error GoTo 0
End Sub


'==== SECTION: RUN LOG =========================================================

Private Sub dp_InitLog()
    Set mLog = CreateObject("Scripting.Dictionary")
    mLog.CompareMode = 1                ' vbTextCompare
    Set mLogOrder = New Collection
End Sub

' Params named logKey/logVal (not label/val) so they don't shadow the VBA
' intrinsics Val() and the MSForms Label class.
Private Sub dp_LogSet(ByVal logKey As String, ByVal logVal As Variant)
    If mLog Is Nothing Then dp_InitLog
    If Not mLog.Exists(logKey) Then mLogOrder.Add logKey
    mLog(logKey) = logVal
End Sub


'==== SECTION: CONFIG ==========================================================
'  Build (or refresh) the Config sheet: labelled input cells, each carrying a
'  workbook-scoped defined name. Defaults are only written into EMPTY cells so a
'  user's prior edits survive a refresh. Dropdowns and the run button are wired.

Private Sub dp_EnsureConfig()
    Dim ws As Worksheet
    Set ws = dp_SheetByName(SH_CONFIG)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SH_CONFIG
    End If

    ws.Range("A1").Value = "Dispatch Pattern Analysis - Configuration"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    ws.Range("A3").Value = "Setting"
    ws.Range("B3").Value = "Value"
    ws.Range("C3").Value = "Notes"
    ws.Range("A3:C3").Font.Bold = True

    ' name | default | notes | dropdown-list (empty = free text)
    Dim spec As Variant
    spec = Array( _
        Array("cfgDataSheet", "RawData", "Source worksheet holding the timeseries", ""), _
        Array("cfgPriceCol", "energy_rt_price", "Price series used for means/spread", "energy_rt_price,energy_da_price"), _
        Array("cfgTopN", 3, "Number of patterns to isolate (rest -> Other)", ""), _
        Array("cfgClusterMethod", "Exact", "Exact 24h signature, or relaxed Window", "Exact,Window"), _
        Array("cfgHEConvention", "Ending", "Ending: 00:00->HE1 | Beginning: shifted", "Ending,Beginning"), _
        Array("cfgTolMWh", 0.001, "MWh threshold for charging / discharging", ""), _
        Array("cfgSimID", "", "Single simulation_id, or blank to pool all", ""), _
        Array("cfgStartDate", "", "Lower date bound (inclusive), or blank", ""), _
        Array("cfgEndDate", "", "Upper date bound (inclusive), or blank", ""), _
        Array("cfgRenewableCost", "Opportunity", "Opportunity: value at market | Zero", "Opportunity,Zero") _
    )

    Dim i As Long, r As Long
    For i = LBound(spec) To UBound(spec)
        r = 4 + i
        Dim nm As String, dflt As Variant, note As String, list As String
        nm = spec(i)(0): dflt = spec(i)(1): note = spec(i)(2): list = spec(i)(3)

        ws.Cells(r, 1).Value = nm
        ws.Cells(r, 3).Value = note

        ' Only seed the default when the value cell is empty (preserve edits).
        If Len(CStr(ws.Cells(r, 2).Value)) = 0 Then ws.Cells(r, 2).Value = dflt

        ' Workbook-scoped defined name on the value cell.
        On Error Resume Next
        ThisWorkbook.Names(nm).Delete
        On Error GoTo 0
        ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & SH_CONFIG & "'!" & ws.Cells(r, 2).Address

        ' Dropdown validation where a list is supplied.
        With ws.Cells(r, 2).Validation
            .Delete
            If Len(list) > 0 Then
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:=list
            End If
        End With
    Next i

    ws.Cells(4, 1).Resize(UBound(spec) - LBound(spec) + 1, 1).Font.Bold = True
    ws.Columns("A:C").AutoFit

    ' Run button, wired to the public entry point.
    dp_EnsureRunButton ws
End Sub

Private Sub dp_EnsureRunButton(ws As Worksheet)
    On Error Resume Next
    Dim b As Object
    For Each b In ws.Buttons
        b.Delete
    Next b
    Dim btn As Object
    Set btn = ws.Buttons.Add(ws.Range("E4").Left, ws.Range("E4").Top, 150, 40)
    btn.OnAction = "RunDispatchAnalysis"
    btn.Caption = "Run Dispatch Analysis"
    btn.Name = "btnRunDispatch"
    On Error GoTo 0
End Sub

Private Function dp_ReadConfig() As Object
    Dim c As Object
    Set c = CreateObject("Scripting.Dictionary")
    c.CompareMode = 1

    c("dataSheet") = dp_NameVal("cfgDataSheet", "RawData")
    c("priceHdr") = LCase$(Trim$(dp_NameVal("cfgPriceCol", H_RT)))
    c("topN") = dp_ToLong(dp_NameVal("cfgTopN", 3), 3)
    c("method") = dp_NormChoice(dp_NameVal("cfgClusterMethod", "Exact"), "exact", "window", "exact")
    c("heConv") = dp_NormChoice(dp_NameVal("cfgHEConvention", "Ending"), "ending", "beginning", "ending")
    c("tol") = dp_ToDouble(dp_NameVal("cfgTolMWh", 0.001), 0.001)
    c("simFilter") = Trim$(CStr(dp_NameVal("cfgSimID", "")))
    c("renCost") = dp_NormChoice(dp_NameVal("cfgRenewableCost", "Opportunity"), "opportunity", "zero", "opportunity")

    Dim sd As Variant, ed As Variant
    sd = dp_NameVal("cfgStartDate", "")
    ed = dp_NameVal("cfgEndDate", "")
    c("startSer") = dp_DateBound(sd)     ' 0 => no bound
    c("endSer") = dp_DateBound(ed)

    ' --- Validation: fail with a specific message, not a deep type mismatch --
    If c("topN") < 1 Then Err.Raise 5001, , "cfgTopN must be >= 1 (got " & c("topN") & ")."
    If c("tol") < 0 Then Err.Raise 5002, , "cfgTolMWh must be >= 0 (got " & c("tol") & ")."
    If c("startSer") > 0 And c("endSer") > 0 And c("startSer") > c("endSer") Then
        Err.Raise 5003, , "cfgStartDate is after cfgEndDate."
    End If
    If c("priceHdr") <> LCase$(H_RT) And c("priceHdr") <> LCase$(H_DA) Then
        Err.Raise 5004, , "cfgPriceCol must be '" & H_RT & "' or '" & H_DA & "'."
    End If
    If dp_SheetByName(CStr(c("dataSheet"))) Is Nothing Then
        Err.Raise 5005, , "Source sheet '" & c("dataSheet") & "' not found."
    End If

    ' Echo config into the run log.
    dp_LogSet "cfg_dataSheet", c("dataSheet")
    dp_LogSet "cfg_priceCol", c("priceHdr")
    dp_LogSet "cfg_topN", c("topN")
    dp_LogSet "cfg_method", c("method")
    dp_LogSet "cfg_heConvention", c("heConv")
    dp_LogSet "cfg_tolMWh", c("tol")
    dp_LogSet "cfg_simFilter", IIf(Len(c("simFilter")) = 0, "(all)", c("simFilter"))
    dp_LogSet "cfg_startDate", IIf(c("startSer") = 0, "(none)", Format$(c("startSer"), FMT_DATE))
    dp_LogSet "cfg_endDate", IIf(c("endSer") = 0, "(none)", Format$(c("endSer"), FMT_DATE))
    dp_LogSet "cfg_renewableCost", c("renCost")

    Set dp_ReadConfig = c
End Function

Private Function dp_NameVal(ByVal nm As String, ByVal fallback As Variant) As Variant
    On Error GoTo useFallback
    dp_NameVal = ThisWorkbook.Names(nm).RefersToRange.Value
    Exit Function
useFallback:
    dp_NameVal = fallback
End Function

Private Function dp_DateBound(ByVal v As Variant) As Double
    ' Returns a date serial, or 0 when blank / unparseable.
    If IsEmpty(v) Then dp_DateBound = 0: Exit Function
    If Len(Trim$(CStr(v))) = 0 Then dp_DateBound = 0: Exit Function
    If IsDate(v) Then
        dp_DateBound = CLng(CDate(v))
    ElseIf IsNumeric(v) Then
        dp_DateBound = Int(CDbl(v))
    Else
        dp_DateBound = 0
    End If
End Function


'==== SECTION: DATA LOAD =======================================================
'  Read the used range once, resolve columns by header, then stream every row
'  into parallel typed arrays. Timestamps arrive as Excel serials (Double) via
'  .Value2; text timestamps are detected and CDate-coerced once, and counted.

Private Function dp_LoadData(cfg As Object) As Object
    Dim ws As Worksheet
    Set ws = dp_SheetByName(CStr(cfg("dataSheet")))

    Dim data As Variant
    data = ws.UsedRange.Value2
    If Not IsArray(data) Then Err.Raise 5010, , "Source sheet is empty."

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1): nCols = UBound(data, 2)
    If nRows <= HDR_ROW Then Err.Raise 5011, , "Source sheet has a header but no data rows."

    ' --- Column map: lower/trimmed header -> column index -------------------
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    colMap.CompareMode = 1
    Dim j As Long, h As String
    For j = 1 To nCols
        h = LCase$(Trim$(CStr(data(HDR_ROW, j))))
        If Len(h) > 0 And Not colMap.Exists(h) Then colMap(h) = j
    Next j

    ' Required columns (price column is whichever config selected).
    Dim missing As String
    missing = ""
    missing = missing & dp_ReqCol(colMap, H_TS)
    missing = missing & dp_ReqCol(colMap, H_SIM)
    missing = missing & dp_ReqCol(colMap, CStr(cfg("priceHdr")))
    missing = missing & dp_ReqCol(colMap, H_DIS)
    missing = missing & dp_ReqCol(colMap, H_GRIDCHG)
    If Len(missing) > 0 Then
        Err.Raise 5012, , "Missing required column(s):" & vbCrLf & missing
    End If

    Dim cTS As Long, cSim As Long, cPx As Long, cDis As Long, cGrid As Long
    cTS = colMap(LCase$(H_TS))
    cSim = colMap(LCase$(H_SIM))
    cPx = colMap(cfg("priceHdr"))
    cDis = colMap(LCase$(H_DIS))
    cGrid = colMap(LCase$(H_GRIDCHG))

    ' Optional columns (0 when absent).
    Dim cRen As Long, cSOC As Long, cDisRev As Long, cGridExp As Long
    cRen = dp_OptCol(colMap, H_RENCHG)
    cSOC = dp_OptCol(colMap, H_SOC)
    cDisRev = dp_OptCol(colMap, H_DISREV)
    cGridExp = dp_OptCol(colMap, H_GRIDEXP)

    Dim renPresent As Boolean: renPresent = (cRen > 0)
    Dim socPresent As Boolean: socPresent = (cSOC > 0)

    ' --- Allocate typed arrays (upper bound = every data row) ---------------
    Dim cap As Long: cap = nRows - HDR_ROW
    Dim rSimKey() As String: ReDim rSimKey(1 To cap)
    Dim rDateSer() As Long:  ReDim rDateSer(1 To cap)
    Dim rHE() As Long:       ReDim rHE(1 To cap)
    Dim rPrice() As Double:  ReDim rPrice(1 To cap)
    Dim rDis() As Double:    ReDim rDis(1 To cap)
    Dim rGrid() As Double:   ReDim rGrid(1 To cap)
    Dim rRen() As Double:    ReDim rRen(1 To cap)
    Dim rSOC() As Double:    ReDim rSOC(1 To cap)
    Dim rDisRev() As Double: ReDim rDisRev(1 To cap)
    Dim rGridExp() As Double: ReDim rGridExp(1 To cap)
    Dim rState() As String:  ReDim rState(1 To cap)

    Dim n As Long: n = 0
    Dim coerced As Long: coerced = 0
    Dim bothHrs As Long: bothHrs = 0
    Dim filtered As Long: filtered = 0

    Dim tol As Double: tol = cfg("tol")
    Dim heConv As String: heConv = cfg("heConv")
    Dim simFilter As String: simFilter = cfg("simFilter")
    Dim startSer As Double: startSer = cfg("startSer")
    Dim endSer As Double: endSer = cfg("endSer")

    Dim i As Long
    Dim ser As Double, dser As Long, hr As Long, he As Long
    Dim vTS As Variant, simKey As String
    Dim disMWh As Double, gridMWh As Double, renMWh As Double, chgMWh As Double

    For i = HDR_ROW + 1 To nRows
        vTS = data(i, cTS)
        If IsEmpty(vTS) Or (VarType(vTS) = vbString And Len(Trim$(CStr(vTS))) = 0) Then
            filtered = filtered + 1: GoTo NextRow
        End If

        ' Timestamp -> serial (Double). Coerce text once, counting it.
        If IsNumeric(vTS) And Not (VarType(vTS) = vbString) Then
            ser = CDbl(vTS)
        Else
            If IsDate(vTS) Then
                ser = CDbl(CDate(vTS))
                coerced = coerced + 1
            ElseIf IsNumeric(vTS) Then
                ser = CDbl(vTS)
            Else
                filtered = filtered + 1: GoTo NextRow
            End If
        End If

        dser = Int(ser)
        hr = Int((ser - dser) * 24 + 0.5)            ' 0..23 (rounded to the hour)
        If hr >= 24 Then hr = hr - 24: dser = dser + 1
        If hr < 0 Then filtered = filtered + 1: GoTo NextRow

        ' Hour-ending label per convention (see spec).
        If heConv = "ending" Then
            he = hr + 1                              ' 00:00 -> HE1 ... 23:00 -> HE24
        Else
            ' "beginning": timestamp marks interval start; shift by one so the
            ' interval starting 23:00 is HE24 and 00:00 rolls to prior-day HE24.
            If hr = 0 Then
                he = 24: dser = dser - 1
            Else
                he = hr
            End If
        End If
        If he < 1 Or he > 24 Then filtered = filtered + 1: GoTo NextRow

        ' Simulation filter (string compare so numeric or text ids both work).
        simKey = Trim$(CStr(data(i, cSim)))
        If Len(simFilter) > 0 Then
            If StrComp(simKey, simFilter, vbTextCompare) <> 0 Then
                filtered = filtered + 1: GoTo NextRow
            End If
        End If

        ' Date range filter.
        If startSer > 0 And dser < startSer Then filtered = filtered + 1: GoTo NextRow
        If endSer > 0 And dser > endSer Then filtered = filtered + 1: GoTo NextRow

        ' MWh components.
        disMWh = dp_Num(data(i, cDis))
        gridMWh = dp_Num(data(i, cGrid))
        If renPresent Then renMWh = dp_Num(data(i, cRen)) Else renMWh = 0#
        chgMWh = gridMWh + renMWh

        n = n + 1
        rSimKey(n) = simKey
        rDateSer(n) = dser
        rHE(n) = he
        rPrice(n) = dp_Num(data(i, cPx))
        rDis(n) = disMWh
        rGrid(n) = gridMWh
        rRen(n) = renMWh
        ' Guard optional columns with If/Else (IIf would evaluate data(i,0)).
        If cDisRev > 0 Then rDisRev(n) = dp_Num(data(i, cDisRev)) Else rDisRev(n) = 0#
        If cGridExp > 0 Then rGridExp(n) = dp_Num(data(i, cGridExp)) Else rGridExp(n) = 0#
        If socPresent Then rSOC(n) = dp_Num(data(i, cSOC)) Else rSOC(n) = 0#

        ' State classification.
        Dim isD As Boolean, isC As Boolean
        isD = (disMWh > tol)
        isC = (chgMWh > tol)
        If isD And isC Then
            rState(n) = "B": bothHrs = bothHrs + 1
        ElseIf isD Then
            rState(n) = "D"
        ElseIf isC Then
            rState(n) = "C"
        Else
            rState(n) = "-"
        End If
NextRow:
    Next i

    If n = 0 Then Err.Raise 5013, , "No rows remain after filtering (check simulation / date filters)."

    ' --- Carrier object -----------------------------------------------------
    Dim rc As Object
    Set rc = CreateObject("Scripting.Dictionary")
    rc.CompareMode = 1
    rc("n") = n
    rc("simKey") = rSimKey
    rc("dateSer") = rDateSer
    rc("he") = rHE
    rc("price") = rPrice
    rc("dis") = rDis
    rc("grid") = rGrid
    rc("ren") = rRen
    rc("soc") = rSOC
    rc("disRev") = rDisRev
    rc("gridExp") = rGridExp
    rc("state") = rState
    rc("renPresent") = renPresent
    rc("socPresent") = socPresent
    rc("hasDisRev") = (cDisRev > 0)
    rc("hasGridExp") = (cGridExp > 0)

    dp_LogSet "rows_loaded", nRows - HDR_ROW
    dp_LogSet "rows_kept", n
    dp_LogSet "rows_filtered", filtered
    dp_LogSet "B_hour_count", bothHrs
    dp_LogSet "text_timestamps_coerced", coerced
    dp_LogSet "renewable_charge_column", IIf(renPresent, "present", "ABSENT (treated as 0)")
    dp_LogSet "soc_column", IIf(socPresent, "present", "absent")

    Set dp_LoadData = rc
End Function

Private Function dp_ReqCol(colMap As Object, ByVal hdr As String) As String
    If colMap.Exists(LCase$(hdr)) Then
        dp_ReqCol = ""
    Else
        dp_ReqCol = "  - """ & hdr & """" & vbCrLf
    End If
End Function

Private Function dp_OptCol(colMap As Object, ByVal hdr As String) As Long
    If colMap.Exists(LCase$(hdr)) Then dp_OptCol = colMap(LCase$(hdr)) Else dp_OptCol = 0
End Function


'==== SECTION: DAYS & SIGNATURES ===============================================
'  Group rows by (simulation_id, date). A day is complete only when it holds
'  exactly 24 distinct hour-ending values. Complete days get a 24-char state
'  signature over HE1->HE24. Incomplete days are excluded from pattern analysis
'  (but their rows still feed the hourly profile) and counted.

Private Sub dp_BuildDays(rc As Object, cfg As Object)
    Dim n As Long: n = rc("n")
    Dim simKey() As String: simKey = rc("simKey")
    Dim dateSer() As Long:  dateSer = rc("dateSer")
    Dim he() As Long:       he = rc("he")
    Dim state() As String:  state = rc("state")

    ' Assign a dense day index to each (sim,date) key.
    Dim dayIdxOf As Object
    Set dayIdxOf = CreateObject("Scripting.Dictionary")
    dayIdxOf.CompareMode = 1

    Dim rDay() As Long: ReDim rDay(1 To n)
    Dim nDays As Long: nDays = 0
    Dim i As Long, k As String
    For i = 1 To n
        k = simKey(i) & "|" & CStr(dateSer(i))
        If Not dayIdxOf.Exists(k) Then
            nDays = nDays + 1
            dayIdxOf(k) = nDays
        End If
        rDay(i) = dayIdxOf(k)
    Next i

    ' Day-level stores.
    Dim daySim() As String:   ReDim daySim(1 To nDays)
    Dim dayDate() As Long:    ReDim dayDate(1 To nDays)
    Dim dayStateGrid() As String  ' (day, HE) modal-by-first state code
    ReDim dayStateGrid(1 To nDays, 1 To 24)
    Dim dayPresent() As Boolean: ReDim dayPresent(1 To nDays, 1 To 24)
    Dim dayDistinct() As Long:   ReDim dayDistinct(1 To nDays)

    Dim di As Long, hh As Long
    For i = 1 To n
        di = rDay(i): hh = he(i)
        If daySim(di) = "" Then
            daySim(di) = simKey(i): dayDate(di) = dateSer(i)
        End If
        If Not dayPresent(di, hh) Then
            dayPresent(di, hh) = True
            dayDistinct(di) = dayDistinct(di) + 1
            dayStateGrid(di, hh) = state(i)
        End If
    Next i

    ' Completeness + signature.
    Dim dayComplete() As Boolean: ReDim dayComplete(1 To nDays)
    Dim daySig() As String:       ReDim daySig(1 To nDays)
    Dim complete As Long: complete = 0
    Dim excluded As Long: excluded = 0
    For di = 1 To nDays
        If dayDistinct(di) = 24 Then
            dayComplete(di) = True
            Dim sg As String: sg = ""
            For hh = 1 To 24
                sg = sg & dayStateGrid(di, hh)
            Next hh
            daySig(di) = sg
            complete = complete + 1
        Else
            excluded = excluded + 1
        End If
    Next di

    rc("nDays") = nDays
    rc("rDay") = rDay
    rc("daySim") = daySim
    rc("dayDate") = dayDate
    rc("dayComplete") = dayComplete
    rc("daySig") = daySig
    rc("dayStateGrid") = dayStateGrid

    dp_LogSet "days_total", nDays
    dp_LogSet "days_complete", complete
    dp_LogSet "days_excluded_incomplete", excluded
End Sub


'==== SECTION: PATTERN IDENTIFICATION ==========================================
'  Group complete days into patterns, rank by day count, keep top-N (rest ->
'  Other). Exact keys on the 24-char signature; Window relaxes each day to a
'  charge/discharge span descriptor. A representative modal signature per rank
'  drives descriptors, heatmaps and PatternHourly.

Private Sub dp_BuildPatterns(rc As Object, cfg As Object)
    Dim nDays As Long: nDays = rc("nDays")
    Dim dayComplete() As Boolean: dayComplete = rc("dayComplete")
    Dim daySig() As String:       daySig = rc("daySig")
    Dim dayStateGrid() As String: dayStateGrid = rc("dayStateGrid")
    Dim method As String: method = cfg("method")
    Dim topN As Long: topN = cfg("topN")

    ' Idle day = complete day whose signature has no charge/discharge/both.
    Dim dayIdle() As Boolean: ReDim dayIdle(1 To dp_Max(nDays, 1))
    Dim di As Long
    For di = 1 To nDays
        If dayComplete(di) Then
            dayIdle(di) = (InStr(daySig(di), "C") = 0 And InStr(daySig(di), "D") = 0 _
                           And InStr(daySig(di), "B") = 0)
        End If
    Next di

    ' Count days per group key (idle days do not compete for a numbered rank).
    Dim keyCount As Object: Set keyCount = CreateObject("Scripting.Dictionary")
    keyCount.CompareMode = 1
    Dim dayKey() As String: ReDim dayKey(1 To dp_Max(nDays, 1))

    Dim gk As String
    Dim totalComplete As Long: totalComplete = 0
    For di = 1 To nDays
        If dayComplete(di) Then
            totalComplete = totalComplete + 1
            If Not dayIdle(di) Then
                If method = "window" Then gk = dp_WindowKey(daySig(di)) Else gk = daySig(di)
                dayKey(di) = gk
                If keyCount.Exists(gk) Then keyCount(gk) = keyCount(gk) + 1 Else keyCount(gk) = 1
            End If
        End If
    Next di

    ' Rank keys by descending day count (ties broken by key for determinism).
    Dim keys() As String, cnts() As Long
    Dim nk As Long: nk = keyCount.Count
    Dim patRank As Object: Set patRank = CreateObject("Scripting.Dictionary")   ' key -> rank
    patRank.CompareMode = 1
    Dim topKeys() As String
    Dim nTop As Long: nTop = 0

    If nk > 0 Then
        ReDim keys(1 To nk): ReDim cnts(1 To nk)
        Dim vk As Variant, ix As Long: ix = 0
        For Each vk In keyCount.keys
            ix = ix + 1
            keys(ix) = CStr(vk)
            cnts(ix) = keyCount(vk)
        Next vk
        dp_SortByCountDesc keys, cnts, 1, nk

        nTop = dp_Min(topN, nk)
        ReDim topKeys(1 To dp_Max(nTop, 1))
        Dim rnk As Long
        For rnk = 1 To nTop
            topKeys(rnk) = keys(rnk)
            patRank(keys(rnk)) = rnk
        Next rnk
    End If

    ' Slot model: 1..nTop patterns, Other = nTop+1, Idle = nTop+2.
    Dim otherRank As Long: otherRank = nTop + 1
    Dim idleRank As Long: idleRank = nTop + 2
    Dim slots As Long: slots = idleRank

    ' Assign each day: rank 1..nTop, 0 Other, -2 Idle, -1 excluded (incomplete).
    Dim dayRank() As Long: ReDim dayRank(1 To dp_Max(nDays, 1))
    Dim topCoverDays As Long: topCoverDays = 0
    For di = 1 To nDays
        If Not dayComplete(di) Then
            dayRank(di) = -1
        ElseIf dayIdle(di) Then
            dayRank(di) = -2
        ElseIf patRank.Exists(dayKey(di)) Then
            dayRank(di) = patRank(dayKey(di))
            topCoverDays = topCoverDays + 1
        Else
            dayRank(di) = 0
        End If
    Next di

    ' Representative modal signature per slot from member days.
    Dim modalGrid() As String: ReDim modalGrid(1 To dp_Max(slots, 1), 1 To 24)
    Dim votes As Object: Set votes = CreateObject("Scripting.Dictionary")
    votes.CompareMode = 1
    Dim hh As Long, slot As Long, st As String, vkey As String
    For di = 1 To nDays
        slot = dp_SlotOf(dayRank(di), nTop)
        If slot > 0 Then
            For hh = 1 To 24
                st = dayStateGrid(di, hh)
                vkey = slot & "|" & hh & "|" & st
                If votes.Exists(vkey) Then votes(vkey) = votes(vkey) + 1 Else votes(vkey) = 1
            Next hh
        End If
    Next di

    Dim states As Variant: states = Array("D", "C", "B", "-")
    Dim s As Long, bestSt As String, bestV As Long, vv As Long
    For slot = 1 To slots
        For hh = 1 To 24
            bestSt = "-": bestV = -1
            For s = LBound(states) To UBound(states)
                vkey = slot & "|" & hh & "|" & states(s)
                If votes.Exists(vkey) Then vv = votes(vkey) Else vv = 0
                If vv > bestV Then bestV = vv: bestSt = CStr(states(s))
            Next s
            modalGrid(slot, hh) = bestSt
        Next hh
    Next slot

    ' Top-N coverage (share of all complete days captured by numbered patterns).
    Dim coverage As Double
    If totalComplete > 0 Then coverage = topCoverDays / totalComplete Else coverage = 0

    rc("nTop") = nTop
    rc("otherRank") = otherRank
    rc("idleRank") = idleRank
    rc("slots") = slots
    rc("dayRank") = dayRank
    rc("dayIdle") = dayIdle
    rc("dayKey") = dayKey
    rc("modalGrid") = modalGrid
    rc("totalComplete") = totalComplete
    rc("coverage") = coverage
    If nTop > 0 Then rc("topKeys") = topKeys

    dp_LogSet "patterns_found_distinct", nk
    dp_LogSet "patterns_topN_used", nTop
    dp_LogSet "topN_coverage_pct", Format$(coverage, "0.0%")
    If cfg("method") = "exact" And coverage < 0.5 And totalComplete > 0 Then
        dp_LogSet "recommendation", "Exact coverage < 50% - consider cfgClusterMethod = Window."
        rc("recommendWindow") = True
    Else
        rc("recommendWindow") = False
    End If
End Sub

Private Function dp_WindowKey(ByVal sig As String) As String
    ' Relax a 24-char signature to:
    '   firstChargeHE-lastChargeHE|firstDischargeHE-lastDischargeHE|nChg|nDis
    Dim fC As Long, lC As Long, fD As Long, lD As Long, nC As Long, nD As Long
    Dim hh As Long, ch As String
    fC = 0: lC = 0: fD = 0: lD = 0: nC = 0: nD = 0
    For hh = 1 To 24
        ch = Mid$(sig, hh, 1)
        If ch = "C" Or ch = "B" Then
            nC = nC + 1
            If fC = 0 Then fC = hh
            lC = hh
        End If
        If ch = "D" Or ch = "B" Then
            nD = nD + 1
            If fD = 0 Then fD = hh
            lD = hh
        End If
    Next hh
    dp_WindowKey = fC & "-" & lC & "|" & fD & "-" & lD & "|" & nC & "|" & nD
End Function

' Map a dayRank to its aggregation slot: 1..nTop patterns, Other=nTop+1,
' Idle=nTop+2, or 0 for an excluded (incomplete) day.
Private Function dp_SlotOf(ByVal rk As Long, ByVal nTop As Long) As Long
    Select Case rk
        Case Is >= 1: dp_SlotOf = rk
        Case 0:       dp_SlotOf = nTop + 1
        Case -2:      dp_SlotOf = nTop + 2
        Case Else:    dp_SlotOf = 0
    End Select
End Function

Private Function dp_SlotLabel(ByVal slot As Long, ByVal nTop As Long) As String
    If slot <= nTop Then
        dp_SlotLabel = "P" & slot
    ElseIf slot = nTop + 1 Then
        dp_SlotLabel = "Other"
    Else
        dp_SlotLabel = "Idle"
    End If
End Function

Private Function dp_SlotColor(ByVal slot As Long, ByVal nTop As Long) As Long
    If slot <= nTop Then
        dp_SlotColor = dp_PatternColor(slot)
    ElseIf slot = nTop + 1 Then
        dp_SlotColor = dp_PatternColor(-1)     ' Other grey
    Else
        dp_SlotColor = dp_PatternColor(-2)     ' Idle light grey
    End If
End Function

' First/last charge and discharge HE within a 24-char signature (0 = none).
Private Sub dp_DaySpan(ByVal sig As String, ByRef fC As Long, ByRef lC As Long, _
                       ByRef fD As Long, ByRef lD As Long)
    Dim hh As Long, ch As String
    fC = 0: lC = 0: fD = 0: lD = 0
    For hh = 1 To 24
        ch = Mid$(sig, hh, 1)
        If ch = "C" Or ch = "B" Then
            If fC = 0 Then fC = hh
            lC = hh
        End If
        If ch = "D" Or ch = "B" Then
            If fD = 0 Then fD = hh
            lD = hh
        End If
    Next hh
End Sub


'==== SECTION: DESCRIPTORS =====================================================
'  Human-readable descriptor from a 24-char signature, e.g.
'  "charge HE2-5, discharge HE18-20", collapsing contiguous runs, commas on gaps.

Private Function dp_Descriptor(ByVal sig As String) As String
    Dim cPart As String, dPart As String
    cPart = dp_RunsOf(sig, "C")     ' C-or-B counts as charging
    dPart = dp_RunsOf(sig, "D")     ' D-or-B counts as discharging

    Dim out As String: out = ""
    If Len(cPart) > 0 Then out = "charge " & cPart
    If Len(dPart) > 0 Then
        If Len(out) > 0 Then out = out & ", "
        out = out & "discharge " & dPart
    End If
    If Len(out) = 0 Then out = "idle (no charge/discharge)"
    dp_Descriptor = out
End Function

Private Function dp_RunsOf(ByVal sig As String, ByVal want As String) As String
    Dim hh As Long, ch As String
    Dim inRun As Boolean, runStart As Long
    Dim parts As String: parts = ""
    inRun = False
    For hh = 1 To 24
        ch = Mid$(sig, hh, 1)
        Dim hit As Boolean
        hit = (ch = want) Or (ch = "B")
        If hit And Not inRun Then
            inRun = True: runStart = hh
        ElseIf (Not hit) And inRun Then
            parts = parts & dp_RunLabel(runStart, hh - 1) & ", "
            inRun = False
        End If
    Next hh
    If inRun Then parts = parts & dp_RunLabel(runStart, 24) & ", "
    If Len(parts) >= 2 Then parts = Left$(parts, Len(parts) - 2)
    dp_RunsOf = parts
End Function

Private Function dp_RunLabel(ByVal a As Long, ByVal b As Long) As String
    If a = b Then
        dp_RunLabel = "HE" & a
    Else
        dp_RunLabel = "HE" & a & "-" & b
    End If
End Function


'==== SECTION: OUTPUT - HOURLYPROFILE ==========================================
'  One row per HE 1-24, aggregated over ALL kept rows (complete or not).
'  Price statistics are UNWEIGHTED: mean and P10/P50/P90 of the market price
'  across the charging / discharging hours at that HE (no MWh weighting).
'  Charge MWh is split into its grid and renewable components.

Private Sub dp_WriteHourlyProfile(rc As Object, cfg As Object)
    Dim n As Long: n = rc("n")
    Dim he() As Long:      he = rc("he")
    Dim price() As Double: price = rc("price")
    Dim dis() As Double:   dis = rc("dis")
    Dim grid() As Double:  grid = rc("grid")
    Dim ren() As Double:   ren = rc("ren")
    Dim state() As String: state = rc("state")

    ' Accumulators per HE.
    Dim chgHrs(1 To 24) As Long, disHrs(1 To 24) As Long, idleHrs(1 To 24) As Long
    Dim dayCountHE(1 To 24) As Long
    Dim chgMWhTot(1 To 24) As Double, disMWhTot(1 To 24) As Double
    Dim gridMWhTot(1 To 24) As Double, renMWhTot(1 To 24) As Double
    Dim chgPxSum(1 To 24) As Double, disPxSum(1 To 24) As Double   ' unweighted price sums
    Dim priceSum(1 To 24) As Double, priceCnt(1 To 24) As Long

    ' Percentile buffers: count per HE, lay out contiguous slices in ONE flat
    ' array (prefix offsets), fill in place, sort each slice once, then read
    ' percentiles. No per-row array copies.
    Dim chgPxN(1 To 24) As Long, disPxN(1 To 24) As Long
    Dim hh As Long, i As Long
    For i = 1 To n
        hh = he(i)
        dayCountHE(hh) = dayCountHE(hh) + 1
        If state(i) = "C" Or state(i) = "B" Then chgPxN(hh) = chgPxN(hh) + 1
        If state(i) = "D" Or state(i) = "B" Then disPxN(hh) = disPxN(hh) + 1
    Next i

    Dim chgOff(1 To 24) As Long, disOff(1 To 24) As Long
    Dim totC As Long, totD As Long
    totC = 0: totD = 0
    For hh = 1 To 24
        chgOff(hh) = totC: totC = totC + chgPxN(hh)
        disOff(hh) = totD: totD = totD + disPxN(hh)
    Next hh
    Dim chgAll() As Double, disAll() As Double
    If totC > 0 Then ReDim chgAll(1 To totC)
    If totD > 0 Then ReDim disAll(1 To totD)
    Dim chgFill(1 To 24) As Long, disFill(1 To 24) As Long

    Dim chg As Double
    For i = 1 To n
        hh = he(i)
        chg = grid(i) + ren(i)
        priceSum(hh) = priceSum(hh) + price(i): priceCnt(hh) = priceCnt(hh) + 1

        If state(i) = "C" Or state(i) = "B" Then
            chgHrs(hh) = chgHrs(hh) + 1
            chgMWhTot(hh) = chgMWhTot(hh) + chg
            gridMWhTot(hh) = gridMWhTot(hh) + grid(i)
            renMWhTot(hh) = renMWhTot(hh) + ren(i)
            chgPxSum(hh) = chgPxSum(hh) + price(i)
            chgFill(hh) = chgFill(hh) + 1
            chgAll(chgOff(hh) + chgFill(hh)) = price(i)
        End If

        If state(i) = "D" Or state(i) = "B" Then
            disHrs(hh) = disHrs(hh) + 1
            disMWhTot(hh) = disMWhTot(hh) + dis(i)
            disPxSum(hh) = disPxSum(hh) + price(i)
            disFill(hh) = disFill(hh) + 1
            disAll(disOff(hh) + disFill(hh)) = price(i)
        End If

        If state(i) = "-" Then idleHrs(hh) = idleHrs(hh) + 1
    Next i

    ' Assemble output array (header + 24).
    Dim nc As Long: nc = 22
    Dim o() As Variant: ReDim o(1 To 25, 1 To nc)
    Dim hdr As Variant
    hdr = Array("he", "n_days", "charge_hours", "charge_freq_pct", "discharge_hours", _
        "discharge_freq_pct", "idle_freq_pct", "charge_mwh_total", "gridcharge_mwh_total", _
        "renewable_charge_mwh_total", "discharge_mwh_total", "mean_charge_mwh_when_charging", _
        "mean_discharge_mwh_when_discharging", "charge_price_mean", "discharge_price_mean", _
        "charge_price_p10", "charge_price_p50", "charge_price_p90", _
        "discharge_price_p10", "discharge_price_p50", "discharge_price_p90", "mean_price_all_hours")
    Dim c As Long
    For c = 1 To nc: o(1, c) = hdr(c - 1): Next c

    Dim netMean(1 To 24) As Variant
    Dim r As Long, dn As Long
    For hh = 1 To 24
        r = hh + 1
        dn = dayCountHE(hh)
        o(r, 1) = hh
        o(r, 2) = dn
        o(r, 3) = chgHrs(hh)
        o(r, 4) = dp_Div(chgHrs(hh), dn)
        o(r, 5) = disHrs(hh)
        o(r, 6) = dp_Div(disHrs(hh), dn)
        o(r, 7) = dp_Div(idleHrs(hh), dn)
        o(r, 8) = chgMWhTot(hh)
        o(r, 9) = gridMWhTot(hh)
        o(r, 10) = renMWhTot(hh)
        o(r, 11) = disMWhTot(hh)
        o(r, 12) = dp_Div(chgMWhTot(hh), chgHrs(hh))
        o(r, 13) = dp_Div(disMWhTot(hh), disHrs(hh))
        If chgHrs(hh) > 0 Then o(r, 14) = dp_Div(chgPxSum(hh), chgHrs(hh)) Else o(r, 14) = ""
        If disHrs(hh) > 0 Then o(r, 15) = dp_Div(disPxSum(hh), disHrs(hh)) Else o(r, 15) = ""
        If chgPxN(hh) > 0 Then
            dp_QuickSortDouble chgAll, chgOff(hh) + 1, chgOff(hh) + chgPxN(hh)
            o(r, 16) = dp_PercentileInc(chgAll, chgOff(hh) + 1, chgPxN(hh), 0.1)
            o(r, 17) = dp_PercentileInc(chgAll, chgOff(hh) + 1, chgPxN(hh), 0.5)
            o(r, 18) = dp_PercentileInc(chgAll, chgOff(hh) + 1, chgPxN(hh), 0.9)
        End If
        If disPxN(hh) > 0 Then
            dp_QuickSortDouble disAll, disOff(hh) + 1, disOff(hh) + disPxN(hh)
            o(r, 19) = dp_PercentileInc(disAll, disOff(hh) + 1, disPxN(hh), 0.1)
            o(r, 20) = dp_PercentileInc(disAll, disOff(hh) + 1, disPxN(hh), 0.5)
            o(r, 21) = dp_PercentileInc(disAll, disOff(hh) + 1, disPxN(hh), 0.9)
        End If
        o(r, 22) = dp_Div(priceSum(hh), priceCnt(hh))
        netMean(hh) = dp_Div(disMWhTot(hh) - chgMWhTot(hh), dn)
    Next hh

    Dim ws As Worksheet
    Set ws = dp_GetOrResetSheet(SH_HOURLY)
    ws.Range("A1").Resize(25, nc).Value = o
    dp_MakeTable ws, "tblHourly", 25, nc
    dp_FormatCols ws, Array(4, 6, 7), FMT_PCT
    dp_FormatCols ws, Array(8, 9, 10, 11, 12, 13), FMT_MWH
    dp_FormatCols ws, Array(14, 15, 16, 17, 18, 19, 20, 21, 22), FMT_PRICE
    ws.Columns.AutoFit

    ' Stash series for the charts (unweighted means and percentiles only).
    rc("hp_chgFreq") = dp_ColFrom(o, 4, 2, 25)
    rc("hp_disFreq") = dp_ColFrom(o, 6, 2, 25)
    rc("hp_meanPrice") = dp_ColFrom(o, 22, 2, 25)
    rc("hp_netMean") = netMean
    rc("hp_chgP10") = dp_ColFrom(o, 16, 2, 25)
    rc("hp_chgP50") = dp_ColFrom(o, 17, 2, 25)
    rc("hp_chgP90") = dp_ColFrom(o, 18, 2, 25)
    rc("hp_disP10") = dp_ColFrom(o, 19, 2, 25)
    rc("hp_disP50") = dp_ColFrom(o, 20, 2, 25)
    rc("hp_disP90") = dp_ColFrom(o, 21, 2, 25)
End Sub


'==== SECTION: OUTPUT - PATTERN AGGREGATES =====================================
'  Compute per-(rank,HE) aggregates over complete days, plus per-rank daily
'  totals. Shared by PatternIndex, PatternHourly and the charts, so it runs once
'  here and caches results on the carrier.

Private Sub dp_AggregatePatterns(rc As Object, cfg As Object)
    If rc.Exists("pa_done") Then Exit Sub

    Dim n As Long: n = rc("n")
    Dim rDay() As Long:    rDay = rc("rDay")
    Dim he() As Long:      he = rc("he")
    Dim price() As Double: price = rc("price")
    Dim dis() As Double:   dis = rc("dis")
    Dim grid() As Double:  grid = rc("grid")
    Dim ren() As Double:   ren = rc("ren")
    Dim soc() As Double:   soc = rc("soc")
    Dim state() As String: state = rc("state")
    Dim disRev() As Double: disRev = rc("disRev")
    Dim gridExp() As Double: gridExp = rc("gridExp")
    Dim dayRank() As Long: dayRank = rc("dayRank")
    Dim nTop As Long: nTop = rc("nTop")
    Dim slots As Long: slots = rc("slots")   ' 1..nTop, Other=nTop+1, Idle=nTop+2
    Dim socPresent As Boolean: socPresent = rc("socPresent")

    ' Per (slot,HE) accumulators. Price stats are UNWEIGHTED means: sum of the
    ' market price over charging / discharging hours, divided by the hour count.
    Dim cCnt() As Long: ReDim cCnt(1 To slots, 1 To 24)
    Dim dCnt() As Long: ReDim dCnt(1 To slots, 1 To 24)
    Dim occ() As Long:  ReDim occ(1 To slots, 1 To 24)      ' day-hours observed
    Dim cM() As Double:  ReDim cM(1 To slots, 1 To 24)
    Dim dM() As Double:  ReDim dM(1 To slots, 1 To 24)
    Dim netM() As Double: ReDim netM(1 To slots, 1 To 24)
    Dim cPxSum() As Double: ReDim cPxSum(1 To slots, 1 To 24)   ' sum price | charging
    Dim dPxSum() As Double: ReDim dPxSum(1 To slots, 1 To 24)   ' sum price | discharging
    Dim pxSum() As Double: ReDim pxSum(1 To slots, 1 To 24)     ' sum price | all hours
    Dim socSum() As Double: ReDim socSum(1 To slots, 1 To 24)
    Dim socCnt() As Long:   ReDim socCnt(1 To slots, 1 To 24)

    Dim i As Long, di As Long, slot As Long, hh As Long, chg As Double
    For i = 1 To n
        di = rDay(i)
        slot = dp_SlotOf(dayRank(di), nTop)
        If slot = 0 Then GoTo NextI        ' incomplete day - excluded from patterns
        hh = he(i)
        chg = grid(i) + ren(i)

        occ(slot, hh) = occ(slot, hh) + 1
        pxSum(slot, hh) = pxSum(slot, hh) + price(i)
        netM(slot, hh) = netM(slot, hh) + (dis(i) - chg)
        If socPresent Then socSum(slot, hh) = socSum(slot, hh) + soc(i): socCnt(slot, hh) = socCnt(slot, hh) + 1

        If state(i) = "C" Or state(i) = "B" Then
            cCnt(slot, hh) = cCnt(slot, hh) + 1
            cM(slot, hh) = cM(slot, hh) + chg
            cPxSum(slot, hh) = cPxSum(slot, hh) + price(i)
        End If
        If state(i) = "D" Or state(i) = "B" Then
            dCnt(slot, hh) = dCnt(slot, hh) + 1
            dM(slot, hh) = dM(slot, hh) + dis(i)
            dPxSum(slot, hh) = dPxSum(slot, hh) + price(i)
        End If
NextI:
    Next i

    ' Per-slot daily totals (charge/discharge MWh and actual $ cost/revenue).
    Dim nDays As Long: nDays = rc("nDays")
    Dim dChg() As Double: ReDim dChg(1 To dp_Max(nDays, 1))
    Dim dDis() As Double: ReDim dDis(1 To dp_Max(nDays, 1))
    Dim dCost() As Double: ReDim dCost(1 To dp_Max(nDays, 1))
    Dim dRev() As Double: ReDim dRev(1 To dp_Max(nDays, 1))
    For i = 1 To n
        di = rDay(i)
        If dp_SlotOf(dayRank(di), nTop) = 0 Then GoTo NextI2
        chg = grid(i) + ren(i)
        dChg(di) = dChg(di) + chg
        dDis(di) = dDis(di) + dis(i)
        dCost(di) = dCost(di) + gridExp(i)
        dRev(di) = dRev(di) + disRev(i)
NextI2:
    Next i
    Dim slotDays() As Long: ReDim slotDays(1 To slots)
    Dim slotChg() As Double: ReDim slotChg(1 To slots)
    Dim slotDis() As Double: ReDim slotDis(1 To slots)
    Dim slotCost() As Double: ReDim slotCost(1 To slots)
    Dim slotRev() As Double: ReDim slotRev(1 To slots)
    For di = 1 To nDays
        slot = dp_SlotOf(dayRank(di), nTop)
        If slot = 0 Then GoTo NextD
        slotDays(slot) = slotDays(slot) + 1
        slotChg(slot) = slotChg(slot) + dChg(di)
        slotDis(slot) = slotDis(slot) + dDis(di)
        slotCost(slot) = slotCost(slot) + dCost(di)
        slotRev(slot) = slotRev(slot) + dRev(di)
NextD:
    Next di

    ' Battery-energy capacity proxy for implied cycles: max observed SOC.
    Dim capProxy As Double: capProxy = 0
    If socPresent Then
        For i = 1 To n
            If soc(i) > capProxy Then capProxy = soc(i)
        Next i
    End If

    rc("pa_slots") = slots
    rc("pa_cCnt") = cCnt: rc("pa_dCnt") = dCnt: rc("pa_occ") = occ
    rc("pa_cM") = cM: rc("pa_dM") = dM: rc("pa_netM") = netM
    rc("pa_cPxSum") = cPxSum: rc("pa_dPxSum") = dPxSum: rc("pa_pxSum") = pxSum
    rc("pa_socSum") = socSum: rc("pa_socCnt") = socCnt
    rc("pa_slotDays") = slotDays: rc("pa_slotChg") = slotChg: rc("pa_slotDis") = slotDis
    rc("pa_slotCost") = slotCost: rc("pa_slotRev") = slotRev
    rc("pa_capProxy") = capProxy
    rc("pa_done") = True
End Sub


'==== SECTION: MONTHLY AGGREGATES ==============================================
'  Per-month day counts by slot (for the seasonal-mix chart) and per-month mean
'  first/last charge & discharge HE (for the timing-drift chart). Months are in
'  chronological order across the whole (filtered) timeseries.

Private Sub dp_AggregateMonthly(rc As Object)
    If rc.Exists("mo_done") Then Exit Sub
    Dim nDays As Long: nDays = rc("nDays")
    Dim dayDate() As Long: dayDate = rc("dayDate")
    Dim daySig() As String: daySig = rc("daySig")
    Dim dayComplete() As Boolean: dayComplete = rc("dayComplete")
    Dim dayRank() As Long: dayRank = rc("dayRank")
    Dim nTop As Long: nTop = rc("nTop")
    Dim slots As Long: slots = rc("slots")

    ' Distinct months keyed Year*12 + (Month-1), collected then sorted ascending.
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim keysArr() As Long: ReDim keysArr(1 To dp_Max(nDays, 1))
    Dim nMonths As Long: nMonths = 0
    Dim di As Long, mk As Long, dt As Date
    For di = 1 To nDays
        If dayComplete(di) Then
            dt = CDate(dayDate(di))
            mk = Year(dt) * 12 + (Month(dt) - 1)
            If Not seen.Exists(mk) Then
                nMonths = nMonths + 1: seen(mk) = 1: keysArr(nMonths) = mk
            End If
        End If
    Next di

    Dim order() As Long: ReDim order(1 To dp_Max(nMonths, 1))
    Dim a As Long, b As Long, tmp As Long
    For a = 1 To nMonths: order(a) = keysArr(a): Next a
    For a = 1 To nMonths - 1
        For b = a + 1 To nMonths
            If order(b) < order(a) Then tmp = order(a): order(a) = order(b): order(b) = tmp
        Next b
    Next a
    Dim posOf As Object: Set posOf = CreateObject("Scripting.Dictionary")
    Dim moLabels() As String: ReDim moLabels(1 To dp_Max(nMonths, 1))
    For a = 1 To nMonths
        posOf(order(a)) = a
        moLabels(a) = Format$(DateSerial(order(a) \ 12, (order(a) Mod 12) + 1, 1), "yyyy-mm")
    Next a

    Dim moCount() As Long: ReDim moCount(1 To dp_Max(nMonths, 1), 1 To slots)
    Dim fcS() As Double: ReDim fcS(1 To dp_Max(nMonths, 1))
    Dim lcS() As Double: ReDim lcS(1 To dp_Max(nMonths, 1))
    Dim fdS() As Double: ReDim fdS(1 To dp_Max(nMonths, 1))
    Dim ldS() As Double: ReDim ldS(1 To dp_Max(nMonths, 1))
    Dim cN() As Long: ReDim cN(1 To dp_Max(nMonths, 1))
    Dim dN() As Long: ReDim dN(1 To dp_Max(nMonths, 1))
    Dim slot As Long, mp As Long, fC As Long, lC As Long, fD As Long, lD As Long
    For di = 1 To nDays
        If Not dayComplete(di) Then GoTo NextDay
        dt = CDate(dayDate(di))
        mp = posOf(Year(dt) * 12 + (Month(dt) - 1))
        slot = dp_SlotOf(dayRank(di), nTop)
        If slot >= 1 Then moCount(mp, slot) = moCount(mp, slot) + 1
        dp_DaySpan daySig(di), fC, lC, fD, lD
        If fC > 0 Then fcS(mp) = fcS(mp) + fC: lcS(mp) = lcS(mp) + lC: cN(mp) = cN(mp) + 1
        If fD > 0 Then fdS(mp) = fdS(mp) + fD: ldS(mp) = ldS(mp) + lD: dN(mp) = dN(mp) + 1
NextDay:
    Next di

    Dim firstChg() As Variant: ReDim firstChg(1 To dp_Max(nMonths, 1))
    Dim lastChg() As Variant:  ReDim lastChg(1 To dp_Max(nMonths, 1))
    Dim firstDis() As Variant: ReDim firstDis(1 To dp_Max(nMonths, 1))
    Dim lastDis() As Variant:  ReDim lastDis(1 To dp_Max(nMonths, 1))
    For a = 1 To nMonths
        If cN(a) > 0 Then
            firstChg(a) = fcS(a) / cN(a): lastChg(a) = lcS(a) / cN(a)
        Else
            firstChg(a) = "": lastChg(a) = ""
        End If
        If dN(a) > 0 Then
            firstDis(a) = fdS(a) / dN(a): lastDis(a) = ldS(a) / dN(a)
        Else
            firstDis(a) = "": lastDis(a) = ""
        End If
    Next a

    rc("mo_nMonths") = nMonths
    rc("mo_labels") = moLabels
    rc("mo_count") = moCount
    rc("mo_firstChg") = firstChg
    rc("mo_lastChg") = lastChg
    rc("mo_firstDis") = firstDis
    rc("mo_lastDis") = lastDis
    rc("mo_done") = True
End Sub


'==== SECTION: OUTPUT - PATTERNINDEX ===========================================

Private Sub dp_WritePatternIndex(rc As Object, cfg As Object)
    dp_AggregatePatterns rc, cfg
    Dim nTop As Long: nTop = rc("nTop")
    Dim otherRank As Long: otherRank = rc("otherRank")
    Dim slots As Long: slots = rc("slots")
    Dim modalGrid() As String: modalGrid = rc("modalGrid")
    Dim slotDays() As Long: slotDays = rc("pa_slotDays")
    Dim slotChg() As Double: slotChg = rc("pa_slotChg")
    Dim slotDis() As Double: slotDis = rc("pa_slotDis")
    Dim totalComplete As Long: totalComplete = rc("totalComplete")
    Dim capProxy As Double: capProxy = rc("pa_capProxy")
    Dim socPresent As Boolean: socPresent = rc("socPresent")
    Dim cCnt() As Long: cCnt = rc("pa_cCnt")
    Dim dCnt() As Long: dCnt = rc("pa_dCnt")
    Dim cPxSum() As Double: cPxSum = rc("pa_cPxSum")
    Dim dPxSum() As Double: dPxSum = rc("pa_dPxSum")

    ' Per-slot rollups for EVERY slot (top-N, Other, Idle); cached for charts.
    Dim slotLabel() As String: ReDim slotLabel(1 To slots)
    Dim slotSig() As String:   ReDim slotSig(1 To slots)
    Dim slotDesc() As String:  ReDim slotDesc(1 To slots)
    Dim slotShare() As Double: ReDim slotShare(1 To slots)
    Dim slotMeanChg() As Variant: ReDim slotMeanChg(1 To slots)   ' mean price | charging
    Dim slotMeanDis() As Variant: ReDim slotMeanDis(1 To slots)   ' mean price | discharging
    Dim slotSpread() As Variant:  ReDim slotSpread(1 To slots)
    Dim slotCycles() As Variant:  ReDim slotCycles(1 To slots)

    Dim slot As Long, hh As Long
    Dim sig As String, cPx As Double, cCt As Long, dPx As Double, dCt As Long
    For slot = 1 To slots
        sig = "": cPx = 0: cCt = 0: dPx = 0: dCt = 0
        For hh = 1 To 24
            sig = sig & modalGrid(slot, hh)
            cPx = cPx + cPxSum(slot, hh): cCt = cCt + cCnt(slot, hh)
            dPx = dPx + dPxSum(slot, hh): dCt = dCt + dCnt(slot, hh)
        Next hh
        slotLabel(slot) = dp_SlotLabel(slot, nTop)
        slotSig(slot) = sig
        slotDesc(slot) = dp_Descriptor(sig)
        slotShare(slot) = dp_Div(slotDays(slot), totalComplete)
        Dim mChg As Double, mDis As Double
        mChg = dp_Div(cPx, cCt): mDis = dp_Div(dPx, dCt)
        If cCt > 0 Then slotMeanChg(slot) = mChg Else slotMeanChg(slot) = ""
        If dCt > 0 Then slotMeanDis(slot) = mDis Else slotMeanDis(slot) = ""
        If cCt > 0 And dCt > 0 Then slotSpread(slot) = mDis - mChg Else slotSpread(slot) = ""
        If socPresent And capProxy > 0 Then
            slotCycles(slot) = dp_Div(dp_Div(slotDis(slot), slotDays(slot)), capProxy)
        Else
            slotCycles(slot) = ""
        End If
    Next slot

    ' Sheet shows top-N + Other (Idle omitted from the index table).
    Dim nRowsOut As Long: nRowsOut = otherRank
    Dim nc As Long: nc = 12
    Dim o() As Variant: ReDim o(1 To nRowsOut + 1, 1 To nc)
    Dim hdr As Variant
    hdr = Array("pattern_rank", "pattern_label", "signature", "descriptor", "n_days", _
        "share_of_days_pct", "avg_daily_charge_mwh", "avg_daily_discharge_mwh", _
        "charge_price_mean", "discharge_price_mean", "price_spread", "implied_cycles_per_day")
    Dim c As Long
    For c = 1 To nc: o(1, c) = hdr(c - 1): Next c

    Dim r As Long
    For slot = 1 To otherRank
        r = slot + 1
        o(r, 1) = IIf(slot = otherRank, "Other", CStr(slot))
        o(r, 2) = slotLabel(slot)
        o(r, 3) = slotSig(slot)
        o(r, 4) = slotDesc(slot)
        o(r, 5) = slotDays(slot)
        o(r, 6) = slotShare(slot)
        o(r, 7) = dp_Div(slotChg(slot), slotDays(slot))
        o(r, 8) = dp_Div(slotDis(slot), slotDays(slot))
        o(r, 9) = slotMeanChg(slot)
        o(r, 10) = slotMeanDis(slot)
        o(r, 11) = slotSpread(slot)
        o(r, 12) = slotCycles(slot)
    Next slot

    Dim ws As Worksheet
    Set ws = dp_GetOrResetSheet(SH_PATIDX)
    ws.Range("A1").Resize(nRowsOut + 1, nc).Value = o
    dp_MakeTable ws, "tblPatternIndex", nRowsOut + 1, nc
    dp_FormatCols ws, Array(6), FMT_PCT
    dp_FormatCols ws, Array(7, 8), FMT_MWH
    dp_FormatCols ws, Array(9, 10, 11), FMT_PRICE
    dp_FormatCols ws, Array(12), FMT_CYC
    ws.Columns.AutoFit

    For slot = 1 To otherRank
        ws.Cells(slot + 1, 2).Interior.Color = dp_SlotColor(slot, nTop)
        ws.Cells(slot + 1, 2).Font.Color = dp_ContrastFont(dp_SlotColor(slot, nTop))
    Next slot

    rc("slotLabel") = slotLabel
    rc("slotDesc") = slotDesc
    rc("slotShare") = slotShare
    rc("slotMeanChg") = slotMeanChg
    rc("slotMeanDis") = slotMeanDis
    rc("slotSpread") = slotSpread
    rc("slotCycles") = slotCycles
    rc("piArray") = o
End Sub


'==== SECTION: OUTPUT - PATTERNHOURLY ==========================================
'  Tidy long: one row per pattern x HE.

Private Sub dp_WritePatternHourly(rc As Object, cfg As Object)
    dp_AggregatePatterns rc, cfg
    Dim nTop As Long: nTop = rc("nTop")
    Dim otherRank As Long: otherRank = rc("otherRank")
    Dim modalGrid() As String: modalGrid = rc("modalGrid")
    Dim occ() As Long: occ = rc("pa_occ")
    Dim cCnt() As Long: cCnt = rc("pa_cCnt")
    Dim dCnt() As Long: dCnt = rc("pa_dCnt")
    Dim cM() As Double: cM = rc("pa_cM")
    Dim dM() As Double: dM = rc("pa_dM")
    Dim netM() As Double: netM = rc("pa_netM")
    Dim cPxSum() As Double: cPxSum = rc("pa_cPxSum")
    Dim dPxSum() As Double: dPxSum = rc("pa_dPxSum")
    Dim pxSum() As Double: pxSum = rc("pa_pxSum")
    Dim socSum() As Double: socSum = rc("pa_socSum")
    Dim socCnt() As Long: socCnt = rc("pa_socCnt")

    Dim nc As Long: nc = 13
    Dim nRowsOut As Long: nRowsOut = otherRank * 24     ' top-N + Other
    Dim o() As Variant: ReDim o(1 To nRowsOut + 1, 1 To nc)
    Dim hdr As Variant
    hdr = Array("pattern_rank", "pattern_label", "he", "modal_state", _
        "charge_freq_pct_within_pattern", "discharge_freq_pct_within_pattern", _
        "mean_charge_mwh", "mean_discharge_mwh", "mean_net_mwh", _
        "charge_price_mean", "discharge_price_mean", "mean_price", "mean_soc")
    Dim c As Long
    For c = 1 To nc: o(1, c) = hdr(c - 1): Next c

    Dim rk As Long, hh As Long, r As Long: r = 1
    For rk = 1 To otherRank
        For hh = 1 To 24
            r = r + 1
            o(r, 1) = IIf(rk = otherRank, "Other", CStr(rk))
            o(r, 2) = dp_SlotLabel(rk, nTop)
            o(r, 3) = hh
            o(r, 4) = modalGrid(rk, hh)
            o(r, 5) = dp_Div(cCnt(rk, hh), occ(rk, hh))
            o(r, 6) = dp_Div(dCnt(rk, hh), occ(rk, hh))
            o(r, 7) = dp_Div(cM(rk, hh), cCnt(rk, hh))
            o(r, 8) = dp_Div(dM(rk, hh), dCnt(rk, hh))
            o(r, 9) = dp_Div(netM(rk, hh), occ(rk, hh))
            If cCnt(rk, hh) > 0 Then o(r, 10) = dp_Div(cPxSum(rk, hh), cCnt(rk, hh)) Else o(r, 10) = ""
            If dCnt(rk, hh) > 0 Then o(r, 11) = dp_Div(dPxSum(rk, hh), dCnt(rk, hh)) Else o(r, 11) = ""
            o(r, 12) = dp_Div(pxSum(rk, hh), occ(rk, hh))
            If socCnt(rk, hh) > 0 Then o(r, 13) = dp_Div(socSum(rk, hh), socCnt(rk, hh)) Else o(r, 13) = ""
        Next hh
    Next rk

    Dim ws As Worksheet
    Set ws = dp_GetOrResetSheet(SH_PATHOURLY)
    ws.Range("A1").Resize(nRowsOut + 1, nc).Value = o
    dp_MakeTable ws, "tblPatternHourly", nRowsOut + 1, nc
    dp_FormatCols ws, Array(5, 6), FMT_PCT
    dp_FormatCols ws, Array(7, 8, 9), FMT_MWH
    dp_FormatCols ws, Array(10, 11, 12), FMT_PRICE
    dp_FormatCols ws, Array(13), FMT_MWH
    ws.Columns.AutoFit

    rc("phArray") = o
End Sub


'==== SECTION: OUTPUT - DAILYASSIGNMENT ========================================

Private Sub dp_WriteDailyAssignment(rc As Object, cfg As Object)
    dp_AggregatePatterns rc, cfg
    Dim nDays As Long: nDays = rc("nDays")
    Dim daySim() As String: daySim = rc("daySim")
    Dim dayDate() As Long: dayDate = rc("dayDate")
    Dim daySig() As String: daySig = rc("daySig")
    Dim dayComplete() As Boolean: dayComplete = rc("dayComplete")
    Dim dayRank() As Long: dayRank = rc("dayRank")
    Dim nTop As Long: nTop = rc("nTop")

    ' Daily totals rebuilt here from rows (kept local; small vs. n).
    Dim n As Long: n = rc("n")
    Dim rDay() As Long: rDay = rc("rDay")
    Dim dis() As Double: dis = rc("dis")
    Dim grid() As Double: grid = rc("grid")
    Dim ren() As Double: ren = rc("ren")
    Dim disRev() As Double: disRev = rc("disRev")
    Dim gridExp() As Double: gridExp = rc("gridExp")

    Dim dChg() As Double: ReDim dChg(1 To dp_Max(nDays, 1))
    Dim dDis() As Double: ReDim dDis(1 To dp_Max(nDays, 1))
    Dim dCost() As Double: ReDim dCost(1 To dp_Max(nDays, 1))
    Dim dRev() As Double: ReDim dRev(1 To dp_Max(nDays, 1))
    Dim i As Long, di As Long
    For i = 1 To n
        di = rDay(i)
        dChg(di) = dChg(di) + grid(i) + ren(i)
        dDis(di) = dDis(di) + dis(i)
        dCost(di) = dCost(di) + gridExp(i)
        dRev(di) = dRev(di) + disRev(i)
    Next i

    ' Only complete days are assigned to patterns; include all days but mark
    ' incomplete ones as "(incomplete)".
    Dim nc As Long: nc = 10
    Dim o() As Variant: ReDim o(1 To nDays + 1, 1 To nc)
    Dim hdr As Variant
    hdr = Array("simulation_id", "date", "signature", "pattern_rank", "pattern_label", _
        "daily_charge_mwh", "daily_discharge_mwh", "daily_charge_cost", _
        "daily_discharge_revenue", "daily_realized_spread")
    Dim c As Long
    For c = 1 To nc: o(1, c) = hdr(c - 1): Next c

    Dim r As Long, rk As Long
    For di = 1 To nDays
        r = di + 1
        o(r, 1) = daySim(di)
        o(r, 2) = CDbl(dayDate(di))
        If dayComplete(di) Then
            o(r, 3) = daySig(di)
            rk = dayRank(di)
            If rk >= 1 Then
                o(r, 4) = rk: o(r, 5) = dp_SlotLabel(rk, nTop)
            ElseIf rk = -2 Then
                o(r, 4) = "Idle": o(r, 5) = "Idle"
            Else
                o(r, 4) = "Other": o(r, 5) = "Other"
            End If
        Else
            o(r, 3) = "(incomplete)"
            o(r, 4) = "": o(r, 5) = "(excluded)"
        End If
        o(r, 6) = dChg(di)
        o(r, 7) = dDis(di)
        o(r, 8) = dCost(di)
        o(r, 9) = dRev(di)
        ' Realized spread (per-MWh) = revenue/discharge - cost/charge
        ' (dp_Div yields 0 for a side with no volume, avoiding a null subtraction).
        o(r, 10) = dp_Div(dRev(di), dDis(di)) - dp_Div(dCost(di), dChg(di))
    Next di

    Dim ws As Worksheet
    Set ws = dp_GetOrResetSheet(SH_DAILY)
    ws.Range("A1").Resize(nDays + 1, nc).Value = o
    dp_MakeTable ws, "tblDailyAssignment", nDays + 1, nc
    dp_FormatCols ws, Array(2), FMT_DATE
    dp_FormatCols ws, Array(6, 7), FMT_MWH
    dp_FormatCols ws, Array(8, 9, 10), FMT_PRICE
    ws.Columns.AutoFit
End Sub


'==== SECTION: OUTPUT - RUNLOG =================================================

Private Sub dp_WriteRunLog(cfg As Object)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = dp_GetOrResetSheet(SH_RUNLOG)

    Dim nRowsOut As Long: nRowsOut = mLogOrder.Count + 1
    Dim o() As Variant: ReDim o(1 To nRowsOut, 1 To 2)
    o(1, 1) = "metric": o(1, 2) = "value"
    Dim i As Long
    For i = 1 To mLogOrder.Count
        o(i + 1, 1) = mLogOrder(i)
        o(i + 1, 2) = CStr(mLog(mLogOrder(i)))
    Next i
    ws.Range("A1").Resize(nRowsOut, 2).Value = o
    dp_MakeTable ws, "tblRunLog", nRowsOut, 2
    ws.Columns.AutoFit
    On Error GoTo 0
End Sub


'==== SECTION: CHARTS ==========================================================
'  Staging on ChartData (xlSheetVeryHidden); charts live on Charts in a
'  non-overlapping grid. PatternColor()/TintColor() are the single source of
'  colour so a pattern is identical in every visual.

Private Sub dp_BuildCharts(rc As Object, cfg As Object)
    Dim wsC As Worksheet, wsD As Worksheet
    Set wsC = dp_GetOrResetSheet(SH_CHARTS)
    Set wsD = dp_GetOrResetSheet(SH_CHARTDATA)

    ' The Charts sheet was reset (its ChartObjects already cleared by the reset
    ' helper) BEFORE ChartData, so no chart holds a reference to a cleared
    ' staging range. ChartData is visible now (reset helper) and is hidden again
    ' only at the very end, after every chart is built.
    dp_AggregatePatterns rc, cfg
    dp_AggregateMonthly rc

    Dim sfx As String
    sfx = " (sim: " & IIf(Len(cfg("simFilter")) = 0, "all", cfg("simFilter")) & _
          ", price: " & cfg("priceHdr") & ")"

    Dim nTop As Long: nTop = rc("nTop")
    Dim otherRank As Long: otherRank = rc("otherRank")
    Dim slots As Long: slots = rc("slots")
    Dim heCol As Variant: heCol = dp_HeCol()

    Dim netM() As Double: netM = rc("pa_netM")
    Dim occ() As Long: occ = rc("pa_occ")
    Dim socSum() As Double: socSum = rc("pa_socSum")
    Dim socCnt() As Long: socCnt = rc("pa_socCnt")

    ' Layout: 2 tiles per row, 460 x 290 pts, 20-pt gutters.
    Const TW As Double = 460, TH As Double = 290, GUT As Double = 20
    Const X0 As Double = 15, Y0 As Double = 15
    Dim rk As Long, hh As Long
    Dim chartFails As Long: chartFails = 0

    ' Each chart is isolated: On Error Resume Next stays active across the whole
    ' block; Err is cleared before each chart and inspected right after, so a
    ' single failing chart is logged to RunLog and skipped instead of taking
    ' down every remaining chart.
    On Error Resume Next

    ' ---- Chart 1: frequency by hour ---------------------------------------
    Err.Clear
    dp_StageBlock wsD, 1, 1, Array("HE", "disFreq", "chgFreqNeg", "meanPrice"), _
        Array(heCol, rc("hp_disFreq"), dp_Negate(rc("hp_chgFreq")), rc("hp_meanPrice")), 24
    dp_Chart1Freq wsC, wsD, 1, X0, Y0, TW, TH, sfx
    dp_ChartErr "1 Frequency by hour", chartFails

    ' ---- Chart 2: mean net MWh by hour ------------------------------------
    Err.Clear
    Dim netPos As Variant, netNeg As Variant
    Dim nm As Variant: nm = rc("hp_netMean")
    ReDim netPos(1 To 24): ReDim netNeg(1 To 24)
    For hh = 1 To 24
        Dim mv As Double: mv = 0
        If IsNumeric(nm(hh)) Then mv = CDbl(nm(hh))
        If mv >= 0 Then
            netPos(hh) = mv: netNeg(hh) = 0
        Else
            netPos(hh) = 0: netNeg(hh) = mv
        End If
    Next hh
    dp_StageBlock wsD, 1, 7, Array("HE", "netDis", "netChg", "meanPrice"), _
        Array(heCol, netPos, netNeg, rc("hp_meanPrice")), 24
    dp_Chart2Net wsC, wsD, 7, X0 + TW + GUT, Y0, TW, TH, sfx
    dp_ChartErr "2 Mean net MWh by hour", chartFails

    ' ---- Chart 3: pattern profiles (one per top-N, comparable scales) ------
    Err.Clear
    Dim lo As Double, hi As Double: lo = 0: hi = 0
    For rk = 1 To nTop
        For hh = 1 To 24
            If occ(rk, hh) > 0 Then
                mv = netM(rk, hh) / occ(rk, hh)
                If mv < lo Then lo = mv
                If mv > hi Then hi = mv
            End If
        Next hh
    Next rk
    If hi <= lo Then hi = lo + 1
    Dim headsP() As String: ReDim headsP(0 To 2 * nTop)
    Dim colsP() As Variant: ReDim colsP(0 To 2 * nTop)
    headsP(0) = "HE": colsP(0) = heCol
    For rk = 1 To nTop
        Dim vp As Variant, vn As Variant: ReDim vp(1 To 24): ReDim vn(1 To 24)
        For hh = 1 To 24
            mv = 0
            If occ(rk, hh) > 0 Then mv = netM(rk, hh) / occ(rk, hh)
            If mv >= 0 Then
                vp(hh) = mv: vn(hh) = 0
            Else
                vp(hh) = 0: vn(hh) = mv
            End If
        Next hh
        headsP(2 * rk - 1) = "P" & rk & " dis": colsP(2 * rk - 1) = vp
        headsP(2 * rk) = "P" & rk & " chg": colsP(2 * rk) = vn
    Next rk
    If nTop > 0 Then
        dp_StageBlock wsD, 1, 13, headsP, colsP, 24
        dp_Chart3Profiles wsC, wsD, 13, rc, lo, hi, X0, Y0 + TH + GUT, sfx
    End If
    dp_ChartErr "3 Pattern profiles", chartFails

    ' ---- Chart 4: price distribution by hour ------------------------------
    Err.Clear
    Dim chgBand As Variant, disBand As Variant
    chgBand = dp_BandCol(rc("hp_chgP10"), rc("hp_chgP90"))
    disBand = dp_BandCol(rc("hp_disP10"), rc("hp_disP90"))
    dp_StageBlock wsD, 1, 13 + (2 * nTop + 2), _
        Array("HE", "chgP10", "chgBand", "disP10", "disBand", "chgP50", "chgP90", "disP50", "disP90", "meanPrice"), _
        Array(heCol, rc("hp_chgP10"), chgBand, rc("hp_disP10"), disBand, _
              rc("hp_chgP50"), rc("hp_chgP90"), rc("hp_disP50"), rc("hp_disP90"), rc("hp_meanPrice")), 24
    dp_Chart4PriceDist wsC, wsD, 13 + (2 * nTop + 2), X0, Y0 + 2 * (TH + GUT), TW, TH, sfx
    dp_ChartErr "4 Price distribution", chartFails

    ' ---- Chart 5: mean SOC by pattern (only if SOC present) ---------------
    If rc("socPresent") Then
        Err.Clear
        Dim cSoc As Long: cSoc = 13 + (2 * nTop + 2) + 11
        Dim headsS() As String: ReDim headsS(0 To otherRank)
        Dim colsS() As Variant: ReDim colsS(0 To otherRank)
        Dim maxSoc As Double: maxSoc = 0
        headsS(0) = "HE": colsS(0) = heCol
        For rk = 1 To otherRank
            headsS(rk) = dp_SlotLabel(rk, nTop)
            Dim vsoc As Variant: ReDim vsoc(1 To 24)
            For hh = 1 To 24
                If socCnt(rk, hh) > 0 Then
                    vsoc(hh) = socSum(rk, hh) / socCnt(rk, hh)
                    If vsoc(hh) > maxSoc Then maxSoc = vsoc(hh)
                Else
                    vsoc(hh) = ""
                End If
            Next hh
            colsS(rk) = vsoc
        Next rk
        dp_StageBlock wsD, 1, cSoc, headsS, colsS, 24
        dp_Chart5Soc wsC, wsD, cSoc, otherRank, nTop, maxSoc, X0 + TW + GUT, Y0 + 2 * (TH + GUT), TW, TH, sfx
        dp_ChartErr "5 Mean SOC by pattern", chartFails
    End If

    ' ---- Charts 6 & 7: monthly views --------------------------------------
    Dim nMonths As Long: nMonths = rc("mo_nMonths")
    If nMonths > 0 Then
        Dim moLabels() As String: moLabels = rc("mo_labels")
        Dim moLabV As Variant: ReDim moLabV(1 To nMonths)
        Dim a As Long
        For a = 1 To nMonths: moLabV(a) = moLabels(a): Next a

        ' Chart 6: monthly pattern mix
        Err.Clear
        Dim moCount() As Long: moCount = rc("mo_count")
        Dim headsM() As String: ReDim headsM(0 To slots)
        Dim colsM() As Variant: ReDim colsM(0 To slots)
        headsM(0) = "month": colsM(0) = moLabV
        For rk = 1 To slots
            headsM(rk) = dp_SlotLabel(rk, nTop)
            Dim vc As Variant: ReDim vc(1 To nMonths)
            For a = 1 To nMonths: vc(a) = moCount(a, rk): Next a
            colsM(rk) = vc
        Next rk
        Dim cMix As Long: cMix = 13 + (2 * nTop + 2) + 11 + (otherRank + 2)
        dp_StageBlock wsD, 1, cMix, headsM, colsM, nMonths
        dp_Chart6MonthlyMix wsC, wsD, cMix, nMonths, slots, nTop, X0, Y0 + 3 * (TH + GUT), TW, TH, sfx
        dp_ChartErr "6 Monthly pattern mix", chartFails

        ' Chart 7: dispatch timing drift by month
        Err.Clear
        Dim cDrift As Long: cDrift = cMix + (slots + 2)
        dp_StageBlock wsD, 1, cDrift, Array("month", "firstChg", "lastChg", "firstDis", "lastDis"), _
            Array(moLabV, rc("mo_firstChg"), rc("mo_lastChg"), rc("mo_firstDis"), rc("mo_lastDis")), nMonths
        dp_Chart7Drift wsC, wsD, cDrift, nMonths, X0 + TW + GUT, Y0 + 3 * (TH + GUT), TW, TH, sfx
        dp_ChartErr "7 Dispatch timing drift", chartFails
    End If

    ' ---- Chart 8: pattern frequency ranking (horizontal bar) --------------
    Err.Clear
    dp_Chart8Ranking wsC, wsD, rc, X0, Y0 + 4 * (TH + GUT), TW, TH, sfx
    dp_ChartErr "8 Pattern frequency ranking", chartFails

    ' ---- Pattern heatmap (cell-based) -------------------------------------
    Err.Clear
    dp_BuildHeatmap wsC, rc
    dp_ChartErr "Pattern heatmap", chartFails

    On Error GoTo 0
    dp_LogSet "charts_failed", chartFails
    rc("chart_fail_count") = chartFails
    wsD.Visible = xlSheetVeryHidden
End Sub

' Log a per-chart failure (if Err is set) to RunLog and bump the failure count,
' then clear Err so the next chart starts clean. Called under On Error Resume Next.
Private Sub dp_ChartErr(ByVal chartName As String, ByRef fails As Long)
    If Err.Number <> 0 Then
        dp_LogSet "chart_error [" & chartName & "]", "[" & Err.Number & "] " & Err.Description
        fails = fails + 1
        Err.Clear
    End If
End Sub

Private Sub dp_Chart1Freq(wsC As Worksheet, wsD As Worksheet, ByVal c0 As Long, _
                          ByVal lft As Double, ByVal tp As Double, ByVal w As Double, _
                          ByVal h As Double, ByVal sfx As String)
    Dim ch As Chart: Set ch = dp_ChartAt(wsC, lft, tp, w, h)
    ch.ChartType = xlColumnClustered
    Dim xr As Range: Set xr = wsD.Cells(2, c0).Resize(24, 1)
    Dim sD As Series, sC As Series, sP As Series
    Set sD = dp_AddSeries(ch, "Discharge freq", xr, wsD.Cells(2, c0 + 1).Resize(24, 1))
    sD.Format.Fill.ForeColor.RGB = dp_StateColor("D")
    Set sC = dp_AddSeries(ch, "Charge freq (neg)", xr, wsD.Cells(2, c0 + 2).Resize(24, 1))
    sC.Format.Fill.ForeColor.RGB = dp_StateColor("C")
    Set sP = dp_AddSeries(ch, "Mean price", xr, wsD.Cells(2, c0 + 3).Resize(24, 1))
    sP.ChartType = xlLine: sP.AxisGroup = xlSecondary
    sP.Format.Line.ForeColor.RGB = RGB(80, 80, 80): sP.Format.Line.Weight = 1.75
    dp_StyleChart ch, "Charge / discharge frequency by hour" & sfx, "Hour ending", "Frequency (share of days)"
    On Error Resume Next
    ch.ChartGroups(1).GapWidth = 40
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0%;0%"
    ch.Axes(xlCategory).Format.Line.Weight = 1.5
    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(90, 90, 90)
    ch.Axes(xlValue, xlSecondary).HasTitle = True
    ch.Axes(xlValue, xlSecondary).AxisTitle.Text = "Price ($/MWh)"
    ch.Axes(xlValue, xlSecondary).AxisTitle.Font.Size = 9
    On Error GoTo 0
End Sub

Private Sub dp_Chart2Net(wsC As Worksheet, wsD As Worksheet, ByVal c0 As Long, _
                         ByVal lft As Double, ByVal tp As Double, ByVal w As Double, _
                         ByVal h As Double, ByVal sfx As String)
    Dim ch As Chart: Set ch = dp_ChartAt(wsC, lft, tp, w, h)
    ch.ChartType = xlColumnClustered
    Dim xr As Range: Set xr = wsD.Cells(2, c0).Resize(24, 1)
    Dim sD As Series, sC As Series, sP As Series
    Set sD = dp_AddSeries(ch, "Net MWh (discharge)", xr, wsD.Cells(2, c0 + 1).Resize(24, 1))
    sD.Format.Fill.ForeColor.RGB = dp_StateColor("D")
    Set sC = dp_AddSeries(ch, "Net MWh (charge)", xr, wsD.Cells(2, c0 + 2).Resize(24, 1))
    sC.Format.Fill.ForeColor.RGB = dp_StateColor("C")
    Set sP = dp_AddSeries(ch, "Mean price", xr, wsD.Cells(2, c0 + 3).Resize(24, 1))
    sP.ChartType = xlLine: sP.AxisGroup = xlSecondary
    sP.Format.Line.ForeColor.RGB = RGB(80, 80, 80): sP.Format.Line.Weight = 1.75
    dp_StyleChart ch, "Mean net MWh by hour (discharge +, charge -)" & sfx, "Hour ending", "Mean net MWh"
    On Error Resume Next
    ch.ChartGroups(1).GapWidth = 40
    ch.Axes(xlValue, xlSecondary).HasTitle = True
    ch.Axes(xlValue, xlSecondary).AxisTitle.Text = "Price ($/MWh)"
    ch.Axes(xlValue, xlSecondary).AxisTitle.Font.Size = 9
    On Error GoTo 0
End Sub

Private Sub dp_Chart3Profiles(wsC As Worksheet, wsD As Worksheet, ByVal c0 As Long, rc As Object, _
                              ByVal lo As Double, ByVal hi As Double, ByVal x0 As Double, _
                              ByVal tp As Double, ByVal sfx As String)
    Dim nTop As Long: nTop = rc("nTop")
    Dim slotDesc() As String: slotDesc = rc("slotDesc")
    Dim slotShare() As Double: slotShare = rc("slotShare")
    Dim slotMeanChg As Variant: slotMeanChg = rc("slotMeanChg")
    Dim slotMeanDis As Variant: slotMeanDis = rc("slotMeanDis")
    Dim slotSpread As Variant: slotSpread = rc("slotSpread")
    Dim slotCycles As Variant: slotCycles = rc("slotCycles")
    Dim piArray As Variant: piArray = rc("piArray")

    Const PW As Double = 300, PH As Double = 290, GUT As Double = 20
    Dim k As Long
    For k = 1 To nTop
        Dim ch As Chart: Set ch = dp_ChartAt(wsC, x0 + (k - 1) * (PW + GUT), tp, PW, PH)
        ch.ChartType = xlColumnClustered
        Dim xr As Range: Set xr = wsD.Cells(2, c0).Resize(24, 1)
        Dim sD As Series, sC As Series
        Set sD = dp_AddSeries(ch, "discharge", xr, wsD.Cells(2, c0 + 2 * k - 1).Resize(24, 1))
        sD.Format.Fill.ForeColor.RGB = dp_PatternColor(k)
        Set sC = dp_AddSeries(ch, "charge", xr, wsD.Cells(2, c0 + 2 * k).Resize(24, 1))
        sC.Format.Fill.ForeColor.RGB = dp_TintColor(dp_PatternColor(k), 0.4)
        dp_StyleChart ch, "P" & k & ": " & slotDesc(k), "Hour ending", "Mean net MWh"
        On Error Resume Next
        ch.ChartGroups(1).Overlap = 100
        ch.ChartGroups(1).GapWidth = 30
        ch.Axes(xlValue).MinimumScale = lo
        ch.Axes(xlValue).MaximumScale = hi
        Dim tb As String
        tb = "n_days=" & piArray(k + 1, 5) & vbCrLf & _
             "share=" & Format$(slotShare(k), "0.0%") & vbCrLf & _
             "chg price=" & dp_FmtNum(slotMeanChg(k)) & vbCrLf & _
             "dis price=" & dp_FmtNum(slotMeanDis(k)) & vbCrLf & _
             "spread=" & dp_FmtNum(slotSpread(k)) & vbCrLf & _
             "cycles=" & dp_FmtNum(slotCycles(k))
        ch.Shapes.AddTextbox(msoTextOrientationHorizontal, 6, 16, 120, 92).TextFrame2.TextRange.Text = tb
        On Error GoTo 0
    Next k
End Sub

Private Sub dp_Chart4PriceDist(wsC As Worksheet, wsD As Worksheet, ByVal c0 As Long, _
                               ByVal lft As Double, ByVal tp As Double, ByVal w As Double, _
                               ByVal h As Double, ByVal sfx As String)
    ' Cols: HE, chgP10, chgBand, disP10, disBand, chgP50, chgP90, disP50, disP90, meanPrice
    Dim ch As Chart: Set ch = dp_ChartAt(wsC, lft, tp, w, h)
    ch.ChartType = xlLine
    Dim xr As Range: Set xr = wsD.Cells(2, c0).Resize(24, 1)

    ' Charge band on the primary axis (stacked area: invisible base + tinted band).
    dp_AddBand ch, xr, wsD.Cells(2, c0 + 1).Resize(24, 1), wsD.Cells(2, c0 + 2).Resize(24, 1), _
               dp_StateColor("C"), xlPrimary
    ' Discharge band on the secondary axis, scale-matched afterwards.
    dp_AddBand ch, xr, wsD.Cells(2, c0 + 3).Resize(24, 1), wsD.Cells(2, c0 + 4).Resize(24, 1), _
               dp_StateColor("D"), xlSecondary

    Dim s As Series
    Set s = dp_AddSeries(ch, "Charge P50", xr, wsD.Cells(2, c0 + 5).Resize(24, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("C")
    Set s = dp_AddSeries(ch, "Charge P10", xr, wsD.Cells(2, c0 + 1).Resize(24, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("C"): s.Format.Line.DashStyle = msoLineDash
    Set s = dp_AddSeries(ch, "Charge P90", xr, wsD.Cells(2, c0 + 6).Resize(24, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("C"): s.Format.Line.DashStyle = msoLineDash
    Set s = dp_AddSeries(ch, "Discharge P50", xr, wsD.Cells(2, c0 + 7).Resize(24, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("D")
    Set s = dp_AddSeries(ch, "Discharge P10", xr, wsD.Cells(2, c0 + 3).Resize(24, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("D"): s.Format.Line.DashStyle = msoLineDash
    Set s = dp_AddSeries(ch, "Discharge P90", xr, wsD.Cells(2, c0 + 8).Resize(24, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("D"): s.Format.Line.DashStyle = msoLineDash
    Set s = dp_AddSeries(ch, "Mean price", xr, wsD.Cells(2, c0 + 9).Resize(24, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = RGB(60, 60, 60): s.Format.Line.Weight = 2

    dp_StyleChart ch, "Price distribution by hour (charge / discharge P10-P50-P90)" & sfx, _
                  "Hour ending", "Price ($/MWh)"
    On Error Resume Next
    ' Match the secondary axis to the primary so the discharge band aligns.
    ch.Axes(xlValue, xlSecondary).MinimumScale = ch.Axes(xlValue, xlPrimary).MinimumScale
    ch.Axes(xlValue, xlSecondary).MaximumScale = ch.Axes(xlValue, xlPrimary).MaximumScale
    ch.Axes(xlValue, xlSecondary).MajorTickMark = xlNone
    ch.Axes(xlValue, xlSecondary).TickLabelPosition = xlNone
    On Error GoTo 0
End Sub

Private Sub dp_Chart5Soc(wsC As Worksheet, wsD As Worksheet, ByVal c0 As Long, ByVal otherRank As Long, _
                         ByVal nTop As Long, ByVal maxSoc As Double, ByVal lft As Double, _
                         ByVal tp As Double, ByVal w As Double, ByVal h As Double, ByVal sfx As String)
    Dim ch As Chart: Set ch = dp_ChartAt(wsC, lft, tp, w, h)
    ch.ChartType = xlLine
    Dim xr As Range: Set xr = wsD.Cells(2, c0).Resize(24, 1)
    Dim rk As Long, s As Series
    For rk = 1 To otherRank
        Set s = dp_AddSeries(ch, dp_SlotLabel(rk, nTop), xr, wsD.Cells(2, c0 + rk).Resize(24, 1))
        s.ChartType = xlLine
        s.Format.Line.ForeColor.RGB = dp_SlotColor(rk, nTop)
        s.Format.Line.Weight = 2
    Next rk
    dp_StyleChart ch, "Mean SOC by hour, per pattern" & sfx, "Hour ending", "State of charge (MWh)"
    On Error Resume Next
    ch.Axes(xlValue).MinimumScale = 0
    If maxSoc > 0 Then ch.Axes(xlValue).MaximumScale = maxSoc
    On Error GoTo 0
End Sub

Private Sub dp_Chart6MonthlyMix(wsC As Worksheet, wsD As Worksheet, ByVal c0 As Long, ByVal nMonths As Long, _
                                ByVal slots As Long, ByVal nTop As Long, ByVal lft As Double, _
                                ByVal tp As Double, ByVal w As Double, ByVal h As Double, ByVal sfx As String)
    Dim ch As Chart: Set ch = dp_ChartAt(wsC, lft, tp, w, h)
    ch.ChartType = xlColumnStacked
    Dim xr As Range: Set xr = wsD.Cells(2, c0).Resize(nMonths, 1)
    Dim rk As Long, s As Series
    For rk = 1 To slots
        Set s = dp_AddSeries(ch, dp_SlotLabel(rk, nTop), xr, wsD.Cells(2, c0 + rk).Resize(nMonths, 1))
        s.Format.Fill.ForeColor.RGB = dp_SlotColor(rk, nTop)
    Next rk
    dp_StyleChart ch, "Monthly pattern mix" & sfx, "Month", "Days"
    On Error Resume Next
    ch.ChartGroups(1).GapWidth = 30
    On Error GoTo 0
End Sub

Private Sub dp_Chart7Drift(wsC As Worksheet, wsD As Worksheet, ByVal c0 As Long, ByVal nMonths As Long, _
                           ByVal lft As Double, ByVal tp As Double, ByVal w As Double, _
                           ByVal h As Double, ByVal sfx As String)
    Dim ch As Chart: Set ch = dp_ChartAt(wsC, lft, tp, w, h)
    ch.ChartType = xlLine
    Dim xr As Range: Set xr = wsD.Cells(2, c0).Resize(nMonths, 1)
    Dim s As Series
    Set s = dp_AddSeries(ch, "First charge HE", xr, wsD.Cells(2, c0 + 1).Resize(nMonths, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("C")
    Set s = dp_AddSeries(ch, "Last charge HE", xr, wsD.Cells(2, c0 + 2).Resize(nMonths, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("C"): s.Format.Line.DashStyle = msoLineDash
    Set s = dp_AddSeries(ch, "First discharge HE", xr, wsD.Cells(2, c0 + 3).Resize(nMonths, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("D")
    Set s = dp_AddSeries(ch, "Last discharge HE", xr, wsD.Cells(2, c0 + 4).Resize(nMonths, 1))
    s.ChartType = xlLine: s.Format.Line.ForeColor.RGB = dp_StateColor("D"): s.Format.Line.DashStyle = msoLineDash
    dp_StyleChart ch, "Dispatch timing drift by month" & sfx, "Month", "Hour ending"
    On Error Resume Next
    ch.Axes(xlValue).MinimumScale = 1
    ch.Axes(xlValue).MaximumScale = 24
    On Error GoTo 0
End Sub

Private Sub dp_Chart8Ranking(wsC As Worksheet, wsD As Worksheet, rc As Object, ByVal lft As Double, _
                             ByVal tp As Double, ByVal w As Double, ByVal h As Double, ByVal sfx As String)
    Dim nTop As Long: nTop = rc("nTop")
    Dim otherRank As Long: otherRank = rc("otherRank")
    Dim slotDesc() As String: slotDesc = rc("slotDesc")
    Dim slotShare() As Double: slotShare = rc("slotShare")
    Dim piArray As Variant: piArray = rc("piArray")

    ' Build (label+descriptor, share, n_days) for top-N + Other, sort by share desc.
    Dim m As Long: m = otherRank
    Dim lab() As String: ReDim lab(1 To m)
    Dim shr() As Double: ReDim shr(1 To m)
    Dim nd() As Long: ReDim nd(1 To m)
    Dim slotIx() As Long: ReDim slotIx(1 To m)
    Dim rk As Long
    For rk = 1 To m
        lab(rk) = dp_SlotLabel(rk, nTop) & ": " & slotDesc(rk)
        shr(rk) = slotShare(rk)
        nd(rk) = CLng(piArray(rk + 1, 5))
        slotIx(rk) = rk
    Next rk
    Dim a As Long, b As Long
    For a = 1 To m - 1
        For b = a + 1 To m
            If shr(b) > shr(a) Then
                dp_SwapS lab, a, b: dp_SwapD shr, a, b: dp_SwapL nd, a, b: dp_SwapL slotIx, a, b
            End If
        Next b
    Next a

    Dim labV As Variant: ReDim labV(1 To m)
    Dim shrV As Variant: ReDim shrV(1 To m)
    For a = 1 To m: labV(a) = lab(a): shrV(a) = shr(a): Next a
    Dim cRank As Long: cRank = 200      ' isolated staging column, well clear of others
    dp_StageBlock wsD, 1, cRank, Array("pattern", "share"), Array(labV, shrV), m

    Dim ch As Chart: Set ch = dp_ChartAt(wsC, lft, tp, w, h)
    ch.ChartType = xlBarClustered
    Dim s As Series
    Set s = dp_AddSeries(ch, "Share of days", wsD.Cells(2, cRank).Resize(m, 1), wsD.Cells(2, cRank + 1).Resize(m, 1))
    dp_StyleChart ch, "Pattern frequency ranking" & sfx, "Share of days", "Pattern"
    On Error Resume Next
    ch.HasLegend = False
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%"
    Dim pt As Point
    For a = 1 To m
        Set pt = s.Points(a)
        pt.Format.Fill.ForeColor.RGB = dp_SlotColor(slotIx(a), nTop)
        pt.HasDataLabel = True
        pt.DataLabel.Text = nd(a) & " (" & Format$(shr(a), "0.0%") & ")"
        pt.DataLabel.Font.Size = 8
    Next a
    On Error GoTo 0
End Sub

Private Sub dp_BuildHeatmap(wsC As Worksheet, rc As Object)
    ' Cell-based heatmap: rows = patterns (top-N, Other, Idle), cols = HE1..24.
    ' Fill = pattern colour (discharge full, charge tinted by within-pattern
    ' frequency, idle white). Text = the pattern's mean price at that HE.
    Dim nTop As Long: nTop = rc("nTop")
    Dim slots As Long: slots = rc("slots")
    Dim modalGrid() As String: modalGrid = rc("modalGrid")
    Dim cCnt() As Long: cCnt = rc("pa_cCnt")
    Dim occ() As Long: occ = rc("pa_occ")
    Dim pxSum() As Double: pxSum = rc("pa_pxSum")

    Dim topR As Long: topR = 100        ' anchor well below the floating charts
    Dim leftC As Long: leftC = 1

    ' Precompute the text array, then write it in one shot.
    Dim txt() As Variant: ReDim txt(1 To slots + 1, 1 To 25)
    txt(1, 1) = "Pattern \ HE"
    Dim hh As Long, rk As Long
    For hh = 1 To 24: txt(1, hh + 1) = "HE" & hh: Next hh
    For rk = 1 To slots
        txt(rk + 1, 1) = dp_SlotLabel(rk, nTop)
        For hh = 1 To 24
            Dim mp As Double: mp = dp_Div(pxSum(rk, hh), occ(rk, hh))
            If occ(rk, hh) > 0 Then txt(rk + 1, hh + 1) = mp Else txt(rk + 1, hh + 1) = ""
        Next hh
    Next rk
    wsC.Cells(topR, leftC).Resize(slots + 1, 25).Value = txt
    wsC.Range(wsC.Cells(topR, leftC), wsC.Cells(topR, leftC + 24)).Font.Bold = True
    wsC.Range(wsC.Cells(topR + 1, leftC), wsC.Cells(topR + slots, leftC)).Font.Bold = True

    ' Colour the value cells.
    Dim st As String, cell As Range, base As Long, tint As Double
    For rk = 1 To slots
        base = dp_SlotColor(rk, nTop)
        For hh = 1 To 24
            Set cell = wsC.Cells(topR + rk, leftC + hh)
            st = modalGrid(rk, hh)
            Select Case st
                Case "D", "B"
                    cell.Interior.Color = base
                Case "C"
                    tint = 0.75 - 0.55 * dp_Div(cCnt(rk, hh), occ(rk, hh))
                    cell.Interior.Color = dp_TintColor(base, tint)
                Case Else
                    cell.Interior.Color = RGB(255, 255, 255)
            End Select
            cell.NumberFormat = "0"
            cell.Font.Color = dp_ContrastFont(cell.Interior.Color)
            cell.HorizontalAlignment = xlCenter
        Next hh
    Next rk

    ' Legend block beneath the grid.
    Dim lr As Long: lr = topR + slots + 2
    wsC.Cells(lr, leftC).Value = "Legend"
    wsC.Cells(lr, leftC).Font.Bold = True
    wsC.Cells(lr + 1, leftC).Value = "Fill = modal state per HE: discharge = solid pattern colour, " & _
        "charge = tinted (lighter = less frequent), idle = white."
    wsC.Cells(lr + 2, leftC).Value = "Cell text = the pattern's mean price at that HE ($/MWh)."
    wsC.Columns.AutoFit
End Sub


'==== SECTION: CHART HELPERS ===================================================

Private Function dp_ChartAt(wsC As Worksheet, ByVal lft As Double, ByVal tp As Double, _
                            ByVal w As Double, ByVal h As Double) As Chart
    Dim co As ChartObject
    Set co = wsC.ChartObjects.Add(lft, tp, w, h)
    Set dp_ChartAt = co.Chart
End Function

Private Function dp_AddSeries(ch As Chart, ByVal nm As String, xRng As Range, yRng As Range) As Series
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.Values = yRng
    s.XValues = xRng
    Set dp_AddSeries = s
End Function

Private Sub dp_StyleChart(ch As Chart, ByVal ttl As String, ByVal xt As String, ByVal yt As String)
    On Error Resume Next
    ch.HasTitle = True
    ch.ChartTitle.Text = ttl
    ch.ChartTitle.Font.Size = 10
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.ChartArea.Format.Line.Visible = msoFalse
    Dim axc As Axis, axv As Axis
    Set axc = ch.Axes(xlCategory)
    axc.HasTitle = True: axc.AxisTitle.Text = xt: axc.AxisTitle.Font.Size = 9
    Set axv = ch.Axes(xlValue)
    axv.HasTitle = True: axv.AxisTitle.Text = yt: axv.AxisTitle.Font.Size = 9
    axv.MajorGridlines.Border.Color = RGB(225, 225, 225)
    On Error GoTo 0
End Sub

' Shaded percentile band via the stacked-area trick: an invisible base series
' at the low percentile plus a tinted band series of (high - low) stacked on it.
Private Sub dp_AddBand(ch As Chart, xRng As Range, loRng As Range, bandRng As Range, _
                       ByVal clr As Long, ByVal grp As XlAxisGroup)
    On Error Resume Next
    Dim sBase As Series, sBand As Series
    Set sBase = ch.SeriesCollection.NewSeries
    sBase.Name = "_base": sBase.Values = loRng: sBase.XValues = xRng
    sBase.ChartType = xlAreaStacked: sBase.AxisGroup = grp
    sBase.Format.Fill.Visible = msoFalse: sBase.Format.Line.Visible = msoFalse
    Set sBand = ch.SeriesCollection.NewSeries
    sBand.Name = "P10-P90 band": sBand.Values = bandRng: sBand.XValues = xRng
    sBand.ChartType = xlAreaStacked: sBand.AxisGroup = grp
    sBand.Format.Fill.ForeColor.RGB = clr: sBand.Format.Fill.Transparency = 0.78
    sBand.Format.Line.Visible = msoFalse
    On Error GoTo 0
End Sub

Private Function dp_BandCol(loCol As Variant, hiCol As Variant) As Variant
    ' (high - low) per HE, or 0 when either bound is blank.
    Dim v As Variant: ReDim v(1 To 24)
    Dim i As Long
    For i = 1 To 24
        If IsNumeric(loCol(i)) And IsNumeric(hiCol(i)) Then
            v(i) = CDbl(hiCol(i)) - CDbl(loCol(i))
        Else
            v(i) = 0
        End If
    Next i
    dp_BandCol = v
End Function

Private Function dp_FmtNum(ByVal v As Variant) As String
    If IsNumeric(v) Then dp_FmtNum = Format$(CDbl(v), "0.00") Else dp_FmtNum = "n/a"
End Function

Private Sub dp_SwapS(arr() As String, ByVal i As Long, ByVal j As Long)
    Dim t As String: t = arr(i): arr(i) = arr(j): arr(j) = t
End Sub
Private Sub dp_SwapD(arr() As Double, ByVal i As Long, ByVal j As Long)
    Dim t As Double: t = arr(i): arr(i) = arr(j): arr(j) = t
End Sub
Private Sub dp_SwapL(arr() As Long, ByVal i As Long, ByVal j As Long)
    Dim t As Long: t = arr(i): arr(i) = arr(j): arr(j) = t
End Sub


'==== SECTION: STAGING HELPERS =================================================

Private Sub dp_StageBlock(ws As Worksheet, ByVal r0 As Long, ByVal c0 As Long, _
                          heads As Variant, cols As Variant, ByVal nRows As Long)
    Dim nCols As Long: nCols = UBound(cols) - LBound(cols) + 1
    Dim o() As Variant: ReDim o(1 To nRows + 1, 1 To nCols)
    Dim j As Long, i As Long
    For j = 1 To nCols
        o(1, j) = heads(LBound(heads) + j - 1)
        Dim col As Variant: col = cols(LBound(cols) + j - 1)
        For i = 1 To nRows
            o(i + 1, j) = col(LBound(col) + i - 1)
        Next i
    Next j
    ws.Cells(r0, c0).Resize(nRows + 1, nCols).Value = o
End Sub

Private Function dp_HeCol() As Variant
    Dim v As Variant: ReDim v(1 To 24)
    Dim i As Long
    For i = 1 To 24: v(i) = i: Next i
    dp_HeCol = v
End Function

Private Function dp_Negate(col As Variant) As Variant
    Dim v As Variant: ReDim v(1 To 24)
    Dim i As Long
    For i = 1 To 24
        If IsNumeric(col(LBound(col) + i - 1)) Then v(i) = -CDbl(col(LBound(col) + i - 1)) Else v(i) = 0
    Next i
    dp_Negate = v
End Function

Private Function dp_ColFrom(o As Variant, ByVal colIx As Long, ByVal r1 As Long, ByVal r2 As Long) As Variant
    Dim v As Variant: ReDim v(1 To r2 - r1 + 1)
    Dim i As Long
    For i = r1 To r2
        v(i - r1 + 1) = o(i, colIx)
    Next i
    dp_ColFrom = v
End Function


'==== SECTION: COLOURS =========================================================
'  Single source of truth so a pattern is the same colour everywhere.

Private Function dp_PatternColor(ByVal rank As Long) As Long
    Select Case rank
        Case 1: dp_PatternColor = RGB(0, 114, 178)     ' P1 blue
        Case 2: dp_PatternColor = RGB(213, 94, 0)      ' P2 vermillion
        Case 3: dp_PatternColor = RGB(0, 158, 115)     ' P3 bluish green
        Case 4: dp_PatternColor = RGB(230, 159, 0)     ' orange (extra ranks)
        Case 5: dp_PatternColor = RGB(204, 121, 167)   ' pink
        Case -2: dp_PatternColor = RGB(220, 220, 220)  ' Idle light grey
        Case Else: dp_PatternColor = CLR_OTHER         ' Other grey RGB(150,150,150)
    End Select
End Function

' State colour for non-pattern charts: discharge blue, charge vermillion.
Private Function dp_StateColor(ByVal stt As String) As Long
    Select Case UCase$(stt)
        Case "D": dp_StateColor = RGB(0, 114, 178)
        Case "C": dp_StateColor = RGB(213, 94, 0)
        Case Else: dp_StateColor = RGB(150, 150, 150)
    End Select
End Function

Private Function dp_TintColor(ByVal baseColor As Long, ByVal pct As Double) As Long
    ' Lighten baseColor toward white by pct (0..1).
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    Dim r As Long, g As Long, b As Long
    r = baseColor Mod 256
    g = (baseColor \ 256) Mod 256
    b = (baseColor \ 65536) Mod 256
    r = r + (255 - r) * pct
    g = g + (255 - g) * pct
    b = b + (255 - b) * pct
    dp_TintColor = RGB(r, g, b)
End Function

Private Function dp_ContrastFont(ByVal bg As Long) As Long
    ' Choose black or white text by luminance of the background.
    Dim r As Long, g As Long, b As Long
    r = bg Mod 256: g = (bg \ 256) Mod 256: b = (bg \ 65536) Mod 256
    Dim lum As Double: lum = 0.299 * r + 0.587 * g + 0.114 * b
    If lum < 140 Then dp_ContrastFont = RGB(255, 255, 255) Else dp_ContrastFont = RGB(0, 0, 0)
End Function


'==== SECTION: NUMERIC HELPERS =================================================

Private Function dp_Num(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        dp_Num = 0
    ElseIf IsNumeric(v) Then
        dp_Num = CDbl(v)
    Else
        dp_Num = 0
    End If
End Function

Private Function dp_Div(ByVal num As Double, ByVal den As Double) As Double
    If den = 0 Then dp_Div = 0 Else dp_Div = num / den
End Function

Private Function dp_Max(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then dp_Max = a Else dp_Max = b
End Function

Private Function dp_Min(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then dp_Min = a Else dp_Min = b
End Function

Private Function dp_ToLong(ByVal v As Variant, ByVal fallback As Long) As Long
    On Error GoTo bad
    If IsNumeric(v) Then dp_ToLong = CLng(v) Else dp_ToLong = fallback
    Exit Function
bad:
    dp_ToLong = fallback
End Function

Private Function dp_ToDouble(ByVal v As Variant, ByVal fallback As Double) As Double
    On Error GoTo bad
    If IsNumeric(v) Then dp_ToDouble = CDbl(v) Else dp_ToDouble = fallback
    Exit Function
bad:
    dp_ToDouble = fallback
End Function

Private Function dp_NormChoice(ByVal v As Variant, ByVal optA As String, _
                               ByVal optB As String, ByVal dflt As String) As String
    Dim s As String: s = LCase$(Trim$(CStr(v)))
    If s = optA Then
        dp_NormChoice = optA
    ElseIf s = optB Then
        dp_NormChoice = optB
    Else
        dp_NormChoice = dflt
    End If
End Function


'==== SECTION: SORT & PERCENTILE ==============================================
'  Own implementations so WorksheetFunction.Percentile is never called in a loop.

Private Sub dp_QuickSortDouble(arr() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, t As Double
    i = lo: j = hi
    p = arr((lo + hi) \ 2)
    Do While i <= j
        Do While arr(i) < p: i = i + 1: Loop
        Do While arr(j) > p: j = j - 1: Loop
        If i <= j Then
            t = arr(i): arr(i) = arr(j): arr(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then dp_QuickSortDouble arr, lo, j
    If i < hi Then dp_QuickSortDouble arr, i, hi
End Sub

Private Function dp_PercentileInc(arr() As Double, ByVal firstIdx As Long, _
                                  ByVal n As Long, ByVal p As Double) As Double
    ' Inclusive percentile (Excel PERCENTILE.INC) over the ascending-sorted slice
    ' arr(firstIdx .. firstIdx + n - 1).
    If n = 1 Then dp_PercentileInc = arr(firstIdx): Exit Function
    Dim rank As Double, k As Long, frac As Double
    rank = p * (n - 1)
    k = Int(rank)
    frac = rank - k
    If k + 1 >= n Then
        dp_PercentileInc = arr(firstIdx + n - 1)
    Else
        dp_PercentileInc = arr(firstIdx + k) + frac * (arr(firstIdx + k + 1) - arr(firstIdx + k))
    End If
End Function

Private Sub dp_SortByCountDesc(keys() As String, cnts() As Long, ByVal lo As Long, ByVal hi As Long)
    ' Sort parallel arrays by cnts descending; ties broken by key ascending
    ' for a deterministic ranking.
    Dim i As Long, j As Long, pc As Long, pk As String, tc As Long, tk As String
    i = lo: j = hi
    pc = cnts((lo + hi) \ 2): pk = keys((lo + hi) \ 2)
    Do While i <= j
        Do While (cnts(i) > pc) Or (cnts(i) = pc And keys(i) < pk): i = i + 1: Loop
        Do While (cnts(j) < pc) Or (cnts(j) = pc And keys(j) > pk): j = j - 1: Loop
        If i <= j Then
            tc = cnts(i): cnts(i) = cnts(j): cnts(j) = tc
            tk = keys(i): keys(i) = keys(j): keys(j) = tk
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then dp_SortByCountDesc keys, cnts, lo, j
    If i < hi Then dp_SortByCountDesc keys, cnts, i, hi
End Sub


'==== SECTION: SHEET / TABLE HELPERS ===========================================

Private Function dp_SheetByName(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then Set dp_SheetByName = ws: Exit Function
    Next ws
    Set dp_SheetByName = Nothing
End Function

' Return an output sheet ready to (re)build, WITHOUT ever deleting a worksheet.
' Deleting sheets trips workbook-structure protection, the last-visible-sheet
' rule, stale references and collection re-indexing - so instead we create the
' sheet on first use and, on later runs, strip its contents in place.
Private Function dp_GetOrResetSheet(ByVal sName As String) As Worksheet
    Dim ws As Worksheet, lo As ListObject, co As ChartObject
    Dim i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:= _
            ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sName
    Else
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
        If ws.ProtectContents Then ws.Unprotect
        For Each co In ws.ChartObjects
            co.Delete
        Next co
        ' Backward: forward iteration skips elements as the collection re-indexes.
        For i = ws.Shapes.Count To 1 Step -1
            ws.Shapes(i).Delete
        Next i
        For Each lo In ws.ListObjects
            lo.Unlist
        Next lo
        ws.Cells.Clear
        ws.Cells.Interior.Pattern = xlNone
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
    End If
    Set dp_GetOrResetSheet = ws
End Function

Private Sub dp_MakeTable(ws As Worksheet, ByVal tblName As String, ByVal nRows As Long, ByVal nCols As Long)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(nRows, nCols), , xlYes)
    lo.Name = tblName
    lo.TableStyle = "TableStyleMedium2"
    ws.Rows(1).Font.Bold = True
    On Error GoTo 0
    ' NB: FreezePanes targets only the ActiveWindow's active sheet, which would
    ' require .Activate - deliberately avoided per the no-Activate rule. The
    ' ListObject header row provides sort/filter in its place.
End Sub

Private Sub dp_FormatCols(ws As Worksheet, cols As Variant, ByVal fmt As String)
    Dim i As Long, c As Long
    For i = LBound(cols) To UBound(cols)
        c = cols(i)
        ws.Range(ws.Cells(2, c), ws.Cells(ws.Rows.Count, c).End(xlUp)).NumberFormat = fmt
    Next i
End Sub


'==== SECTION: SUMMARY DIALOG ==================================================

Private Sub dp_ShowSummary(rc As Object, cfg As Object)
    Dim piArray As Variant: piArray = rc("piArray")
    Dim nTop As Long: nTop = rc("nTop")
    Dim slotDesc() As String: slotDesc = rc("slotDesc")

    Dim msg As String
    msg = "Dispatch pattern analysis complete." & vbCrLf & _
          "Method: " & cfg("method") & "   Price: " & cfg("priceHdr") & vbCrLf & _
          "Top-N coverage: " & Format$(rc("coverage"), "0.0%") & vbCrLf & String(42, "-") & vbCrLf

    Dim k As Long
    For k = 1 To nTop
        msg = msg & "P" & k & "  " & slotDesc(k) & vbCrLf & _
              "     days=" & piArray(k + 1, 5) & _
              "  share=" & Format$(piArray(k + 1, 6), "0.0%") & _
              "  chg price=" & dp_FmtNum(piArray(k + 1, 9)) & _
              "  dis price=" & dp_FmtNum(piArray(k + 1, 10)) & _
              "  spread=" & dp_FmtNum(piArray(k + 1, 11)) & vbCrLf
    Next k

    msg = msg & String(42, "-") & vbCrLf & _
          "Rows kept: " & mLog("rows_kept") & "   filtered: " & mLog("rows_filtered") & vbCrLf & _
          "Days complete: " & mLog("days_complete") & "   excluded: " & mLog("days_excluded_incomplete") & vbCrLf & _
          "'B' (charge+discharge) hours: " & mLog("B_hour_count") & vbCrLf & _
          "Text timestamps coerced: " & mLog("text_timestamps_coerced")

    Dim nFail As Long: nFail = 0
    If rc.Exists("chart_fail_count") Then nFail = CLng(rc("chart_fail_count"))
    If nFail > 0 Then
        msg = msg & vbCrLf & "Charts failed: " & nFail & " (see RunLog for details)"
    End If

    If rc("recommendWindow") Then
        msg = msg & vbCrLf & vbCrLf & ">> Exact coverage < 50%. Consider cfgClusterMethod = Window."
    End If

    MsgBox msg, vbInformation, "modDispatchPatterns"
End Sub
