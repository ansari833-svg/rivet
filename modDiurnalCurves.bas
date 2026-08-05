Attribute VB_Name = "modDiurnalCurves"
Option Explicit

' ============================================================================
'  modDiurnalCurves
'  ---------------------------------------------------------------------------
'  Reads a 5-year hourly day-ahead price timeseries from the `prices` sheet and
'  produces one chart per calendar year, each with 12 lines (the average
'  diurnal price curve for every month of that year), plus an all-years overlay
'  and a year-over-year drift chart.
'
'  Public entry point:  RunDiurnalAnalysis
'  All private helpers are prefixed  dn_
'
'  Design notes
'   - Single array read per source sheet (.Value2), single array write per
'     output sheet. No cell-by-cell loops over the data.
'   - Long (never Integer) for every row counter / index.
'   - Late-bound Scripting.Dictionary via CreateObject (no project reference).
'   - App-state is saved on entry and always restored through one error handler
'     that reports the failing step via the module-level mProc string.
'   - No .Select / .Activate anywhere.
'   - No identifier collides (case-insensitively) with a VBA reserved word or
'     built-in function: month/year/hour/date/min/max are spelled
'     lMonthNum / lYearNum / lHourEnding / dMinPrice / dMaxPrice, etc.
' ============================================================================

' ---- module-level shared state --------------------------------------------
Private mProc      As String            ' name of the step currently running
Private mLog       As Object            ' Collection of RunLog lines

' config (populated by dn_ReadConfig)
Private mCfgPriceSheet   As String
Private mCfgPriceColName As String
Private mCfgHeaderRow    As Long
Private mCfgFirstDataRow As Long
Private mCfgHEConv       As String
Private mCfgAggMethod    As String
Private mCfgSimID        As String
Private mCfgYAxisMode    As String
Private mCfgClipOutliers As Boolean
Private mCfgShowBands    As Boolean

' loaded observations (parallel arrays, 1..mObs)
Private mObsYear()  As Long
Private mObsMonth() As Long
Private mObsHE()    As Long
Private mObsPrice() As Double
Private mObs        As Long

' discovered years
Private mYears()    As Long
Private mNYears     As Long
Private mYearIndex  As Object            ' Dictionary lYearNum -> 1-based index

' aggregated results, dimensioned (1..mNYears, 1..12, 1..24)
Private mAgg()      As Double            ' Mean or Median per (year, month, HE)
Private mCnt()      As Long             ' observation count per cell
Private mBucket()   As Variant           ' per-cell Double() values (median/bands)

' reporting counters
Private mRowsLoaded    As Long
Private mRowsFiltered  As Long
Private mCoerced       As Long
Private mWinsorized    As Long
Private mEmptyCells    As Long
Private mIncompMonths  As Long
Private mHdrRow        As Long
Private mFirstDataRow  As Long
Private mRowsSkipped   As Long
Private mHasSim        As Boolean
Private mRawMin        As Double
Private mRawMax        As Double
Private mAggMin        As Double
Private mAggMax        As Double

' ============================================================================
'  SECTION: PUBLIC ENTRY POINT
' ============================================================================

Public Sub RunDiurnalAnalysis()

    ' error/state variables declared once, up top
    Dim lErrNum      As Long
    Dim sErrDesc     As String
    Dim bScreen      As Boolean
    Dim bEvents      As Boolean
    Dim bAlerts      As Boolean
    Dim lCalcSave    As Long
    Dim dStart       As Double

    ' capture and set application state
    bScreen = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    bAlerts = Application.DisplayAlerts
    lCalcSave = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Handler

    dStart = Timer
    Set mLog = CreateObject("Scripting.Dictionary")   ' used as an ordered list
    mLog.RemoveAll

    mProc = "EnsureConfig"
    dn_EnsureConfigSheet

    mProc = "ReadConfig"
    dn_ReadConfig

    mProc = "LoadData"
    dn_LoadData

    mProc = "Winsorize"
    If mCfgClipOutliers Then dn_Winsorize

    mProc = "Aggregate"
    dn_Aggregate

    mProc = "WriteDiurnalData"
    dn_WriteDiurnalData

    mProc = "BuildCharts"
    dn_BuildCharts

    mProc = "WriteRunLog"
    dn_LogLine "Runtime (s):|" & Format$(Timer - dStart, "0.00")
    dn_WriteRunLog

    mProc = "Summary"
    dn_ShowSummary

    ' restore state, then leave BEFORE the handler label
    Application.Calculation = lCalcSave
    Application.DisplayAlerts = bAlerts
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreen
    Exit Sub

Handler:
    lErrNum = Err.Number
    sErrDesc = Err.Description

    ' always restore application state
    On Error Resume Next
    Application.Calculation = lCalcSave
    Application.DisplayAlerts = bAlerts
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreen

    ' record the trapped error in RunLog if we can
    If Not mLog Is Nothing Then
        dn_LogLine "TRAPPED ERROR in [" & mProc & "]:|" & sErrDesc & " (#" & lErrNum & ")"
        dn_WriteRunLog
    End If
    On Error GoTo 0

    MsgBox "RunDiurnalAnalysis failed in step [" & mProc & "]:" & vbCrLf & vbCrLf & _
           sErrDesc & " (error " & lErrNum & ")", vbCritical, "Diurnal Curves"
End Sub

' ============================================================================
'  SECTION: CONFIG SHEET (build / refresh)
' ============================================================================

Private Sub dn_EnsureConfigSheet()
    Dim ws As Worksheet
    Dim bNew As Boolean

    Set ws = dn_SheetByName("ConfigDiurnal")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "ConfigDiurnal"
        bNew = True
    End If

    ' title
    ws.Range("A1").Value = "Diurnal Curve Analysis - Configuration"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:B1").Interior.Color = RGB(230, 235, 245)

    ' labels + defaults (only write defaults when the sheet is new or blank)
    dn_CfgRow ws, 2, "cfgPriceSheet", "prices", bNew, "source worksheet"
    dn_CfgRow ws, 3, "cfgPriceColName", "energy_da", bNew, "header name of the price column"
    dn_CfgRow ws, 4, "cfgHeaderRow", 0, bNew, "0 = auto-detect; else row holding column names"
    dn_CfgRow ws, 5, "cfgFirstDataRow", 0, bNew, "0 = auto-detect; else first row of data"
    dn_CfgRow ws, 6, "cfgHEConvention", "Ending", bNew, "Ending (00:00->HE1) / Beginning"
    dn_CfgRow ws, 7, "cfgAggMethod", "Mean", bNew, "Mean / Median"
    dn_CfgRow ws, 8, "cfgSimID", "", bNew, "blank = pool all; ignored if no simulation_id column"
    dn_CfgRow ws, 9, "cfgYAxisMode", "Common", bNew, "Common (same scale on all 5) / PerYear"
    dn_CfgRow ws, 10, "cfgClipOutliers", False, bNew, "TRUE = winsorize at P1/P99 before averaging"
    dn_CfgRow ws, 11, "cfgShowP10P90", False, bNew, "TRUE = also emit seasonal band charts"

    ' dropdown validations
    dn_AddList ws.Range("B6"), "Ending,Beginning"
    dn_AddList ws.Range("B7"), "Mean,Median"
    dn_AddList ws.Range("B9"), "Common,PerYear"
    dn_AddList ws.Range("B10"), "TRUE,FALSE"
    dn_AddList ws.Range("B11"), "TRUE,FALSE"

    ws.Columns("A:C").AutoFit

    ' run button (remove any prior one first so refresh does not stack them)
    Dim shp As Object
    For Each shp In ws.Buttons
        shp.Delete
    Next shp
    Dim btn As Object
    Set btn = ws.Buttons.Add(ws.Range("A13").Left, ws.Range("A13").Top, 170, 28)
    btn.Caption = "Run Diurnal Analysis"
    btn.OnAction = "RunDiurnalAnalysis"
End Sub

' write one labeled config row and give the value cell a workbook-scoped name
Private Sub dn_CfgRow(ws As Worksheet, lRow As Long, sName As String, _
                      vDefault As Variant, bWriteDefault As Boolean, sNote As String)
    ws.Cells(lRow, 1).Value = sName
    If bWriteDefault Or Len(CStr(ws.Cells(lRow, 2).Value)) = 0 Then
        ws.Cells(lRow, 2).Value = vDefault
    End If
    ws.Cells(lRow, 3).Value = sNote
    ws.Cells(lRow, 3).Font.Italic = True
    ws.Cells(lRow, 3).Font.Color = RGB(120, 120, 120)

    ' (re)create the workbook-scoped defined name
    On Error Resume Next
    ThisWorkbook.Names(sName).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=sName, _
        RefersTo:="='" & ws.Name & "'!" & ws.Cells(lRow, 2).Address(True, True)
End Sub

Private Sub dn_AddList(rng As Range, sList As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                       Operator:=xlBetween, Formula1:=sList
    On Error GoTo 0
End Sub

' ============================================================================
'  SECTION: CONFIG READ + VALIDATE
' ============================================================================

Private Sub dn_ReadConfig()
    mCfgPriceSheet = dn_NameStr("cfgPriceSheet", "prices")
    mCfgPriceColName = dn_NameStr("cfgPriceColName", "energy_da")
    mCfgHeaderRow = dn_NameLng("cfgHeaderRow", 0)
    mCfgFirstDataRow = dn_NameLng("cfgFirstDataRow", 0)
    mCfgHEConv = dn_NameStr("cfgHEConvention", "Ending")
    mCfgAggMethod = dn_NameStr("cfgAggMethod", "Mean")
    mCfgSimID = dn_NameStr("cfgSimID", "")
    mCfgYAxisMode = dn_NameStr("cfgYAxisMode", "Common")
    mCfgClipOutliers = dn_NameBool("cfgClipOutliers", False)
    mCfgShowBands = dn_NameBool("cfgShowP10P90", False)

    ' --- validate ---
    If dn_SheetByName(mCfgPriceSheet) Is Nothing Then
        dn_Fail "Source sheet '" & mCfgPriceSheet & "' (cfgPriceSheet) was not found."
    End If
    If Len(Trim$(mCfgPriceColName)) = 0 Then
        dn_Fail "cfgPriceColName is blank - it must name the price column header (e.g. energy_da)."
    End If
    If UCase$(mCfgAggMethod) <> "MEAN" And UCase$(mCfgAggMethod) <> "MEDIAN" Then
        dn_Fail "cfgAggMethod must be Mean or Median (found '" & mCfgAggMethod & "')."
    End If
    If UCase$(mCfgHEConv) <> "ENDING" And UCase$(mCfgHEConv) <> "BEGINNING" Then
        dn_Fail "cfgHEConvention must be Ending or Beginning (found '" & mCfgHEConv & "')."
    End If
    If UCase$(mCfgYAxisMode) <> "COMMON" And UCase$(mCfgYAxisMode) <> "PERYEAR" Then
        dn_Fail "cfgYAxisMode must be Common or PerYear (found '" & mCfgYAxisMode & "')."
    End If
    If mCfgHeaderRow < 0 Or mCfgFirstDataRow < 0 Then
        dn_Fail "cfgHeaderRow / cfgFirstDataRow must be 0 (auto) or a positive row number."
    End If
    If mCfgHeaderRow > 0 And mCfgFirstDataRow > 0 And mCfgFirstDataRow <= mCfgHeaderRow Then
        dn_Fail "cfgFirstDataRow (" & mCfgFirstDataRow & ") must be below cfgHeaderRow (" & _
                mCfgHeaderRow & ")."
    End If

    dn_LogLine "Config echo: sheet|" & mCfgPriceSheet
    dn_LogLine "Config echo: priceCol|" & mCfgPriceColName
    dn_LogLine "Config echo: HEConvention|" & mCfgHEConv
    dn_LogLine "Config echo: aggMethod|" & mCfgAggMethod
    dn_LogLine "Config echo: yAxisMode|" & mCfgYAxisMode
    dn_LogLine "Config echo: clipOutliers|" & CStr(mCfgClipOutliers)
    dn_LogLine "Config echo: showP10P90|" & CStr(mCfgShowBands)
    dn_LogLine "Config echo: simID|" & IIf(Len(mCfgSimID) = 0, "(blank)", mCfgSimID)
End Sub

' ============================================================================
'  SECTION: DATA LOAD  (header detection + single array read)
' ============================================================================

Private Sub dn_LoadData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mCfgPriceSheet)

    ' --- one read of the whole used range ---
    Dim rngU As Range
    Set rngU = ws.UsedRange
    Dim lBaseRow As Long, lBaseCol As Long
    lBaseRow = rngU.Row
    lBaseCol = rngU.Column

    Dim vAll As Variant
    vAll = rngU.Value2                     ' 1-based Variant array, no re-reads after this
    If Not IsArray(vAll) Then dn_Fail "Sheet '" & mCfgPriceSheet & "' appears to be empty."

    Dim lNRows As Long, lNCols As Long
    lNRows = UBound(vAll, 1)
    lNCols = UBound(vAll, 2)

    ' --- 1) find the header row (scan sheet rows 1..10) ---
    Dim sPriceHdr As String: sPriceHdr = LCase$(Trim$(mCfgPriceColName))
    Dim lHdrSheet As Long: lHdrSheet = 0
    Dim lHdrAi    As Long
    Dim sr As Long, c As Long, ai As Long, s As String

    If mCfgHeaderRow > 0 Then
        lHdrSheet = mCfgHeaderRow           ' explicit override
    Else
        For sr = 1 To 10
            ai = sr - lBaseRow + 1
            If ai >= 1 And ai <= lNRows Then
                Dim bTs As Boolean, bPr As Boolean
                bTs = False: bPr = False
                For c = 1 To lNCols
                    s = LCase$(Trim$(CStr(vAll(ai, c))))
                    If s = "timestamp" Then bTs = True
                    If s = sPriceHdr Then bPr = True
                Next c
                If bTs And bPr Then lHdrSheet = sr: Exit For
            End If
        Next sr
    End If

    If lHdrSheet = 0 Then
        dn_Fail "Could not find a header row containing both 'timestamp' and '" & _
                mCfgPriceColName & "' in rows 1-10 of '" & mCfgPriceSheet & "'." & vbCrLf & _
                vbCrLf & "Row 1: " & dn_RowDump(vAll, 1 - lBaseRow + 1, lNRows, lNCols) & vbCrLf & _
                "Row 2: " & dn_RowDump(vAll, 2 - lBaseRow + 1, lNRows, lNCols)
    End If
    lHdrAi = lHdrSheet - lBaseRow + 1
    If lHdrAi < 1 Or lHdrAi > lNRows Then
        dn_Fail "cfgHeaderRow (" & lHdrSheet & ") is outside the used range of '" & mCfgPriceSheet & "'."
    End If

    ' --- map header names -> array column indexes ---
    Dim dHdr As Object: Set dHdr = CreateObject("Scripting.Dictionary")
    dHdr.CompareMode = 1                    ' TextCompare (case-insensitive)
    For c = 1 To lNCols
        s = Trim$(CStr(vAll(lHdrAi, c)))
        If Len(s) > 0 And Not dHdr.Exists(s) Then dHdr.Add s, c
    Next c

    If Not dHdr.Exists("timestamp") Then dn_Fail "Header 'timestamp' not found in header row " & lHdrSheet & "."
    If Not dHdr.Exists(mCfgPriceColName) Then _
        dn_Fail "Price header '" & mCfgPriceColName & "' not found in header row " & lHdrSheet & "."

    Dim lTsCol As Long, lPrCol As Long, lSimCol As Long
    lTsCol = dHdr("timestamp")
    lPrCol = dHdr(mCfgPriceColName)
    mHasSim = dHdr.Exists("simulation_id")
    If mHasSim Then lSimCol = dHdr("simulation_id")

    ' --- 2) first data row: first row below header whose timestamp is a real date ---
    Dim lFirstSheet As Long: lFirstSheet = 0
    If mCfgFirstDataRow > 0 Then
        lFirstSheet = mCfgFirstDataRow
    Else
        For ai = lHdrAi + 1 To lNRows
            If dn_LooksLikeDate(vAll(ai, lTsCol)) Then
                lFirstSheet = lBaseRow + ai - 1
                Exit For
            End If
        Next ai
    End If
    If lFirstSheet = 0 Then dn_Fail "No data row with a valid timestamp was found below header row " & lHdrSheet & "."

    ' --- guard the last row via the timestamp column's last populated cell ---
    Dim lTsColAbs As Long: lTsColAbs = lBaseCol + lTsCol - 1
    Dim lLastSheet As Long
    lLastSheet = ws.Cells(ws.Rows.Count, lTsColAbs).End(xlUp).Row
    Dim lUsedLast As Long: lUsedLast = lBaseRow + lNRows - 1
    If lLastSheet > lUsedLast Then lLastSheet = lUsedLast

    mHdrRow = lHdrSheet
    mFirstDataRow = lFirstSheet
    mRowsSkipped = mFirstDataRow - mHdrRow - 1
    dn_LogLine "Header row detected|" & mHdrRow
    dn_LogLine "First data row detected|" & mFirstDataRow
    dn_LogLine "Rows skipped in header block|" & mRowsSkipped

    ' --- iterate rows, deriving year/month/HE straight from the serial ---
    Dim lFirstAi As Long, lLastAi As Long
    lFirstAi = lFirstSheet - lBaseRow + 1
    lLastAi = lLastSheet - lBaseRow + 1
    If lFirstAi < 1 Then lFirstAi = 1
    If lLastAi > lNRows Then lLastAi = lNRows

    Dim lCap As Long: lCap = lLastAi - lFirstAi + 1
    If lCap < 1 Then dn_Fail "No data rows between the header block and the last timestamp."
    ReDim mObsYear(1 To lCap)
    ReDim mObsMonth(1 To lCap)
    ReDim mObsHE(1 To lCap)
    ReDim mObsPrice(1 To lCap)

    mObs = 0
    mRowsFiltered = 0
    mCoerced = 0
    mRawMin = 1E+308
    mRawMax = -1E+308

    Dim vTs As Variant, vPr As Variant
    Dim dSerial As Double, dPrice As Double
    Dim dFrac As Double, lHourRaw As Long
    Dim lYearNum As Long, lMonthNum As Long, lHE As Long
    Dim bUseSimFilter As Boolean
    bUseSimFilter = (mHasSim And Len(Trim$(mCfgSimID)) > 0)

    For ai = lFirstAi To lLastAi
        vTs = vAll(ai, lTsCol)
        vPr = vAll(ai, lPrCol)

        ' timestamp -> serial (Double), coercing text once if needed
        If IsEmpty(vTs) Or (VarType(vTs) = vbString And Len(Trim$(CStr(vTs))) = 0) Then
            mRowsFiltered = mRowsFiltered + 1
            GoTo NextRow
        ElseIf IsNumeric(vTs) Then
            dSerial = CDbl(vTs)
            If dSerial <= 1 Then mRowsFiltered = mRowsFiltered + 1: GoTo NextRow
        ElseIf VarType(vTs) = vbString Then
            If IsDate(vTs) Then
                dSerial = CDbl(CDate(vTs))
                mCoerced = mCoerced + 1
            Else
                mRowsFiltered = mRowsFiltered + 1
                GoTo NextRow
            End If
        Else
            mRowsFiltered = mRowsFiltered + 1
            GoTo NextRow
        End If

        ' price must be numeric
        If Not IsNumeric(vPr) Then mRowsFiltered = mRowsFiltered + 1: GoTo NextRow
        dPrice = CDbl(vPr)

        ' optional simulation filter
        If bUseSimFilter Then
            If CStr(vAll(ai, lSimCol)) <> mCfgSimID Then
                mRowsFiltered = mRowsFiltered + 1
                GoTo NextRow
            End If
        End If

        ' derive fields from the serial
        dFrac = dSerial - Int(dSerial)
        lHourRaw = Int(dFrac * 24# + 0.0000001)
        If lHourRaw < 0 Then lHourRaw = 0
        If lHourRaw > 23 Then lHourRaw = 23
        lYearNum = VBA.Year(CDate(dSerial))
        lMonthNum = VBA.Month(CDate(dSerial))
        lHE = dn_HourToHE(lHourRaw)

        mObs = mObs + 1
        mObsYear(mObs) = lYearNum
        mObsMonth(mObs) = lMonthNum
        mObsHE(mObs) = lHE
        mObsPrice(mObs) = dPrice
        If dPrice < mRawMin Then mRawMin = dPrice
        If dPrice > mRawMax Then mRawMax = dPrice
NextRow:
    Next ai

    If mObs = 0 Then dn_Fail "No valid observations parsed from '" & mCfgPriceSheet & "'."
    mRowsLoaded = mObs

    dn_LogLine "Rows loaded (valid obs)|" & mRowsLoaded
    dn_LogLine "Rows filtered out|" & mRowsFiltered
    dn_LogLine "Hours coerced from text|" & mCoerced
    dn_LogLine "Raw price min|" & Format$(mRawMin, "0.0000")
    dn_LogLine "Raw price max|" & Format$(mRawMax, "0.0000")
End Sub

' Ending: 00:00 -> HE1 ... 23:00 -> HE24.  Data is interval-start-stamped, so the
' Beginning convention buckets identically (the timestamp already marks the start);
' the convention drives only the axis label (see dn_BuildCharts). Kept explicit so
' end-stamped data could be shifted here if ever needed.
Private Function dn_HourToHE(ByVal lHourRaw As Long) As Long
    dn_HourToHE = lHourRaw + 1
End Function

' ============================================================================
'  SECTION: WINSORIZE (optional P1/P99 clamp)
' ============================================================================

Private Sub dn_Winsorize()
    Dim dCopy() As Double
    ReDim dCopy(1 To mObs)
    Dim i As Long
    For i = 1 To mObs
        dCopy(i) = mObsPrice(i)
    Next i
    dn_QuickSortDouble dCopy, 1, mObs

    Dim dP1 As Double, dP99 As Double
    dP1 = dn_PercentileInc(dCopy, mObs, 0.01)
    dP99 = dn_PercentileInc(dCopy, mObs, 0.99)

    mWinsorized = 0
    For i = 1 To mObs
        If mObsPrice(i) < dP1 Then
            mObsPrice(i) = dP1: mWinsorized = mWinsorized + 1
        ElseIf mObsPrice(i) > dP99 Then
            mObsPrice(i) = dP99: mWinsorized = mWinsorized + 1
        End If
    Next i

    dn_LogLine "Winsorized at P1|" & Format$(dP1, "0.0000") & " P99|" & Format$(dP99, "0.0000")
    dn_LogLine "Hours winsorized|" & mWinsorized
End Sub

' ============================================================================
'  SECTION: AGGREGATE  ([year, month, HE] mean or median)
' ============================================================================

Private Sub dn_Aggregate()
    ' --- discover distinct years ---
    Set mYearIndex = CreateObject("Scripting.Dictionary")
    Dim i As Long, lYearNum As Long
    For i = 1 To mObs
        lYearNum = mObsYear(i)
        If Not mYearIndex.Exists(lYearNum) Then mYearIndex.Add lYearNum, 0
    Next i

    ' sort the discovered years ascending
    mNYears = mYearIndex.Count
    ReDim mYears(1 To mNYears)
    Dim vKeys As Variant: vKeys = mYearIndex.Keys
    Dim dTmp() As Double: ReDim dTmp(1 To mNYears)
    For i = 0 To mNYears - 1
        dTmp(i + 1) = CDbl(vKeys(i))
    Next i
    dn_QuickSortDouble dTmp, 1, mNYears
    For i = 1 To mNYears
        mYears(i) = CLng(dTmp(i))
        mYearIndex(mYears(i)) = i          ' year -> 1-based index
    Next i
    dn_LogLine "Distinct years found|" & mNYears & " (" & dn_JoinYears() & ")"

    ' --- accumulate sum/count always ---
    ReDim mCnt(1 To mNYears, 1 To 12, 1 To 24)
    ReDim mAgg(1 To mNYears, 1 To 12, 1 To 24)
    Dim dSum() As Double: ReDim dSum(1 To mNYears, 1 To 12, 1 To 24)

    Dim yi As Long, lM As Long, lHE As Long
    For i = 1 To mObs
        yi = mYearIndex(mObsYear(i))
        lM = mObsMonth(i)
        lHE = mObsHE(i)
        mCnt(yi, lM, lHE) = mCnt(yi, lM, lHE) + 1
        dSum(yi, lM, lHE) = dSum(yi, lM, lHE) + mObsPrice(i)
    Next i

    Dim bMedian As Boolean: bMedian = (UCase$(mCfgAggMethod) = "MEDIAN")
    Dim bNeedBuckets As Boolean: bNeedBuckets = (bMedian Or mCfgShowBands)

    If bNeedBuckets Then dn_CollectBuckets

    ' --- reduce to the aggregated value per cell ---
    Dim y As Long, m As Long, h As Long
    Dim dArr() As Double, lN As Long
    For y = 1 To mNYears
        For m = 1 To 12
            For h = 1 To 24
                If mCnt(y, m, h) = 0 Then
                    mAgg(y, m, h) = 0
                ElseIf bMedian Then
                    dArr = mBucket(y, m, h)
                    lN = mCnt(y, m, h)
                    dn_QuickSortDouble dArr, 1, lN
                    mAgg(y, m, h) = dn_PercentileInc(dArr, lN, 0.5)
                Else
                    mAgg(y, m, h) = dSum(y, m, h) / mCnt(y, m, h)
                End If
            Next h
        Next m
    Next y

    ' --- quality reporting ---
    mEmptyCells = 0
    mIncompMonths = 0
    Dim lDistinctHE As Long
    For y = 1 To mNYears
        For m = 1 To 12
            lDistinctHE = 0
            For h = 1 To 24
                If mCnt(y, m, h) = 0 Then
                    mEmptyCells = mEmptyCells + 1
                Else
                    lDistinctHE = lDistinctHE + 1
                End If
            Next h
            If lDistinctHE > 0 And lDistinctHE < 24 Then
                mIncompMonths = mIncompMonths + 1
                dn_LogLine "Incomplete month|" & mYears(y) & "-" & dn_MonthName(m) & _
                           " has " & lDistinctHE & "/24 HE"
            End If
        Next m
    Next y
    dn_LogLine "Empty (year,month,HE) cells|" & mEmptyCells
    dn_LogLine "Incomplete months (<24 HE)|" & mIncompMonths

    ' --- global aggregated min/max for common axis scaling ---
    mAggMin = 1E+308: mAggMax = -1E+308
    For y = 1 To mNYears
        For m = 1 To 12
            For h = 1 To 24
                If mCnt(y, m, h) > 0 Then
                    If mAgg(y, m, h) < mAggMin Then mAggMin = mAgg(y, m, h)
                    If mAgg(y, m, h) > mAggMax Then mAggMax = mAgg(y, m, h)
                End If
            Next h
        Next m
    Next y
    dn_LogLine "Aggregated curve min|" & Format$(mAggMin, "0.0000")
    dn_LogLine "Aggregated curve max|" & Format$(mAggMax, "0.0000")
End Sub

' two-pass exact allocation of per-cell value arrays (no ReDim Preserve churn)
Private Sub dn_CollectBuckets()
    ReDim mBucket(1 To mNYears, 1 To 12, 1 To 24)
    Dim y As Long, m As Long, h As Long
    For y = 1 To mNYears
        For m = 1 To 12
            For h = 1 To 24
                If mCnt(y, m, h) > 0 Then
                    Dim d() As Double
                    ReDim d(1 To mCnt(y, m, h))
                    mBucket(y, m, h) = d
                End If
            Next h
        Next m
    Next y

    Dim lFill() As Long: ReDim lFill(1 To mNYears, 1 To 12, 1 To 24)
    Dim i As Long, yi As Long, lM As Long, lHE As Long
    Dim dTmp() As Double
    For i = 1 To mObs
        yi = mYearIndex(mObsYear(i))
        lM = mObsMonth(i)
        lHE = mObsHE(i)
        lFill(yi, lM, lHE) = lFill(yi, lM, lHE) + 1
        dTmp = mBucket(yi, lM, lHE)
        dTmp(lFill(yi, lM, lHE)) = mObsPrice(i)
        mBucket(yi, lM, lHE) = dTmp
    Next i
End Sub

' ============================================================================
'  SECTION: OUTPUT SHEET  DiurnalData  (single array write)
' ============================================================================

Private Sub dn_WriteDiurnalData()
    Dim ws As Worksheet
    Set ws = dn_FreshSheet("DiurnalData")

    Const NCOL As Long = 13                 ' HE label + 12 months

    ' pre-compute layout: per-year (27 rows) + all-years (27) + summary
    Dim lSummaryRows As Long: lSummaryRows = mNYears * 12
    Dim lTotal As Long
    lTotal = mNYears * 27 + 27 + (2 + lSummaryRows) + 2
    Dim vOut() As Variant
    ReDim vOut(1 To lTotal, 1 To NCOL)

    ' record ListObject placements to build after the write
    Dim colTblTop As Object: Set colTblTop = CreateObject("Scripting.Dictionary")

    Dim r As Long: r = 0
    Dim y As Long, m As Long, h As Long

    ' ---- one table per year ----
    For y = 1 To mNYears
        r = r + 1
        vOut(r, 1) = "Year " & mYears(y) & " - Average Diurnal Day-Ahead Price ($/MWh) [" & mCfgAggMethod & "]"
        r = r + 1                            ' header row of the ListObject
        Dim lHdrRowData As Long: lHdrRowData = r
        vOut(r, 1) = "HE"
        For m = 1 To 12
            vOut(r, 1 + m) = dn_MonthName(m)
        Next m
        For h = 1 To 24
            r = r + 1
            vOut(r, 1) = "HE" & h
            For m = 1 To 12
                If mCnt(y, m, h) > 0 Then vOut(r, 1 + m) = mAgg(y, m, h) Else vOut(r, 1 + m) = Empty
            Next m
        Next h
        colTblTop.Add "tblYear_" & mYears(y), lHdrRowData & "|" & NCOL
        r = r + 1                            ' spacer
    Next y

    ' ---- all-years reference table ----
    r = r + 1
    vOut(r, 1) = "All Years - Reference Diurnal Curve (mean of yearly curves)"
    r = r + 1
    Dim lHdrAllYears As Long: lHdrAllYears = r
    vOut(r, 1) = "HE"
    For m = 1 To 12
        vOut(r, 1 + m) = dn_MonthName(m)
    Next m
    For h = 1 To 24
        r = r + 1
        vOut(r, 1) = "HE" & h
        For m = 1 To 12
            vOut(r, 1 + m) = dn_AllYearsValue(m, h)
        Next m
    Next h
    colTblTop.Add "tblAllYears", lHdrAllYears & "|" & NCOL
    r = r + 1

    ' ---- summary table ----
    r = r + 1
    vOut(r, 1) = "Summary - one row per (year, month)"
    r = r + 1
    Dim lHdrSummary As Long: lHdrSummary = r
    Dim vSumHdr As Variant
    vSumHdr = Array("year", "month", "mean_da_price", "min_he", "min_price", "max_he", _
                    "max_price", "daily_spread", "morning_peak_he", "evening_peak_he", _
                    "hours_below_zero", "hours_above_2x_mean")
    Dim c As Long
    For c = 0 To UBound(vSumHdr)
        vOut(r, c + 1) = vSumHdr(c)
    Next c

    For y = 1 To mNYears
        For m = 1 To 12
            r = r + 1
            dn_FillSummaryRow vOut, r, y, m
        Next m
    Next y
    colTblTop.Add "tblSummary", lHdrSummary & "|12"

    ' ---- single write ----
    ws.Range("A1").Resize(lTotal, NCOL).Value = vOut

    ' ---- build ListObjects + formatting (post-write, no cell loops over data) ----
    Dim vKey As Variant, sVal As String, lTop As Long, lCols As Long
    For Each vKey In colTblTop.Keys
        sVal = colTblTop(vKey)
        lTop = CLng(Split(sVal, "|")(0))
        lCols = CLng(Split(sVal, "|")(1))
        Dim lRows As Long
        If CStr(vKey) = "tblSummary" Then lRows = mNYears * 12 Else lRows = 24
        Dim rngT As Range
        Set rngT = ws.Range(ws.Cells(lTop, 1), ws.Cells(lTop + lRows, lCols))
        Dim lo As ListObject
        Set lo = ws.ListObjects.Add(xlSrcRange, rngT, , xlYes)
        lo.Name = CStr(vKey)
        lo.TableStyle = "TableStyleLight9"

        If CStr(vKey) = "tblSummary" Then
            ws.Range(ws.Cells(lTop + 1, 3), ws.Cells(lTop + lRows, 3)).NumberFormat = "$ #,##0.00"
            ws.Range(ws.Cells(lTop + 1, 5), ws.Cells(lTop + lRows, 5)).NumberFormat = "$ #,##0.00"
            ws.Range(ws.Cells(lTop + 1, 7), ws.Cells(lTop + lRows, 7)).NumberFormat = "$ #,##0.00"
            ws.Range(ws.Cells(lTop + 1, 8), ws.Cells(lTop + lRows, 8)).NumberFormat = "$ #,##0.00"
        Else
            ws.Range(ws.Cells(lTop + 1, 2), ws.Cells(lTop + lRows, lCols)).NumberFormat = "$ #,##0.00"
        End If
    Next vKey

    ws.Rows(1).Font.Bold = True
    ws.Columns("A:M").AutoFit
    ' (Header-row freeze intentionally omitted: freezing panes requires the sheet
    '  to be the active window, which the "no .Activate/.Select" constraint forbids.)
End Sub

Private Function dn_AllYearsValue(ByVal m As Long, ByVal h As Long) As Variant
    Dim y As Long, dS As Double, lC As Long
    dS = 0: lC = 0
    For y = 1 To mNYears
        If mCnt(y, m, h) > 0 Then dS = dS + mAgg(y, m, h): lC = lC + 1
    Next y
    If lC = 0 Then dn_AllYearsValue = Empty Else dn_AllYearsValue = dS / lC
End Function

Private Sub dn_FillSummaryRow(ByRef vOut() As Variant, ByVal r As Long, ByVal y As Long, ByVal m As Long)
    Dim h As Long, dS As Double, lC As Long
    Dim dMinPrice As Double, dMaxPrice As Double
    Dim lMinHE As Long, lMaxHE As Long
    Dim lMornHE As Long, lEveHE As Long
    Dim dMorn As Double, dEve As Double
    Dim lBelowZero As Long, lAbove2x As Long

    dS = 0: lC = 0
    dMinPrice = 1E+308: dMaxPrice = -1E+308
    dMorn = -1E+308: dEve = -1E+308
    lMinHE = 0: lMaxHE = 0: lMornHE = 0: lEveHE = 0
    lBelowZero = 0: lAbove2x = 0

    For h = 1 To 24
        If mCnt(y, m, h) > 0 Then
            Dim dV As Double: dV = mAgg(y, m, h)
            dS = dS + dV: lC = lC + 1
            If dV < dMinPrice Then dMinPrice = dV: lMinHE = h
            If dV > dMaxPrice Then dMaxPrice = dV: lMaxHE = h
            If h <= 12 Then
                If dV > dMorn Then dMorn = dV: lMornHE = h
            Else
                If dV > dEve Then dEve = dV: lEveHE = h
            End If
        End If
    Next h

    Dim dMean As Double
    If lC > 0 Then dMean = dS / lC Else dMean = 0

    For h = 1 To 24
        If mCnt(y, m, h) > 0 Then
            If mAgg(y, m, h) < 0 Then lBelowZero = lBelowZero + 1
            If mAgg(y, m, h) > 2 * dMean Then lAbove2x = lAbove2x + 1
        End If
    Next h

    vOut(r, 1) = mYears(y)
    vOut(r, 2) = dn_MonthName(m)
    If lC = 0 Then Exit Sub
    vOut(r, 3) = dMean
    vOut(r, 4) = lMinHE
    vOut(r, 5) = dMinPrice
    vOut(r, 6) = lMaxHE
    vOut(r, 7) = dMaxPrice
    vOut(r, 8) = dMaxPrice - dMinPrice
    vOut(r, 9) = lMornHE
    vOut(r, 10) = lEveHE
    vOut(r, 11) = lBelowZero
    vOut(r, 12) = lAbove2x
End Sub

' ============================================================================
'  SECTION: CHARTS
' ============================================================================

Private Sub dn_BuildCharts()
    Dim wsCharts As Worksheet
    Set wsCharts = dn_FreshSheet("DiurnalCharts")
    dn_DeleteCharts wsCharts

    Dim wsStg As Worksheet
    Set wsStg = dn_SheetByName("DiurnalChartData")
    If wsStg Is Nothing Then
        Set wsStg = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsStg.Name = "DiurnalChartData"
    End If
    wsStg.Cells.Clear
    wsStg.Visible = xlSheetVeryHidden

    Dim lStgCol As Long: lStgCol = 1        ' running column cursor on the staging sheet

    ' layout geometry
    Const CW As Double = 460, CH As Double = 300, GUT As Double = 20
    Dim dLeft As Double, dTop As Double
    Dim lPos As Long: lPos = 0

    ' common axis bounds: validated 5% headroom around the aggregated global range
    Dim dYMin As Double, dYMax As Double
    Dim bDoScale As Boolean
    bDoScale = dn_CommonScale(dYMin, dYMax)

    Dim colCharts As Object: Set colCharts = CreateObject("Scripting.Dictionary")

    ' ---- Charts 1..N : one per year, 12 month lines ----
    Dim y As Long
    For y = 1 To mNYears
        dLeft = (lPos Mod 2) * (CW + GUT) + GUT
        dTop = (lPos \ 2) * (CH + GUT) + GUT
        Dim co As ChartObject
        Set co = dn_BuildYearChart(wsCharts, wsStg, lStgCol, y, dLeft, dTop, CW, CH)
        colCharts.Add "yr" & y, co.Name
        lPos = lPos + 1
    Next y

    ' ---- All-years overlay ----
    dLeft = (lPos Mod 2) * (CW + GUT) + GUT
    dTop = (lPos \ 2) * (CH + GUT) + GUT
    Dim coAll As ChartObject
    Set coAll = dn_BuildAllYearsChart(wsCharts, wsStg, lStgCol, dLeft, dTop, CW, CH)
    colCharts.Add "all", coAll.Name
    lPos = lPos + 1

    ' ---- Year-over-year drift ----
    dLeft = (lPos Mod 2) * (CW + GUT) + GUT
    dTop = (lPos \ 2) * (CH + GUT) + GUT
    dn_BuildDriftChart wsCharts, wsStg, lStgCol, dLeft, dTop, CW, CH
    lPos = lPos + 1

    ' ---- apply common Y scale to the annual + overlay charts ----
    ' (all series are already plotted on every chart at this point, so the value
    '  axis exists and scale properties will take)
    If UCase$(mCfgYAxisMode) = "COMMON" And bDoScale Then
        Dim vKey As Variant
        For Each vKey In colCharts.Keys
            Dim cc As ChartObject
            Set cc = wsCharts.ChartObjects(colCharts(vKey))
            ' MinimumScale BEFORE MaximumScale: the reverse order can transiently
            ' invert the range and fail. Both routed through the safe setter.
            dn_SafeAxis cc.Chart.Axes(xlValue), "MinimumScale", dYMin
            dn_SafeAxis cc.Chart.Axes(xlValue), "MaximumScale", dYMax
        Next vKey
        dn_LogLine "Y-axis mode|Common [" & Format$(dYMin, "0.00") & " , " & Format$(dYMax, "0.00") & "]"
    ElseIf UCase$(mCfgYAxisMode) = "COMMON" Then
        dn_LogLine "Y-axis mode|Common requested but scaling skipped (invalid/zero range)"
    Else
        dn_LogLine "Y-axis mode|PerYear (auto-scaled)"
    End If

    ' ---- optional seasonal band charts ----
    If mCfgShowBands Then
        dn_BuildBandCharts wsCharts, wsStg, lStgCol
    End If
End Sub

' -- build one annual chart (12 month series) ---------------------------------
Private Function dn_BuildYearChart(wsCharts As Worksheet, wsStg As Worksheet, _
        ByRef lStgCol As Long, ByVal y As Long, _
        ByVal dLeft As Double, ByVal dTop As Double, _
        ByVal CW As Double, ByVal CH As Double) As ChartObject

    ' stage a 24 x 13 block: col0 = HE, col1..12 = month values
    Dim vBlk() As Variant: ReDim vBlk(1 To 24, 1 To 13)
    Dim h As Long, m As Long
    For h = 1 To 24
        vBlk(h, 1) = h
        For m = 1 To 12
            If mCnt(y, m, h) > 0 Then vBlk(h, 1 + m) = mAgg(y, m, h) Else vBlk(h, 1 + m) = 0
        Next m
    Next h
    Dim lC0 As Long: lC0 = lStgCol
    wsStg.Cells(1, lC0).Resize(24, 13).Value = vBlk
    lStgCol = lStgCol + 14

    ' which months are the year's high/low mean (emphasized)
    Dim lHiM As Long, lLoM As Long
    dn_YearExtremeMonths y, lHiM, lLoM

    Dim co As ChartObject
    Set co = dn_NewLineChart(wsCharts, dLeft, dTop, CW, CH)

    Dim bAnyNeg As Boolean: bAnyNeg = False
    For m = 1 To 12
        Dim ser As Series
        Set ser = co.Chart.SeriesCollection.NewSeries
        ser.XValues = wsStg.Range(wsStg.Cells(1, lC0), wsStg.Cells(24, lC0))
        ser.Values = wsStg.Range(wsStg.Cells(1, lC0 + m), wsStg.Cells(24, lC0 + m))
        ser.Name = dn_MonthName(m)
        ser.ChartType = xlLine
        ser.MarkerStyle = xlMarkerStyleNone
        ser.Format.Line.ForeColor.RGB = dn_MonthColor(m)
        If m = lHiM Or m = lLoM Then
            ser.Format.Line.Weight = 3
        Else
            ser.Format.Line.Weight = 1.75
        End If
        For h = 1 To 24
            If mCnt(y, m, h) > 0 Then If mAgg(y, m, h) < 0 Then bAnyNeg = True
        Next h
    Next m

    dn_StyleChart co, _
        "Average Diurnal Day-Ahead Price Curve - " & mYears(y), _
        dn_Subtitle(), bAnyNeg, wsStg, lC0

    Set dn_BuildYearChart = co
End Function

' -- all-years overlay --------------------------------------------------------
Private Function dn_BuildAllYearsChart(wsCharts As Worksheet, wsStg As Worksheet, _
        ByRef lStgCol As Long, ByVal dLeft As Double, ByVal dTop As Double, _
        ByVal CW As Double, ByVal CH As Double) As ChartObject

    Dim vBlk() As Variant: ReDim vBlk(1 To 24, 1 To 13)
    Dim h As Long, m As Long
    Dim bAnyNeg As Boolean
    For h = 1 To 24
        vBlk(h, 1) = h
        For m = 1 To 12
            Dim v As Variant: v = dn_AllYearsValue(m, h)
            If IsEmpty(v) Then v = 0
            vBlk(h, 1 + m) = v
            If v < 0 Then bAnyNeg = True
        Next m
    Next h
    Dim lC0 As Long: lC0 = lStgCol
    wsStg.Cells(1, lC0).Resize(24, 13).Value = vBlk
    lStgCol = lStgCol + 14

    Dim co As ChartObject
    Set co = dn_NewLineChart(wsCharts, dLeft, dTop, CW, CH)
    For m = 1 To 12
        Dim ser As Series
        Set ser = co.Chart.SeriesCollection.NewSeries
        ser.XValues = wsStg.Range(wsStg.Cells(1, lC0), wsStg.Cells(24, lC0))
        ser.Values = wsStg.Range(wsStg.Cells(1, lC0 + m), wsStg.Cells(24, lC0 + m))
        ser.Name = dn_MonthName(m)
        ser.ChartType = xlLine
        ser.MarkerStyle = xlMarkerStyleNone
        ser.Format.Line.Weight = 1.75
        ser.Format.Line.ForeColor.RGB = dn_MonthColor(m)
    Next m

    dn_StyleChart co, "All-Years Reference Diurnal Curve", dn_Subtitle(), bAnyNeg, wsStg, lC0
    Set dn_BuildAllYearsChart = co
End Function

' -- year-over-year drift (one series per year) -------------------------------
Private Sub dn_BuildDriftChart(wsCharts As Worksheet, wsStg As Worksheet, _
        ByRef lStgCol As Long, ByVal dLeft As Double, ByVal dTop As Double, _
        ByVal CW As Double, ByVal CH As Double)

    Dim vBlk() As Variant: ReDim vBlk(1 To 24, 1 To (1 + mNYears))
    Dim h As Long, y As Long
    For h = 1 To 24
        vBlk(h, 1) = h
        For y = 1 To mNYears
            vBlk(h, 1 + y) = dn_YearAnnualCurve(y, h)
        Next y
    Next h
    Dim lC0 As Long: lC0 = lStgCol
    wsStg.Cells(1, lC0).Resize(24, 1 + mNYears).Value = vBlk
    lStgCol = lStgCol + mNYears + 2

    Dim co As ChartObject
    Set co = dn_NewLineChart(wsCharts, dLeft, dTop, CW, CH)
    For y = 1 To mNYears
        Dim ser As Series
        Set ser = co.Chart.SeriesCollection.NewSeries
        ser.XValues = wsStg.Range(wsStg.Cells(1, lC0), wsStg.Cells(24, lC0))
        ser.Values = wsStg.Range(wsStg.Cells(1, lC0 + y), wsStg.Cells(24, lC0 + y))
        ser.Name = CStr(mYears(y))
        ser.ChartType = xlLine
        ser.MarkerStyle = xlMarkerStyleNone
        ser.Format.Line.Weight = 2#
        ser.Format.Line.ForeColor.RGB = dn_DriftColor(y, mNYears)
    Next y

    dn_StyleChart co, "Year-over-Year Diurnal Drift (annual average curve)", _
                  dn_Subtitle(), False, wsStg, lC0
End Sub

' -- shared line-chart skeleton ----------------------------------------------
Private Function dn_NewLineChart(ws As Worksheet, ByVal dLeft As Double, ByVal dTop As Double, _
        ByVal CW As Double, ByVal CH As Double) As ChartObject
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(dLeft, dTop, CW, CH)
    co.Chart.ChartType = xlLine
    Set dn_NewLineChart = co
End Function

' -- validated common Y-scale bounds; returns False if scaling should be skipped
Private Function dn_CommonScale(ByRef dYMin As Double, ByRef dYMax As Double) As Boolean
    dn_CommonScale = False
    dYMin = 0: dYMax = 0

    If mAggMin > mAggMax Then Exit Function                 ' no valid aggregated values
    If mAggMin = 0 And mAggMax = 0 Then
        dn_LogLine "Y-axis|global range 0 to 0 - explicit scaling skipped"
        Exit Function
    End If

    Dim dRange As Double: dRange = mAggMax - mAggMin
    If dRange <= 0 Then dRange = IIf(mAggMax = 0, 1, Abs(mAggMax) * 0.1)
    dYMin = mAggMin - 0.05 * dRange
    dYMax = mAggMax + 0.05 * dRange

    ' equal or inverted -> pad by +/-10% of the absolute value
    If dYMin >= dYMax Then
        Dim dPad As Double: dPad = IIf(dYMax = 0, 1, Abs(dYMax) * 0.1)
        dYMin = dYMax - dPad
        dYMax = dYMax + dPad
    End If
    If dYMin >= dYMax Then Exit Function                    ' still bad -> auto-scale

    dn_CommonScale = True
End Function

' -- route every axis property assignment through here: one unsupported property
'    logs a warning instead of aborting the whole run
Private Sub dn_SafeAxis(ByVal ax As Object, ByVal sProp As String, ByVal vVal As Variant)
    On Error Resume Next
    CallByName ax, sProp, VbLet, vVal
    If Err.Number <> 0 Then
        dn_LogWarn "Axis property " & sProp & " not supported: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' -- titles, axes, legend, gridlines, optional zero line ---------------------
Private Sub dn_StyleChart(co As ChartObject, ByVal sTitle As String, ByVal sSub As String, _
        ByVal bAddZero As Boolean, wsStg As Worksheet, ByVal lC0 As Long)

    With co.Chart
        .HasTitle = True
        .ChartTitle.Text = sTitle & IIf(Len(sSub) > 0, vbLf & sSub, "")
        On Error Resume Next
        .ChartTitle.Font.Size = 11
        On Error GoTo 0

        ' X axis is a CATEGORY axis (HE1..HE24 are text labels): use tick SPACING,
        ' never MajorUnit / MinimumScale / MaximumScale (those are value-axis only).
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = IIf(UCase$(mCfgHEConv) = "BEGINNING", "Hour Beginning", "Hour Ending")
            On Error Resume Next
            .MajorGridlines.Delete        ' category gridlines may not exist
            On Error GoTo 0
        End With
        dn_SafeAxis .Axes(xlCategory), "TickLabelSpacing", 2
        dn_SafeAxis .Axes(xlCategory), "TickMarkSpacing", 2

        ' Y axis
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Day-Ahead Price ($/MWh)"
            On Error Resume Next
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
            .MajorGridlines.Format.Line.Weight = 0.5
            On Error GoTo 0
        End With

        ' legend at bottom, small font (3x4 auto-flows on supporting versions)
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        On Error Resume Next
        .Legend.Font.Size = 8
        On Error GoTo 0

        ' no chart-area border
        On Error Resume Next
        .ChartArea.Format.Line.Visible = msoFalse
        On Error GoTo 0
    End With

    ' heavy zero line when any plotted value is negative
    If bAddZero Then
        Dim zeros(1 To 24) As Double
        Dim serZ As Series
        Set serZ = co.Chart.SeriesCollection.NewSeries
        serZ.XValues = wsStg.Range(wsStg.Cells(1, lC0), wsStg.Cells(24, lC0))
        serZ.Values = zeros                 ' all 0
        serZ.Name = "0"
        serZ.ChartType = xlLine
        serZ.MarkerStyle = xlMarkerStyleNone
        serZ.Format.Line.Weight = 2#
        serZ.Format.Line.ForeColor.RGB = RGB(20, 20, 20)
        On Error Resume Next
        co.Chart.Legend.LegendEntries(co.Chart.SeriesCollection.Count).Delete
        On Error GoTo 0
    End If
End Sub

' -- optional seasonal P10-P90 band charts (one per year per season) ----------
Private Sub dn_BuildBandCharts(wsCharts As Worksheet, wsStg As Worksheet, ByRef lStgCol As Long)
    Dim wsBands As Worksheet
    Set wsBands = dn_FreshSheet("DiurnalBands")
    dn_DeleteCharts wsBands

    Const CW As Double = 460, CH As Double = 300, GUT As Double = 20
    Dim lPos As Long: lPos = 0
    Dim vSeasonName As Variant
    vSeasonName = Array("Winter (DJF)", "Spring (MAM)", "Summer (JJA)", "Fall (SON)")
    Dim vSeasonM As Variant
    vSeasonM = Array(Array(12, 1, 2), Array(3, 4, 5), Array(6, 7, 8), Array(9, 10, 11))

    Dim y As Long, s As Long, h As Long
    For y = 1 To mNYears
        For s = 0 To 3
            ' stage 24 x 4: HE, P10, band=(P90-P10), mean
            Dim vBlk() As Variant: ReDim vBlk(1 To 24, 1 To 4)
            For h = 1 To 24
                Dim dP10 As Double, dP90 As Double, dMeanV As Double
                dn_SeasonStats y, vSeasonM(s), h, dP10, dP90, dMeanV
                vBlk(h, 1) = h
                vBlk(h, 2) = dP10
                vBlk(h, 3) = dP90 - dP10
                vBlk(h, 4) = dMeanV
            Next h
            Dim lC0 As Long: lC0 = lStgCol
            wsStg.Cells(1, lC0).Resize(24, 4).Value = vBlk
            lStgCol = lStgCol + 5

            Dim dLeft As Double, dTop As Double
            dLeft = (lPos Mod 2) * (CW + GUT) + GUT
            dTop = (lPos \ 2) * (CH + GUT) + GUT

            Dim co As ChartObject
            Set co = wsBands.ChartObjects.Add(dLeft, dTop, CW, CH)
            co.Chart.ChartType = xlAreaStacked

            ' base (P10) - transparent
            Dim sB As Series
            Set sB = co.Chart.SeriesCollection.NewSeries
            sB.XValues = wsStg.Range(wsStg.Cells(1, lC0), wsStg.Cells(24, lC0))
            sB.Values = wsStg.Range(wsStg.Cells(1, lC0 + 1), wsStg.Cells(24, lC0 + 1))
            sB.Name = "P10"
            sB.ChartType = xlAreaStacked
            On Error Resume Next
            sB.Format.Fill.Visible = msoFalse
            sB.Format.Line.Visible = msoFalse
            On Error GoTo 0

            ' band (P90-P10) stacked on top, ~25% transparency
            Dim sBand As Series
            Set sBand = co.Chart.SeriesCollection.NewSeries
            sBand.XValues = wsStg.Range(wsStg.Cells(1, lC0), wsStg.Cells(24, lC0))
            sBand.Values = wsStg.Range(wsStg.Cells(1, lC0 + 2), wsStg.Cells(24, lC0 + 2))
            sBand.Name = "P10-P90 band"
            sBand.ChartType = xlAreaStacked
            On Error Resume Next
            sBand.Format.Fill.ForeColor.RGB = RGB(120, 150, 200)
            sBand.Format.Fill.Transparency = 0.75
            sBand.Format.Line.Visible = msoFalse
            On Error GoTo 0

            ' mean curve as an overlaid line
            Dim sMean As Series
            Set sMean = co.Chart.SeriesCollection.NewSeries
            sMean.XValues = wsStg.Range(wsStg.Cells(1, lC0), wsStg.Cells(24, lC0))
            sMean.Values = wsStg.Range(wsStg.Cells(1, lC0 + 3), wsStg.Cells(24, lC0 + 3))
            sMean.Name = "Mean"
            sMean.ChartType = xlLine
            sMean.MarkerStyle = xlMarkerStyleNone
            sMean.Format.Line.Weight = 2.25
            sMean.Format.Line.ForeColor.RGB = RGB(200, 40, 40)

            co.Chart.HasTitle = True
            co.Chart.ChartTitle.Text = mYears(y) & " " & vSeasonName(s) & _
                " - Diurnal Mean with P10-P90 Band"
            co.Chart.Axes(xlCategory).HasTitle = True
            co.Chart.Axes(xlCategory).AxisTitle.Text = _
                IIf(UCase$(mCfgHEConv) = "BEGINNING", "Hour Beginning", "Hour Ending")
            co.Chart.Axes(xlValue).HasTitle = True
            co.Chart.Axes(xlValue).AxisTitle.Text = "Day-Ahead Price ($/MWh)"
            lPos = lPos + 1
        Next s
    Next y
    dn_LogLine "Seasonal band charts emitted|" & (mNYears * 4)
End Sub

Private Sub dn_SeasonStats(ByVal y As Long, vMonths As Variant, ByVal h As Long, _
        ByRef dP10 As Double, ByRef dP90 As Double, ByRef dMeanV As Double)
    ' pool the season's month buckets at this HE
    Dim lTot As Long, i As Long, m As Long
    For i = LBound(vMonths) To UBound(vMonths)
        m = vMonths(i)
        lTot = lTot + mCnt(y, m, h)
    Next i
    If lTot = 0 Then dP10 = 0: dP90 = 0: dMeanV = 0: Exit Sub

    Dim dPool() As Double: ReDim dPool(1 To lTot)
    Dim k As Long: k = 0
    Dim dArr() As Double, j As Long
    Dim dSum As Double
    For i = LBound(vMonths) To UBound(vMonths)
        m = vMonths(i)
        If mCnt(y, m, h) > 0 Then
            dArr = mBucket(y, m, h)
            For j = 1 To mCnt(y, m, h)
                k = k + 1
                dPool(k) = dArr(j)
                dSum = dSum + dArr(j)
            Next j
        End If
    Next i
    dn_QuickSortDouble dPool, 1, lTot
    dP10 = dn_PercentileInc(dPool, lTot, 0.1)
    dP90 = dn_PercentileInc(dPool, lTot, 0.9)
    dMeanV = dSum / lTot
End Sub

' ============================================================================
'  SECTION: SUMMARY MsgBox
' ============================================================================

Private Sub dn_ShowSummary()
    Dim s As String
    s = "Diurnal analysis complete." & vbCrLf & vbCrLf
    s = s & "Years charted: " & mNYears & " (" & dn_JoinYears() & ")" & vbCrLf
    s = s & "Global price range: " & Format$(mRawMin, "$#,##0.00") & " to " & _
            Format$(mRawMax, "$#,##0.00") & vbCrLf & vbCrLf

    Dim y As Long
    For y = 1 To mNYears
        Dim dAnnMean As Double, lDeepM As Long, dDeep As Double
        Dim lShalM As Long, dShal As Double
        dn_YearSpreadExtremes y, dAnnMean, lDeepM, dDeep, lShalM, dShal
        s = s & mYears(y) & ": mean " & Format$(dAnnMean, "$#,##0.00") & _
                " | deepest spread " & dn_MonthName(lDeepM) & " (" & Format$(dDeep, "$#,##0.00") & ")" & _
                " | shallowest " & dn_MonthName(lShalM) & " (" & Format$(dShal, "$#,##0.00") & ")" & vbCrLf
    Next y

    Dim sWarn As String
    If mNYears <> 5 Then sWarn = sWarn & "- Year count is " & mNYears & ", not 5." & vbCrLf
    If mEmptyCells > 0 Then sWarn = sWarn & "- " & mEmptyCells & " (year,month,HE) cell(s) had no observations." & vbCrLf
    If Len(sWarn) > 0 Then s = s & vbCrLf & "WARNINGS:" & vbCrLf & sWarn

    MsgBox s, IIf(Len(sWarn) > 0, vbExclamation, vbInformation), "Diurnal Curves"
End Sub

' ============================================================================
'  SECTION: RunLog  (single array write)
' ============================================================================

Private Sub dn_WriteRunLog()
    Dim ws As Worksheet
    Set ws = dn_FreshSheet("RunLog")

    Dim lN As Long: lN = mLog.Count
    Dim vOut() As Variant
    ReDim vOut(1 To lN + 1, 1 To 2)
    vOut(1, 1) = "Metric": vOut(1, 2) = "Value"

    Dim i As Long, sLine As String, p As Long
    For i = 1 To lN
        sLine = mLog(i)                     ' "label|value"
        p = InStr(sLine, "|")
        If p > 0 Then
            vOut(i + 1, 1) = Left$(sLine, p - 1)
            vOut(i + 1, 2) = Mid$(sLine, p + 1)
        Else
            vOut(i + 1, 1) = sLine
        End If
    Next i

    ws.Range("A1").Resize(lN + 1, 2).Value = vOut
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:B").AutoFit
End Sub

' ============================================================================
'  SECTION: HELPERS - aggregation lookups
' ============================================================================

Private Function dn_YearAnnualCurve(ByVal y As Long, ByVal h As Long) As Double
    Dim m As Long, dS As Double, lC As Long
    For m = 1 To 12
        If mCnt(y, m, h) > 0 Then dS = dS + mAgg(y, m, h): lC = lC + 1
    Next m
    If lC > 0 Then dn_YearAnnualCurve = dS / lC
End Function

Private Sub dn_YearExtremeMonths(ByVal y As Long, ByRef lHiM As Long, ByRef lLoM As Long)
    Dim m As Long, h As Long, dS As Double, lC As Long, dMeanM As Double
    Dim dHi As Double, dLo As Double
    dHi = -1E+308: dLo = 1E+308: lHiM = 1: lLoM = 1
    For m = 1 To 12
        dS = 0: lC = 0
        For h = 1 To 24
            If mCnt(y, m, h) > 0 Then dS = dS + mAgg(y, m, h): lC = lC + 1
        Next h
        If lC > 0 Then
            dMeanM = dS / lC
            If dMeanM > dHi Then dHi = dMeanM: lHiM = m
            If dMeanM < dLo Then dLo = dMeanM: lLoM = m
        End If
    Next m
End Sub

Private Sub dn_YearSpreadExtremes(ByVal y As Long, ByRef dAnnMean As Double, _
        ByRef lDeepM As Long, ByRef dDeep As Double, _
        ByRef lShalM As Long, ByRef dShal As Double)
    Dim m As Long, h As Long
    Dim dGrand As Double, lGrandC As Long
    dDeep = -1E+308: dShal = 1E+308: lDeepM = 1: lShalM = 1
    For m = 1 To 12
        Dim dMn As Double, dMx As Double, lC As Long
        dMn = 1E+308: dMx = -1E+308: lC = 0
        For h = 1 To 24
            If mCnt(y, m, h) > 0 Then
                If mAgg(y, m, h) < dMn Then dMn = mAgg(y, m, h)
                If mAgg(y, m, h) > dMx Then dMx = mAgg(y, m, h)
                dGrand = dGrand + mAgg(y, m, h): lGrandC = lGrandC + 1: lC = lC + 1
            End If
        Next h
        If lC > 0 Then
            Dim dSpread As Double: dSpread = dMx - dMn
            If dSpread > dDeep Then dDeep = dSpread: lDeepM = m
            If dSpread < dShal Then dShal = dSpread: lShalM = m
        End If
    Next m
    If lGrandC > 0 Then dAnnMean = dGrand / lGrandC
End Sub

' ============================================================================
'  SECTION: HELPERS - sorting / percentiles / colors
' ============================================================================

Private Sub dn_QuickSortDouble(arr() As Double, ByVal lLo As Long, ByVal lHi As Long)
    Dim i As Long, j As Long
    Dim dPivot As Double, dSwap As Double
    i = lLo: j = lHi
    dPivot = arr((lLo + lHi) \ 2)
    Do While i <= j
        Do While arr(i) < dPivot: i = i + 1: Loop
        Do While arr(j) > dPivot: j = j - 1: Loop
        If i <= j Then
            dSwap = arr(i): arr(i) = arr(j): arr(j) = dSwap
            i = i + 1: j = j - 1
        End If
    Loop
    If lLo < j Then dn_QuickSortDouble arr, lLo, j
    If i < lHi Then dn_QuickSortDouble arr, i, lHi
End Sub

' PERCENTILE.INC on an ascending-sorted array (indices 1..lN)
Private Function dn_PercentileInc(arr() As Double, ByVal lN As Long, ByVal dPct As Double) As Double
    If lN <= 0 Then Exit Function
    If lN = 1 Then dn_PercentileInc = arr(1): Exit Function
    Dim dRank As Double, lLo As Long, dFrac As Double
    dRank = dPct * (lN - 1) + 1
    lLo = Int(dRank)
    dFrac = dRank - lLo
    If lLo >= lN Then
        dn_PercentileInc = arr(lN)
    Else
        dn_PercentileInc = arr(lLo) + dFrac * (arr(lLo + 1) - arr(lLo))
    End If
End Function

' cyclical seasonal ramp: adjacent months are adjacent in color
Private Function dn_MonthColor(ByVal lMonthNum As Long) As Long
    Select Case lMonthNum
        Case 1:  dn_MonthColor = RGB(49, 54, 149)
        Case 2:  dn_MonthColor = RGB(69, 117, 180)
        Case 3:  dn_MonthColor = RGB(116, 173, 209)
        Case 4:  dn_MonthColor = RGB(171, 217, 233)
        Case 5:  dn_MonthColor = RGB(224, 243, 248)
        Case 6:  dn_MonthColor = RGB(254, 224, 144)
        Case 7:  dn_MonthColor = RGB(253, 174, 97)
        Case 8:  dn_MonthColor = RGB(244, 109, 67)
        Case 9:  dn_MonthColor = RGB(215, 48, 39)
        Case 10: dn_MonthColor = RGB(190, 90, 80)
        Case 11: dn_MonthColor = RGB(120, 100, 150)
        Case Else: dn_MonthColor = RGB(70, 70, 140)
    End Select
End Function

' sequential light->dark ramp so drift direction over time is legible
Private Function dn_DriftColor(ByVal y As Long, ByVal lNY As Long) As Long
    Dim t As Double
    If lNY <= 1 Then t = 0 Else t = (y - 1) / (lNY - 1)
    Dim r As Long, g As Long, b As Long
    r = CLng(200 + (20 - 200) * t)
    g = CLng(210 + (30 - 210) * t)
    b = CLng(230 + (90 - 230) * t)
    dn_DriftColor = RGB(r, g, b)
End Function

Private Function dn_MonthName(ByVal lMonthNum As Long) As String
    Dim v As Variant
    v = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    dn_MonthName = v(lMonthNum - 1)
End Function

' ============================================================================
'  SECTION: HELPERS - sheets / names / misc
' ============================================================================

Private Function dn_SheetByName(ByVal sName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sName, vbTextCompare) = 0 Then Set dn_SheetByName = ws: Exit Function
    Next ws
End Function

' return an existing sheet cleared, or a freshly added one
Private Function dn_FreshSheet(ByVal sName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = dn_SheetByName(sName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sName
    Else
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
        dn_DeleteListObjects ws
        ws.Cells.Clear
    End If
    Set dn_FreshSheet = ws
End Function

Private Sub dn_DeleteListObjects(ws As Worksheet)
    Dim i As Long
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Unlist
    Next i
End Sub

Private Sub dn_DeleteCharts(ws As Worksheet)
    Dim i As Long
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
End Sub

Private Function dn_NameStr(ByVal sName As String, ByVal sDefault As String) As String
    On Error GoTo Fallback
    dn_NameStr = CStr(ThisWorkbook.Names(sName).RefersToRange.Value)
    Exit Function
Fallback:
    dn_NameStr = sDefault
End Function

Private Function dn_NameLng(ByVal sName As String, ByVal lDefault As Long) As Long
    On Error GoTo Fallback
    Dim v As Variant: v = ThisWorkbook.Names(sName).RefersToRange.Value
    If IsNumeric(v) Then dn_NameLng = CLng(v) Else dn_NameLng = lDefault
    Exit Function
Fallback:
    dn_NameLng = lDefault
End Function

Private Function dn_NameBool(ByVal sName As String, ByVal bDefault As Boolean) As Boolean
    On Error GoTo Fallback
    Dim v As Variant: v = ThisWorkbook.Names(sName).RefersToRange.Value
    dn_NameBool = dn_ToBool(v, bDefault)
    Exit Function
Fallback:
    dn_NameBool = bDefault
End Function

Private Function dn_ToBool(ByVal v As Variant, ByVal bDefault As Boolean) As Boolean
    Select Case UCase$(Trim$(CStr(v)))
        Case "TRUE", "-1", "1", "YES", "Y": dn_ToBool = True
        Case "FALSE", "0", "NO", "N", "": dn_ToBool = False
        Case Else: dn_ToBool = bDefault
    End Select
End Function

Private Function dn_LooksLikeDate(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        dn_LooksLikeDate = (CDbl(v) > 1)
    ElseIf VarType(v) = vbString Then
        dn_LooksLikeDate = IsDate(v)
    End If
End Function

Private Function dn_RowDump(vAll As Variant, ByVal ai As Long, ByVal lNRows As Long, ByVal lNCols As Long) As String
    If ai < 1 Or ai > lNRows Then dn_RowDump = "(row not in used range)": Exit Function
    Dim c As Long, s As String
    For c = 1 To lNCols
        s = s & "[" & CStr(vAll(ai, c)) & "] "
        If c >= 12 Then s = s & "...": Exit For
    Next c
    dn_RowDump = Trim$(s)
End Function

Private Function dn_Subtitle() As String
    Dim s As String
    s = "Price col: " & mCfgPriceColName & "  |  Aggregation: " & mCfgAggMethod
    If mHasSim And Len(Trim$(mCfgSimID)) > 0 Then s = s & "  |  Sim: " & mCfgSimID
    dn_Subtitle = s
End Function

Private Function dn_JoinYears() As String
    Dim i As Long, s As String
    For i = 1 To mNYears
        s = s & IIf(i > 1, ", ", "") & mYears(i)
    Next i
    dn_JoinYears = s
End Function

' ordered log: store as a 1-based Dictionary keyed by insertion index
Private Sub dn_LogLine(ByVal sLine As String)
    If mLog Is Nothing Then Set mLog = CreateObject("Scripting.Dictionary")
    mLog.Add mLog.Count + 1, sLine
End Sub

' non-fatal warning collected into RunLog rather than raised
Private Sub dn_LogWarn(ByVal sMsg As String)
    dn_LogLine "WARN|" & sMsg
End Sub

Private Sub dn_Fail(ByVal sMsg As String)
    Err.Raise vbObjectError + 513, "modDiurnalCurves." & mProc, sMsg
End Sub
