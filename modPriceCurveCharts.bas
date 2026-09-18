Attribute VB_Name = "modPriceCurveCharts"
Option Explicit

' ==========================================================================
'  modPriceCurveCharts
'  --------------------------------------------------------------------------
'  Reads a timestamp column and a price column, aggregates prices into annual
'  "shape" curves, and emits a batch of line charts. Each chart holds a
'  user-specified number of year curves so many years can be compared side by
'  side on a shared y-axis.
'
'  Public entry point:  BuildPriceCurveCharts
'  Target:              Excel for Windows, VBA 7.x
'  Dependencies:        none (no external references, no Scripting.Dictionary)
'
'  Years are indexed by integer offset from the minimum detected year, so plain
'  arrays are sufficient and no dictionary is required.
'
'  --------------------------------------------------------------------------
'  ACCEPTANCE TESTS (expected behavior — the implementation satisfies each)
'  --------------------------------------------------------------------------
'  1. Ten years of hourly data, 4 years per chart -> chartCount =
'     -Int(-10/4) = 3 charts holding 4, 4, and 2 curves. (See CreateChartBatch:
'     the last chart's end index is clamped to the final year, so it carries
'     the remainder.)
'
'  2. Seven years of monthly data, 7 years per chart -> -Int(-7/7) = 1 chart
'     with all 7 curves.
'
'  3. Three years of hourly data, 5 years per chart. Step 3 validation clamps
'     the allowed maximum to the total year count (3), so 5 is rejected and the
'     user is re-prompted; once a valid value <= 3 is entered (e.g. 3),
'     -Int(-3/3) = 1 chart holds all 3 curves.
'
'  4. An hourly series whose spring-forward day is missing the 02:00 hour still
'     yields a full 24-point curve: every other day of the year supplies an
'     02:00 observation, so that month-hour cell's count is non-zero and no
'     divide-by-zero occurs. Count-based averaging (not an assumed 8760) makes
'     this automatic.
'
'  5. A series containing negative prices renders with a y-axis minimum below
'     zero on every chart. The shared scale is computed from the actual data
'     min/max (padded and rounded) and is never floored to zero, so negative
'     wholesale prices are shown honestly.
'
'  6. A series whose first year begins in July and whose last year ends in
'     March still produces valid curves for those partial years. Months with no
'     observations have a zero count and are excluded from the average (left as
'     a gap in the plotted series) rather than zero-filled, so the partial-year
'     shape is not dragged toward zero.
'  --------------------------------------------------------------------------

' ==========================================================================
'  CONFIGURATION
' ==========================================================================

' When True (the two-stage method), hourly curves are built by first averaging
' every observation within each (month, hour) cell, then averaging those 12
' monthly means for each hour. This weights every month equally, so a 31-day
' month does not pull the annual diurnal shape harder than a 28-day month.
' When False, a single-stage mean of every observation at that hour of day
' within the year is used instead.
Private Const USE_EQUAL_MONTH_WEIGHTING As Boolean = True

Private Const SH_STAGING As String = "_CurveData"
Private Const SH_CHARTS  As String = "PriceCurves"

' Chart layout geometry (points).
Private Const CHART_W    As Double = 480
Private Const CHART_H    As Double = 300
Private Const CHART_GAP  As Double = 20
Private Const GRID_COLS  As Long = 2

' ==========================================================================
'  RECORD TYPES (data is passed between procedures as these, never as globals)
' ==========================================================================

' Parsed, validated observations. ts/price are 1-based and sized 1..n.
Private Type TParsed
    ts()      As Date
    price()   As Double
    n         As Long        ' count of valid rows kept
    rowsRead  As Long        ' total input rows examined
    skipped   As Long        ' rows dropped as unparseable / non-numeric
End Type

' The aggregated curve table ready for staging and charting.
' vals/present are dimensioned (0 .. xCount-1, 0 .. yearCount-1).
Private Type TCurves
    isHourly  As Boolean
    xCount    As Long        ' 24 (hours) or 12 (months)
    yearCount As Long        ' distinct years with data
    years()   As Long        ' actual four-digit years, ascending (0-based)
    vals()    As Double      ' curve value at (xIndex, yearIndex)
    present() As Boolean     ' True where a value exists; False = gap
End Type

' ==========================================================================
'  ORCHESTRATOR
' ==========================================================================

Public Sub BuildPriceCurveCharts()

    Dim savedScreen As Boolean, savedCalc As XlCalculation, savedAlerts As Boolean
    Dim stateSaved As Boolean

    On Error GoTo ErrHandler

    ' Snapshot and suppress application state; restored in Cleanup.
    savedScreen = Application.ScreenUpdating
    savedCalc = Application.Calculation
    savedAlerts = Application.DisplayAlerts
    stateSaved = True
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' --- Input ranges ---------------------------------------------------
    Dim tsRng As Range, prRng As Range
    If Not GetInputRanges(tsRng, prRng) Then GoTo Cleanup

    Dim tsArr As Variant, prArr As Variant, rowsCount As Long
    tsArr = ReadColumn(tsRng)
    prArr = ReadColumn(prRng)
    rowsCount = tsRng.Rows.Count

    ' --- Parse & validate ----------------------------------------------
    Dim p As TParsed
    If Not ParseSeries(tsArr, prArr, rowsCount, p) Then
        MsgBox "No valid rows could be parsed from the selected ranges." & vbCrLf & _
               "Rows examined: " & p.rowsRead & ", all skipped.", _
               vbExclamation, "Price Curve Charts"
        GoTo Cleanup
    End If

    ' --- Step 1: resolution --------------------------------------------
    Dim isHourly As Boolean
    If Not DetectResolution(p, isHourly) Then GoTo Cleanup

    ' --- Step 2: year detection ----------------------------------------
    Dim minYear As Long, maxYear As Long, span As Long
    DetectYearSpan p, minYear, maxYear, span

    ' --- Aggregate (also tells us which years actually carry data) ------
    Dim curves As TCurves
    If isHourly Then
        If Not AggregateHourly(p, minYear, span, curves) Then GoTo Cleanup
    Else
        If Not AggregateMonthly(p, minYear, span, curves) Then GoTo Cleanup
    End If

    If curves.yearCount = 0 Then
        MsgBox "No years with usable data were found after aggregation.", _
               vbExclamation, "Price Curve Charts"
        GoTo Cleanup
    End If

    ' Report the detected years (Step 2 message).
    MsgBox "Detected " & curves.yearCount & " distinct year(s): " & _
           curves.years(0) & " through " & curves.years(curves.yearCount - 1) & ".", _
           vbInformation, "Price Curve Charts"

    ' --- Step 3: years per chart ---------------------------------------
    Dim yearsPerChart As Long
    If Not AskYearsPerChart(curves.yearCount, yearsPerChart) Then GoTo Cleanup

    ' --- Step 4: chart count (integer ceiling of yearCount/yearsPerChart)
    ' -Int(-a/b) is the ceiling because Int() truncates toward negative
    ' infinity; negating both operands turns that floor into a ceiling.
    Dim chartCount As Long
    chartCount = -Int(-curves.yearCount / yearsPerChart)

    ' --- Staging sheet --------------------------------------------------
    If Not WriteStagingSheet(curves) Then GoTo Cleanup

    ' --- Charts ---------------------------------------------------------
    If Not CreateChartBatch(curves, yearsPerChart, chartCount) Then GoTo Cleanup

    ' --- Shared y-axis scale -------------------------------------------
    Dim yLo As Double, yHi As Double
    If Not ApplySharedScale(curves, yLo, yHi) Then GoTo Cleanup

    ' --- Completion report ---------------------------------------------
    ReportCompletion p, isHourly, curves, yearsPerChart, chartCount, yLo, yHi

Cleanup:
    If stateSaved Then
        Application.ScreenUpdating = savedScreen
        Application.Calculation = savedCalc
        Application.DisplayAlerts = savedAlerts
    End If
    Exit Sub

ErrHandler:
    MsgBox "Unexpected error in BuildPriceCurveCharts:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, _
           vbCritical, "Price Curve Charts"
    Resume Cleanup
End Sub

' ==========================================================================
'  INPUT
' ==========================================================================

' Prompts for the timestamp and price ranges and validates them. Returns True
' on success with the ranges set ByRef; False if the user cancelled or the
' selection is invalid.
Private Function GetInputRanges(ByRef tsRng As Range, ByRef prRng As Range) As Boolean
    On Error GoTo ErrHandler
    GetInputRanges = False

    Set tsRng = AskRange("Select the TIMESTAMP range (single column).", "Timestamps")
    If tsRng Is Nothing Then Exit Function

    Set prRng = AskRange("Select the PRICE range (single column).", "Prices")
    If prRng Is Nothing Then Exit Function

    If tsRng.Columns.Count <> 1 Or prRng.Columns.Count <> 1 Then
        MsgBox "Both selections must be a single column.", vbExclamation, "Price Curve Charts"
        Exit Function
    End If

    If tsRng.Rows.Count <> prRng.Rows.Count Then
        MsgBox "The timestamp range has " & tsRng.Rows.Count & " row(s) but the price " & _
               "range has " & prRng.Rows.Count & " row(s). They must match.", _
               vbExclamation, "Price Curve Charts"
        Exit Function
    End If

    GetInputRanges = True
    Exit Function

ErrHandler:
    MsgBox "Error while collecting input ranges:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, vbCritical, "Price Curve Charts"
    GetInputRanges = False
End Function

' Thin wrapper over Application.InputBox(Type:=8). Returns Nothing on cancel.
Private Function AskRange(ByVal prompt As String, ByVal title As String) As Range
    Dim r As Range
    On Error Resume Next
    Set r = Application.InputBox(prompt:=prompt, title:=title, Type:=8)
    On Error GoTo 0
    Set AskRange = r
End Function

' Reads a single-column range into a 1-based 2D array, coping with the
' single-cell case where .Value2 returns a scalar rather than an array.
Private Function ReadColumn(ByVal rng As Range) As Variant
    Dim v As Variant
    v = rng.Value2
    If IsArray(v) Then
        ReadColumn = v
    Else
        Dim a(1 To 1, 1 To 1) As Variant
        a(1, 1) = v
        ReadColumn = a
    End If
End Function

' ==========================================================================
'  PARSING
' ==========================================================================

' Parses the raw timestamp/price arrays into typed, validated observations.
' Rows whose timestamp cannot be parsed, or whose price is blank/non-numeric,
' are skipped and counted. Returns True if at least one valid row survived.
Private Function ParseSeries(ByVal tsArr As Variant, ByVal prArr As Variant, _
                             ByVal rowsCount As Long, ByRef outP As TParsed) As Boolean
    On Error GoTo ErrHandler
    ParseSeries = False

    ReDim outP.ts(1 To rowsCount)
    ReDim outP.price(1 To rowsCount)
    outP.rowsRead = rowsCount
    outP.n = 0
    outP.skipped = 0

    Dim i As Long, d As Date
    Dim tv As Variant, pv As Variant

    For i = 1 To rowsCount
        tv = tsArr(i, 1)
        pv = prArr(i, 1)

        ' Timestamp must parse.
        If Not TryParseTimestamp(tv, d) Then
            outP.skipped = outP.skipped + 1
        ' Price must be a real number (blank/text/error -> skip).
        ElseIf IsError(pv) Then
            outP.skipped = outP.skipped + 1
        ElseIf Not IsNumeric(pv) Then
            outP.skipped = outP.skipped + 1
        Else
            outP.n = outP.n + 1
            outP.ts(outP.n) = d
            outP.price(outP.n) = CDbl(pv)
        End If
    Next i

    ParseSeries = (outP.n > 0)
    Exit Function

ErrHandler:
    MsgBox "Error while parsing the series:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, vbCritical, "Price Curve Charts"
    ParseSeries = False
End Function

' Defensive timestamp parser. Numerics are treated as Excel date serials; text
' is parsed via CDate, with an explicit path for bare "YYYY-MM" values that
' CDate handles inconsistently across locales. Returns True and sets d on
' success.
Private Function TryParseTimestamp(ByVal v As Variant, ByRef d As Date) As Boolean
    TryParseTimestamp = False

    Select Case VarType(v)

        Case vbDate
            d = v
            TryParseTimestamp = True

        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
            ' Numeric cell -> Excel date serial.
            If CDbl(v) > 0 Then
                d = CDate(CDbl(v))
                TryParseTimestamp = True
            End If

        Case vbString
            Dim s As String
            s = Trim$(CStr(v))
            If Len(s) = 0 Then Exit Function

            ' Bare year-month, e.g. "2014-03": build the date explicitly so we
            ' do not depend on CDate's locale handling of a missing day.
            If Len(s) = 7 Then
                If Mid$(s, 5, 1) = "-" Then
                    Dim yPart As String, mPart As String
                    yPart = Left$(s, 4)
                    mPart = Right$(s, 2)
                    If IsNumeric(yPart) And IsNumeric(mPart) Then
                        Dim yy As Long, mm As Long
                        yy = CLng(yPart)
                        mm = CLng(mPart)
                        If yy >= 1900 And mm >= 1 And mm <= 12 Then
                            d = DateSerial(yy, mm, 1)
                            TryParseTimestamp = True
                        End If
                    End If
                    Exit Function
                End If
            End If

            ' General case: let CDate handle "YYYY-MM-DD HH:MM",
            ' "M/D/YYYY H:MM", and friends. Swallow parse failures.
            Dim tmp As Date
            On Error Resume Next
            tmp = CDate(s)
            If Err.Number = 0 Then
                d = tmp
                TryParseTimestamp = True
            End If
            On Error GoTo 0

    End Select
End Function

' ==========================================================================
'  RESOLUTION & YEAR DETECTION
' ==========================================================================

' Step 1. Auto-detects hourly vs monthly from the modal time delta between the
' first several dozen valid timestamps, presents that as the default, and asks
' the user to confirm (1 = Hourly, 2 = Monthly). Returns False on cancel.
Private Function DetectResolution(ByRef p As TParsed, ByRef isHourly As Boolean) As Boolean
    On Error GoTo ErrHandler
    DetectResolution = False

    Dim guess As Long
    guess = GuessResolution(p)   ' 1 = hourly, 2 = monthly

    Dim defTxt As String
    If guess = 1 Then defTxt = "1" Else defTxt = "2"

    Dim ans As String
    Do
        ans = InputBox( _
            "Time resolution of the series?" & vbCrLf & vbCrLf & _
            "  1 = Hourly (x-axis = hour of day, 0-23)" & vbCrLf & _
            "  2 = Monthly (x-axis = month, Jan-Dec)" & vbCrLf & vbCrLf & _
            "Auto-detected default is pre-filled below.", _
            "Step 1 - Resolution", defTxt)

        If StrPtr(ans) = 0 Then Exit Function   ' Cancel pressed

        ans = Trim$(ans)
        If ans = "1" Then
            isHourly = True
            DetectResolution = True
            Exit Function
        ElseIf ans = "2" Then
            isHourly = False
            DetectResolution = True
            Exit Function
        End If

        MsgBox "Please enter 1 (Hourly) or 2 (Monthly).", vbExclamation, "Price Curve Charts"
    Loop

ErrHandler:
    MsgBox "Error during resolution detection:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, vbCritical, "Price Curve Charts"
    DetectResolution = False
End Function

' Returns 1 (hourly) or 2 (monthly) based on the modal whole-hour gap between
' consecutive valid timestamps over a leading sample.
Private Function GuessResolution(ByRef p As TParsed) As Long
    Dim sampleN As Long
    sampleN = p.n - 1
    If sampleN > 60 Then sampleN = 60
    If sampleN < 1 Then
        GuessResolution = 1
        Exit Function
    End If

    ' Tally rounded hour-gaps in parallel arrays (no dictionary).
    Dim keys() As Long, cnts() As Long, used As Long
    ReDim keys(1 To sampleN)
    ReDim cnts(1 To sampleN)
    used = 0

    Dim i As Long, gapHrs As Long, j As Long, found As Boolean
    For i = 1 To sampleN
        gapHrs = CLng(Abs(CDbl(p.ts(i + 1) - p.ts(i)) * 24#))
        found = False
        For j = 1 To used
            If keys(j) = gapHrs Then
                cnts(j) = cnts(j) + 1
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            used = used + 1
            keys(used) = gapHrs
            cnts(used) = 1
        End If
    Next i

    ' Pick the most frequent gap.
    Dim bestIdx As Long, bestCnt As Long
    bestIdx = 1: bestCnt = 0
    For j = 1 To used
        If cnts(j) > bestCnt Then
            bestCnt = cnts(j)
            bestIdx = j
        End If
    Next j

    ' A modal gap of a day or less reads as intraday (hourly); anything larger
    ' (a monthly series steps ~28-31 days) reads as monthly.
    If keys(bestIdx) <= 48 Then
        GuessResolution = 1
    Else
        GuessResolution = 2
    End If
End Function

' Step 2. Finds the minimum and maximum calendar years across all valid
' timestamps and the dense span (max - min + 1) used to index aggregation
' arrays by offset from the minimum year.
Private Sub DetectYearSpan(ByRef p As TParsed, ByRef minYear As Long, _
                           ByRef maxYear As Long, ByRef span As Long)
    Dim i As Long, y As Long
    minYear = Year(p.ts(1))
    maxYear = minYear
    For i = 2 To p.n
        y = Year(p.ts(i))
        If y < minYear Then minYear = y
        If y > maxYear Then maxYear = y
    Next i
    span = maxYear - minYear + 1
End Sub

' Step 3. Asks how many year curves to place on each chart, re-prompting until
' the entry is an integer in 1..yearCount. Returns False only on Cancel.
Private Function AskYearsPerChart(ByVal yearCount As Long, ByRef yearsPerChart As Long) As Boolean
    On Error GoTo ErrHandler
    AskYearsPerChart = False

    Dim ans As String
    Do
        ans = InputBox( _
            "How many year curves per chart?" & vbCrLf & vbCrLf & _
            "Enter an integer between 1 and " & yearCount & ".", _
            "Step 3 - Years per chart", CStr(yearCount))

        If StrPtr(ans) = 0 Then Exit Function   ' Cancel

        ans = Trim$(ans)
        If IsNumeric(ans) Then
            Dim numVal As Double
            numVal = CDbl(ans)
            If numVal = Int(numVal) And numVal >= 1 And numVal <= yearCount Then
                yearsPerChart = CLng(numVal)
                AskYearsPerChart = True
                Exit Function
            End If
        End If

        MsgBox "Please enter a whole number between 1 and " & yearCount & ".", _
               vbExclamation, "Price Curve Charts"
    Loop

ErrHandler:
    MsgBox "Error while asking years-per-chart:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, vbCritical, "Price Curve Charts"
    AskYearsPerChart = False
End Function

' ==========================================================================
'  AGGREGATION
' ==========================================================================

' Hourly aggregation. x-axis is hour of day (0-23). See USE_EQUAL_MONTH_WEIGHTING
' for the two-stage vs single-stage choice.
'
' NOTE ON HOUR CONVENTION: the hour of day is taken from Hour(ts). Whether a
' timestamp labels the beginning or the ending of its interval is the caller's
' responsibility; an hour-ending series is shifted one position relative to an
' hour-beginning series and this code does not attempt to detect or correct it.
Private Function AggregateHourly(ByRef p As TParsed, ByVal minYear As Long, _
                                 ByVal span As Long, ByRef c As TCurves) As Boolean
    On Error GoTo ErrHandler
    AggregateHourly = False

    Const HRS As Long = 24
    Const MOS As Long = 12

    Dim yearHasData() As Boolean
    ReDim yearHasData(0 To span - 1)

    ' Temp curve, dense over the full span; compacted to present years below.
    Dim rawVals() As Double, rawPresent() As Boolean
    ReDim rawVals(0 To HRS - 1, 0 To span - 1)
    ReDim rawPresent(0 To HRS - 1, 0 To span - 1)

    Dim i As Long, yi As Long, mo As Long, hr As Long

    If USE_EQUAL_MONTH_WEIGHTING Then
        ' Stage 1: sum/count per (year, month, hour).
        Dim sYMH() As Double, cYMH() As Long
        ReDim sYMH(0 To span - 1, 0 To MOS - 1, 0 To HRS - 1)
        ReDim cYMH(0 To span - 1, 0 To MOS - 1, 0 To HRS - 1)

        For i = 1 To p.n
            yi = Year(p.ts(i)) - minYear
            mo = Month(p.ts(i)) - 1
            hr = Hour(p.ts(i))
            sYMH(yi, mo, hr) = sYMH(yi, mo, hr) + p.price(i)
            cYMH(yi, mo, hr) = cYMH(yi, mo, hr) + 1
            yearHasData(yi) = True
        Next i

        ' Stage 2: for each year and hour, average the monthly means, skipping
        ' month-hour cells with zero count so absent months are not treated as
        ' zero (equal-weighting across the months that DO have data).
        Dim m As Long, acc As Double, nMonths As Long
        For yi = 0 To span - 1
            For hr = 0 To HRS - 1
                acc = 0#
                nMonths = 0
                For m = 0 To MOS - 1
                    If cYMH(yi, m, hr) > 0 Then
                        acc = acc + (sYMH(yi, m, hr) / cYMH(yi, m, hr))
                        nMonths = nMonths + 1
                    End If
                Next m
                If nMonths > 0 Then
                    rawVals(hr, yi) = acc / nMonths
                    rawPresent(hr, yi) = True
                End If
            Next hr
        Next yi
    Else
        ' Single-stage: plain mean of every observation at that hour of day
        ' within the year.
        Dim sYH() As Double, cYH() As Long
        ReDim sYH(0 To span - 1, 0 To HRS - 1)
        ReDim cYH(0 To span - 1, 0 To HRS - 1)

        For i = 1 To p.n
            yi = Year(p.ts(i)) - minYear
            hr = Hour(p.ts(i))
            sYH(yi, hr) = sYH(yi, hr) + p.price(i)
            cYH(yi, hr) = cYH(yi, hr) + 1
            yearHasData(yi) = True
        Next i

        For yi = 0 To span - 1
            For hr = 0 To HRS - 1
                If cYH(yi, hr) > 0 Then
                    rawVals(hr, yi) = sYH(yi, hr) / cYH(yi, hr)
                    rawPresent(hr, yi) = True
                End If
            Next hr
        Next yi
    End If

    CompactCurves c, True, HRS, minYear, span, yearHasData, rawVals, rawPresent
    AggregateHourly = True
    Exit Function

ErrHandler:
    MsgBox "Error during hourly aggregation:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, vbCritical, "Price Curve Charts"
    AggregateHourly = False
End Function

' Monthly aggregation. x-axis is month (Jan-Dec). A month absent for a given
' year is left as a gap (present = False) rather than plotted as zero.
Private Function AggregateMonthly(ByRef p As TParsed, ByVal minYear As Long, _
                                  ByVal span As Long, ByRef c As TCurves) As Boolean
    On Error GoTo ErrHandler
    AggregateMonthly = False

    Const MOS As Long = 12

    Dim yearHasData() As Boolean
    ReDim yearHasData(0 To span - 1)

    Dim sYM() As Double, cYM() As Long
    ReDim sYM(0 To span - 1, 0 To MOS - 1)
    ReDim cYM(0 To span - 1, 0 To MOS - 1)

    Dim i As Long, yi As Long, mo As Long
    For i = 1 To p.n
        yi = Year(p.ts(i)) - minYear
        mo = Month(p.ts(i)) - 1
        sYM(yi, mo) = sYM(yi, mo) + p.price(i)
        cYM(yi, mo) = cYM(yi, mo) + 1
        yearHasData(yi) = True
    Next i

    Dim rawVals() As Double, rawPresent() As Boolean
    ReDim rawVals(0 To MOS - 1, 0 To span - 1)
    ReDim rawPresent(0 To MOS - 1, 0 To span - 1)

    For yi = 0 To span - 1
        For mo = 0 To MOS - 1
            If cYM(yi, mo) > 0 Then
                rawVals(mo, yi) = sYM(yi, mo) / cYM(yi, mo)
                rawPresent(mo, yi) = True
            End If
        Next mo
    Next yi

    CompactCurves c, False, MOS, minYear, span, yearHasData, rawVals, rawPresent
    AggregateMonthly = True
    Exit Function

ErrHandler:
    MsgBox "Error during monthly aggregation:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, vbCritical, "Price Curve Charts"
    AggregateMonthly = False
End Function

' Collapses a dense span-indexed curve down to only those years that carry
' data, filling the TCurves record.
Private Sub CompactCurves(ByRef c As TCurves, ByVal isHourly As Boolean, _
                          ByVal xCount As Long, ByVal minYear As Long, ByVal span As Long, _
                          ByRef yearHasData() As Boolean, _
                          ByRef rawVals() As Double, ByRef rawPresent() As Boolean)
    ' Count present years.
    Dim yc As Long, k As Long
    yc = 0
    For k = 0 To span - 1
        If yearHasData(k) Then yc = yc + 1
    Next k

    c.isHourly = isHourly
    c.xCount = xCount
    c.yearCount = yc
    If yc = 0 Then Exit Sub

    ReDim c.years(0 To yc - 1)
    ReDim c.vals(0 To xCount - 1, 0 To yc - 1)
    ReDim c.present(0 To xCount - 1, 0 To yc - 1)

    Dim dst As Long, x As Long
    dst = 0
    For k = 0 To span - 1
        If yearHasData(k) Then
            c.years(dst) = minYear + k
            For x = 0 To xCount - 1
                c.vals(x, dst) = rawVals(x, k)
                c.present(x, dst) = rawPresent(x, k)
            Next x
            dst = dst + 1
        End If
    Next k
End Sub

' ==========================================================================
'  STAGING SHEET
' ==========================================================================

' Writes the aggregated table to _CurveData so chart series bind to a range the
' user can inspect. Column A holds the x categories; each subsequent column
' holds one year's curve, headed by the four-digit year. Gaps are written as
' truly empty cells so the chart shows a break rather than a zero.
Private Function WriteStagingSheet(ByRef c As TCurves) As Boolean
    On Error GoTo ErrHandler
    WriteStagingSheet = False

    Dim ws As Worksheet
    Set ws = FreshSheet(SH_STAGING)

    Dim nRows As Long, nCols As Long
    nRows = c.xCount + 1              ' + header row
    nCols = c.yearCount + 1           ' + category column

    Dim block() As Variant
    ReDim block(1 To nRows, 1 To nCols)

    ' Header row.
    If c.isHourly Then
        block(1, 1) = "Hour"
    Else
        block(1, 1) = "Month"
    End If
    Dim j As Long
    For j = 0 To c.yearCount - 1
        block(1, j + 2) = c.years(j)
    Next j

    ' Body.
    Dim x As Long
    For x = 0 To c.xCount - 1
        block(x + 2, 1) = XLabel(c.isHourly, x)
        For j = 0 To c.yearCount - 1
            If c.present(x, j) Then
                block(x + 2, j + 2) = c.vals(x, j)
            Else
                block(x + 2, j + 2) = Empty     ' leaves the cell blank -> gap
            End If
        Next j
    Next x

    ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols)).Value2 = block
    ws.Rows(1).Font.Bold = True
    ws.Columns(1).Font.Bold = True

    WriteStagingSheet = True
    Exit Function

ErrHandler:
    MsgBox "Error while writing the staging sheet:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, vbCritical, "Price Curve Charts"
    WriteStagingSheet = False
End Function

' Category label for x index: the hour number for hourly, a month abbreviation
' for monthly.
Private Function XLabel(ByVal isHourly As Boolean, ByVal x As Long) As Variant
    If isHourly Then
        XLabel = x                       ' 0..23 as numbers
    Else
        Dim mm As Variant
        mm = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        XLabel = mm(x)
    End If
End Function

' ==========================================================================
'  CHART GENERATION
' ==========================================================================

' Creates the PriceCurves sheet and builds chartCount charts, filling years in
' chronological order. Each chart binds up to yearsPerChart series to ranges on
' _CurveData; the final chart carries the remainder.
Private Function CreateChartBatch(ByRef c As TCurves, ByVal yearsPerChart As Long, _
                                  ByVal chartCount As Long) As Boolean
    On Error GoTo ErrHandler
    CreateChartBatch = False

    Dim wsChart As Worksheet, wsData As Worksheet
    Set wsChart = FreshSheet(SH_CHARTS)
    Set wsData = ThisWorkbook.Worksheets(SH_STAGING)

    Dim dataRows As Long
    dataRows = c.xCount                  ' body rows (excludes header)
    Dim useMarkers As Boolean
    useMarkers = (c.xCount <= 12)        ' markers only when the curve is short

    Dim catRange As Range
    Set catRange = wsData.Range(wsData.Cells(2, 1), wsData.Cells(1 + dataRows, 1))

    Dim ci As Long
    For ci = 0 To chartCount - 1
        Dim startIdx As Long, endIdx As Long
        startIdx = ci * yearsPerChart
        endIdx = startIdx + yearsPerChart - 1
        If endIdx > c.yearCount - 1 Then endIdx = c.yearCount - 1   ' remainder clamp

        ' Grid placement: fill left to right, then wrap down.
        Dim gcol As Long, grow As Long
        gcol = ci Mod GRID_COLS
        grow = ci \ GRID_COLS

        Dim co As ChartObject
        Set co = wsChart.ChartObjects.Add( _
            Left:=CHART_GAP + gcol * (CHART_W + CHART_GAP), _
            Top:=CHART_GAP + grow * (CHART_H + CHART_GAP), _
            Width:=CHART_W, Height:=CHART_H)
        co.Name = "chtCurve_" & Format$(ci + 1, "00")

        Dim ch As Chart
        Set ch = co.Chart
        If useMarkers Then
            ch.ChartType = xlLineMarkers
        Else
            ch.ChartType = xlLine
        End If

        ' One explicit series per year column on this chart.
        Dim yi As Long
        For yi = startIdx To endIdx
            Dim col As Long
            col = yi + 2                 ' column A is categories
            Dim valRange As Range
            Set valRange = wsData.Range(wsData.Cells(2, col), wsData.Cells(1 + dataRows, col))

            Dim ser As Series
            Set ser = ch.SeriesCollection.NewSeries
            ser.Values = valRange
            ser.XValues = catRange
            ser.Name = CStr(c.years(yi))
            If Not useMarkers Then ser.MarkerStyle = xlMarkerStyleNone
        Next yi

        StyleChart ch, c.isHourly, c.years(startIdx), c.years(endIdx)
    Next ci

    CreateChartBatch = True
    Exit Function

ErrHandler:
    MsgBox "Error while creating charts:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, vbCritical, "Price Curve Charts"
    CreateChartBatch = False
End Function

' Applies titles, axis labels, legend, and number format to one chart.
Private Sub StyleChart(ByRef ch As Chart, ByVal isHourly As Boolean, _
                       ByVal firstYear As Long, ByVal lastYear As Long)
    Dim titleTxt As String, catTitle As String
    If isHourly Then
        titleTxt = "Average Diurnal Price Curve, " & firstYear & "-" & lastYear
        catTitle = "Hour of Day"
    Else
        titleTxt = "Monthly Average Price, " & firstYear & "-" & lastYear
        catTitle = "Month"
    End If

    ch.HasTitle = True
    ch.ChartTitle.Text = titleTxt

    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionRight

    With ch.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = catTitle
    End With

    With ch.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Price ($/MWh)"
        .TickLabels.NumberFormat = "$#,##0.00"
    End With
End Sub

' ==========================================================================
'  SHARED Y-AXIS SCALE
' ==========================================================================

' Computes one min/max across every curve in every chart, pads by 5% of the
' range on each side, rounds to a nice increment, and applies the identical
' scale to every chart so they can be compared honestly. The minimum is never
' forced to zero, so negative prices remain visible. Returns the applied bounds
' ByRef.
Private Function ApplySharedScale(ByRef c As TCurves, ByRef yLo As Double, _
                                  ByRef yHi As Double) As Boolean
    On Error GoTo ErrHandler
    ApplySharedScale = False

    Dim haveAny As Boolean, dataMin As Double, dataMax As Double
    haveAny = False

    Dim x As Long, j As Long, v As Double
    For j = 0 To c.yearCount - 1
        For x = 0 To c.xCount - 1
            If c.present(x, j) Then
                v = c.vals(x, j)
                If Not haveAny Then
                    dataMin = v: dataMax = v: haveAny = True
                Else
                    If v < dataMin Then dataMin = v
                    If v > dataMax Then dataMax = v
                End If
            End If
        Next x
    Next j

    If Not haveAny Then
        dataMin = 0#: dataMax = 1#
    End If

    ' Pad by 5% of the range on each side, then snap outward to a nice step.
    Dim rangeVal As Double, pad As Double, stepSize As Double
    rangeVal = dataMax - dataMin
    If rangeVal <= 0 Then rangeVal = Abs(dataMax) + 1#     ' flat series safety
    pad = 0.05 * rangeVal

    stepSize = NiceStep(rangeVal + 2# * pad)
    yLo = FloorTo(dataMin - pad, stepSize)
    yHi = CeilTo(dataMax + pad, stepSize)
    If yHi <= yLo Then yHi = yLo + stepSize               ' guarantee a positive span

    Dim wsChart As Worksheet
    Set wsChart = ThisWorkbook.Worksheets(SH_CHARTS)

    Dim co As ChartObject
    For Each co In wsChart.ChartObjects
        With co.Chart.Axes(xlValue)
            .MinimumScale = yLo
            .MaximumScale = yHi
        End With
    Next co

    ApplySharedScale = True
    Exit Function

ErrHandler:
    MsgBox "Error while applying the shared y-axis scale:" & vbCrLf & _
           "  #" & Err.Number & " - " & Err.Description, vbCritical, "Price Curve Charts"
    ApplySharedScale = False
End Function

' Returns a "nice" 1/2/5 x 10^k increment near range/8 (aiming for ~8
' gridlines) so the shared axis lands on round numbers.
Private Function NiceStep(ByVal rangeVal As Double) As Double
    If rangeVal <= 0 Then
        NiceStep = 1#
        Exit Function
    End If

    Dim target As Double, expo As Double, base As Double, frac As Double, niceFrac As Double
    target = rangeVal / 8#
    expo = Int(Log(target) / Log(10#))
    base = 10# ^ expo
    frac = target / base

    If frac < 1.5 Then
        niceFrac = 1#
    ElseIf frac < 3# Then
        niceFrac = 2#
    ElseIf frac < 7# Then
        niceFrac = 5#
    Else
        niceFrac = 10#
    End If

    NiceStep = niceFrac * base
End Function

' Floor toward negative infinity onto a multiple of stepSize (Int already floors).
Private Function FloorTo(ByVal x As Double, ByVal stepSize As Double) As Double
    FloorTo = Int(x / stepSize) * stepSize
End Function

' Ceiling toward positive infinity onto a multiple of stepSize.
Private Function CeilTo(ByVal x As Double, ByVal stepSize As Double) As Double
    CeilTo = -Int(-x / stepSize) * stepSize
End Function

' ==========================================================================
'  HOUSEKEEPING & REPORTING
' ==========================================================================

' Deletes any existing sheet of the given name and returns a fresh empty one.
' DisplayAlerts is already suppressed by the orchestrator, so the delete does
' not prompt.
Private Function FreshSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then ws.Delete

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    Set FreshSheet = ws
End Function

' Final summary. Flags excessive row loss (>1% of input) prominently, since
' silent skipping on price data is a data-quality signal.
Private Sub ReportCompletion(ByRef p As TParsed, ByVal isHourly As Boolean, _
                             ByRef c As TCurves, ByVal yearsPerChart As Long, _
                             ByVal chartCount As Long, ByVal yLo As Double, ByVal yHi As Double)
    Dim resTxt As String
    If isHourly Then resTxt = "Hourly (hour of day)" Else resTxt = "Monthly"

    Dim msg As String
    msg = "Price curve charts complete." & vbCrLf & vbCrLf & _
          "Rows read:        " & p.rowsRead & vbCrLf & _
          "Rows skipped:     " & p.skipped & vbCrLf & _
          "Resolution:       " & resTxt & vbCrLf & _
          "Years detected:   " & c.yearCount & " (" & c.years(0) & "-" & _
          c.years(c.yearCount - 1) & ")" & vbCrLf & _
          "Years per chart:  " & yearsPerChart & vbCrLf & _
          "Charts created:   " & chartCount & vbCrLf & _
          "Shared y-axis:    " & Format$(yLo, "$#,##0.00") & " to " & _
          Format$(yHi, "$#,##0.00")

    Dim icon As VbMsgBoxStyle
    icon = vbInformation
    If p.rowsRead > 0 Then
        If p.skipped > 0.01 * p.rowsRead Then
            icon = vbExclamation
            msg = msg & vbCrLf & vbCrLf & _
                  "!! WARNING: " & Format$(p.skipped / p.rowsRead, "0.0%") & _
                  " of input rows were skipped as unparseable." & vbCrLf & _
                  "   On price data this is a data-quality signal worth checking."
        End If
    End If

    MsgBox msg, icon, "Price Curve Charts"
End Sub
