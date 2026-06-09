Attribute VB_Name = "MatchParcels"
Option Explicit

' ============================================================
'  CONFIG — only this block needs to change if headers or
'  sheet names differ; nothing else below requires editing
' ============================================================

' Sheet (tab) names — must match exactly, including spaces
Private Const SH_PARCELS      As String = "Sub-Only Proximal (5mi)"
Private Const SH_SUBSTATIONS  As String = "Substations in Market Footprint"
Private Const SH_LMP_HIST     As String = "LMP 5-yr Historical Metrics"
Private Const SH_LMP_FCST5    As String = "LMP 5-yr Forecast Metrics"
Private Const SH_LMP_FCST20   As String = "LMP 20-yr Forecast Metrics"
Private Const SH_OUTPUT       As String = "Parcel_Matches"

' Header row number (1-based, same on all sheets)
Private Const HDR_ROW         As Long = 1

' Parcel column headers
Private Const HDR_PARCEL_ID   As String = "uuid"
Private Const HDR_PARCEL_LAT  As String = "Latitude"
Private Const HDR_PARCEL_LON  As String = "Longitude"

' Substation column headers
Private Const HDR_SUB_ID      As String = "ENVSubstationID"
Private Const HDR_SUB_LAT     As String = "Latitude"
Private Const HDR_SUB_LON     As String = "Longitude"

' LMP node column headers (identical on all three LMP tabs)
Private Const HDR_LMP_ID      As String = "ENVNodeID"
Private Const HDR_LMP_LAT     As String = "y"
Private Const HDR_LMP_LON     As String = "x"

' Earth radius for HaversineMiles.
' Change to 6371.0 to get kilometres everywhere; also update MAX_MATCH_MILES.
Private Const EARTH_RADIUS_MI As Double = 3958.8

' Maximum allowable match distance (same units as EARTH_RADIUS_MI).
' If the nearest result exceeds this, the ID cell shows "No match within X mi".
' Set to 0 to disable the check entirely.
Private Const MAX_MATCH_MILES As Double = 0

' ============================================================
'  MAIN
' ============================================================

Public Sub MatchParcels()

    ' ── Stage 1: Resolve worksheet references ────────────────
    Dim wsParcels As Worksheet, wsSubs  As Worksheet
    Dim wsHist    As Worksheet, wsFcst5 As Worksheet
    Dim wsFcst20  As Worksheet, wsOut   As Worksheet

    On Error GoTo ErrSheets
    Set wsParcels = ThisWorkbook.Sheets(SH_PARCELS)
    Set wsSubs    = ThisWorkbook.Sheets(SH_SUBSTATIONS)
    Set wsHist    = ThisWorkbook.Sheets(SH_LMP_HIST)
    Set wsFcst5   = ThisWorkbook.Sheets(SH_LMP_FCST5)
    Set wsFcst20  = ThisWorkbook.Sheets(SH_LMP_FCST20)
    On Error GoTo 0

    ' ── Stage 2: Pull every used range into memory at once ───
    Dim aParcels As Variant: aParcels = wsParcels.UsedRange.Value
    Dim aSubs    As Variant: aSubs    = wsSubs.UsedRange.Value
    Dim aHist    As Variant: aHist    = wsHist.UsedRange.Value
    Dim aFcst5   As Variant: aFcst5   = wsFcst5.UsedRange.Value
    Dim aFcst20  As Variant: aFcst20  = wsFcst20.UsedRange.Value

    ' ── Stage 3: Locate required columns by header text ──────
    Dim pIdCol   As Long, pLatCol  As Long, pLonCol  As Long
    Dim sIdCol   As Long, sLatCol  As Long, sLonCol  As Long
    Dim hIdCol   As Long, hLatCol  As Long, hLonCol  As Long
    Dim f5IdCol  As Long, f5LatCol As Long, f5LonCol As Long
    Dim f20IdCol As Long, f20LatCol As Long, f20LonCol As Long

    pIdCol    = FindCol(aParcels, HDR_ROW, HDR_PARCEL_ID)
    pLatCol   = FindCol(aParcels, HDR_ROW, HDR_PARCEL_LAT)
    pLonCol   = FindCol(aParcels, HDR_ROW, HDR_PARCEL_LON)
    sIdCol    = FindCol(aSubs,    HDR_ROW, HDR_SUB_ID)
    sLatCol   = FindCol(aSubs,    HDR_ROW, HDR_SUB_LAT)
    sLonCol   = FindCol(aSubs,    HDR_ROW, HDR_SUB_LON)
    hIdCol    = FindCol(aHist,    HDR_ROW, HDR_LMP_ID)
    hLatCol   = FindCol(aHist,    HDR_ROW, HDR_LMP_LAT)
    hLonCol   = FindCol(aHist,    HDR_ROW, HDR_LMP_LON)
    f5IdCol   = FindCol(aFcst5,   HDR_ROW, HDR_LMP_ID)
    f5LatCol  = FindCol(aFcst5,   HDR_ROW, HDR_LMP_LAT)
    f5LonCol  = FindCol(aFcst5,   HDR_ROW, HDR_LMP_LON)
    f20IdCol  = FindCol(aFcst20,  HDR_ROW, HDR_LMP_ID)
    f20LatCol = FindCol(aFcst20,  HDR_ROW, HDR_LMP_LAT)
    f20LonCol = FindCol(aFcst20,  HDR_ROW, HDR_LMP_LON)

    Dim missing As String: missing = ""
    If pIdCol    = 0 Then missing = missing & "  Parcels: """ & HDR_PARCEL_ID & """" & vbCrLf
    If pLatCol   = 0 Then missing = missing & "  Parcels: """ & HDR_PARCEL_LAT & """" & vbCrLf
    If pLonCol   = 0 Then missing = missing & "  Parcels: """ & HDR_PARCEL_LON & """" & vbCrLf
    If sIdCol    = 0 Then missing = missing & "  Substations: """ & HDR_SUB_ID & """" & vbCrLf
    If sLatCol   = 0 Then missing = missing & "  Substations: """ & HDR_SUB_LAT & """" & vbCrLf
    If sLonCol   = 0 Then missing = missing & "  Substations: """ & HDR_SUB_LON & """" & vbCrLf
    If hIdCol    = 0 Then missing = missing & "  LMP Hist: """ & HDR_LMP_ID & """" & vbCrLf
    If hLatCol   = 0 Then missing = missing & "  LMP Hist: """ & HDR_LMP_LAT & """" & vbCrLf
    If hLonCol   = 0 Then missing = missing & "  LMP Hist: """ & HDR_LMP_LON & """" & vbCrLf
    If f5IdCol   = 0 Then missing = missing & "  LMP Fcst5yr: """ & HDR_LMP_ID & """" & vbCrLf
    If f5LatCol  = 0 Then missing = missing & "  LMP Fcst5yr: """ & HDR_LMP_LAT & """" & vbCrLf
    If f5LonCol  = 0 Then missing = missing & "  LMP Fcst5yr: """ & HDR_LMP_LON & """" & vbCrLf
    If f20IdCol  = 0 Then missing = missing & "  LMP Fcst20yr: """ & HDR_LMP_ID & """" & vbCrLf
    If f20LatCol = 0 Then missing = missing & "  LMP Fcst20yr: """ & HDR_LMP_LAT & """" & vbCrLf
    If f20LonCol = 0 Then missing = missing & "  LMP Fcst20yr: """ & HDR_LMP_LON & """" & vbCrLf
    If Len(missing) > 0 Then
        MsgBox "Header(s) not found — update the Config block:" & vbCrLf & missing, _
               vbCritical, "MatchParcels"
        Exit Sub
    End If

    ' ── Stage 4: Load substations into parallel typed arrays ─
    ' Using typed Double arrays for lat/lon avoids per-iteration CVar overhead
    Dim rr As Long, tLat As Double, tLon As Double
    Dim validSubs As Long: validSubs = 0
    Dim subIds()  As Variant
    Dim subLats() As Double
    Dim subLons() As Double
    Dim nSubRows  As Long: nSubRows = UBound(aSubs, 1) - HDR_ROW

    If nSubRows > 0 Then
        ReDim subIds(1 To nSubRows)
        ReDim subLats(1 To nSubRows)
        ReDim subLons(1 To nSubRows)
        For rr = HDR_ROW + 1 To UBound(aSubs, 1)
            If IsValidCoord(aSubs(rr, sLatCol), aSubs(rr, sLonCol), tLat, tLon) Then
                validSubs = validSubs + 1
                subIds(validSubs)  = aSubs(rr, sIdCol)
                subLats(validSubs) = tLat
                subLons(validSubs) = tLon
            End If
        Next rr
    End If

    ' ── Stage 5: Load LMP node arrays (one set per tab) ──────
    Dim validHist  As Long: validHist = 0
    Dim histIds()  As Variant
    Dim histLats() As Double
    Dim histLons() As Double

    If UBound(aHist, 1) > HDR_ROW Then
        ReDim histIds(1 To UBound(aHist, 1) - HDR_ROW)
        ReDim histLats(1 To UBound(aHist, 1) - HDR_ROW)
        ReDim histLons(1 To UBound(aHist, 1) - HDR_ROW)
        For rr = HDR_ROW + 1 To UBound(aHist, 1)
            If IsValidCoord(aHist(rr, hLatCol), aHist(rr, hLonCol), tLat, tLon) Then
                validHist = validHist + 1
                histIds(validHist)  = aHist(rr, hIdCol)
                histLats(validHist) = tLat
                histLons(validHist) = tLon
            End If
        Next rr
    End If

    Dim validFcst5 As Long: validFcst5 = 0
    Dim fcst5Ids()  As Variant
    Dim fcst5Lats() As Double
    Dim fcst5Lons() As Double

    If UBound(aFcst5, 1) > HDR_ROW Then
        ReDim fcst5Ids(1 To UBound(aFcst5, 1) - HDR_ROW)
        ReDim fcst5Lats(1 To UBound(aFcst5, 1) - HDR_ROW)
        ReDim fcst5Lons(1 To UBound(aFcst5, 1) - HDR_ROW)
        For rr = HDR_ROW + 1 To UBound(aFcst5, 1)
            If IsValidCoord(aFcst5(rr, f5LatCol), aFcst5(rr, f5LonCol), tLat, tLon) Then
                validFcst5 = validFcst5 + 1
                fcst5Ids(validFcst5)  = aFcst5(rr, f5IdCol)
                fcst5Lats(validFcst5) = tLat
                fcst5Lons(validFcst5) = tLon
            End If
        Next rr
    End If

    Dim validFcst20 As Long: validFcst20 = 0
    Dim fcst20Ids()  As Variant
    Dim fcst20Lats() As Double
    Dim fcst20Lons() As Double

    If UBound(aFcst20, 1) > HDR_ROW Then
        ReDim fcst20Ids(1 To UBound(aFcst20, 1) - HDR_ROW)
        ReDim fcst20Lats(1 To UBound(aFcst20, 1) - HDR_ROW)
        ReDim fcst20Lons(1 To UBound(aFcst20, 1) - HDR_ROW)
        For rr = HDR_ROW + 1 To UBound(aFcst20, 1)
            If IsValidCoord(aFcst20(rr, f20LatCol), aFcst20(rr, f20LonCol), tLat, tLon) Then
                validFcst20 = validFcst20 + 1
                fcst20Ids(validFcst20)  = aFcst20(rr, f20IdCol)
                fcst20Lats(validFcst20) = tLat
                fcst20Lons(validFcst20) = tLon
            End If
        Next rr
    End If

    ' ── Stage 6: Allocate output array (header row + one per parcel) ──
    Dim nParcels As Long: nParcels = UBound(aParcels, 1) - HDR_ROW
    Dim aOut()   As Variant
    ReDim aOut(1 To nParcels + 1, 1 To 9)

    aOut(1, 1) = "uuid"
    aOut(1, 2) = "Closest_SubstationID"
    aOut(1, 3) = "Parcel_to_Sub_Dist_mi"
    aOut(1, 4) = "Hist_NodeID"
    aOut(1, 5) = "Sub_to_Hist_Dist_mi"
    aOut(1, 6) = "Fcst5yr_NodeID"
    aOut(1, 7) = "Sub_to_Fcst5yr_Dist_mi"
    aOut(1, 8) = "Fcst20yr_NodeID"
    aOut(1, 9) = "Sub_to_Fcst20yr_Dist_mi"

    ' ── Stage 7: Loop every parcel and perform three-level match ──
    Dim matchedCount As Long: matchedCount = 0
    Dim skippedCount As Long: skippedCount = 0
    Dim outRow       As Long: outRow = 1   ' row 1 is the header

    Dim pRow         As Long
    Dim pLat         As Double, pLon        As Double
    Dim k            As Long,   d           As Double
    Dim bestSubDist  As Double, bestSubLat  As Double, bestSubLon As Double
    Dim bestSubId    As Variant
    Dim bestHistDist As Double, bestHistId  As Variant
    Dim bestF5Dist   As Double, bestF5Id    As Variant
    Dim bestF20Dist  As Double, bestF20Id   As Variant

    For pRow = HDR_ROW + 1 To UBound(aParcels, 1)

        ' Validate parcel lat/lon; skip the row if bad or missing
        If Not IsValidCoord(aParcels(pRow, pLatCol), aParcels(pRow, pLonCol), pLat, pLon) Then
            skippedCount = skippedCount + 1
            GoTo NextParcel
        End If

        outRow = outRow + 1
        aOut(outRow, 1) = aParcels(pRow, pIdCol)

        ' ── Nearest substation ───────────────────────────────
        bestSubDist = 1E+308
        bestSubId   = Empty
        bestSubLat  = 0
        bestSubLon  = 0
        For k = 1 To validSubs
            d = HaversineMiles(pLat, pLon, subLats(k), subLons(k))
            If d < bestSubDist Then
                bestSubDist = d
                bestSubId   = subIds(k)
                bestSubLat  = subLats(k)
                bestSubLon  = subLons(k)
            End If
        Next k

        If IsEmpty(bestSubId) Then
            aOut(outRow, 2) = "No substations available"
            aOut(outRow, 3) = ""
            aOut(outRow, 4) = "N/A": aOut(outRow, 5) = ""
            aOut(outRow, 6) = "N/A": aOut(outRow, 7) = ""
            aOut(outRow, 8) = "N/A": aOut(outRow, 9) = ""
            matchedCount = matchedCount + 1
            GoTo NextParcel
        End If

        If MAX_MATCH_MILES > 0 And bestSubDist > MAX_MATCH_MILES Then
            aOut(outRow, 2) = "No match within " & MAX_MATCH_MILES & " mi"
            aOut(outRow, 3) = bestSubDist
            aOut(outRow, 4) = "N/A (substation out of range)": aOut(outRow, 5) = ""
            aOut(outRow, 6) = "N/A (substation out of range)": aOut(outRow, 7) = ""
            aOut(outRow, 8) = "N/A (substation out of range)": aOut(outRow, 9) = ""
            matchedCount = matchedCount + 1
            GoTo NextParcel
        End If

        aOut(outRow, 2) = bestSubId
        aOut(outRow, 3) = bestSubDist

        ' ── Nearest Historical LMP node to the substation ────
        bestHistDist = 1E+308
        bestHistId   = Empty
        For k = 1 To validHist
            d = HaversineMiles(bestSubLat, bestSubLon, histLats(k), histLons(k))
            If d < bestHistDist Then
                bestHistDist = d
                bestHistId   = histIds(k)
            End If
        Next k

        If IsEmpty(bestHistId) Then
            aOut(outRow, 4) = "No nodes available": aOut(outRow, 5) = ""
        ElseIf MAX_MATCH_MILES > 0 And bestHistDist > MAX_MATCH_MILES Then
            aOut(outRow, 4) = "No match within " & MAX_MATCH_MILES & " mi"
            aOut(outRow, 5) = bestHistDist
        Else
            aOut(outRow, 4) = bestHistId
            aOut(outRow, 5) = bestHistDist
        End If

        ' ── Nearest 5yr Forecast LMP node to the substation ──
        bestF5Dist = 1E+308
        bestF5Id   = Empty
        For k = 1 To validFcst5
            d = HaversineMiles(bestSubLat, bestSubLon, fcst5Lats(k), fcst5Lons(k))
            If d < bestF5Dist Then
                bestF5Dist = d
                bestF5Id   = fcst5Ids(k)
            End If
        Next k

        If IsEmpty(bestF5Id) Then
            aOut(outRow, 6) = "No nodes available": aOut(outRow, 7) = ""
        ElseIf MAX_MATCH_MILES > 0 And bestF5Dist > MAX_MATCH_MILES Then
            aOut(outRow, 6) = "No match within " & MAX_MATCH_MILES & " mi"
            aOut(outRow, 7) = bestF5Dist
        Else
            aOut(outRow, 6) = bestF5Id
            aOut(outRow, 7) = bestF5Dist
        End If

        ' ── Nearest 20yr Forecast LMP node to the substation ─
        bestF20Dist = 1E+308
        bestF20Id   = Empty
        For k = 1 To validFcst20
            d = HaversineMiles(bestSubLat, bestSubLon, fcst20Lats(k), fcst20Lons(k))
            If d < bestF20Dist Then
                bestF20Dist = d
                bestF20Id   = fcst20Ids(k)
            End If
        Next k

        If IsEmpty(bestF20Id) Then
            aOut(outRow, 8) = "No nodes available": aOut(outRow, 9) = ""
        ElseIf MAX_MATCH_MILES > 0 And bestF20Dist > MAX_MATCH_MILES Then
            aOut(outRow, 8) = "No match within " & MAX_MATCH_MILES & " mi"
            aOut(outRow, 9) = bestF20Dist
        Else
            aOut(outRow, 8) = bestF20Id
            aOut(outRow, 9) = bestF20Dist
        End If

        matchedCount = matchedCount + 1

NextParcel:
    Next pRow

    ' ── Stage 8: Write results to Parcel_Matches ─────────────
    Dim ws2 As Worksheet
    Dim shExists As Boolean: shExists = False
    For Each ws2 In ThisWorkbook.Sheets
        If ws2.Name = SH_OUTPUT Then
            shExists = True
            Set wsOut = ws2
            Exit For
        End If
    Next ws2

    If Not shExists Then
        Set wsOut = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOut.Name = SH_OUTPUT
    Else
        wsOut.Cells.Clear
    End If

    ' Trim output array to rows actually written, then paste in one shot
    Dim aFinal() As Variant
    ReDim aFinal(1 To outRow, 1 To 9)
    Dim r As Long, c As Long
    For r = 1 To outRow
        For c = 1 To 9
            aFinal(r, c) = aOut(r, c)
        Next c
    Next r

    wsOut.Range("A1").Resize(outRow, 9).Value = aFinal
    wsOut.Rows(1).Font.Bold = True
    wsOut.Columns("A:I").AutoFit

    ' ── Stage 9: Summary dialog ──────────────────────────────
    MsgBox "MatchParcels complete." & vbCrLf & vbCrLf & _
           "Parcels processed:               " & matchedCount & vbCrLf & _
           "Rows skipped (bad/missing coords): " & skippedCount, _
           vbInformation, "MatchParcels"

    Exit Sub

ErrSheets:
    MsgBox "Sheet not found: " & Err.Description & vbCrLf & _
           "Check the sheet name constants in the Config block.", _
           vbCritical, "MatchParcels"
End Sub

' ============================================================
'  HaversineMiles — great-circle distance in miles.
'  Inputs are decimal-degree lat/lon; formula converts to radians.
'  Swap EARTH_RADIUS_MI to 6371.0 to get km.
' ============================================================

Public Function HaversineMiles(lat1 As Double, lon1 As Double, _
                                lat2 As Double, lon2 As Double) As Double
    Const PI As Double = 3.14159265358979

    Dim dLat  As Double: dLat  = (lat2 - lat1) * PI / 180#
    Dim dLon  As Double: dLon  = (lon2 - lon1) * PI / 180#
    Dim rlat1 As Double: rlat1 = lat1 * PI / 180#
    Dim rlat2 As Double: rlat2 = lat2 * PI / 180#

    Dim a As Double
    a = Sin(dLat / 2) * Sin(dLat / 2) + _
        Cos(rlat1) * Cos(rlat2) * Sin(dLon / 2) * Sin(dLon / 2)

    ' Clamp to [0,1] to guard against floating-point drift at identical points
    If a > 1 Then a = 1
    If a < 0 Then a = 0

    ' Antipodal guard: a=1 makes Sqr(1-a)=0 (division by zero)
    If a >= 1 Then
        HaversineMiles = EARTH_RADIUS_MI * PI
    Else
        HaversineMiles = EARTH_RADIUS_MI * 2 * Atn(Sqr(a) / Sqr(1 - a))
    End If
End Function

' ============================================================
'  FindCol — 1-based column index of hdrText in the given row
'  of a 2-D Variant array. Returns 0 if not found.
'  Match is case-insensitive and ignores leading/trailing spaces.
' ============================================================

Private Function FindCol(arr As Variant, hdrRow As Long, hdrText As String) As Long
    Dim j      As Long
    Dim target As String: target = LCase(Trim(hdrText))
    For j = LBound(arr, 2) To UBound(arr, 2)
        If LCase(Trim(CStr(arr(hdrRow, j)))) = target Then
            FindCol = j
            Exit Function
        End If
    Next j
    FindCol = 0
End Function

' ============================================================
'  IsValidCoord — True when latVal/lonVal are numeric, non-null,
'  non-empty, non-error, and within ±90/±180 decimal-degree ranges.
'  Sets outLat and outLon on success; both zeroed on failure.
' ============================================================

Private Function IsValidCoord(latVal As Variant, lonVal As Variant, _
                               outLat As Double, outLon As Double) As Boolean
    IsValidCoord = False
    outLat = 0: outLon = 0

    If IsError(latVal)  Or IsError(lonVal)  Then Exit Function
    If IsEmpty(latVal)  Or IsEmpty(lonVal)  Then Exit Function
    If IsNull(latVal)   Or IsNull(lonVal)   Then Exit Function
    If Not IsNumeric(latVal) Or Not IsNumeric(lonVal) Then Exit Function

    Dim lat As Double: lat = CDbl(latVal)
    Dim lon As Double: lon = CDbl(lonVal)

    If lat < -90  Or lat > 90  Then Exit Function
    If lon < -180 Or lon > 180 Then Exit Function

    outLat = lat
    outLon = lon
    IsValidCoord = True
End Function
