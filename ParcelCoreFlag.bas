Attribute VB_Name = "ParcelCoreFlag"
'===============================================================================
' ParcelCoreFlag.bas
' Flags each parcel 1 (core/interior) or 0 based on two conditions:
'   (a) the parcel has at least one contiguous neighbor in each of N, S, E, W;
'   (b) the parcel's connected block totals >= MIN_BLOCK_ACRES.
' Rows with bad data receive a blank rather than a 0.
'
' CAVEAT: centroid + acreage only approximates contiguity and direction. It
' assumes parcels are reasonably compact and near-square. Long, thin, or
' irregular parcels can misclassify. True adjacency and direction require
' actual parcel polygons (GIS). This is a fast screen, not ground truth.
'
' HOW TO IMPORT AND RUN:
'   1. Press Alt+F11, right-click workbook > Import File > select this .bas.
'   2. Close the editor, press Alt+F8, select FlagCoreParcels, click Run.
'   3. Open the Immediate window (Ctrl+G) before running to see debug output.
'===============================================================================
Option Explicit

' -----------------------------------------------------------------------
' CONFIG BLOCK — tune these constants for your dataset.
' Columns and data-start row are chosen at runtime via prompts.
' -----------------------------------------------------------------------

' A parcel's connected block must total at least this many acres for
' condition (b) to pass.
Private Const MIN_BLOCK_ACRES As Double = 1000#

' Contiguity tolerance multiplier.
'   1.0 = exactly edge-sharing: two adjacent squares of sides s1, s2 have
'         centroid distance s1/2 + s2/2, so TOLERANCE=1.0 accepts exactly
'         that distance. Raise toward 1.2–1.4 to also catch near-misses.
Private Const TOLERANCE As Double = 1.0

' Print diagnostics to the Immediate window after the run?
Private Const DEBUG_MODE As Boolean = True

' -----------------------------------------------------------------------
' Unit and geometry constants — do not change.
' -----------------------------------------------------------------------
Private Const ACRES_TO_M2  As Double = 4046.86
Private Const EARTH_RADIUS As Double = 6371000#
Private Const PI           As Double = 3.14159265358979

' Cardinal-direction bitmasks packed into one Long per parcel.
Private Const DIR_NORTH As Long = 1
Private Const DIR_SOUTH As Long = 2
Private Const DIR_EAST  As Long = 4
Private Const DIR_WEST  As Long = 8
Private Const DIR_ALL   As Long = 15   ' N OR S OR E OR W = all four

'===============================================================================
' Main entry point
'===============================================================================
Public Sub FlagCoreParcels()

    ' -----------------------------------------------------------------------
    ' STAGE 1: Collect runtime inputs — data-start row and column picks.
    ' -----------------------------------------------------------------------

    Dim ws        As Worksheet
    Dim rPick     As Range
    Dim colID     As Long
    Dim colLat    As Long
    Dim colLon    As Long
    Dim colAcre   As Long
    Dim colOut    As Long
    Dim dataStart As Long
    Dim userInput As String

    userInput = InputBox( _
        "Enter the row number where parcel DATA starts." & vbCrLf & _
        "(Row 1 is typically the header row, so data starts at row 2.)", _
        "Data Start Row", "2")
    If userInput = "" Then
        MsgBox "Cancelled.", vbInformation, "FlagCoreParcels"
        Exit Sub
    End If
    If Not IsNumeric(userInput) Or CLng(userInput) < 1 Then
        MsgBox "Invalid row number. Macro cancelled.", vbExclamation, "FlagCoreParcels"
        Exit Sub
    End If
    dataStart = CLng(userInput)

    On Error GoTo UserCancelled

    Set rPick = Application.InputBox( _
        "Click any cell in the PARCEL ID column.", "Select Parcel ID Column", Type:=8)
    Set ws = rPick.Worksheet
    colID = rPick.Column

    Set rPick = Application.InputBox( _
        "Click any cell in the LATITUDE column (decimal degrees).", _
        "Select Latitude Column", Type:=8)
    If Not rPick.Worksheet Is ws Then GoTo WrongSheet
    colLat = rPick.Column

    Set rPick = Application.InputBox( _
        "Click any cell in the LONGITUDE column (decimal degrees).", _
        "Select Longitude Column", Type:=8)
    If Not rPick.Worksheet Is ws Then GoTo WrongSheet
    colLon = rPick.Column

    Set rPick = Application.InputBox( _
        "Click any cell in the ACREAGE column.", "Select Acreage Column", Type:=8)
    If Not rPick.Worksheet Is ws Then GoTo WrongSheet
    colAcre = rPick.Column

    Set rPick = Application.InputBox( _
        "Click any cell in the OUTPUT column (the 1/0 flag will be written here).", _
        "Select Output Column", Type:=8)
    If Not rPick.Worksheet Is ws Then GoTo WrongSheet
    colOut = rPick.Column

    On Error GoTo 0

    ' -----------------------------------------------------------------------
    ' STAGE 2: Determine the data extent on the chosen sheet.
    ' -----------------------------------------------------------------------

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colID).End(xlUp).Row
    If lastRow < dataStart Then
        MsgBox "No data rows found at or below row " & dataStart & ". Macro cancelled.", _
               vbExclamation, "FlagCoreParcels"
        Exit Sub
    End If

    Dim nRows As Long
    nRows = lastRow - dataStart + 1

    ' -----------------------------------------------------------------------
    ' STAGE 3: Read each needed column into memory — one Range.Value call each.
    ' -----------------------------------------------------------------------

    Dim arrID()   As Variant
    Dim arrLat()  As Variant
    Dim arrLon()  As Variant
    Dim arrAcre() As Variant

    arrID   = ws.Range(ws.Cells(dataStart, colID),   ws.Cells(lastRow, colID)).Value
    arrLat  = ws.Range(ws.Cells(dataStart, colLat),  ws.Cells(lastRow, colLat)).Value
    arrLon  = ws.Range(ws.Cells(dataStart, colLon),  ws.Cells(lastRow, colLon)).Value
    arrAcre = ws.Range(ws.Cells(dataStart, colAcre), ws.Cells(lastRow, colAcre)).Value

    ' -----------------------------------------------------------------------
    ' STAGE 4: Guard — warn before overwriting the output column.
    ' -----------------------------------------------------------------------

    Dim outRange As Range
    Set outRange = ws.Range(ws.Cells(dataStart, colOut), ws.Cells(lastRow, colOut))
    If Application.WorksheetFunction.CountA(outRange) > 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox( _
            "The output column already contains data in rows " & _
            dataStart & ":" & lastRow & "." & vbCrLf & vbCrLf & _
            "Do you want to OVERWRITE it?", _
            vbQuestion + vbYesNo, "Output Column Not Empty")
        If ans = vbNo Then
            MsgBox "Macro cancelled — output column left untouched.", _
                   vbInformation, "FlagCoreParcels"
            Exit Sub
        End If
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 5: Validate rows; build compact in-memory working arrays.
    '
    ' Only valid rows enter the algorithm. srcRow(k) maps compact index k
    ' back to the 1-based position within arrLat/etc. (= sheet row
    ' dataStart + srcRow(k) - 1). Skipped rows stay blank in the output.
    ' -----------------------------------------------------------------------

    Dim latArr()  As Double
    Dim lonArr()  As Double
    Dim acreArr() As Double
    Dim heArr()   As Double   ' half-extent = sqrt(area_m2) / 2  (half-side)
    Dim srcRow()  As Long

    ReDim latArr(1 To nRows)
    ReDim lonArr(1 To nRows)
    ReDim acreArr(1 To nRows)
    ReDim heArr(1 To nRows)
    ReDim srcRow(1 To nRows)

    Dim n       As Long
    Dim skipped As Long
    Dim k       As Long
    Dim latV    As Double, lonV As Double, acreV As Double, areaM2 As Double
    Dim firstOK As Boolean

    For k = 1 To nRows

        If IsEmpty(arrLat(k, 1)) Or Not IsNumeric(arrLat(k, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        latV = CDbl(arrLat(k, 1))
        If latV < -90# Or latV > 90# Then skipped = skipped + 1: GoTo NextRow

        If IsEmpty(arrLon(k, 1)) Or Not IsNumeric(arrLon(k, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        lonV = CDbl(arrLon(k, 1))
        If lonV < -180# Or lonV > 180# Then skipped = skipped + 1: GoTo NextRow

        If IsEmpty(arrAcre(k, 1)) Or Not IsNumeric(arrAcre(k, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        acreV = CDbl(arrAcre(k, 1))
        If acreV <= 0# Then skipped = skipped + 1: GoTo NextRow

        n = n + 1
        srcRow(n)  = k
        latArr(n)  = latV
        lonArr(n)  = lonV
        acreArr(n) = acreV
        areaM2     = acreV * ACRES_TO_M2
        heArr(n)   = Sqr(areaM2) / 2#   ' half-side of the equivalent square

        If DEBUG_MODE And Not firstOK Then
            firstOK = True
            Debug.Print "--- FlagCoreParcels sanity (first valid parcel) ---"
            Debug.Print "  Parcel ID  : " & CStr(arrID(k, 1))
            Debug.Print "  Acreage    : " & acreV & " acres"
            Debug.Print "  Area       : " & Format(areaM2, "#,##0.0") & " m²"
            Debug.Print "  Width (s)  : " & Format(Sqr(areaM2), "#,##0.0") & " m" & _
                        "   (expect ~100–1 000 m; if >10 000 m, check acreage units)"
        End If

NextRow:
    Next k

    If n = 0 Then
        MsgBox "No valid parcel rows found. Check your column selections and data.", _
               vbExclamation, "FlagCoreParcels"
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 6: Initialise union-find and direction-coverage arrays.
    '
    ' ufParent(i) / ufRank(i) — standard union-find for connected blocks.
    ' dirBits(i)  — bitmask of the cardinal directions covered by parcel i's
    '               contiguous neighbours. Set per-bit during Stage 7 and
    '               tested in Stage 9: (dirBits(i) AND DIR_ALL) = DIR_ALL
    '               means all four cardinals are covered.
    ' -----------------------------------------------------------------------

    Dim ufParent() As Long
    Dim ufRank()   As Long
    Dim dirBits()  As Long
    ReDim ufParent(1 To n)
    ReDim ufRank(1 To n)
    ReDim dirBits(1 To n)

    Dim i As Long, j As Long
    For i = 1 To n
        ufParent(i) = i
        ufRank(i)   = 0
        dirBits(i)  = 0
    Next i

    ' -----------------------------------------------------------------------
    ' STAGE 7: O(n²) contiguity pass — merge blocks and accumulate directions.
    '
    ' For each ordered pair (i, j) with j > i:
    '   1. Compute the Haversine centroid-to-centroid distance.
    '   2. If dist <= TOLERANCE * (he_i + he_j), the parcels are contiguous:
    '        • Merge their connected blocks with UnionSets.
    '        • Compute the north-south and east-west offsets in metres:
    '            ns_m = (latJ - latI) * 111320
    '            ew_m = (lonJ - lonI) * 111320 * cos(latI)
    '          The larger-magnitude axis determines whether j lies to the
    '          N/S or E/W of i (±45° quadrant rule). Both i's and j's
    '          dirBits are updated — i sees j in some direction, j sees i
    '          in the opposite direction.
    '
    ' cosLatI is precomputed outside the inner loop to avoid repeating the
    ' trig call n times for the same parcel i.
    ' -----------------------------------------------------------------------

    Dim dist    As Double
    Dim ns_m    As Double, ew_m As Double
    Dim cosLatI As Double

    For i = 1 To n - 1
        cosLatI = Cos(latArr(i) * PI / 180#)

        For j = i + 1 To n
            dist = HaversineMeters(latArr(i), lonArr(i), latArr(j), lonArr(j))

            If dist <= TOLERANCE * (heArr(i) + heArr(j)) Then
                ' Merge the two parcels into one connected block.
                Call UnionSets(ufParent, ufRank, i, j)

                ' Determine which cardinal direction j lies in from i.
                ns_m = (latArr(j) - latArr(i)) * 111320#
                ew_m = (lonArr(j) - lonArr(i)) * 111320# * cosLatI

                If Abs(ns_m) >= Abs(ew_m) Then
                    ' North-South is the dominant axis.
                    If ns_m >= 0# Then
                        dirBits(i) = dirBits(i) Or DIR_NORTH   ' j is north of i
                        dirBits(j) = dirBits(j) Or DIR_SOUTH   ' i is south of j
                    Else
                        dirBits(i) = dirBits(i) Or DIR_SOUTH   ' j is south of i
                        dirBits(j) = dirBits(j) Or DIR_NORTH   ' i is north of j
                    End If
                Else
                    ' East-West is the dominant axis.
                    If ew_m >= 0# Then
                        dirBits(i) = dirBits(i) Or DIR_EAST    ' j is east of i
                        dirBits(j) = dirBits(j) Or DIR_WEST    ' i is west of j
                    Else
                        dirBits(i) = dirBits(i) Or DIR_WEST    ' j is west of i
                        dirBits(j) = dirBits(j) Or DIR_EAST    ' i is east of j
                    End If
                End If
            End If
        Next j
    Next i

    ' -----------------------------------------------------------------------
    ' STAGE 8: Compute connected-block acreage totals.
    '
    ' After all unions are complete, walk every valid parcel, find its root,
    ' and accumulate its acreage into rootAcres(root). This is the one-pass
    ' efficient method — O(n) after the O(n²) contiguity pass above.
    ' Each parcel's block total is then rootAcres(FindRoot(ufParent, i)).
    ' -----------------------------------------------------------------------

    Dim rootAcres() As Double
    ReDim rootAcres(1 To n)   ' sparse: only entries at root indices matter

    Dim r As Long
    For i = 1 To n
        r = FindRoot(ufParent, i)
        rootAcres(r) = rootAcres(r) + acreArr(i)
    Next i

    ' -----------------------------------------------------------------------
    ' STAGE 9: Assign 1/0 flags.
    '
    ' Parcel i gets 1 if and only if both conditions hold:
    '   (a) (dirBits(i) AND DIR_ALL) = DIR_ALL  — N, S, E, W all covered
    '   (b) rootAcres(FindRoot(i)) >= MIN_BLOCK_ACRES
    ' -----------------------------------------------------------------------

    Dim flagArr() As Long
    ReDim flagArr(1 To n)

    Dim oneCount As Long, zeroCount As Long
    Dim maxBlockAcres As Double

    For i = 1 To n
        r = FindRoot(ufParent, i)
        If rootAcres(r) > maxBlockAcres Then maxBlockAcres = rootAcres(r)

        If (dirBits(i) And DIR_ALL) = DIR_ALL And _
           rootAcres(r) >= MIN_BLOCK_ACRES Then
            flagArr(i) = 1
            oneCount   = oneCount + 1
        Else
            flagArr(i) = 0
            zeroCount  = zeroCount + 1
        End If
    Next i

    ' -----------------------------------------------------------------------
    ' STAGE 10: Build the output array and write to the sheet in one shot.
    '
    ' Valid parcels get their flagArr value (1 or 0).
    ' Skipped rows retain the initialised empty string — blank cell, not 0.
    ' This lets the user distinguish "evaluated to 0" from "couldn't evaluate."
    ' -----------------------------------------------------------------------

    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To 1)
    For k = 1 To nRows
        outArr(k, 1) = ""     ' default: blank (skipped rows)
    Next k
    For i = 1 To n
        outArr(srcRow(i), 1) = flagArr(i)
    Next i

    ws.Range(ws.Cells(dataStart, colOut), ws.Cells(lastRow, colOut)).Value = outArr

    ' Write header if the row above dataStart exists and is currently blank.
    Dim hdrRow As Long
    hdrRow = dataStart - 1
    If hdrRow >= 1 Then
        If Trim(CStr(ws.Cells(hdrRow, colOut).Value)) = "" Then
            ws.Cells(hdrRow, colOut).Value = "Core (0/1)"
        End If
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 11: Debug report.
    ' -----------------------------------------------------------------------

    If DEBUG_MODE Then
        ' Count distinct connected blocks.
        Dim blockSeen() As Boolean
        ReDim blockSeen(1 To n)
        Dim blockCount As Long
        For i = 1 To n
            r = FindRoot(ufParent, i)
            If Not blockSeen(r) Then
                blockSeen(r) = True
                blockCount   = blockCount + 1
            End If
        Next i

        Debug.Print ""
        Debug.Print "=== FlagCoreParcels debug report ==="
        Debug.Print "Settings: MIN_BLOCK_ACRES=" & MIN_BLOCK_ACRES & _
                    "  TOLERANCE=" & TOLERANCE
        Debug.Print "Valid parcels      : " & n & "  |  Skipped: " & skipped
        Debug.Print "Flagged 1 (core)   : " & oneCount
        Debug.Print "Flagged 0          : " & zeroCount
        Debug.Print "Connected blocks   : " & blockCount
        Debug.Print "Largest block      : " & Format(maxBlockAcres, "#,##0.0") & " ac"
        Debug.Print "==================================="
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 12: Summary MsgBox.
    ' -----------------------------------------------------------------------

    MsgBox "FlagCoreParcels complete." & vbCrLf & vbCrLf & _
           "  Valid parcels    : " & n & vbCrLf & _
           "  Rows skipped     : " & skipped & vbCrLf & _
           "  Flagged 1 (core) : " & oneCount & vbCrLf & _
           "  Flagged 0        : " & zeroCount & vbCrLf & vbCrLf & _
           "Results written to column " & ColLetter(colOut) & "." & vbCrLf & _
           "(Skipped rows are blank, not 0.)", _
           vbInformation, "FlagCoreParcels"

    Exit Sub

' -----------------------------------------------------------------------
' Error / cancel handlers
' -----------------------------------------------------------------------
UserCancelled:
    MsgBox "Column selection cancelled. Macro aborted.", vbInformation, "FlagCoreParcels"
    Exit Sub

WrongSheet:
    MsgBox "All columns must be on the same worksheet. Macro cancelled.", _
           vbExclamation, "FlagCoreParcels"
    Exit Sub

End Sub

'===============================================================================
' Haversine great-circle distance
' Inputs  : two points in decimal degrees
' Returns : distance in metres
'===============================================================================
Public Function HaversineMeters(lat1 As Double, lon1 As Double, _
                                 lat2 As Double, lon2 As Double) As Double
    Dim dLat As Double, dLon As Double, a As Double
    dLat = (lat2 - lat1) * PI / 180#
    dLon = (lon2 - lon1) * PI / 180#
    a = Sin(dLat / 2#) ^ 2 + _
        Cos(lat1 * PI / 180#) * Cos(lat2 * PI / 180#) * Sin(dLon / 2#) ^ 2
    HaversineMeters = EARTH_RADIUS * 2# * Atan2(Sqr(a), Sqr(1# - a))
End Function

'===============================================================================
' Atan2(y, x) — two-argument arctangent in radians (VBA lacks a native one)
'===============================================================================
Private Function Atan2(y As Double, x As Double) As Double
    If x > 0# Then
        Atan2 = Atn(y / x)
    ElseIf x < 0# Then
        Atan2 = Atn(y / x) + IIf(y >= 0#, PI, -PI)
    Else
        If y > 0# Then
            Atan2 = PI / 2#
        ElseIf y < 0# Then
            Atan2 = -PI / 2#
        Else
            Atan2 = 0#    ' degenerate: undefined, return 0
        End If
    End If
End Function

'===============================================================================
' Union-Find: FindRoot with iterative path compression
'===============================================================================
Private Function FindRoot(ufParent() As Long, ByVal x As Long) As Long
    Dim root As Long, nxt As Long
    root = x
    Do While ufParent(root) <> root: root = ufParent(root): Loop
    Do While ufParent(x) <> root
        nxt = ufParent(x): ufParent(x) = root: x = nxt
    Loop
    FindRoot = root
End Function

'===============================================================================
' Union-Find: UnionSets by rank
'===============================================================================
Private Sub UnionSets(ufParent() As Long, ufRank() As Long, _
                       ByVal a As Long, ByVal b As Long)
    Dim rA As Long, rB As Long
    rA = FindRoot(ufParent, a)
    rB = FindRoot(ufParent, b)
    If rA = rB Then Exit Sub
    If ufRank(rA) < ufRank(rB) Then
        ufParent(rA) = rB
    ElseIf ufRank(rA) > ufRank(rB) Then
        ufParent(rB) = rA
    Else
        ufParent(rB) = rA
        ufRank(rA) = ufRank(rA) + 1
    End If
End Sub

'===============================================================================
' Utility: 1-based column number → letter(s), e.g. 28 → "AB"
'===============================================================================
Private Function ColLetter(ByVal c As Long) As String
    Dim s As String
    Do While c > 0
        Dim r As Long
        r = (c - 1) Mod 26
        s = Chr(65 + r) & s
        c = (c - 1 - r) \ 26
    Loop
    ColLetter = s
End Function
