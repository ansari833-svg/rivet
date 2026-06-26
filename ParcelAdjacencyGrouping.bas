Attribute VB_Name = "ParcelAdjacencyGrouping"
'===============================================================================
' ParcelAdjacencyGrouping.bas
' Groups parcels by spatial adjacency using centroid lat/lon and acreage.
'
' CAVEAT: Centroid + acreage only approximates contiguity. This method assumes
' parcels are reasonably compact and near-square. Long, thin, or irregular
' parcels can have centroids far from a neighbor they actually border, and small
' parcels clustered near a large one can appear adjacent when they aren't. For
' ground-truth contiguity, actual parcel polygons (GIS data) are needed. This
' macro is a fast screen, not a definitive answer.
'
' HOW TO IMPORT AND RUN:
'   1. In Excel, press Alt+F11 to open the VBA editor.
'   2. In the Project window, right-click your workbook name > Import File,
'      then select this .bas file.
'   3. Close the editor, return to Excel, and press Alt+F8.
'   4. Select "GroupParcels" and click Run. Follow the on-screen prompts.
'   5. To see debug output, open the Immediate window first: Ctrl+G in the editor.
'===============================================================================
Option Explicit

' -----------------------------------------------------------------------
' CONFIG BLOCK — adjust these constants to tune grouping behaviour.
' Column selections and the data-start row are picked at runtime via prompts.
' -----------------------------------------------------------------------

' Adjacency threshold multiplier applied to PAIR_SIZE (see PAIR_RULE).
'   1.0 = strict: two same-size squares sharing an edge have centroids
'         exactly s apart, so 1.0 accepts only true edge-sharers.
'   1.25–1.4 progressively catches diagonal/corner and near neighbours,
'   but risks chaining unrelated parcels — lower first if groups are too big.
Private Const TOLERANCE As Double = 1.0

' How to compute PAIR_SIZE from the two parcels' equivalent sizes (s1, s2).
'   "min" — uses Min(s1, s2).  A large parcel cannot reach out and link
'           distant small ones; the small parcel's own size governs.
'           DEFAULT — best defence against runaway groups.
'   "avg" — uses (s1 + s2) / 2.  More permissive; matches the classic
'           "shared-edge centroid distance" formula for equal-size squares.
Private Const PAIR_RULE As String = "min"

' Shape model used to derive each parcel's equivalent size from its acreage.
'   "square" — side s = Sqrt(area_m2); adjacent when dist <= TOLERANCE * PAIR_SIZE(s1,s2)
'   "circle" — radius r = Sqrt(area_m2 / Pi); s is reinterpreted as r
Private Const SHAPE_MODEL As String = "square"

' Write a Group Size column immediately to the right of the Group ID column?
'   True  = yes (skipped automatically if that neighbour column has data).
'   False = skip entirely.
Private Const WRITE_GROUP_SIZE As Boolean = True

' Diagnostic / debug mode.
'   True  = after the run, print a group-size distribution, the largest group,
'           the 10 accepted links with the greatest centroid distances, and a
'           sanity check on the first valid parcel's computed size.
'           Output goes to the Immediate window (Ctrl+G in the VBA editor)
'           and a summary MsgBox.
'   False = silent; only the final summary MsgBox is shown.
Private Const DEBUG_MODE As Boolean = True

' Unit-conversion and geometry constants.
Private Const ACRES_TO_M2  As Double = 4046.86
Private Const EARTH_RADIUS As Double = 6371000#
Private Const PI           As Double = 3.14159265358979

' Number of "longest accepted links" to report in debug mode.
Private Const DEBUG_TOP_LINKS As Long = 10

'===============================================================================
' Main entry point
'===============================================================================
Public Sub GroupParcels()

    ' -----------------------------------------------------------------------
    ' STAGE 1: Collect runtime inputs — data-start row, then column picks.
    ' -----------------------------------------------------------------------

    Dim ws         As Worksheet
    Dim rPick      As Range
    Dim colID      As Long
    Dim colLat     As Long
    Dim colLon     As Long
    Dim colAcre    As Long
    Dim colOut     As Long
    Dim dataStart  As Long
    Dim userInput  As String

    ' Data-start row (default 2 → row 1 is the header).
    userInput = InputBox( _
        "Enter the row number where parcel DATA starts." & vbCrLf & _
        "(Row 1 is typically the header row, so data starts at row 2.)", _
        "Data Start Row", "2")
    If userInput = "" Then
        MsgBox "Cancelled.", vbInformation, "GroupParcels"
        Exit Sub
    End If
    If Not IsNumeric(userInput) Or CLng(userInput) < 1 Then
        MsgBox "Invalid row number. Macro cancelled.", vbExclamation, "GroupParcels"
        Exit Sub
    End If
    dataStart = CLng(userInput)

    ' Column pickers — each uses Application.InputBox Type:=8 (range picker).
    On Error GoTo UserCancelled

    Set rPick = Application.InputBox( _
        "Click any cell in the PARCEL ID column.", _
        "Select Parcel ID Column", Type:=8)
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
        "Click any cell in the ACREAGE column.", _
        "Select Acreage Column", Type:=8)
    If Not rPick.Worksheet Is ws Then GoTo WrongSheet
    colAcre = rPick.Column

    Set rPick = Application.InputBox( _
        "Click any cell in the OUTPUT column where Group IDs should be written." & vbCrLf & _
        "(Can be an empty column or an existing one — you'll be warned before overwriting.)", _
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
               vbExclamation, "GroupParcels"
        Exit Sub
    End If

    Dim nRows As Long
    nRows = lastRow - dataStart + 1

    ' -----------------------------------------------------------------------
    ' STAGE 3: Read each needed column into memory in one Range.Value call.
    ' Reading non-contiguous columns separately avoids pulling unused data.
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
            "The selected output column already contains data in rows " & _
            dataStart & ":" & lastRow & "." & vbCrLf & vbCrLf & _
            "Do you want to OVERWRITE it?", _
            vbQuestion + vbYesNo, "Output Column Not Empty")
        If ans = vbNo Then
            MsgBox "Macro cancelled — output column left untouched.", vbInformation, "GroupParcels"
            Exit Sub
        End If
    End If

    ' Decide on the Group Size neighbour column.
    Dim colSize   As Long
    Dim writeSize As Boolean
    colSize   = colOut + 1
    writeSize = WRITE_GROUP_SIZE
    If writeSize Then
        Dim sizeRange As Range
        Set sizeRange = ws.Range(ws.Cells(dataStart, colSize), ws.Cells(lastRow, colSize))
        If Application.WorksheetFunction.CountA(sizeRange) > 0 Then
            MsgBox "WRITE_GROUP_SIZE is on, but column " & ColLetter(colSize) & _
                   " already contains data — Group Size will NOT be written.", _
                   vbInformation, "GroupParcels"
            writeSize = False
        End If
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 5: Validate rows and build in-memory working arrays.
    ' -----------------------------------------------------------------------

    Dim latArr()    As Double   ' valid parcels only
    Dim lonArr()    As Double
    Dim sizeArr()   As Double   ' equivalent "size" (half-side or radius per SHAPE_MODEL)
    Dim srcRow()    As Long     ' maps valid index -> 1-based position in arrID/arrLat/…

    ReDim latArr(1 To nRows)
    ReDim lonArr(1 To nRows)
    ReDim sizeArr(1 To nRows)
    ReDim srcRow(1 To nRows)

    Dim n       As Long     ' count of valid parcels
    Dim skipped As Long
    Dim i       As Long
    Dim latV    As Double
    Dim lonV    As Double
    Dim acreV   As Double
    Dim areaM2  As Double
    Dim firstValid As Boolean
    firstValid = False

    For i = 1 To nRows

        ' --- Latitude ---
        If IsEmpty(arrLat(i, 1)) Or Not IsNumeric(arrLat(i, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        latV = CDbl(arrLat(i, 1))
        If latV < -90 Or latV > 90 Then skipped = skipped + 1: GoTo NextRow

        ' --- Longitude ---
        If IsEmpty(arrLon(i, 1)) Or Not IsNumeric(arrLon(i, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        lonV = CDbl(arrLon(i, 1))
        If lonV < -180 Or lonV > 180 Then skipped = skipped + 1: GoTo NextRow

        ' --- Acreage ---
        If IsEmpty(arrAcre(i, 1)) Or Not IsNumeric(arrAcre(i, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        acreV = CDbl(arrAcre(i, 1))
        If acreV <= 0 Then skipped = skipped + 1: GoTo NextRow

        ' --- Store valid parcel ---
        n = n + 1
        srcRow(n) = i
        latArr(n) = latV
        lonArr(n) = lonV
        areaM2 = acreV * ACRES_TO_M2

        Select Case LCase(Trim(SHAPE_MODEL))
            Case "circle"
                sizeArr(n) = Sqr(areaM2 / PI)      ' radius
            Case Else                               ' "square" (default)
                sizeArr(n) = Sqr(areaM2)            ' full side length
        End Select

        ' Sanity line for the first valid parcel (debug mode).
        If DEBUG_MODE And Not firstValid Then
            firstValid = True
            Debug.Print "--- GroupParcels sanity check (first valid parcel) ---"
            Debug.Print "  Parcel ID : " & CStr(arrID(i, 1))
            Debug.Print "  Acreage   : " & acreV & " acres"
            Debug.Print "  Area      : " & Format(areaM2, "#,##0.0") & " m²"
            Debug.Print "  Size (s)  : " & Format(sizeArr(n), "#,##0.0") & " m  " & _
                        "(expect ~100–1000 m for typical parcels; " & _
                        "if >10,000 m, acreage column may not be in acres)"
        End If

NextRow:
    Next i

    If n = 0 Then
        MsgBox "No valid parcel rows found. Check your column selections and data.", _
               vbExclamation, "GroupParcels"
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 6: Initialise union-find.
    '   ufParent(i) = root of the set that contains i  (path-compressed lazily)
    '   ufRank(i)   = rank of tree rooted at i          (union by rank)
    ' -----------------------------------------------------------------------

    Dim ufParent() As Long
    Dim ufRank()   As Long
    ReDim ufParent(1 To n)
    ReDim ufRank(1 To n)

    For i = 1 To n
        ufParent(i) = i
        ufRank(i)   = 0
    Next i

    ' -----------------------------------------------------------------------
    ' STAGE 7: O(n²) adjacency pass with in-memory union-find.
    '
    ' For DEBUG_MODE we also track the accepted links sorted by distance so
    ' we can report the top-N longest ones (most likely chaining culprits).
    ' -----------------------------------------------------------------------

    Dim j         As Long
    Dim dist      As Double
    Dim threshold As Double
    Dim pairSize  As Double

    ' Debug link tracking: parallel arrays for the top accepted links.
    Dim dbgDist()    As Double   ' distance in metres
    Dim dbgThresh()  As Double   ' threshold that accepted it
    Dim dbgIdxA()    As Long     ' index into valid-parcel arrays
    Dim dbgIdxB()    As Long
    Dim dbgCount     As Long     ' how many links recorded so far
    Dim dbgCap       As Long     ' capacity (grows as needed)

    If DEBUG_MODE Then
        dbgCap = 64
        ReDim dbgDist(1 To dbgCap)
        ReDim dbgThresh(1 To dbgCap)
        ReDim dbgIdxA(1 To dbgCap)
        ReDim dbgIdxB(1 To dbgCap)
        dbgCount = 0
    End If

    For i = 1 To n - 1
        For j = i + 1 To n
            dist = HaversineMeters(latArr(i), lonArr(i), latArr(j), lonArr(j))

            ' Compute PAIR_SIZE according to PAIR_RULE.
            Select Case LCase(Trim(PAIR_RULE))
                Case "avg"
                    pairSize = (sizeArr(i) + sizeArr(j)) / 2
                Case Else   ' "min" (default)
                    If sizeArr(i) < sizeArr(j) Then
                        pairSize = sizeArr(i)
                    Else
                        pairSize = sizeArr(j)
                    End If
            End Select

            threshold = TOLERANCE * pairSize

            If dist <= threshold Then
                Call UnionSets(ufParent, ufRank, i, j)

                ' Record link for debug reporting.
                If DEBUG_MODE Then
                    dbgCount = dbgCount + 1
                    If dbgCount > dbgCap Then
                        dbgCap = dbgCap * 2
                        ReDim Preserve dbgDist(1 To dbgCap)
                        ReDim Preserve dbgThresh(1 To dbgCap)
                        ReDim Preserve dbgIdxA(1 To dbgCap)
                        ReDim Preserve dbgIdxB(1 To dbgCap)
                    End If
                    dbgDist(dbgCount)   = dist
                    dbgThresh(dbgCount) = threshold
                    dbgIdxA(dbgCount)   = i
                    dbgIdxB(dbgCount)   = j
                End If
            End If
        Next j
    Next i

    ' -----------------------------------------------------------------------
    ' STAGE 8: Assign sequential group IDs and tally sizes.
    ' -----------------------------------------------------------------------

    Dim rootToGroup() As Long   ' rootToGroup(root) = group number (1-based)
    Dim groupID()     As Long   ' groupID(i) = group number for valid parcel i
    ReDim rootToGroup(1 To n)
    ReDim groupID(1 To n)

    Dim groupCount As Long
    groupCount = 0
    Dim root As Long

    For i = 1 To n
        root = FindRoot(ufParent, i)
        If rootToGroup(root) = 0 Then
            groupCount = groupCount + 1
            rootToGroup(root) = groupCount
        End If
        groupID(i) = rootToGroup(root)
    Next i

    Dim groupSizes() As Long
    ReDim groupSizes(1 To groupCount)
    Dim g As Long
    For i = 1 To n
        groupSizes(groupID(i)) = groupSizes(groupID(i)) + 1
    Next i

    Dim maxSize  As Long
    Dim maxGrpID As Long
    For g = 1 To groupCount
        If groupSizes(g) > maxSize Then
            maxSize  = groupSizes(g)
            maxGrpID = g
        End If
    Next g

    ' -----------------------------------------------------------------------
    ' STAGE 9: Build output arrays and write back in one shot.
    ' -----------------------------------------------------------------------

    Dim outGrp()  As Variant
    Dim outSz()   As Variant
    ReDim outGrp(1 To nRows, 1 To 1)
    If writeSize Then ReDim outSz(1 To nRows, 1 To 1)

    For i = 1 To nRows
        outGrp(i, 1) = ""
        If writeSize Then outSz(i, 1) = ""
    Next i

    For i = 1 To n
        outGrp(srcRow(i), 1) = "GRP-" & Format(groupID(i), "0000")
        If writeSize Then outSz(srcRow(i), 1) = groupSizes(groupID(i))
    Next i

    ws.Range(ws.Cells(dataStart, colOut), ws.Cells(lastRow, colOut)).Value = outGrp
    If writeSize Then
        ws.Range(ws.Cells(dataStart, colSize), ws.Cells(lastRow, colSize)).Value = outSz
    End If

    ' Write headers if the row above dataStart exists and is blank.
    Dim hdrRow As Long
    hdrRow = dataStart - 1
    If hdrRow >= 1 Then
        If Trim(CStr(ws.Cells(hdrRow, colOut).Value)) = "" Then
            ws.Cells(hdrRow, colOut).Value = "Group ID"
        End If
        If writeSize Then
            If Trim(CStr(ws.Cells(hdrRow, colSize).Value)) = "" Then
                ws.Cells(hdrRow, colSize).Value = "Group Size"
            End If
        End If
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 10: Debug report (when DEBUG_MODE = True).
    ' -----------------------------------------------------------------------

    Dim debugMsg As String
    debugMsg = ""

    If DEBUG_MODE Then

        ' -- Group-size distribution --
        Dim b1 As Long, b2 As Long, b3 As Long, b4 As Long, b5 As Long
        For g = 1 To groupCount
            Select Case groupSizes(g)
                Case 1:             b1 = b1 + 1
                Case 2 To 5:        b2 = b2 + 1
                Case 6 To 20:       b3 = b3 + 1
                Case 21 To 100:     b4 = b4 + 1
                Case Else:          b5 = b5 + 1
            End Select
        Next g

        Debug.Print ""
        Debug.Print "=== GroupParcels debug report ==="
        Debug.Print "Parcels: " & n & "  |  Groups: " & groupCount & _
                    "  |  Skipped: " & skipped
        Debug.Print "Tolerance: " & TOLERANCE & "  |  PairRule: " & PAIR_RULE & _
                    "  |  ShapeModel: " & SHAPE_MODEL
        Debug.Print ""
        Debug.Print "Group-size distribution:"
        Debug.Print "  Size 1        : " & b1
        Debug.Print "  Size 2–5      : " & b2
        Debug.Print "  Size 6–20     : " & b3
        Debug.Print "  Size 21–100   : " & b4
        Debug.Print "  Size 101+     : " & b5
        Debug.Print ""
        Debug.Print "Largest group   : GRP-" & Format(maxGrpID, "0000") & _
                    "  (" & maxSize & " parcels)"

        debugMsg = "Group-size distribution:" & vbCrLf & _
                   "  1       : " & b1 & vbCrLf & _
                   "  2–5     : " & b2 & vbCrLf & _
                   "  6–20    : " & b3 & vbCrLf & _
                   "  21–100  : " & b4 & vbCrLf & _
                   "  101+    : " & b5 & vbCrLf & vbCrLf & _
                   "Largest group: GRP-" & Format(maxGrpID, "0000") & _
                   " (" & maxSize & " parcels)"

        ' -- Top accepted links by distance (insertion-sort into a small array) --
        If dbgCount > 0 Then
            Dim topN    As Long
            topN = DEBUG_TOP_LINKS
            If dbgCount < topN Then topN = dbgCount

            ' Partial selection sort to find the topN largest distances.
            Dim sortDist()   As Double
            Dim sortThresh() As Double
            Dim sortA()      As Long
            Dim sortB()      As Long
            ReDim sortDist(1 To topN)
            ReDim sortThresh(1 To topN)
            ReDim sortA(1 To topN)
            ReDim sortB(1 To topN)

            Dim k As Long, minPos As Long, minVal As Double
            ' Fill first topN slots.
            For k = 1 To topN
                sortDist(k)   = dbgDist(k)
                sortThresh(k) = dbgThresh(k)
                sortA(k)      = dbgIdxA(k)
                sortB(k)      = dbgIdxB(k)
            Next k
            ' Sort those topN descending by distance (selection sort — small N).
            Dim p As Long, q As Long
            Dim tmpD As Double, tmpT As Double, tmpA As Long, tmpB As Long
            For p = 1 To topN - 1
                Dim maxPos As Long
                maxPos = p
                For q = p + 1 To topN
                    If sortDist(q) > sortDist(maxPos) Then maxPos = q
                Next q
                If maxPos <> p Then
                    tmpD = sortDist(p): sortDist(p) = sortDist(maxPos): sortDist(maxPos) = tmpD
                    tmpT = sortThresh(p): sortThresh(p) = sortThresh(maxPos): sortThresh(maxPos) = tmpT
                    tmpA = sortA(p): sortA(p) = sortA(maxPos): sortA(maxPos) = tmpA
                    tmpB = sortB(p): sortB(p) = sortB(maxPos): sortB(maxPos) = tmpB
                End If
            Next p
            ' Now scan remaining links, replacing the minimum in our top-N list if larger.
            For k = topN + 1 To dbgCount
                minPos = 1
                minVal = sortDist(1)
                For p = 2 To topN
                    If sortDist(p) < minVal Then minVal = sortDist(p): minPos = p
                Next p
                If dbgDist(k) > minVal Then
                    sortDist(minPos)   = dbgDist(k)
                    sortThresh(minPos) = dbgThresh(k)
                    sortA(minPos)      = dbgIdxA(k)
                    sortB(minPos)      = dbgIdxB(k)
                End If
            Next k
            ' Final sort descending.
            For p = 1 To topN - 1
                maxPos = p
                For q = p + 1 To topN
                    If sortDist(q) > sortDist(maxPos) Then maxPos = q
                Next q
                If maxPos <> p Then
                    tmpD = sortDist(p): sortDist(p) = sortDist(maxPos): sortDist(maxPos) = tmpD
                    tmpT = sortThresh(p): sortThresh(p) = sortThresh(maxPos): sortThresh(maxPos) = tmpT
                    tmpA = sortA(p): sortA(p) = sortA(maxPos): sortA(maxPos) = tmpA
                    tmpB = sortB(p): sortB(p) = sortB(maxPos): sortB(maxPos) = tmpB
                End If
            Next p

            Debug.Print ""
            Debug.Print "Top " & topN & " accepted links by distance " & _
                        "(these are the most likely chaining culprits):"
            Debug.Print "  " & PadR("ParcelA", 20) & PadR("ParcelB", 20) & _
                        PadL("Dist(m)", 10) & PadL("Thresh(m)", 12) & _
                        PadL("sA(m)", 10) & PadL("sB(m)", 10)

            Dim linkLine As String
            For k = 1 To topN
                Dim idxA As Long, idxB As Long
                idxA = sortA(k)
                idxB = sortB(k)
                linkLine = "  " & _
                    PadR(CStr(arrID(srcRow(idxA), 1)), 20) & _
                    PadR(CStr(arrID(srcRow(idxB), 1)), 20) & _
                    PadL(Format(sortDist(k), "#,##0.0"), 10) & _
                    PadL(Format(sortThresh(k), "#,##0.0"), 12) & _
                    PadL(Format(sizeArr(idxA), "#,##0.0"), 10) & _
                    PadL(Format(sizeArr(idxB), "#,##0.0"), 10)
                Debug.Print linkLine
            Next k

            debugMsg = debugMsg & vbCrLf & vbCrLf & _
                       "Top " & topN & " longest accepted links printed" & vbCrLf & _
                       "to the Immediate window (Ctrl+G in the VBA editor)." & vbCrLf & _
                       "If groups look too large, those links are the first" & vbCrLf & _
                       "place to look — try lowering TOLERANCE or switching" & vbCrLf & _
                       "PAIR_RULE from ""avg"" to ""min""."
        End If

        Debug.Print "==================================="
        MsgBox debugMsg, vbInformation, "GroupParcels — Debug Report"
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 11: Final summary.
    ' -----------------------------------------------------------------------

    MsgBox "GroupParcels complete." & vbCrLf & vbCrLf & _
           "  Parcels processed : " & n & vbCrLf & _
           "  Groups formed     : " & groupCount & vbCrLf & _
           "  Largest group     : " & maxSize & " parcel(s)  (GRP-" & Format(maxGrpID, "0000") & ")" & vbCrLf & _
           "  Rows skipped      : " & skipped & " (bad / missing data)" & vbCrLf & vbCrLf & _
           "Results written to column " & ColLetter(colOut) & _
           IIf(writeSize, " (Group ID) and " & ColLetter(colSize) & " (Group Size).", "."), _
           vbInformation, "GroupParcels"

    Exit Sub

' -----------------------------------------------------------------------
' Error / cancel handlers
' -----------------------------------------------------------------------
UserCancelled:
    MsgBox "Column selection cancelled. Macro aborted.", vbInformation, "GroupParcels"
    Exit Sub

WrongSheet:
    MsgBox "All columns must be on the same worksheet. Macro cancelled.", _
           vbExclamation, "GroupParcels"
    Exit Sub

End Sub

'===============================================================================
' Haversine great-circle distance
' Inputs : two points in decimal degrees
' Returns: distance in metres
'===============================================================================
Public Function HaversineMeters(lat1 As Double, lon1 As Double, _
                                 lat2 As Double, lon2 As Double) As Double
    Dim dLat As Double, dLon As Double
    Dim a    As Double, c    As Double

    dLat = (lat2 - lat1) * PI / 180#
    dLon = (lon2 - lon1) * PI / 180#

    a = Sin(dLat / 2) ^ 2 + _
        Cos(lat1 * PI / 180#) * Cos(lat2 * PI / 180#) * Sin(dLon / 2) ^ 2

    ' atan2(sqrt(a), sqrt(1-a)) via Atn — avoids domain error when a ≈ 0.
    c = 2# * Atn(Sqr(a) / Sqr(1# - a))

    HaversineMeters = EARTH_RADIUS * c
End Function

'===============================================================================
' Union-Find: Find root with iterative path compression
'===============================================================================
Private Function FindRoot(ufParent() As Long, ByVal x As Long) As Long
    Dim root As Long
    Dim nxt  As Long

    ' Walk to root.
    root = x
    Do While ufParent(root) <> root
        root = ufParent(root)
    Loop

    ' Path compression — flatten every node on the path directly to root.
    Do While ufParent(x) <> root
        nxt         = ufParent(x)
        ufParent(x) = root
        x           = nxt
    Loop

    FindRoot = root
End Function

'===============================================================================
' Union-Find: Union by rank — keeps trees shallow
'===============================================================================
Private Sub UnionSets(ufParent() As Long, ufRank() As Long, _
                       ByVal a As Long, ByVal b As Long)
    Dim rA As Long, rB As Long
    rA = FindRoot(ufParent, a)
    rB = FindRoot(ufParent, b)
    If rA = rB Then Exit Sub

    ' Attach lower-rank tree under higher-rank root.
    If ufRank(rA) < ufRank(rB) Then
        ufParent(rA) = rB
    ElseIf ufRank(rA) > ufRank(rB) Then
        ufParent(rB) = rA
    Else
        ufParent(rB) = rA
        ufRank(rA)   = ufRank(rA) + 1
    End If
End Sub

'===============================================================================
' Utility: 1-based column number -> column letter(s)  (e.g. 28 -> "AB")
'===============================================================================
Private Function ColLetter(ByVal c As Long) As String
    Dim s As String
    s = ""
    Do While c > 0
        Dim r As Long
        r = (c - 1) Mod 26
        s = Chr(65 + r) & s
        c = (c - 1 - r) \ 26
    Loop
    ColLetter = s
End Function

'===============================================================================
' Utility: right-pad a string to width w (for Immediate-window table alignment)
'===============================================================================
Private Function PadR(ByVal s As String, ByVal w As Long) As String
    If Len(s) >= w Then
        PadR = Left(s, w)
    Else
        PadR = s & Space(w - Len(s))
    End If
End Function

'===============================================================================
' Utility: left-pad a string to width w
'===============================================================================
Private Function PadL(ByVal s As String, ByVal w As Long) As String
    If Len(s) >= w Then
        PadL = Left(s, w)
    Else
        PadL = Space(w - Len(s)) & s
    End If
End Function
