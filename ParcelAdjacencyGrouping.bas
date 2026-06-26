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
'===============================================================================
Option Explicit

' -----------------------------------------------------------------------
' CONFIG BLOCK — adjust these constants to change grouping behaviour.
' Column selections and the data-start row are picked at runtime via prompts.
' -----------------------------------------------------------------------

' Adjacency threshold multiplier.
'   1.0 = strict edge-sharing only (centroids exactly (s1+s2)/2 apart).
'   1.25 = default, allows slight gaps / positional imprecision.
'   ~1.4 also catches diagonal / corner neighbours.
Private Const TOLERANCE As Double = 1.25

' Shape model used to estimate each parcel's extent from its acreage.
'   "square"  — side length = Sqrt(area_m2);
'               adjacent when dist <= TOLERANCE * (s1 + s2) / 2
'   "circle"  — radius = Sqrt(area_m2 / Pi);
'               adjacent when dist <= TOLERANCE * (r1 + r2)
Private Const SHAPE_MODEL As String = "square"

' Write a Group Size column immediately to the right of the Group ID column?
'   True  = yes (skipped automatically if that neighbour column has data).
'   False = skip entirely.
Private Const WRITE_GROUP_SIZE As Boolean = True

' Unit-conversion and geometry constants — no need to touch these.
Private Const ACRES_TO_M2 As Double = 4046.86
Private Const EARTH_RADIUS_M As Double = 6371000#
Private Const PI As Double = 3.14159265358979

'===============================================================================
' Main entry point
'===============================================================================
Public Sub GroupParcels()

    ' -----------------------------------------------------------------------
    ' STAGE 1: Collect runtime inputs — column picks and data-start row.
    ' -----------------------------------------------------------------------

    Dim ws As Worksheet
    Dim rPick As Range
    Dim colID As Long, colLat As Long, colLon As Long
    Dim colAcre As Long, colOut As Long
    Dim dataStartRow As Long
    Dim userInput As String

    ' Ask for the data-start row (default 2, meaning row 1 is the header).
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
    dataStartRow = CLng(userInput)

    ' Helper: pick a column by clicking any cell in it.
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
    ' STAGE 2: Determine the data extent.
    ' -----------------------------------------------------------------------

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colID).End(xlUp).Row
    If lastRow < dataStartRow Then
        MsgBox "No data rows found below row " & dataStartRow & ". Macro cancelled.", _
               vbExclamation, "GroupParcels"
        Exit Sub
    End If

    Dim nRows As Long
    nRows = lastRow - dataStartRow + 1   ' total rows including potential blanks

    ' -----------------------------------------------------------------------
    ' STAGE 3: Read the entire used block into memory in one shot.
    ' We read columns for ID, Lat, Lon, Acre individually (each as a 2-D
    ' single-column array) to avoid pulling the whole sheet width.
    ' -----------------------------------------------------------------------

    ' Find the widest column index we need to read; we'll read that column
    ' plus potentially colOut+1 for the size column.  Since columns are
    ' non-contiguous we read each column separately.

    Dim arrID()   As Variant
    Dim arrLat()  As Variant
    Dim arrLon()  As Variant
    Dim arrAcre() As Variant

    arrID   = ws.Range(ws.Cells(dataStartRow, colID),   ws.Cells(lastRow, colID)).Value
    arrLat  = ws.Range(ws.Cells(dataStartRow, colLat),  ws.Cells(lastRow, colLat)).Value
    arrLon  = ws.Range(ws.Cells(dataStartRow, colLon),  ws.Cells(lastRow, colLon)).Value
    arrAcre = ws.Range(ws.Cells(dataStartRow, colAcre), ws.Cells(lastRow, colAcre)).Value

    ' -----------------------------------------------------------------------
    ' STAGE 4: Guard — check if the output column already has data.
    ' -----------------------------------------------------------------------

    Dim outColRange As Range
    Set outColRange = ws.Range(ws.Cells(dataStartRow, colOut), ws.Cells(lastRow, colOut))
    If Application.WorksheetFunction.CountA(outColRange) > 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox( _
            "The selected output column already contains data in rows " & _
            dataStartRow & ":" & lastRow & "." & vbCrLf & vbCrLf & _
            "Do you want to OVERWRITE it?", _
            vbQuestion + vbYesNo, "Output Column Not Empty")
        If ans = vbNo Then
            MsgBox "Macro cancelled — output column left untouched.", vbInformation, "GroupParcels"
            Exit Sub
        End If
    End If

    ' Check the Group Size neighbour column.
    Dim colSize As Long
    colSize = colOut + 1
    Dim writeSize As Boolean
    writeSize = WRITE_GROUP_SIZE
    If writeSize Then
        Dim sizeColRange As Range
        Set sizeColRange = ws.Range(ws.Cells(dataStartRow, colSize), ws.Cells(lastRow, colSize))
        If Application.WorksheetFunction.CountA(sizeColRange) > 0 Then
            MsgBox "WRITE_GROUP_SIZE is enabled, but column " & _
                   ColLetter(colSize) & " already contains data. " & _
                   "Group Size will NOT be written to avoid overwriting.", _
                   vbInformation, "GroupParcels"
            writeSize = False
        End If
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 5: Parse rows — validate and build working arrays.
    ' -----------------------------------------------------------------------

    ' Working arrays (1-based, length nRows; valid rows only tracked via mask).
    Dim lat()   As Double
    Dim lon()   As Double
    Dim halfExt() As Double   ' half-extent: s/2 for square, r for circle
    Dim valid() As Boolean
    Dim srcRow() As Long      ' maps working index -> original array row

    ReDim lat(1 To nRows)
    ReDim lon(1 To nRows)
    ReDim halfExt(1 To nRows)
    ReDim valid(1 To nRows)
    ReDim srcRow(1 To nRows)

    Dim skipped As Long
    skipped = 0
    Dim n As Long   ' count of valid parcels
    n = 0

    Dim i As Long
    Dim latV As Double, lonV As Double, acreV As Double, areaM2 As Double

    For i = 1 To nRows
        ' Validate lat
        If IsEmpty(arrLat(i, 1)) Or Not IsNumeric(arrLat(i, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        latV = CDbl(arrLat(i, 1))
        If latV < -90 Or latV > 90 Then skipped = skipped + 1: GoTo NextRow

        ' Validate lon
        If IsEmpty(arrLon(i, 1)) Or Not IsNumeric(arrLon(i, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        lonV = CDbl(arrLon(i, 1))
        If lonV < -180 Or lonV > 180 Then skipped = skipped + 1: GoTo NextRow

        ' Validate acreage
        If IsEmpty(arrAcre(i, 1)) Or Not IsNumeric(arrAcre(i, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        acreV = CDbl(arrAcre(i, 1))
        If acreV <= 0 Then skipped = skipped + 1: GoTo NextRow

        ' All good — store.
        n = n + 1
        srcRow(n) = i
        lat(n) = latV
        lon(n) = lonV
        areaM2 = acreV * ACRES_TO_M2

        Select Case LCase(Trim(SHAPE_MODEL))
            Case "circle"
                halfExt(n) = Sqr(areaM2 / PI)       ' radius
            Case Else                                 ' "square" (default)
                halfExt(n) = Sqr(areaM2) / 2         ' half-side = s/2
        End Select

NextRow:
    Next i

    If n = 0 Then
        MsgBox "No valid parcel rows found. Check your column selections and data.", _
               vbExclamation, "GroupParcels"
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 6: Initialise union-find structures.
    '   parent(i) = root of set containing i  (path-compressed lazily)
    '   ufSize(i) = size of set rooted at i   (union by size)
    ' -----------------------------------------------------------------------

    Dim parent() As Long
    Dim ufSize() As Long
    ReDim parent(1 To n)
    ReDim ufSize(1 To n)

    For i = 1 To n
        parent(i) = i
        ufSize(i) = 1
    Next i

    ' -----------------------------------------------------------------------
    ' STAGE 7: O(n²) adjacency pass — union adjacent parcels.
    ' -----------------------------------------------------------------------

    Dim j As Long
    Dim dist As Double, threshold As Double

    For i = 1 To n - 1
        For j = i + 1 To n
            dist = HaversineMeters(lat(i), lon(i), lat(j), lon(j))

            Select Case LCase(Trim(SHAPE_MODEL))
                Case "circle"
                    ' threshold = TOLERANCE * (r_i + r_j)
                    threshold = TOLERANCE * (halfExt(i) + halfExt(j))
                Case Else
                    ' threshold = TOLERANCE * (s_i/2 + s_j/2)
                    threshold = TOLERANCE * (halfExt(i) + halfExt(j))
            End Select

            If dist <= threshold Then
                Call UnionSets(parent, ufSize, i, j)
            End If
        Next j
    Next i

    ' -----------------------------------------------------------------------
    ' STAGE 8: Assign sequential group IDs and compute group sizes.
    ' -----------------------------------------------------------------------

    ' Map each root -> sequential group number (1, 2, 3, ...).
    Dim rootToGroup() As Long
    ReDim rootToGroup(1 To n)   ' sparse; index is root index

    Dim groupCount As Long
    groupCount = 0

    Dim groupID() As Long     ' groupID(i) = group number for valid parcel i
    ReDim groupID(1 To n)

    Dim root As Long
    For i = 1 To n
        root = FindRoot(parent, i)
        If rootToGroup(root) = 0 Then
            groupCount = groupCount + 1
            rootToGroup(root) = groupCount
        End If
        groupID(i) = rootToGroup(root)
    Next i

    ' Tally group sizes.
    Dim groupSizes() As Long
    ReDim groupSizes(1 To groupCount)
    For i = 1 To n
        groupSizes(groupID(i)) = groupSizes(groupID(i)) + 1
    Next i

    Dim maxGroupSize As Long
    maxGroupSize = 0
    Dim g As Long
    For g = 1 To groupCount
        If groupSizes(g) > maxGroupSize Then maxGroupSize = groupSizes(g)
    Next g

    ' -----------------------------------------------------------------------
    ' STAGE 9: Build output arrays and write back in one operation.
    ' -----------------------------------------------------------------------

    ' Output arrays span all nRows (including skipped); skipped rows get "".
    Dim outGrp()  As Variant
    Dim outSize() As Variant
    ReDim outGrp(1 To nRows, 1 To 1)
    If writeSize Then ReDim outSize(1 To nRows, 1 To 1)

    ' Initialise with empty string (skipped rows stay blank).
    For i = 1 To nRows
        outGrp(i, 1) = ""
        If writeSize Then outSize(i, 1) = ""
    Next i

    For i = 1 To n
        outGrp(srcRow(i), 1)  = "GRP-" & Format(groupID(i), "0000")
        If writeSize Then outSize(srcRow(i), 1) = groupSizes(groupID(i))
    Next i

    ' Write Group ID column.
    ws.Range(ws.Cells(dataStartRow, colOut), ws.Cells(lastRow, colOut)).Value = outGrp

    ' Write Group Size column if enabled and safe.
    If writeSize Then
        ws.Range(ws.Cells(dataStartRow, colSize), ws.Cells(lastRow, colSize)).Value = outSize
    End If

    ' Write headers on header row (row dataStartRow - 1) if that row exists.
    Dim headerRow As Long
    headerRow = dataStartRow - 1
    If headerRow >= 1 Then
        If Trim(CStr(ws.Cells(headerRow, colOut).Value)) = "" Then
            ws.Cells(headerRow, colOut).Value = "Group ID"
        End If
        If writeSize Then
            If Trim(CStr(ws.Cells(headerRow, colSize).Value)) = "" Then
                ws.Cells(headerRow, colSize).Value = "Group Size"
            End If
        End If
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 10: Summary report.
    ' -----------------------------------------------------------------------

    MsgBox "GroupParcels complete." & vbCrLf & vbCrLf & _
           "  Parcels processed : " & n & vbCrLf & _
           "  Groups formed     : " & groupCount & vbCrLf & _
           "  Largest group     : " & maxGroupSize & " parcel(s)" & vbCrLf & _
           "  Rows skipped      : " & skipped & " (bad / missing data)" & vbCrLf & vbCrLf & _
           "Results written to column " & ColLetter(colOut) & _
           IIf(writeSize, " (Group ID) and " & ColLetter(colSize) & " (Group Size).", "."), _
           vbInformation, "GroupParcels"

    Exit Sub

' -----------------------------------------------------------------------
' Error handlers
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
    Dim a As Double, c As Double

    dLat = (lat2 - lat1) * PI / 180#
    dLon = (lon2 - lon1) * PI / 180#

    a = Sin(dLat / 2) ^ 2 + _
        Cos(lat1 * PI / 180#) * Cos(lat2 * PI / 180#) * Sin(dLon / 2) ^ 2

    c = 2 * Atn(Sqr(a) / Sqr(1 - a))   ' atan2 equivalent: 2*atan2(sqrt(a),sqrt(1-a))

    HaversineMeters = EARTH_RADIUS_M * c
End Function

'===============================================================================
' Union-Find: Find root with path compression (iterative)
'===============================================================================
Private Function FindRoot(parent() As Long, ByVal x As Long) As Long
    ' Walk to the root.
    Dim root As Long
    root = x
    Do While parent(root) <> root
        root = parent(root)
    Loop
    ' Path compression — point every node on the path directly to root.
    Dim next As Long
    Do While parent(x) <> root
        next = parent(x)
        parent(x) = root
        x = next
    Loop
    FindRoot = root
End Function

'===============================================================================
' Union-Find: Union by size — attach smaller tree under larger
'===============================================================================
Private Sub UnionSets(parent() As Long, ufSize() As Long, _
                       ByVal a As Long, ByVal b As Long)
    Dim rootA As Long, rootB As Long
    rootA = FindRoot(parent, a)
    rootB = FindRoot(parent, b)
    If rootA = rootB Then Exit Sub   ' already in the same set

    ' Merge smaller into larger.
    If ufSize(rootA) < ufSize(rootB) Then
        parent(rootA) = rootB
        ufSize(rootB) = ufSize(rootB) + ufSize(rootA)
    Else
        parent(rootB) = rootA
        ufSize(rootA) = ufSize(rootA) + ufSize(rootB)
    End If
End Sub

'===============================================================================
' Utility: convert a 1-based column number to its letter(s) (e.g. 28 -> "AB")
'===============================================================================
Private Function ColLetter(ByVal colNum As Long) As String
    Dim s As String
    s = ""
    Do While colNum > 0
        Dim rem As Long
        rem = (colNum - 1) Mod 26
        s = Chr(65 + rem) & s
        colNum = (colNum - 1 - rem) \ 26
    Loop
    ColLetter = s
End Function
