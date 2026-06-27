Attribute VB_Name = "ParcelLeadAssembly"
'===============================================================================
' ParcelLeadAssembly.bas
' Assembles land-siting prospecting leads from a parcel list.
'
' CAVEAT: centroid + acreage only approximates contiguity. It assumes parcels
' are reasonably compact and near-square. Long, thin, or irregular parcels can
' misclassify, and true contiguity requires actual parcel polygons (GIS). This
' is a fast prospecting screen, not ground truth.
'
' HOW TO IMPORT AND RUN:
'   1. In Excel, press Alt+F11 to open the VBA editor.
'   2. Right-click your workbook in the Project window > Import File >
'      select this .bas file.
'   3. Close the editor, press Alt+F8, select AssembleLeads, click Run.
'   4. To see debug output, open the Immediate window first (Ctrl+G in editor).
'===============================================================================
Option Explicit

' -----------------------------------------------------------------------
' CONFIG BLOCK — tune these constants for your campaign.
' Columns and data-start row are chosen at runtime via prompts.
' -----------------------------------------------------------------------

' A lead must reach this total assembled acreage to qualify for output.
Private Const MIN_LEAD_ACRES As Double = 1000#

' Maximum straight-line distance (miles) from an anchor's centroid.
' Parcels beyond this radius are excluded from that anchor's assemblage
' even if they touch a member parcel.
Private Const REACH_MILES As Double = 2#

' Adjacency tolerance multiplier.
'   1.0 ≈ strict edge-sharing; raise toward 1.2–1.4 to assemble more loosely
'   but risk chaining parcels that only nearly touch.
Private Const TOLERANCE As Double = 1.0

' How to derive PAIR_BASE from the two parcels' half-extents (he1, he2).
'   "sum" — he1 + he2:  PAIR_BASE equals the centroid-to-centroid distance
'            between two parcels that share an edge, regardless of their
'            relative sizes. Two parcels that are truly edge-adjacent will
'            always satisfy dist <= TOLERANCE * (he1+he2) at TOLERANCE=1.0.
'            DEFAULT — connects mixed-size neighbours correctly.
'   "min" — 2 × Min(he1, he2):  stricter; only links similarly-sized
'            neighbours. Use only if you see runaway over-merging. It
'            under-connects parcel sets where sizes vary widely, so it
'            is not the default.
Private Const PAIR_RULE As String = "sum"

' Shape model used to derive each parcel's half-extent from its acreage.
'   "square" — half-extent = Sqrt(area_m2) / 2   (half-side of equiv. square)
'   "circle" — half-extent = Sqrt(area_m2 / Pi)  (radius of equiv. circle)
Private Const SHAPE_MODEL As String = "square"

' Print diagnostics to the Immediate window after the run?
Private Const DEBUG_MODE As Boolean = True

' -----------------------------------------------------------------------
' Unit-conversion constants — do not change these.
' -----------------------------------------------------------------------
Private Const ACRES_TO_M2  As Double = 4046.86
Private Const EARTH_RADIUS As Double = 6371000#
Private Const MILES_TO_M   As Double = 1609.344
Private Const PI           As Double = 3.14159265358979

'===============================================================================
' Main entry point
'===============================================================================
Public Sub AssembleLeads()

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

    ' Data-start row (default 2 — row 1 is the header).
    userInput = InputBox( _
        "Enter the row number where parcel DATA starts." & vbCrLf & _
        "(Row 1 is typically the header row, so data starts at row 2.)", _
        "Data Start Row", "2")
    If userInput = "" Then
        MsgBox "Cancelled.", vbInformation, "AssembleLeads"
        Exit Sub
    End If
    If Not IsNumeric(userInput) Or CLng(userInput) < 1 Then
        MsgBox "Invalid row number. Macro cancelled.", vbExclamation, "AssembleLeads"
        Exit Sub
    End If
    dataStart = CLng(userInput)

    ' Column pickers — Application.InputBox Type:=8 returns a Range object.
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
        "Click any cell in the first OUTPUT column (Lead ID goes here;" & vbCrLf & _
        "Anchor, Lead Acres, and Lead Parcels fill the next 3 columns to the right).", _
        "Select First Output Column", Type:=8)
    If Not rPick.Worksheet Is ws Then GoTo WrongSheet
    colOut = rPick.Column

    On Error GoTo 0

    ' -----------------------------------------------------------------------
    ' STAGE 2: Determine the data extent.
    ' -----------------------------------------------------------------------

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colID).End(xlUp).Row
    If lastRow < dataStart Then
        MsgBox "No data rows found at or below row " & dataStart & ". Macro cancelled.", _
               vbExclamation, "AssembleLeads"
        Exit Sub
    End If

    Dim nRows As Long
    nRows = lastRow - dataStart + 1

    ' -----------------------------------------------------------------------
    ' STAGE 3: Read each needed column into memory in one Range.Value call.
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
    ' STAGE 4: Guard — warn before overwriting any of the 4 output columns.
    ' -----------------------------------------------------------------------

    Dim c As Long
    For c = colOut To colOut + 3
        Dim chkRange As Range
        Set chkRange = ws.Range(ws.Cells(dataStart, c), ws.Cells(lastRow, c))
        If Application.WorksheetFunction.CountA(chkRange) > 0 Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox( _
                "One or more of the 4 output columns (" & _
                ColLetter(colOut) & "–" & ColLetter(colOut + 3) & _
                ") already contains data in rows " & dataStart & ":" & lastRow & "." & _
                vbCrLf & vbCrLf & "Do you want to OVERWRITE all four columns?", _
                vbQuestion + vbYesNo, "Output Columns Not Empty")
            If ans = vbNo Then
                MsgBox "Macro cancelled — output columns left untouched.", _
                       vbInformation, "AssembleLeads"
                Exit Sub
            End If
            Exit For
        End If
    Next c

    ' -----------------------------------------------------------------------
    ' STAGE 5: Validate rows and build compact in-memory working arrays.
    '
    ' Only valid rows enter the algorithm. srcRow(i) maps the compact index i
    ' back to the 1-based position within arrLat/arrLon/etc. (= sheet row
    ' dataStart + srcRow(i) - 1).
    ' -----------------------------------------------------------------------

    Dim latArr()  As Double   ' valid parcels only
    Dim lonArr()  As Double
    Dim acreArr() As Double
    Dim heArr()   As Double   ' half-extent: side/2 for square, radius for circle
    Dim srcRow()  As Long     ' srcRow(i) = 1-based index into arrXxx arrays

    ReDim latArr(1 To nRows)
    ReDim lonArr(1 To nRows)
    ReDim acreArr(1 To nRows)
    ReDim heArr(1 To nRows)
    ReDim srcRow(1 To nRows)

    Dim n       As Long   ' count of valid parcels
    Dim skipped As Long
    Dim i       As Long
    Dim latV    As Double, lonV As Double, acreV As Double, areaM2 As Double
    Dim firstOK As Boolean

    For i = 1 To nRows

        If IsEmpty(arrLat(i, 1)) Or Not IsNumeric(arrLat(i, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        latV = CDbl(arrLat(i, 1))
        If latV < -90# Or latV > 90# Then skipped = skipped + 1: GoTo NextRow

        If IsEmpty(arrLon(i, 1)) Or Not IsNumeric(arrLon(i, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        lonV = CDbl(arrLon(i, 1))
        If lonV < -180# Or lonV > 180# Then skipped = skipped + 1: GoTo NextRow

        If IsEmpty(arrAcre(i, 1)) Or Not IsNumeric(arrAcre(i, 1)) Then
            skipped = skipped + 1: GoTo NextRow
        End If
        acreV = CDbl(arrAcre(i, 1))
        If acreV <= 0# Then skipped = skipped + 1: GoTo NextRow

        n = n + 1
        srcRow(n)  = i
        latArr(n)  = latV
        lonArr(n)  = lonV
        acreArr(n) = acreV
        areaM2     = acreV * ACRES_TO_M2

        Select Case LCase(Trim(SHAPE_MODEL))
            Case "circle"
                heArr(n) = Sqr(areaM2 / PI)     ' radius
            Case Else                            ' "square" (default)
                heArr(n) = Sqr(areaM2) / 2#      ' half-side
        End Select

        ' Debug sanity line for the first valid parcel.
        If DEBUG_MODE And Not firstOK Then
            firstOK = True
            Debug.Print "--- AssembleLeads sanity (first valid parcel) ---"
            Debug.Print "  Parcel ID  : " & CStr(arrID(i, 1))
            Debug.Print "  Acreage    : " & acreV & " acres"
            Debug.Print "  Area       : " & Format(areaM2, "#,##0.0") & " m²"
            Debug.Print "  Width (2×he): " & Format(heArr(n) * 2#, "#,##0.0") & " m" & _
                        "   (expect ~100–1 000 m for typical parcels;" & _
                        " if >10 000 m, check acreage units)"
        End If

NextRow:
    Next i

    If n = 0 Then
        MsgBox "No valid parcel rows found. Check your column selections and data.", _
               vbExclamation, "AssembleLeads"
        Exit Sub
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 6: Sort valid-parcel indices by acreage descending.
    '
    ' sortIdx(1..n) is the order in which the algorithm considers anchors.
    ' The largest unassigned parcel becomes each new anchor.
    ' -----------------------------------------------------------------------

    Dim sortIdx() As Long
    ReDim sortIdx(1 To n)
    For i = 1 To n
        sortIdx(i) = i
    Next i
    Call QSortByKeyDesc(sortIdx, acreArr, 1, n)

    ' -----------------------------------------------------------------------
    ' STAGE 7: Greedy BFS lead assembly.
    '
    ' We walk sortIdx (largest acreage first). Each unassigned parcel becomes
    ' an ANCHOR. From the anchor we do BFS, pulling in any unassigned parcel
    ' that is:
    '   (a) within REACH_MILES of the anchor's centroid, AND
    '   (b) contiguous (per ContiguityThresh) to the current BFS frontier node.
    '
    ' A parcel is claimed (assigned and labeled) the instant the BFS reaches
    ' it — no later resolution pass needed. Because anchors are processed
    ' largest-first, a large anchor always wins contested ground over a smaller
    ' one that might have claimed the same parcels.
    '
    ' After the BFS, a lead that doesn't reach MIN_LEAD_ACRES is dropped
    ' (Stage 8) and its parcels stay blank in the output.
    ' -----------------------------------------------------------------------

    Dim assigned()     As Boolean   ' assigned(i) = True once parcel i joins any lead
    Dim parcelLead()   As Long      ' parcelLead(i) = internal lead number (0 = none)
    Dim isAnchorFlag() As Boolean   ' True if parcel i is an anchor
    ReDim assigned(1 To n)
    ReDim parcelLead(1 To n)
    ReDim isAnchorFlag(1 To n)

    ' Lead metadata — allocated to worst-case n leads; only first nLeads used.
    Dim leadAcres()  As Double   ' total assembled acreage
    Dim leadParcels() As Long    ' parcel count
    Dim leadAnchor() As Long     ' valid-parcel index of the anchor
    ReDim leadAcres(1 To n)
    ReDim leadParcels(1 To n)
    ReDim leadAnchor(1 To n)
    Dim nLeads As Long           ' number of leads formed
    nLeads = 0

    ' BFS queue (one Long per valid parcel; reused across leads).
    Dim bfsQ() As Long
    ReDim bfsQ(1 To n)

    ' Anchor-to-parcel reach distances: precomputed once per anchor so the
    ' reach check inside the BFS inner loop is a single array lookup.
    Dim reachDist() As Double
    ReDim reachDist(1 To n)

    Dim reachM As Double
    reachM = REACH_MILES * MILES_TO_M

    Dim si As Long, anchorIdx As Long
    Dim qHead As Long, qTail As Long
    Dim curr As Long, j As Long
    Dim contThresh As Double

    For si = 1 To n
        anchorIdx = sortIdx(si)
        If assigned(anchorIdx) Then GoTo NextAnchor

        ' --- New lead: initialise with anchor ---
        nLeads = nLeads + 1
        assigned(anchorIdx)       = True
        isAnchorFlag(anchorIdx)   = True
        parcelLead(anchorIdx)     = nLeads
        leadAcres(nLeads)         = acreArr(anchorIdx)
        leadParcels(nLeads)       = 1
        leadAnchor(nLeads)        = anchorIdx

        ' Precompute Haversine distance from this anchor to every valid parcel.
        ' This makes the per-frontier-node reach check a simple comparison.
        For j = 1 To n
            If Not assigned(j) Then
                reachDist(j) = HaversineMeters( _
                    latArr(anchorIdx), lonArr(anchorIdx), latArr(j), lonArr(j))
            End If
        Next j

        ' --- BFS from anchor ---
        qHead = 1: qTail = 1
        bfsQ(1) = anchorIdx

        Do While qHead <= qTail
            curr = bfsQ(qHead)
            qHead = qHead + 1

            For j = 1 To n
                If Not assigned(j) Then

                    ' (a) Reach check — use precomputed distance from anchor.
                    If reachDist(j) <= reachM Then

                        ' (b) Contiguity check — parcel j must touch frontier node curr.
                        contThresh = ContiguityThresh(heArr(curr), heArr(j))
                        If HaversineMeters(latArr(curr), lonArr(curr), _
                                           latArr(j), lonArr(j)) <= contThresh Then

                            ' Claim parcel j for this lead immediately.
                            assigned(j)          = True
                            parcelLead(j)        = nLeads
                            leadAcres(nLeads)    = leadAcres(nLeads) + acreArr(j)
                            leadParcels(nLeads)  = leadParcels(nLeads) + 1
                            qTail = qTail + 1
                            bfsQ(qTail) = j

                        End If
                    End If
                End If
            Next j
        Loop

NextAnchor:
    Next si

    ' -----------------------------------------------------------------------
    ' STAGE 8: Filter to qualifying leads; rank by total acreage descending.
    '
    ' Leads below MIN_LEAD_ACRES are dropped — their parcels get blank output.
    ' Qualifying leads are sorted and numbered LEAD-001, LEAD-002, … so
    ' LEAD-001 is the largest prospect.
    ' -----------------------------------------------------------------------

    ' Count qualifying leads.
    Dim qualCount As Long
    qualCount = 0
    Dim L As Long
    For L = 1 To nLeads
        If leadAcres(L) >= MIN_LEAD_ACRES Then qualCount = qualCount + 1
    Next L

    ' Extract qualifying leads into parallel arrays for sorting.
    Dim dim1 As Long
    dim1 = IIf(qualCount > 0, qualCount, 1)   ' avoid ReDim with upper = 0
    Dim qOrig()  As Long    ' original lead numbers of qualifying leads
    Dim qAcres() As Double  ' their total acreage (sorted in place)
    ReDim qOrig(1 To dim1)
    ReDim qAcres(1 To dim1)

    Dim qi As Long
    qi = 0
    For L = 1 To nLeads
        If leadAcres(L) >= MIN_LEAD_ACRES Then
            qi = qi + 1
            qOrig(qi)  = L
            qAcres(qi) = leadAcres(L)
        End If
    Next L

    ' Sort qualifying leads by total acreage descending.
    If qualCount > 1 Then Call QSortPairedDesc(qAcres, qOrig, 1, qualCount)

    ' Build finalLabel(L): "LEAD-001", "LEAD-002", … or "" for non-qualifying.
    Dim finalLabel() As String
    ReDim finalLabel(1 To IIf(nLeads > 0, nLeads, 1))

    For qi = 1 To qualCount
        finalLabel(qOrig(qi)) = "LEAD-" & Format(qi, "000")
    Next qi

    ' -----------------------------------------------------------------------
    ' STAGE 9: Compute summary statistics (used by debug report and MsgBox).
    ' -----------------------------------------------------------------------

    Dim parcelsInLeads As Long
    Dim parcelsUngrouped As Long
    For i = 1 To n
        L = parcelLead(i)
        If L > 0 Then
            If finalLabel(L) <> "" Then
                parcelsInLeads = parcelsInLeads + 1
            Else
                parcelsUngrouped = parcelsUngrouped + 1
            End If
        Else
            parcelsUngrouped = parcelsUngrouped + 1
        End If
    Next i

    ' -----------------------------------------------------------------------
    ' STAGE 10: Build output arrays (4 columns) and write to sheet in one shot.
    ' -----------------------------------------------------------------------

    Dim outLeadID()  As Variant   ' colOut     — "LEAD-001" or ""
    Dim outAnchor()  As Variant   ' colOut + 1 — "ANCHOR" or ""
    Dim outAcres()   As Variant   ' colOut + 2 — total lead acreage or ""
    Dim outParcels() As Variant   ' colOut + 3 — lead parcel count or ""
    ReDim outLeadID(1 To nRows, 1 To 1)
    ReDim outAnchor(1 To nRows, 1 To 1)
    ReDim outAcres(1 To nRows, 1 To 1)
    ReDim outParcels(1 To nRows, 1 To 1)

    ' Initialise to empty string — skipped and ungrouped rows stay blank.
    For i = 1 To nRows
        outLeadID(i, 1)  = ""
        outAnchor(i, 1)  = ""
        outAcres(i, 1)   = ""
        outParcels(i, 1) = ""
    Next i

    ' Fill in qualifying parcel rows.
    Dim r As Long
    For i = 1 To n
        L = parcelLead(i)
        If L > 0 Then
            If finalLabel(L) <> "" Then
                r = srcRow(i)
                outLeadID(r, 1)  = finalLabel(L)
                outAnchor(r, 1)  = IIf(isAnchorFlag(i), "ANCHOR", "")
                outAcres(r, 1)   = leadAcres(L)
                outParcels(r, 1) = leadParcels(L)
            End If
        End If
    Next i

    ' Write all four output columns in single Range.Value assignments.
    ws.Range(ws.Cells(dataStart, colOut),     ws.Cells(lastRow, colOut)).Value     = outLeadID
    ws.Range(ws.Cells(dataStart, colOut + 1), ws.Cells(lastRow, colOut + 1)).Value = outAnchor
    ws.Range(ws.Cells(dataStart, colOut + 2), ws.Cells(lastRow, colOut + 2)).Value = outAcres
    ws.Range(ws.Cells(dataStart, colOut + 3), ws.Cells(lastRow, colOut + 3)).Value = outParcels

    ' Write column headers into the row above dataStart, if that row exists
    ' and the header cells are currently blank.
    Dim hdrRow As Long
    hdrRow = dataStart - 1
    If hdrRow >= 1 Then
        Dim hdr(0 To 3) As String
        hdr(0) = "Lead ID"
        hdr(1) = "Anchor"
        hdr(2) = "Lead Acres"
        hdr(3) = "Lead Parcels"
        For c = 0 To 3
            If Trim(CStr(ws.Cells(hdrRow, colOut + c).Value)) = "" Then
                ws.Cells(hdrRow, colOut + c).Value = hdr(c)
            End If
        Next c
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 11: Debug report (when DEBUG_MODE = True).
    ' -----------------------------------------------------------------------

    If DEBUG_MODE Then
        Debug.Print ""
        Debug.Print "=== AssembleLeads debug report ==="
        Debug.Print "Settings:  MIN_LEAD_ACRES=" & MIN_LEAD_ACRES & _
                    "  REACH_MILES=" & REACH_MILES & _
                    "  TOLERANCE=" & TOLERANCE & _
                    "  PAIR_RULE=" & PAIR_RULE & _
                    "  SHAPE_MODEL=" & SHAPE_MODEL
        Debug.Print "Parcels valid: " & n & "  |  Skipped: " & skipped
        Debug.Print "Leads formed: " & nLeads & "  |  Qualifying (>= " & _
                    MIN_LEAD_ACRES & " ac): " & qualCount
        Debug.Print "Parcels in qualifying leads: " & parcelsInLeads & _
                    "  |  Ungrouped / dropped: " & parcelsUngrouped

        If qualCount > 0 Then
            Dim topN As Long
            topN = IIf(qualCount < 10, qualCount, 10)
            Debug.Print ""
            Debug.Print "Top " & topN & " qualifying leads (ranked by total acreage):"
            Debug.Print "  " & PadR("Lead ID", 10) & _
                        PadL("Total Acres", 14) & _
                        PadL("Parcels", 10) & _
                        "  Anchor Parcel ID"
            Dim k As Long
            For k = 1 To topN
                L = qOrig(k)
                Dim ancIdx As Long
                ancIdx = leadAnchor(L)
                Debug.Print "  " & PadR(finalLabel(L), 10) & _
                            PadL(Format(leadAcres(L), "#,##0.0"), 14) & _
                            PadL(CStr(leadParcels(L)), 10) & "  " & _
                            CStr(arrID(srcRow(ancIdx), 1))
            Next k
        End If
        Debug.Print "==================================="
    End If

    ' -----------------------------------------------------------------------
    ' STAGE 12: Final summary MsgBox.
    ' -----------------------------------------------------------------------

    MsgBox "AssembleLeads complete." & vbCrLf & vbCrLf & _
           "  Valid parcels          : " & n & vbCrLf & _
           "  Rows skipped           : " & skipped & vbCrLf & _
           "  Total leads formed     : " & nLeads & vbCrLf & _
           "  Qualifying leads       : " & qualCount & _
               "  (>= " & MIN_LEAD_ACRES & " acres)" & vbCrLf & _
           "  Parcels in leads       : " & parcelsInLeads & vbCrLf & _
           "  Ungrouped / dropped    : " & parcelsUngrouped & vbCrLf & vbCrLf & _
           "Results written to columns " & ColLetter(colOut) & _
           "–" & ColLetter(colOut + 3) & "." & vbCrLf & _
           "Sort by Lead Acres descending to get a ranked prospecting list.", _
           vbInformation, "AssembleLeads"

    Exit Sub

' -----------------------------------------------------------------------
' Error / cancel handlers
' -----------------------------------------------------------------------
UserCancelled:
    MsgBox "Column selection cancelled. Macro aborted.", vbInformation, "AssembleLeads"
    Exit Sub

WrongSheet:
    MsgBox "All columns must be on the same worksheet. Macro cancelled.", _
           vbExclamation, "AssembleLeads"
    Exit Sub

End Sub

'===============================================================================
' Contiguity threshold
' Returns the maximum centroid-to-centroid distance for two parcels to be
' considered touching, given their half-extents he1 and he2.
'===============================================================================
Private Function ContiguityThresh(he1 As Double, he2 As Double) As Double
    Dim pairBase As Double
    If LCase(Trim(PAIR_RULE)) = "sum" Then
        pairBase = he1 + he2
    ElseIf he1 < he2 Then
        pairBase = 2# * he1
    Else
        pairBase = 2# * he2
    End If
    ContiguityThresh = TOLERANCE * pairBase
End Function

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
' Atan2(y, x) — standard two-argument arctangent in radians.
' VBA provides only single-argument Atn, so we derive Atan2 here.
'===============================================================================
Private Function Atan2(y As Double, x As Double) As Double
    If x > 0# Then
        Atan2 = Atn(y / x)
    ElseIf x < 0# Then
        If y >= 0# Then
            Atan2 = Atn(y / x) + PI
        Else
            Atan2 = Atn(y / x) - PI
        End If
    Else
        If y > 0# Then
            Atan2 = PI / 2#
        ElseIf y < 0# Then
            Atan2 = -PI / 2#
        Else
            Atan2 = 0#       ' degenerate: both args zero
        End If
    End If
End Function

'===============================================================================
' QSortByKeyDesc — quicksort idx(lo..hi) by key(idx(i)) descending.
' idx holds integer pointers into key; key is read-only.
' Used in Stage 6 to order valid parcels by acreage.
'===============================================================================
Private Sub QSortByKeyDesc(idx() As Long, key() As Double, _
                            ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub

    Dim pivot As Double
    Dim i As Long, j As Long, tmp As Long

    pivot = key(idx((lo + hi) \ 2))
    i = lo: j = hi

    Do While i <= j
        Do While key(idx(i)) > pivot: i = i + 1: Loop
        Do While key(idx(j)) < pivot: j = j - 1: Loop
        If i <= j Then
            tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    Call QSortByKeyDesc(idx, key, lo, j)
    Call QSortByKeyDesc(idx, key, i, hi)
End Sub

'===============================================================================
' QSortPairedDesc — quicksort key(lo..hi) descending, keeping secondary(i)
' in sync. Used in Stage 8 to rank qualifying leads by total acreage.
'===============================================================================
Private Sub QSortPairedDesc(key() As Double, secondary() As Long, _
                             ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub

    Dim pivot As Double
    Dim i As Long, j As Long
    Dim tmpD As Double, tmpL As Long

    pivot = key((lo + hi) \ 2)
    i = lo: j = hi

    Do While i <= j
        Do While key(i) > pivot: i = i + 1: Loop
        Do While key(j) < pivot: j = j - 1: Loop
        If i <= j Then
            tmpD = key(i): key(i) = key(j): key(j) = tmpD
            tmpL = secondary(i): secondary(i) = secondary(j): secondary(j) = tmpL
            i = i + 1: j = j - 1
        End If
    Loop

    Call QSortPairedDesc(key, secondary, lo, j)
    Call QSortPairedDesc(key, secondary, i, hi)
End Sub

'===============================================================================
' ColLetter — 1-based column number to column letter(s), e.g. 28 → "AB"
'===============================================================================
Private Function ColLetter(ByVal c As Long) As String
    Dim s As String
    s = ""
    Do While c > 0
        Dim rr As Long
        rr = (c - 1) Mod 26
        s = Chr(65 + rr) & s
        c = (c - 1 - rr) \ 26
    Loop
    ColLetter = s
End Function

'===============================================================================
' PadR / PadL — string padding for Immediate-window table alignment
'===============================================================================
Private Function PadR(ByVal s As String, ByVal w As Long) As String
    If Len(s) >= w Then PadR = Left(s, w) Else PadR = s & Space(w - Len(s))
End Function

Private Function PadL(ByVal s As String, ByVal w As Long) As String
    If Len(s) >= w Then PadL = Left(s, w) Else PadL = Space(w - Len(s)) & s
End Function
