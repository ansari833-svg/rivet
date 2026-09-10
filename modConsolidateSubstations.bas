Attribute VB_Name = "modConsolidateSubstations"
Option Explicit

' ============================================================
'  modConsolidateSubstations.bas
'  Multi-Workbook Substation Consolidator (Windows Excel)
'
'  Public entry point:  ConsolidateSubstationData
'
'  Prompts the user to select multiple source workbooks, each of
'  which is expected to hold exactly three sheets:
'
'      1. Cover Page
'      2. Gen Assumptions
'      3. A substation sheet whose tab name encodes the substation
'         name and (usually) its voltage class.
'
'  Only the third sheet carries data to consolidate. Every selected
'  file is validated up front; nothing is written to the output sheet
'  until the whole validation pass completes and the user confirms.
'  The data from each third sheet is then block-copied (values only)
'  into a single "Consolidated" sheet in ThisWorkbook, prefixed with
'  four metadata columns, and a per-file run log is written alongside.
'
'  Target platform: Windows Excel only. No Mac fallback.
' ============================================================

' ---------- Configuration block (edit here only) ------------

Private Const HEADER_ROW        As Long = 1              ' Header row on each source data sheet
Private Const DATA_SHEET_INDEX  As Long = 3              ' 1-based index of the sheet to consolidate
Private Const OUTPUT_SHEET_NAME As String = "Consolidated"
Private Const LOG_SHEET_NAME    As String = "_Import Log"
Private Const RESERVED_SHEET_1  As String = "Cover Page"
Private Const RESERVED_SHEET_2  As String = "Gen Assumptions"

' ---------- Fixed metadata layout (do not change lightly) ---

Private Const META_COL_COUNT    As Long = 4              ' Source File | Source Sheet | Substation Name | Voltage (kV)
Private Const EXCEL_MAX_ROWS     As Long = 1048576
Private Const EXCEL_MAX_TAB_LEN  As Long = 31

' ---------- Per-file working record --------------------------

Private Type TFileRecord
    FilePath        As String
    SheetName       As String   ' raw third-sheet tab name, verbatim
    Substation      As String   ' parsed substation name
    Voltage         As Double   ' parsed voltage
    VoltageParsed   As Boolean  ' True when a voltage was identified
    HeaderSig       As String   ' header signature for this file
    UsedCols        As Long     ' number of used columns on the data sheet
    FirstDataRow    As Long
    LastDataRow     As Long
    IsValid         As Boolean
    Status          As String   ' Imported | Skipped | Failed
    Message         As String   ' failure / note reason
    RowsImported    As Long
End Type

' ============================================================
'  PUBLIC ENTRY POINT
' ============================================================

'--------------------------------------------------------------
' ConsolidateSubstationData
'   Orchestrates the full run: file selection, validation pass,
'   user confirmation, consolidation pass, formatting and logging.
'   Parameters: none.  Returns: nothing (side effects only).
'--------------------------------------------------------------
Public Sub ConsolidateSubstationData()

    Dim files() As String
    Dim recs() As TFileRecord
    Dim fileCount As Long
    Dim proceed As Boolean
    Dim includeMismatched As Boolean
    Dim savedScreen As Boolean, savedEvents As Boolean
    Dim savedAlerts As Boolean, savedCalc As XlCalculation

    ' --- Select files -------------------------------------------------
    If Not SelectSourceFiles(files, fileCount) Then
        MsgBox "No files were selected. Nothing to do.", vbInformation, "Consolidate Substations"
        Exit Sub
    End If

    ' --- Capture and set application state ----------------------------
    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedAlerts = Application.DisplayAlerts
    savedCalc = Application.Calculation

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' --- Validation pass (no output written yet) ----------------------
    ValidateAllFiles files, fileCount, recs

    ' --- Summarise and confirm with the user --------------------------
    proceed = ConfirmValidation(recs, includeMismatched)
    If Not proceed Then
        WriteLogSheet recs                 ' still log what was inspected
        GoTo ExitHandler
    End If

    ' --- Consolidation pass -------------------------------------------
    RunConsolidation recs, includeMismatched
    WriteLogSheet recs

ExitHandler:
    Application.ScreenUpdating = savedScreen
    Application.EnableEvents = savedEvents
    Application.DisplayAlerts = savedAlerts
    Application.Calculation = savedCalc
    Exit Sub

ErrHandler:
    MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Consolidate Substations"
    Resume ExitHandler

End Sub

' ============================================================
'  FILE SELECTION
' ============================================================

'--------------------------------------------------------------
' SelectSourceFiles
'   Shows a multi-select Excel file picker rooted at ThisWorkbook's
'   folder.
'   Parameters:
'     outFiles () - receives the selected full paths (1-based)
'     outCount    - receives the number of files selected
'   Returns: True when at least one file was chosen, else False.
'--------------------------------------------------------------
Private Function SelectSourceFiles(ByRef outFiles() As String, _
                                   ByRef outCount As Long) As Boolean

    Dim fd As FileDialog
    Dim i As Long

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select substation workbooks to consolidate"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"
        .InitialFileName = ThisWorkbook.Path & Application.PathSeparator
    End With

    If fd.Show <> -1 Then
        outCount = 0
        SelectSourceFiles = False
        Exit Function
    End If

    outCount = fd.SelectedItems.Count
    If outCount = 0 Then
        SelectSourceFiles = False
        Exit Function
    End If

    ReDim outFiles(1 To outCount)
    For i = 1 To outCount
        outFiles(i) = fd.SelectedItems(i)
    Next i

    SelectSourceFiles = True

End Function

' ============================================================
'  VALIDATION PASS
' ============================================================

'--------------------------------------------------------------
' ValidateAllFiles
'   Validates every selected file and populates a TFileRecord array.
'   The first file to pass all structural checks sets the reference
'   header signature; later files are compared against it.
'   Parameters:
'     files ()   - selected full paths (1-based)
'     fileCount  - number of selected files
'     recs ()    - receives one populated TFileRecord per file
'   Returns: nothing (recs filled by reference).
'--------------------------------------------------------------
Private Sub ValidateAllFiles(ByRef files() As String, _
                             ByVal fileCount As Long, _
                             ByRef recs() As TFileRecord)

    Dim i As Long
    Dim refSig As String
    Dim haveRef As Boolean

    ReDim recs(1 To fileCount)
    haveRef = False

    For i = 1 To fileCount
        recs(i).FilePath = files(i)
        ValidateOneFile recs(i)

        If recs(i).IsValid Then
            If Not haveRef Then
                refSig = recs(i).HeaderSig
                haveRef = True
            ElseIf StrComp(recs(i).HeaderSig, refSig, vbTextCompare) <> 0 Then
                recs(i).IsValid = False
                recs(i).Status = "Failed"
                recs(i).Message = "Header signature mismatch vs first valid file"
            End If
        End If
    Next i

End Sub

'--------------------------------------------------------------
' ValidateOneFile
'   Opens one workbook read-only and performs every structural check.
'   Populates the record's validity, status, message, and (on success)
'   its parsed name/voltage, header signature and data-row range.
'   Parameters:
'     rec - a TFileRecord whose FilePath is already set; filled here.
'   Returns: nothing (rec updated by reference).
'--------------------------------------------------------------
Private Sub ValidateOneFile(ByRef rec As TFileRecord)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As String

    rec.IsValid = False
    rec.Status = "Failed"

    On Error GoTo OpenFail
    Set wb = Application.Workbooks.Open( _
                Filename:=rec.FilePath, _
                ReadOnly:=True, _
                UpdateLinks:=0, _
                Password:="", _
                WriteResPassword:="")
    On Error GoTo CloseFail

    If wb.Sheets.Count < DATA_SHEET_INDEX Then
        rec.Message = "Workbook has fewer than " & DATA_SHEET_INDEX & " sheets"
        GoTo CloseAndExit
    End If

    If Not TypeOf wb.Sheets(DATA_SHEET_INDEX) Is Worksheet Then
        rec.Message = "Sheet " & DATA_SHEET_INDEX & " is not a worksheet (chart sheet?)"
        GoTo CloseAndExit
    End If
    Set ws = wb.Sheets(DATA_SHEET_INDEX)

    rec.SheetName = ws.Name
    nm = Trim$(ws.Name)
    If StrComp(nm, RESERVED_SHEET_1, vbTextCompare) = 0 _
       Or StrComp(nm, RESERVED_SHEET_2, vbTextCompare) = 0 Then
        rec.Message = "Third sheet name is a reserved name (" & nm & ")"
        GoTo CloseAndExit
    End If

    If Not MeasureDataSheet(ws, rec) Then GoTo CloseAndExit

    rec.HeaderSig = ComputeHeaderSignature(ws, rec.UsedCols)
    ParseSheetName rec.SheetName, rec.Substation, rec.Voltage, rec.VoltageParsed

    rec.IsValid = True
    rec.Status = "Imported"     ' provisional; may become Skipped later
    rec.Message = ""

CloseAndExit:
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Exit Sub

OpenFail:
    rec.Message = "Could not open (open elsewhere, password-protected, or corrupt): " _
                  & Err.Description
    rec.IsValid = False
    Exit Sub

CloseFail:
    rec.Message = "Error while inspecting workbook: " & Err.Description
    rec.IsValid = False
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0

End Sub

'--------------------------------------------------------------
' MeasureDataSheet
'   Determines the used-column count and the data-row range on the
'   third sheet, and verifies the header row is non-empty with at
'   least one data row below it. Also spans any data columns that
'   extend further right than the header row.
'   Parameters:
'     ws  - the third worksheet
'     rec - record to receive UsedCols / FirstDataRow / LastDataRow
'   Returns: True when the sheet has a usable header + data, else
'            False (rec.Message set to the reason).
'--------------------------------------------------------------
Private Function MeasureDataSheet(ByVal ws As Worksheet, _
                                  ByRef rec As TFileRecord) As Boolean

    Dim ur As Range
    Dim lastCol As Long, lastRow As Long
    Dim headerCells As Range

    Set ur = ws.UsedRange
    If ur Is Nothing Then
        rec.Message = "Third sheet is empty"
        MeasureDataSheet = False
        Exit Function
    End If

    ' Rightmost used column across the whole used range (data may run
    ' further right than the header row).
    lastCol = ur.Column + ur.Columns.Count - 1
    lastRow = ur.Row + ur.Rows.Count - 1

    If lastRow <= HEADER_ROW Then
        rec.Message = "No data rows below the header"
        MeasureDataSheet = False
        Exit Function
    End If

    ' Header row must contain at least one non-blank value.
    Set headerCells = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, lastCol))
    If Application.WorksheetFunction.CountA(headerCells) = 0 Then
        rec.Message = "Header row " & HEADER_ROW & " is empty"
        MeasureDataSheet = False
        Exit Function
    End If

    rec.UsedCols = lastCol
    rec.FirstDataRow = HEADER_ROW + 1
    rec.LastDataRow = lastRow
    MeasureDataSheet = True

End Function

'--------------------------------------------------------------
' ComputeHeaderSignature
'   Builds a stable signature from the trimmed header values across
'   all used columns. Merged header cells report their value only in
'   the top-left cell, which is the natural behaviour here.
'   Parameters:
'     ws       - the third worksheet
'     usedCols - number of columns to include
'   Returns: a delimited signature string (case preserved; comparison
'            is done case-insensitively by the caller).
'--------------------------------------------------------------
Private Function ComputeHeaderSignature(ByVal ws As Worksheet, _
                                        ByVal usedCols As Long) As String

    Dim c As Long
    Dim parts() As String

    ReDim parts(1 To usedCols)
    For c = 1 To usedCols
        parts(c) = Trim$(CStr(ws.Cells(HEADER_ROW, c).Value))
    Next c

    ComputeHeaderSignature = Join(parts, "|")

End Function

' ============================================================
'  SHEET-NAME PARSING  (lenient, best-effort, tunable in isolation)
' ============================================================

'--------------------------------------------------------------
' ParseSheetName
'   Extracts a substation name and voltage from a semi-structured
'   tab name. Parsing NEVER fails: when no voltage can be identified
'   the full (trimmed) tab name becomes the substation name, voltage
'   is left unset, and parsed = False.
'
'   Rule: scan every numeric token in the tab name and keep the LAST
'   one that plausibly reads as a voltage (a bare number, optionally
'   followed by a "kV" marker in any casing). That token becomes the
'   voltage; everything else, with separators and the kV marker
'   stripped and whitespace collapsed, becomes the substation name.
'
'   Parameters:
'     rawName    - the verbatim tab name
'     outName    - receives the parsed substation name
'     outVoltage - receives the parsed voltage (Double)
'     outParsed  - True when a voltage was identified, else False
'   Returns: nothing (out params set by reference).
'--------------------------------------------------------------
Private Sub ParseSheetName(ByVal rawName As String, _
                           ByRef outName As String, _
                           ByRef outVoltage As Double, _
                           ByRef outParsed As Boolean)

    Dim work As String
    Dim i As Long, n As Long
    Dim ch As String
    Dim tokStart As Long, tokEnd As Long
    Dim bestStart As Long, bestEnd As Long
    Dim bestVal As Double
    Dim found As Boolean
    Dim token As String

    outParsed = False
    outVoltage = 0
    work = Trim$(rawName)
    n = Len(work)
    found = False
    i = 1

    ' Walk the string, isolating numeric tokens (digits and one dot).
    Do While i <= n
        ch = Mid$(work, i, 1)
        If IsDigitOrDot(ch) Then
            tokStart = i
            Do While i <= n And IsDigitOrDot(Mid$(work, i, 1))
                i = i + 1
            Loop
            tokEnd = i - 1
            token = Mid$(work, tokStart, tokEnd - tokStart + 1)
            If IsPlausibleVoltage(token) Then
                bestStart = tokStart
                bestEnd = tokEnd
                bestVal = CDbl(token)
                found = True     ' keep overwriting -> ends on the LAST match
            End If
        Else
            i = i + 1
        End If
    Loop

    If Not found Then
        outName = CollapseName(work)
        If Len(outName) = 0 Then outName = work
        Exit Sub
    End If

    ' Voltage found: pull it out, drop a trailing "kV" marker if present,
    ' and collapse whatever remains into the substation name.
    outVoltage = bestVal
    outParsed = True

    Dim remainder As String
    remainder = Left$(work, bestStart - 1) & " " & StripKvMarker(Mid$(work, bestEnd + 1))
    outName = CollapseName(remainder)
    If Len(outName) = 0 Then outName = "(unnamed)"

End Sub

'--------------------------------------------------------------
' IsDigitOrDot
'   Helper: True when a single character is 0-9 or a period.
'   Parameters: ch - a one-character string.  Returns: Boolean.
'--------------------------------------------------------------
Private Function IsDigitOrDot(ByVal ch As String) As Boolean
    IsDigitOrDot = (ch >= "0" And ch <= "9") Or (ch = ".")
End Function

'--------------------------------------------------------------
' IsPlausibleVoltage
'   Helper: True when a numeric token reads as a real number that
'   could be a voltage (numeric, not a bare or trailing dot, within
'   a sane 0 < v <= 2000 kV range).
'   Parameters: token - a candidate numeric string.  Returns: Boolean.
'--------------------------------------------------------------
Private Function IsPlausibleVoltage(ByVal token As String) As Boolean

    Dim v As Double

    IsPlausibleVoltage = False
    If Len(token) = 0 Then Exit Function
    If Left$(token, 1) = "." Or Right$(token, 1) = "." Then Exit Function
    If InStr(token, ".") <> InStrRev(token, ".") Then Exit Function   ' more than one dot
    If Not IsNumeric(token) Then Exit Function

    v = CDbl(token)
    If v > 0 And v <= 2000 Then IsPlausibleVoltage = True

End Function

'--------------------------------------------------------------
' StripKvMarker
'   Helper: removes a leading "kV" marker (any casing, with optional
'   surrounding separators) from the text that follows the voltage.
'   Parameters: s - the post-voltage remainder.  Returns: cleaned text.
'--------------------------------------------------------------
Private Function StripKvMarker(ByVal s As String) As String

    Dim t As String
    Dim j As Long

    t = s
    ' Skip any leading separators/whitespace.
    j = 1
    Do While j <= Len(t) And IsSeparator(Mid$(t, j, 1))
        j = j + 1
    Loop
    t = Mid$(t, j)

    If Len(t) >= 2 Then
        If StrComp(Left$(t, 2), "kV", vbTextCompare) = 0 Then
            t = Mid$(t, 3)
        End If
    End If

    StripKvMarker = t

End Function

'--------------------------------------------------------------
' IsSeparator
'   Helper: True when a character is treated as a name/voltage
'   separator (space, underscore, hyphen, parentheses, comma, dot).
'   Parameters: ch - a one-character string.  Returns: Boolean.
'--------------------------------------------------------------
Private Function IsSeparator(ByVal ch As String) As Boolean
    Select Case ch
        Case " ", "_", "-", "(", ")", ",", ".", vbTab
            IsSeparator = True
        Case Else
            IsSeparator = False
    End Select
End Function

'--------------------------------------------------------------
' CollapseName
'   Helper: converts separators to spaces and collapses runs of
'   whitespace to single spaces, trimming the result. Used to turn a
'   raw remainder into a clean substation name.
'   Parameters: s - raw text.  Returns: cleaned, single-spaced name.
'--------------------------------------------------------------
Private Function CollapseName(ByVal s As String) As String

    Dim i As Long
    Dim ch As String
    Dim sb As String
    Dim lastWasSpace As Boolean

    lastWasSpace = True     ' suppress leading spaces
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If IsSeparator(ch) Then
            If Not lastWasSpace Then
                sb = sb & " "
                lastWasSpace = True
            End If
        Else
            sb = sb & ch
            lastWasSpace = False
        End If
    Next i

    CollapseName = Trim$(sb)

End Function

' ============================================================
'  USER CONFIRMATION
' ============================================================

'--------------------------------------------------------------
' ConfirmValidation
'   Presents a pass/fail summary and asks the user to confirm. When
'   the only failures are header-signature mismatches, offers abort /
'   skip / include-anyway. Files chosen to be skipped are marked here.
'   Parameters:
'     recs ()          - validated records (updated by reference)
'     includeMismatched- receives True when the user opts to include
'                        signature-mismatched files aligned by column
'   Returns: True to proceed with consolidation, False to abort.
'--------------------------------------------------------------
Private Function ConfirmValidation(ByRef recs() As TFileRecord, _
                                   ByRef includeMismatched As Boolean) As Boolean

    Dim i As Long
    Dim passCount As Long, failCount As Long, mismatchCount As Long
    Dim summary As String
    Dim ans As VbMsgBoxResult

    includeMismatched = False

    For i = LBound(recs) To UBound(recs)
        If recs(i).IsValid Then
            passCount = passCount + 1
        Else
            failCount = failCount + 1
            If InStr(1, recs(i).Message, "Header signature mismatch", vbTextCompare) > 0 Then
                mismatchCount = mismatchCount + 1
            End If
        End If
    Next i

    summary = "Validation complete." & vbCrLf & _
              "Passed: " & passCount & vbCrLf & _
              "Failed: " & failCount & vbCrLf & vbCrLf & _
              BuildFailureList(recs)

    If passCount = 0 Then
        MsgBox summary & vbCrLf & "No valid files to consolidate.", _
               vbExclamation, "Consolidate Substations"
        ConfirmValidation = False
        Exit Function
    End If

    ' Header-signature mismatches get a dedicated three-way choice.
    If mismatchCount > 0 Then
        ans = MsgBox(summary & vbCrLf & _
                     mismatchCount & " file(s) have a mismatched header signature." & vbCrLf & vbCrLf & _
                     "Yes  = INCLUDE them anyway, aligned by column position" & vbCrLf & _
                     "No   = SKIP the mismatched files and consolidate the rest" & vbCrLf & _
                     "Cancel = ABORT the whole run", _
                     vbYesNoCancel + vbQuestion, "Header mismatch")
        Select Case ans
            Case vbCancel
                ConfirmValidation = False
                Exit Function
            Case vbYes
                includeMismatched = True
                MarkMismatchedIncluded recs
            Case vbNo
                includeMismatched = False   ' left as Failed -> skipped
        End Select
    Else
        ans = MsgBox(summary & vbCrLf & "Proceed with consolidation?", _
                     vbOKCancel + vbQuestion, "Consolidate Substations")
        If ans <> vbOK Then
            ConfirmValidation = False
            Exit Function
        End If
    End If

    ConfirmValidation = True

End Function

'--------------------------------------------------------------
' BuildFailureList
'   Helper: builds a human-readable list of failed files and reasons.
'   Parameters: recs () - validated records.  Returns: a text block
'   (empty string when nothing failed).
'--------------------------------------------------------------
Private Function BuildFailureList(ByRef recs() As TFileRecord) As String

    Dim i As Long
    Dim s As String

    For i = LBound(recs) To UBound(recs)
        If Not recs(i).IsValid Then
            s = s & "  - " & GetFileName(recs(i).FilePath) & ": " & recs(i).Message & vbCrLf
        End If
    Next i

    If Len(s) > 0 Then s = "Failures:" & vbCrLf & s
    BuildFailureList = s

End Function

'--------------------------------------------------------------
' MarkMismatchedIncluded
'   Helper: flips header-signature-mismatch failures back to valid so
'   they are consolidated aligned by column position, noting the fact.
'   Parameters: recs () - records (updated by reference). Returns: none.
'--------------------------------------------------------------
Private Sub MarkMismatchedIncluded(ByRef recs() As TFileRecord)

    Dim i As Long

    For i = LBound(recs) To UBound(recs)
        If (Not recs(i).IsValid) _
           And InStr(1, recs(i).Message, "Header signature mismatch", vbTextCompare) > 0 Then
            recs(i).IsValid = True
            recs(i).Status = "Imported"
            recs(i).Message = "Included despite header mismatch (aligned by column position)"
        End If
    Next i

End Sub

' ============================================================
'  CONSOLIDATION PASS
' ============================================================

'--------------------------------------------------------------
' RunConsolidation
'   Re-opens each valid file, block-reads its data sheet and writes
'   the values (with four metadata columns) into the output sheet.
'   Stops cleanly if the worksheet row limit would be exceeded.
'   Parameters:
'     recs ()           - validated records (updated with row counts)
'     includeMismatched - True when mismatched files are aligned by col
'   Returns: nothing (side effects: output sheet populated).
'--------------------------------------------------------------
Private Sub RunConsolidation(ByRef recs() As TFileRecord, _
                             ByVal includeMismatched As Boolean)

    Dim outWs As Worksheet
    Dim firstValid As Long
    Dim headerCols As Long
    Dim nextRow As Long
    Dim i As Long
    Dim totalRows As Long, skipped As Long, processed As Long
    Dim hitRowLimit As Boolean

    firstValid = FirstValidIndex(recs)
    If firstValid = 0 Then Exit Sub

    Set outWs = CreateOutputSheet()
    If outWs Is Nothing Then Exit Sub      ' user declined overwrite path

    headerCols = recs(firstValid).UsedCols
    nextRow = WriteOutputHeader(outWs, recs(firstValid))
    hitRowLimit = False

    For i = LBound(recs) To UBound(recs)
        If recs(i).IsValid Then
            processed = processed + 1
            If Not AppendFileData(outWs, recs(i), headerCols, nextRow) Then
                hitRowLimit = True
                Exit For
            End If
            totalRows = totalRows + recs(i).RowsImported
        Else
            skipped = skipped + 1
            If recs(i).Status <> "Failed" Then recs(i).Status = "Skipped"
        End If
    Next i

    FormatOutputSheet outWs, headerCols

    If hitRowLimit Then
        MsgBox "Row limit reached: output stopped at " & (nextRow - 1) & _
               " rows to avoid exceeding the worksheet maximum." & vbCrLf & _
               "Some files were only partially imported or skipped.", _
               vbExclamation, "Consolidate Substations"
    End If

    MsgBox "Done." & vbCrLf & _
           "Files processed: " & processed & vbCrLf & _
           "Files skipped/failed: " & (UBound(recs) - LBound(recs) + 1 - processed) & vbCrLf & _
           "Total data rows written: " & totalRows, _
           vbInformation, "Consolidate Substations"

End Sub

'--------------------------------------------------------------
' FirstValidIndex
'   Helper: returns the index of the first valid record, or 0 if none.
'   Parameters: recs ().  Returns: 1-based index or 0.
'--------------------------------------------------------------
Private Function FirstValidIndex(ByRef recs() As TFileRecord) As Long
    Dim i As Long
    For i = LBound(recs) To UBound(recs)
        If recs(i).IsValid Then
            FirstValidIndex = i
            Exit Function
        End If
    Next i
    FirstValidIndex = 0
End Function

'--------------------------------------------------------------
' CreateOutputSheet
'   Creates the output sheet in ThisWorkbook. If a sheet with the
'   configured name exists, asks whether to overwrite it or create a
'   new, timestamp-suffixed sheet.
'   Parameters: none.
'   Returns: the output Worksheet, or Nothing if the user cancels.
'--------------------------------------------------------------
Private Function CreateOutputSheet() As Worksheet

    Dim ws As Worksheet
    Dim existing As Worksheet
    Dim ans As VbMsgBoxResult
    Dim newName As String

    On Error Resume Next
    Set existing = ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME)
    On Error GoTo 0

    If existing Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = OUTPUT_SHEET_NAME
        Set CreateOutputSheet = ws
        Exit Function
    End If

    ans = MsgBox("A sheet named '" & OUTPUT_SHEET_NAME & "' already exists." & vbCrLf & _
                 "Yes = overwrite it" & vbCrLf & _
                 "No  = create a new timestamped sheet" & vbCrLf & _
                 "Cancel = abort", _
                 vbYesNoCancel + vbQuestion, "Output sheet exists")

    Select Case ans
        Case vbCancel
            Set CreateOutputSheet = Nothing
        Case vbYes
            existing.Cells.Clear
            Set CreateOutputSheet = existing
        Case vbNo
            newName = OUTPUT_SHEET_NAME & " " & Format$(Now, "yyyymmdd_hhnnss")
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = Left$(newName, EXCEL_MAX_TAB_LEN)
            Set CreateOutputSheet = ws
    End Select

End Function

'--------------------------------------------------------------
' WriteOutputHeader
'   Writes the four metadata headers plus the data headers taken from
'   the first valid file's header row.
'   Parameters:
'     outWs - the output worksheet
'     rec   - the first valid record (its file supplies the headers)
'   Returns: the row number where data writing should begin.
'--------------------------------------------------------------
Private Function WriteOutputHeader(ByVal outWs As Worksheet, _
                                   ByRef rec As TFileRecord) As Long

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim c As Long

    outWs.Cells(1, 1).Value = "Source File"
    outWs.Cells(1, 2).Value = "Source Sheet"
    outWs.Cells(1, 3).Value = "Substation Name"
    outWs.Cells(1, 4).Value = "Voltage (kV)"

    On Error Resume Next
    Set wb = Application.Workbooks.Open(Filename:=rec.FilePath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0

    If Not wb Is Nothing Then
        Set ws = wb.Sheets(DATA_SHEET_INDEX)
        For c = 1 To rec.UsedCols
            outWs.Cells(1, META_COL_COUNT + c).Value = ws.Cells(HEADER_ROW, c).Value
        Next c
        wb.Close SaveChanges:=False
    End If

    WriteOutputHeader = 2

End Function

'--------------------------------------------------------------
' AppendFileData
'   Opens one valid file, block-reads its data range into a Variant
'   array (values only), and writes the non-blank rows to the output
'   sheet with metadata columns. Closes the source without saving on
'   every path.
'   Parameters:
'     outWs      - the output worksheet
'     rec        - the record to import (RowsImported set here)
'     headerCols - the reference header column count
'     nextRow    - the next free output row (advanced by reference)
'   Returns: True when written (or nothing to write); False when the
'            worksheet row limit would be exceeded.
'--------------------------------------------------------------
Private Function AppendFileData(ByVal outWs As Worksheet, _
                                ByRef rec As TFileRecord, _
                                ByVal headerCols As Long, _
                                ByRef nextRow As Long) As Boolean

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim srcVals As Variant
    Dim outBlock() As Variant
    Dim srcRows As Long, srcCols As Long
    Dim r As Long, c As Long
    Dim outR As Long
    Dim writeCols As Long

    AppendFileData = True
    rec.RowsImported = 0

    On Error GoTo Fail
    Set wb = Application.Workbooks.Open(Filename:=rec.FilePath, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Sheets(DATA_SHEET_INDEX)

    srcVals = BlockReadData(ws, rec)
    If IsEmpty(srcVals) Then
        wb.Close SaveChanges:=False
        Exit Function
    End If

    srcRows = UBound(srcVals, 1) - LBound(srcVals, 1) + 1
    srcCols = UBound(srcVals, 2) - LBound(srcVals, 2) + 1
    writeCols = headerCols                      ' align to reference width
    If srcCols < writeCols Then writeCols = srcCols

    ' Row-limit guard before allocating / writing.
    If nextRow + srcRows - 1 > EXCEL_MAX_ROWS Then
        wb.Close SaveChanges:=False
        AppendFileData = False
        Exit Function
    End If

    ReDim outBlock(1 To srcRows, 1 To META_COL_COUNT + writeCols)
    outR = 0
    For r = 1 To srcRows
        If Not RowIsBlank(srcVals, r, srcCols) Then
            outR = outR + 1
            outBlock(outR, 1) = rec.FilePath
            outBlock(outR, 2) = rec.SheetName
            outBlock(outR, 3) = rec.Substation
            If rec.VoltageParsed Then outBlock(outR, 4) = rec.Voltage Else outBlock(outR, 4) = ""
            For c = 1 To writeCols
                outBlock(outR, META_COL_COUNT + c) = srcVals(r, c)
            Next c
        End If
    Next r

    If outR > 0 Then
        outWs.Range(outWs.Cells(nextRow, 1), _
                    outWs.Cells(nextRow + outR - 1, META_COL_COUNT + writeCols)).Value = _
            TrimBlock(outBlock, outR, META_COL_COUNT + writeCols)
        nextRow = nextRow + outR
        rec.RowsImported = outR
    End If

    wb.Close SaveChanges:=False
    Exit Function

Fail:
    rec.Status = "Failed"
    rec.Message = "Import error: " & Err.Description
    rec.RowsImported = 0
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    AppendFileData = True    ' one file's failure must not stop the run

End Function

'--------------------------------------------------------------
' BlockReadData
'   Helper: reads a source sheet's data range (below the header, out
'   to the used column count) into a Variant array in one operation.
'   Always returns a 2-D array (a single cell is normalised to 1x1).
'   Parameters:
'     ws  - the source data worksheet
'     rec - record supplying the row/column extents
'   Returns: a Variant 2-D array, or Empty when there is no data.
'--------------------------------------------------------------
Private Function BlockReadData(ByVal ws As Worksheet, _
                               ByRef rec As TFileRecord) As Variant

    Dim rng As Range
    Dim tmp(1 To 1, 1 To 1) As Variant

    If rec.LastDataRow < rec.FirstDataRow Then
        BlockReadData = Empty
        Exit Function
    End If

    Set rng = ws.Range(ws.Cells(rec.FirstDataRow, 1), _
                       ws.Cells(rec.LastDataRow, rec.UsedCols))

    If rng.Cells.Count = 1 Then
        tmp(1, 1) = rng.Value            ' normalise single cell to a 2-D array
        BlockReadData = tmp
    Else
        BlockReadData = rng.Value
    End If

End Function

'--------------------------------------------------------------
' RowIsBlank
'   Helper: True when every cell in one row of the source array is
'   empty or whitespace-only.
'   Parameters:
'     arr  - the source Variant array
'     r    - the row index within arr
'     cols - number of columns to inspect
'   Returns: Boolean.
'--------------------------------------------------------------
Private Function RowIsBlank(ByRef arr As Variant, _
                            ByVal r As Long, _
                            ByVal cols As Long) As Boolean

    Dim c As Long
    Dim v As Variant

    For c = 1 To cols
        v = arr(r, c)
        If Not IsEmpty(v) Then
            If Len(Trim$(CStr(v))) > 0 Then
                RowIsBlank = False
                Exit Function
            End If
        End If
    Next c

    RowIsBlank = True

End Function

'--------------------------------------------------------------
' TrimBlock
'   Helper: returns a rows x cols slice of a working array so the
'   single block write only covers populated rows.
'   Parameters:
'     arr  - the oversized working array
'     rows - number of populated rows
'     cols - number of columns
'   Returns: a rows x cols Variant array.
'--------------------------------------------------------------
Private Function TrimBlock(ByRef arr() As Variant, _
                           ByVal rows As Long, _
                           ByVal cols As Long) As Variant

    Dim outArr() As Variant
    Dim r As Long, c As Long

    ReDim outArr(1 To rows, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols
            outArr(r, c) = arr(r, c)
        Next c
    Next r

    TrimBlock = outArr

End Function

' ============================================================
'  OUTPUT FORMATTING
' ============================================================

'--------------------------------------------------------------
' FormatOutputSheet
'   Applies the finishing formatting: bold header, frozen top row,
'   AutoFilter across the header, and column autofit.
'   Parameters:
'     outWs      - the output worksheet
'     headerCols - the data-header column count
'   Returns: nothing (formatting side effects only).
'--------------------------------------------------------------
Private Sub FormatOutputSheet(ByVal outWs As Worksheet, _
                              ByVal headerCols As Long)

    Dim lastCol As Long
    Dim hdr As Range

    lastCol = META_COL_COUNT + headerCols
    Set hdr = outWs.Range(outWs.Cells(1, 1), outWs.Cells(1, lastCol))

    hdr.Font.Bold = True

    On Error Resume Next
    outWs.AutoFilterMode = False
    hdr.AutoFilter
    outWs.Columns.AutoFit

    ' Freeze the top row via the sheet's window.
    Application.GoTo outWs.Cells(1, 1), Scroll:=True
    outWs.Activate
    With ActiveWindow
        .SplitRow = 1
        .SplitColumn = 0
        .FreezePanes = True
    End With
    On Error GoTo 0

End Sub

' ============================================================
'  RUN LOG
' ============================================================

'--------------------------------------------------------------
' WriteLogSheet
'   Writes (or replaces) the run-log sheet: one row per selected file
'   capturing path, sheet name, parsed name/voltage, rows imported,
'   status and message.
'   Parameters:
'     recs () - the validated / processed records
'   Returns: nothing (log sheet populated).
'--------------------------------------------------------------
Private Sub WriteLogSheet(ByRef recs() As TFileRecord)

    Dim ws As Worksheet
    Dim i As Long, r As Long
    Dim logArr() As Variant
    Dim n As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
    Else
        ws.Cells.Clear
    End If

    n = UBound(recs) - LBound(recs) + 1
    ReDim logArr(1 To n + 1, 1 To 7)

    logArr(1, 1) = "File Path"
    logArr(1, 2) = "Third Sheet"
    logArr(1, 3) = "Substation Name"
    logArr(1, 4) = "Voltage (kV)"
    logArr(1, 5) = "Rows Imported"
    logArr(1, 6) = "Status"
    logArr(1, 7) = "Message"

    r = 1
    For i = LBound(recs) To UBound(recs)
        r = r + 1
        logArr(r, 1) = recs(i).FilePath
        logArr(r, 2) = recs(i).SheetName
        logArr(r, 3) = recs(i).Substation
        If recs(i).VoltageParsed Then logArr(r, 4) = recs(i).Voltage Else logArr(r, 4) = ""
        logArr(r, 5) = recs(i).RowsImported
        logArr(r, 6) = recs(i).Status
        logArr(r, 7) = BuildLogMessage(recs(i))
    Next i

    ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 7)).Value = logArr
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit

End Sub

'--------------------------------------------------------------
' BuildLogMessage
'   Helper: composes the log message, appending a "Voltage not
'   parsed" note when the parser could not identify a voltage.
'   Parameters: rec - one record.  Returns: the message string.
'--------------------------------------------------------------
Private Function BuildLogMessage(ByRef rec As TFileRecord) As String

    Dim msg As String

    msg = rec.Message
    If rec.IsValid And (Not rec.VoltageParsed) And Len(rec.SheetName) > 0 Then
        If Len(msg) > 0 Then msg = msg & "; "
        msg = msg & "Voltage not parsed"
    End If
    ' Note names truncated at Excel's tab-length limit (voltage may be cut off).
    If Len(rec.SheetName) = EXCEL_MAX_TAB_LEN Then
        If Len(msg) > 0 Then msg = msg & "; "
        msg = msg & "Tab name at 31-char limit (may be truncated)"
    End If

    BuildLogMessage = msg

End Function

' ============================================================
'  SMALL SHARED HELPERS
' ============================================================

'--------------------------------------------------------------
' GetFileName
'   Helper: returns the file name portion of a full path.
'   Parameters: fullPath - a full file path.  Returns: the file name.
'--------------------------------------------------------------
Private Function GetFileName(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, Application.PathSeparator)
    If p = 0 Then p = InStrRev(fullPath, "\")
    If p = 0 Then
        GetFileName = fullPath
    Else
        GetFileName = Mid$(fullPath, p + 1)
    End If
End Function
