Attribute VB_Name = "ModelChecks"
Option Explicit

' Governance snapshot of each imported input file, written to tblModelChecks on
' sheet 01_model_checks. UpdateModelCheckForKey is called at the end of each
' import and only touches the row for that key: a key that was not re-imported
' keeps the values recorded at its last import.
'
' Columns written (table must also hold a Description column, user-maintained):
'   Key              input-file key, matches the key used elsewhere in AppState
'   Last_Updated     the source file's modification time on disk at import time
'   Last_Refreshed   when the import ran (same value as SourceLastImportedOn)
'   Data_Up_To       the newest date in the imported local table
'   Refreshed_By       Application.UserName at the time of the import
'
' ConfirmModelCheck records a manual sign-off for checks not tied to an import
' (e.g. "rates reviewed"). It stamps Last_Updated, Last_Refreshed and
' Refreshed_By with now; use keys that don't collide with the import keys in
' LookupConfig.
'
' Review columns (all optional; rows are skipped if Review_Frequency is absent):
'   Review_Frequency user-maintained, e.g. "Annual", "Quarterly", "6 months", "1y"
'   Next_Review_Due  Last_Updated + Review_Frequency
'   Review_Status    OK / Review due / No valid Last_Updated / Invalid frequency
' CheckModelReviewsDue is called from Workbook_Open and only warns: it never
' blocks the workbook opening or any scenario from running.

Private Const CHECKS_TABLE As String = "tblModelChecks"

Private Const COL_KEY As String = "Key"
Private Const COL_LAST_UPDATED As String = "Last_Updated"
Private Const COL_LAST_REFRESHED As String = "Last_Refreshed"
Private Const COL_DATA_UP_TO As String = "Data_Up_To"
Private Const COL_REFRESHED_BY As String = "Refreshed_By"
Private Const COL_REVIEW_FREQUENCY As String = "Review_Frequency"
Private Const COL_NEXT_REVIEW_DUE As String = "Next_Review_Due"
Private Const COL_REVIEW_STATUS As String = "Review_Status"

Private Const STATUS_OK As String = "OK"
Private Const STATUS_DUE As String = "Review due"
Private Const STATUS_NO_DATE As String = "No valid Last_Updated"
Private Const STATUS_BAD_FREQ As String = "Invalid frequency"

Public Sub UpdateModelCheckForKey(key As String)
    Dim localTable As String, dateColumn As String
    Dim lo As ListObject
    Dim rowIdx As Long
    Dim path As String

    If Not LookupConfig(key, localTable, dateColumn) Then Exit Sub

    Set lo = TableOrNothing(CHECKS_TABLE)
    If lo Is Nothing Then Exit Sub

    rowIdx = FindOrAppendRow(lo, key)

    path = SourceFilePath(key)
    If Len(path) > 0 And Len(Dir$(path)) > 0 Then
        CellInRow(lo, rowIdx, COL_LAST_UPDATED).value = FileDateTime(path)
    End If

    CellInRow(lo, rowIdx, COL_LAST_REFRESHED).value = SourceLastImportedOn(key)
    CellInRow(lo, rowIdx, COL_DATA_UP_TO).value = MaxDateInTable(localTable, dateColumn)
    CellInRow(lo, rowIdx, COL_REFRESHED_BY).value = Application.UserName
    RefreshReviewForRow lo, rowIdx
End Sub

' Workbook_Open entry point: refreshes every row's review columns and warns
' about any that need revalidation. Never raises: a broken or missing table
' must not stop the workbook opening or scenarios running.
Public Sub CheckModelReviewsDue()
    Dim lo As ListObject
    Dim r As Long
    Dim issue As String
    Dim issues As String

    On Error GoTo Done

    Set lo = TableOrNothing(CHECKS_TABLE)
    If lo Is Nothing Then Exit Sub

    For r = 1 To TableRowCount(lo)
        issue = RefreshReviewForRow(lo, r)
        If Len(issue) > 0 Then issues = issues & vbNewLine & "  - " & issue
    Next r

    If Len(issues) > 0 Then
        MsgBox "The following model inputs or assumptions require revalidation:" & vbNewLine & _
               issues & vbNewLine & vbNewLine & _
               "The model can still be run. Re-import the input or confirm the check on " & _
               lo.Parent.Name & " to clear this warning.", vbExclamation, "Model checks"
    End If
Done:
End Sub

' Asks the user to confirm, then stamps the row for key with the current time
' and user, which also restarts its review period. Returns False if the user
' declines or the table is missing, so a calling process can stop:
'   If Not ConfirmModelCheck("X") Then Exit Sub
Public Function ConfirmModelCheck(key As String, _
                                  Optional prompt As String = "") As Boolean
    Dim lo As ListObject
    Dim rowIdx As Long

    If Len(prompt) = 0 Then prompt = "Confirm '" & key & "' has been checked?"
    If MsgBox(prompt, vbYesNo + vbQuestion, "Model check") <> vbYes Then Exit Function

    Set lo = TableOrNothing(CHECKS_TABLE)
    If lo Is Nothing Then
        MsgBox "Table " & CHECKS_TABLE & " not found; check not recorded.", vbExclamation
        Exit Function
    End If

    rowIdx = FindOrAppendRow(lo, key)
    CellInRow(lo, rowIdx, COL_LAST_UPDATED).value = Now
    CellInRow(lo, rowIdx, COL_LAST_REFRESHED).value = Now
    CellInRow(lo, rowIdx, COL_REFRESHED_BY).value = Application.UserName
    RefreshReviewForRow lo, rowIdx
    ConfirmModelCheck = True
End Function

' Button entry point: confirms the tblModelChecks row the active cell is in.
Public Sub ConfirmSelectedModelCheck()
    Dim lo As ListObject
    Dim keyCol As Long
    Dim rowIdx As Long
    Dim key As String

    Set lo = TableOrNothing(CHECKS_TABLE)
    If lo Is Nothing Then Exit Sub

    If lo.DataBodyRange Is Nothing Then Exit Sub
    If ActiveCell.Worksheet.Name <> lo.Parent.Name Then
        MsgBox "Select a row in " & CHECKS_TABLE & " first.", vbInformation
        Exit Sub
    End If
    If Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then
        MsgBox "Select a row in " & CHECKS_TABLE & " first.", vbInformation
        Exit Sub
    End If

    keyCol = ColumnIndexOrZero(lo, COL_KEY)
    If keyCol = 0 Then Exit Sub

    rowIdx = ActiveCell.Row - lo.DataBodyRange.Row + 1
    key = CStr(lo.DataBodyRange.Cells(rowIdx, keyCol).value)
    If Len(key) = 0 Then Exit Sub

    ConfirmModelCheck key
End Sub

' Recomputes Next_Review_Due and Review_Status for one row from Last_Updated and
' Review_Frequency. Returns a one-line description when the row needs
' revalidation, or "" when it is in date or has no frequency set.
Private Function RefreshReviewForRow(lo As ListObject, rowIdx As Long) As String
    Dim key As String
    Dim freqText As String
    Dim lastUpdated As Variant
    Dim interval As String
    Dim count As Long
    Dim nextDue As Variant
    Dim status As String
    Dim issue As String

    If ColumnIndexOrZero(lo, COL_REVIEW_FREQUENCY) = 0 Then Exit Function

    key = CStr(CellValue(lo, rowIdx, COL_KEY))
    freqText = Trim$(CStr(CellValue(lo, rowIdx, COL_REVIEW_FREQUENCY)))
    lastUpdated = CellValue(lo, rowIdx, COL_LAST_UPDATED)

    If Len(freqText) = 0 Then
        ' No review required for this row.
    ElseIf Not ParseFrequency(freqText, interval, count) Then
        status = STATUS_BAD_FREQ
        issue = key & ": review frequency '" & freqText & "' not recognised"
    ElseIf Not IsDate(lastUpdated) Then
        status = STATUS_NO_DATE
        issue = key & ": Last_Updated is missing or not a date"
    Else
        nextDue = DateAdd(interval, count, DateValue(CDate(lastUpdated)))
        If Date >= nextDue Then
            status = STATUS_DUE
            issue = key & ": review was due " & Format$(nextDue, "dd-mmm-yyyy")
        Else
            status = STATUS_OK
        End If
    End If

    SetCellIfColumn lo, rowIdx, COL_NEXT_REVIEW_DUE, nextDue
    SetCellIfColumn lo, rowIdx, COL_REVIEW_STATUS, status
    RefreshReviewForRow = issue
End Function

' Accepts named frequencies (Monthly, Quarterly, Annual, ...) or a count and
' unit ("6 months", "1y", "every 2 weeks", "once every year"), returning the
' DateAdd interval and count.
Private Function ParseFrequency(ByVal text As String, ByRef interval As String, _
                                ByRef count As Long) As Boolean
    Dim i As Long
    Dim unit As String

    text = LCase$(Trim$(text))
    text = Replace(text, "once ", "")
    text = Trim$(Replace(text, "every ", ""))

    Select Case text
        Case "daily": interval = "d": count = 1
        Case "weekly": interval = "ww": count = 1
        Case "fortnightly": interval = "ww": count = 2
        Case "monthly": interval = "m": count = 1
        Case "quarterly": interval = "q": count = 1
        Case "half-yearly", "half yearly", "semi-annual", "semi-annually"
            interval = "m": count = 6
        Case "annual", "annually", "yearly"
            interval = "yyyy": count = 1
        Case Else
            i = 1
            Do While i <= Len(text) And Mid$(text, i, 1) Like "#"
                i = i + 1
            Loop
            If i > 5 Then Exit Function
            count = 1
            If i > 1 Then count = CLng(Left$(text, i - 1))
            If count < 1 Then Exit Function

            unit = Trim$(Mid$(text, i))
            Select Case unit
                Case "d", "day", "days": interval = "d"
                Case "w", "wk", "wks", "week", "weeks": interval = "ww"
                Case "m", "mo", "mth", "mths", "month", "months": interval = "m"
                Case "q", "qtr", "quarter", "quarters": interval = "q"
                Case "y", "yr", "yrs", "year", "years": interval = "yyyy"
                Case Else: Exit Function
            End Select
    End Select
    ParseFrequency = True
End Function

' Empty when the column is absent or the cell holds an error value.
Private Function CellValue(lo As ListObject, rowIdx As Long, header As String) As Variant
    Dim col As Long

    col = ColumnIndexOrZero(lo, header)
    If col = 0 Then Exit Function

    CellValue = lo.DataBodyRange.Cells(rowIdx, col).value
    If IsError(CellValue) Then CellValue = Empty
End Function

' Writes only when the column exists and the value has changed, so optional
' columns can be left out and opening the workbook doesn't mark it dirty.
Private Sub SetCellIfColumn(lo As ListObject, rowIdx As Long, header As String, value As Variant)
    Dim col As Long
    Dim cell As Range

    col = ColumnIndexOrZero(lo, header)
    If col = 0 Then Exit Sub

    Set cell = lo.DataBodyRange.Cells(rowIdx, col)
    If Not IsError(cell.value) Then
        If CStr(cell.value) = CStr(value) Then Exit Sub
    End If
    cell.value = value
End Sub

' Keys without a config entry are ignored, so a new import type that forgets to
' register here is non-fatal rather than breaking the import.
Private Function LookupConfig(key As String, ByRef localTable As String, _
                              ByRef dateColumn As String) As Boolean
    Select Case key
        Case "Applications"
            localTable = "tblApplications"
            dateColumn = "Application_Week_Ending"
        Case "Offsets"
            localTable = "tblOffsets"
            dateColumn = "Latest_Record_Date"
        Case "Canstar"
            localTable = "tblCanstar"
            dateColumn = "Effective Date"
        Case "RFC"
            localTable = "tblRFC"
            dateColumn = "Month"
        Case Else
            Exit Function
    End Select
    LookupConfig = True
End Function

Private Function FindOrAppendRow(lo As ListObject, key As String) As Long
    Dim idx As Long

    idx = FindRowIndex(lo, COL_KEY, key)
    If idx > 0 Then
        FindOrAppendRow = idx
        Exit Function
    End If

    lo.ListRows.Add
    idx = TableRowCount(lo)
    CellInRow(lo, idx, COL_KEY).value = key
    FindOrAppendRow = idx
End Function

Private Function MaxDateInTable(tableName As String, dateColumn As String) As Variant
    Dim lo As ListObject
    Dim body As Variant
    Dim col As Long
    Dim r As Long
    Dim best As Date, cur As Date
    Dim found As Boolean

    Set lo = TableOrNothing(tableName)
    If lo Is Nothing Then Exit Function

    col = ColumnIndexOrZero(lo, dateColumn)
    If col = 0 Then Exit Function

    body = TableBody(lo)
    If IsEmpty(body) Then Exit Function

    For r = 1 To UBound(body, 1)
        If IsDate(body(r, col)) Then
            cur = CDate(body(r, col))
            If Not found Or cur > best Then
                best = cur
                found = True
            End If
        End If
    Next r

    If found Then MaxDateInTable = best
End Function

