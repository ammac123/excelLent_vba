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

Private Const CHECKS_TABLE As String = "tblModelChecks"

Private Const COL_KEY As String = "Key"
Private Const COL_LAST_UPDATED As String = "Last_Updated"
Private Const COL_LAST_REFRESHED As String = "Last_Refreshed"
Private Const COL_DATA_UP_TO As String = "Data_Up_To"
Private Const COL_REFRESHED_BY As String = "Refreshed_By"

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

