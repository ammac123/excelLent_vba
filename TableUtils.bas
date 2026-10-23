Attribute VB_Name = "TableUtils"
Option Explicit

' Shared worksheet and ListObject helpers for the scenario modules.
'
' Existing modules keep their own Private copies of FindTable, ColIndex and friends.
' That is deliberate: VBA resolves a name to the module-local Private first, so adding
' these Public versions cannot change how any existing module behaves.

Public Function SheetOrNothing(sheetName As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function TableOrNothing(tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                If MustBeUnfiltered(lo.Name) Then ClearTableFilter lo
                Set TableOrNothing = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' Shows every row again if the table's AutoFilter is hiding any. Leaves the filter
' buttons in place so the user can filter again afterwards.
Public Sub ClearTableFilter(lo As ListObject)
    If lo.AutoFilter Is Nothing Then Exit Sub
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
End Sub

Private Function MustBeUnfiltered(tableName As String) As Boolean
    Dim names As Variant
    Dim i As Long

    names = UnfilteredTableNames()
    For i = LBound(names) To UBound(names)
        If StrComp(CStr(names(i)), tableName, vbTextCompare) = 0 Then
            MustBeUnfiltered = True
            Exit Function
        End If
    Next i
End Function

Public Function RequireTable(tableName As String) As ListObject
    Set RequireTable = TableOrNothing(tableName)
    If RequireTable Is Nothing Then
        Err.Raise vbObjectError + 600, "TableUtils", _
                  "Table '" & tableName & "' was not found in this workbook."
    End If
End Function

Public Function NamedRangeOrNothing(nameRef As String) As Range
    On Error Resume Next
    Set NamedRangeOrNothing = ThisWorkbook.names(nameRef).RefersToRange
    On Error GoTo 0
End Function

Public Function RequireNamedRange(nameRef As String) As Range
    Set RequireNamedRange = NamedRangeOrNothing(nameRef)
    If RequireNamedRange Is Nothing Then
        Err.Raise vbObjectError + 601, "TableUtils", _
                  "Named range '" & nameRef & "' was not found in this workbook."
    End If
End Function

' Zero when the table has no column with that header.
Public Function ColumnIndexOrZero(lo As ListObject, header As String) As Long
    Dim lc As ListColumn

    On Error Resume Next
    Set lc = lo.ListColumns(header)
    On Error GoTo 0

    If Not lc Is Nothing Then ColumnIndexOrZero = lc.Index
End Function

Public Function RequireColumnIndex(lo As ListObject, header As String) As Long
    RequireColumnIndex = ColumnIndexOrZero(lo, header)
    If RequireColumnIndex = 0 Then
        Err.Raise vbObjectError + 602, "TableUtils", _
                  "Column '" & header & "' was not found in table '" & lo.Name & "'."
    End If
End Function

Public Function TableRowCount(lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    TableRowCount = lo.DataBodyRange.Rows.count
End Function

' The whole body as a 1-based two dimensional array, or Empty when the table has no rows.
Public Function TableBody(lo As ListObject) As Variant
    Dim one(1 To 1, 1 To 1) As Variant

    If TableRowCount(lo) = 0 Then Exit Function

    ' A single cell comes back as a scalar, so force the array shape.
    If lo.DataBodyRange.Cells.count = 1 Then
        one(1, 1) = lo.DataBodyRange.value
        TableBody = one
    Else
        TableBody = lo.DataBodyRange.value
    End If
End Function

' Grows or shrinks the table to exactly rowCount data rows. Growing absorbs the cells
' immediately below the table, so a table resized this way must own its region outright.
Public Sub ResizeTableBody(lo As ListObject, ByVal rowCount As Long)
    Dim current As Long

    If rowCount < 0 Then rowCount = 0
    current = TableRowCount(lo)
    If rowCount = current Then Exit Sub

    If rowCount = 0 Then
        lo.DataBodyRange.Delete
        Exit Sub
    End If

    If rowCount < current Then
        ' Delete the surplus outright rather than resizing over it, so nothing is left
        ' behind in the cells the table stops owning.
        lo.ListRows(rowCount + 1).Range.Resize(current - rowCount).Delete xlShiftUp
    Else
        lo.Resize lo.HeaderRowRange.Resize(rowCount + 1)
    End If
End Sub

' True for each column whose first data cell holds a formula, so callers can leave
' calculated columns alone. All False when the table is empty.
Public Function FormulaColumnFlags(lo As ListObject) As Variant
    Dim flags() As Boolean
    Dim i As Long

    ReDim flags(1 To lo.ListColumns.count)

    If TableRowCount(lo) > 0 Then
        For i = 1 To lo.ListColumns.count
            flags(i) = lo.ListColumns(i).DataBodyRange.Cells(1).HasFormula
        Next i
    End If

    FormulaColumnFlags = flags
End Function

' Row index of the first body row whose column equals value, or zero.
Public Function FindRowIndex(lo As ListObject, header As String, value As String) As Long
    Dim body As Variant
    Dim col As Long
    Dim r As Long

    body = TableBody(lo)
    If IsEmpty(body) Then Exit Function

    col = RequireColumnIndex(lo, header)
    For r = 1 To UBound(body, 1)
        If StrComp(CStr(body(r, col)), value, vbTextCompare) = 0 Then
            FindRowIndex = r
            Exit Function
        End If
    Next r
End Function

Public Function CellInRow(lo As ListObject, ByVal rowIndex As Long, header As String) As Range
    Set CellInRow = lo.DataBodyRange.Cells(rowIndex, RequireColumnIndex(lo, header))
End Function


