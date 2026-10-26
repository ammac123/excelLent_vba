Attribute VB_Name = "CopyMarginalFTP"
Option Explicit

' Input blocks: labels in column D, one data column per month starting in column E.
Private Const INPUT_LABEL_COL As String = "D"
Private Const INPUT_FIRST_DATA_COL As String = "E"
Private Const BLOCK_ROWS As Long = 5
Private Const RESI_FIRST_ROW As Long = 5
Private Const OA_FIRST_ROW As Long = 12
Private Const TABLE_NAME As String = "tblMarginalFTP"

Private Type InputBlock
    firstRow As Long
    rateHeader As String
    lpHeader As String
    clcHeader As String
    sumHeader As String
End Type

Public Sub PostMarginalFTP()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim blocks(1 To 2) As InputBlock
    Dim i As Long
    Dim added As Long, updated As Long

    Set ws = ActiveSheet

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_NAME)
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Table '" & TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbCritical
        Exit Sub
    End If

    blocks(1) = MakeBlock(RESI_FIRST_ROW, "BBSW 1mth", "RESI LP", "RESI CLC", "RESI LP+CLC")
    blocks(2) = MakeBlock(OA_FIRST_ROW, "BBSW 1mth", "OE LP", "OE CLC", "OE LP+CLC")

    ' Validate every block before writing anything.
    For i = LBound(blocks) To UBound(blocks)
        If Not BlockIsValid(ws, lo, blocks(i)) Then Exit Sub
    Next i

    Application.ScreenUpdating = False

    For i = LBound(blocks) To UBound(blocks)
        PostBlock ws, lo, blocks(i), added, updated
    Next i

    Application.ScreenUpdating = True

    If added + updated = 0 Then
        MsgBox "There is no data to post.", vbInformation
    Else
        UpdateModelCheckForKey "Marginal FTP"
        MsgBox "Posted " & (added + updated) & " month(s) to " & TABLE_NAME & _
               " (" & added & " added, " & updated & " overwritten).", vbInformation
    End If
End Sub

Private Function MakeBlock(firstRow As Long, rateHeader As String, lpHeader As String, _
                           clcHeader As String, sumHeader As String) As InputBlock
    With MakeBlock
        .firstRow = firstRow
        .rateHeader = rateHeader
        .lpHeader = lpHeader
        .clcHeader = clcHeader
        .sumHeader = sumHeader
    End With
End Function

Private Function BlockIsValid(ws As Worksheet, lo As ListObject, blk As InputBlock) As Boolean
    Dim labelCol As Long, firstDataCol As Long, lastDataCol As Long
    Dim col As Long, lastRow As Long

    labelCol = ws.Columns(INPUT_LABEL_COL).column
    firstDataCol = ws.Columns(INPUT_FIRST_DATA_COL).column
    lastRow = blk.firstRow + BLOCK_ROWS - 1

    If Not InputLabelsValid(ws, labelCol, blk.firstRow) Then
        MsgBox "The labels in " & INPUT_LABEL_COL & blk.firstRow & ":" & INPUT_LABEL_COL & lastRow & _
               " must be Month, Interest Rate, Liquidity Premium, CLC, Marginal FTP.", vbCritical
        Exit Function
    End If

    If HeaderIndex(lo, "Month") = 0 Or HeaderIndex(lo, blk.rateHeader) = 0 _
        Or HeaderIndex(lo, blk.lpHeader) = 0 Or HeaderIndex(lo, blk.clcHeader) = 0 Then
        MsgBox "Table '" & TABLE_NAME & "' is missing one or more of the columns Month, " & _
               blk.rateHeader & ", " & blk.lpHeader & ", " & blk.clcHeader & ".", vbCritical
        Exit Function
    End If

    lastDataCol = LastDataColumn(ws, blk.firstRow, firstDataCol)

    For col = firstDataCol To lastDataCol
        If Not IsDate(ws.Cells(blk.firstRow, col).value) Then
            MsgBox "The month in cell " & ws.Cells(blk.firstRow, col).Address(False, False) & _
                   " is not a valid date.", vbCritical
            Exit Function
        End If
        If Not IsNumeric(ws.Cells(blk.firstRow + 1, col).value) _
            Or Not IsNumeric(ws.Cells(blk.firstRow + 2, col).value) _
            Or Not IsNumeric(ws.Cells(blk.firstRow + 3, col).value) Then
            MsgBox "Interest Rate, Liquidity Premium and CLC must all be numeric for " & _
                   Format$(ws.Cells(blk.firstRow, col).value, "mmm yyyy") & " in row " & _
                   blk.firstRow & ".", vbCritical
            Exit Function
        End If
    Next col

    BlockIsValid = True
End Function

Private Sub PostBlock(ws As Worksheet, lo As ListObject, blk As InputBlock, _
                      ByRef added As Long, ByRef updated As Long)
    Dim firstDataCol As Long, lastDataCol As Long, col As Long
    Dim colMonth As Long, colRate As Long, colLP As Long, colCLC As Long, colSum As Long
    Dim monthValue As Date
    Dim targetRow As ListRow

    firstDataCol = ws.Columns(INPUT_FIRST_DATA_COL).column
    lastDataCol = LastDataColumn(ws, blk.firstRow, firstDataCol)
    If lastDataCol < firstDataCol Then Exit Sub

    colMonth = HeaderIndex(lo, "Month")
    colRate = HeaderIndex(lo, blk.rateHeader)
    colLP = HeaderIndex(lo, blk.lpHeader)
    colCLC = HeaderIndex(lo, blk.clcHeader)
    colSum = HeaderIndex(lo, blk.sumHeader)

    For col = firstDataCol To lastDataCol
        monthValue = NormaliseMonth(CDate(ws.Cells(blk.firstRow, col).value))

        Set targetRow = FindMonthRow(lo, colMonth, monthValue)
        If targetRow Is Nothing Then
            Set targetRow = lo.ListRows.Add
            added = added + 1
        Else
            updated = updated + 1
        End If

        targetRow.Range.Cells(1, colMonth).value = monthValue
        targetRow.Range.Cells(1, colRate).value = ws.Cells(blk.firstRow + 1, col).value
        targetRow.Range.Cells(1, colLP).value = ws.Cells(blk.firstRow + 2, col).value
        targetRow.Range.Cells(1, colCLC).value = ws.Cells(blk.firstRow + 3, col).value
        If colSum > 0 Then
            targetRow.Range.Cells(1, colSum).Formula = _
                "=[@[" & lo.ListColumns(colLP).Name & "]]+[@[" & lo.ListColumns(colCLC).Name & "]]"
        End If
    Next col

    ws.Range(ws.Cells(blk.firstRow, firstDataCol), _
             ws.Cells(blk.firstRow + BLOCK_ROWS - 1, lastDataCol)).ClearContents
End Sub

' Scans right from the first data column until an empty month cell is reached.
Private Function LastDataColumn(ws As Worksheet, monthRow As Long, firstDataCol As Long) As Long
    Dim col As Long

    col = firstDataCol - 1
    Do While Len(Trim$(CStr(ws.Cells(monthRow, col + 1).value))) > 0
        col = col + 1
    Loop
    LastDataColumn = col
End Function

Private Function InputLabelsValid(ws As Worksheet, labelCol As Long, firstRow As Long) As Boolean
    Dim expected As Variant
    Dim i As Long

    expected = Array("Month", "Interest Rate", "Liquidity Premium", "CLC", "Marginal FTP")
    For i = 0 To UBound(expected)
        If StrComp(Trim$(CStr(ws.Cells(firstRow + i, labelCol).value)), _
                   expected(i), vbTextCompare) <> 0 Then
            Exit Function
        End If
    Next i
    InputLabelsValid = True
End Function

Private Function HeaderIndex(lo As ListObject, header As String) As Long
    Dim i As Long

    If Len(header) = 0 Then Exit Function
    For i = 1 To lo.ListColumns.count
        If StrComp(Trim$(lo.ListColumns(i).Name), header, vbTextCompare) = 0 Then
            HeaderIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function NormaliseMonth(value As Date) As Date
    NormaliseMonth = DateSerial(Year(value), Month(value), 1)
End Function

Private Function FindMonthRow(lo As ListObject, monthCol As Long, monthValue As Date) As ListRow
    Dim lr As ListRow
    Dim cellValue As Variant

    For Each lr In lo.ListRows
        cellValue = lr.Range.Cells(1, monthCol).value
        If IsDate(cellValue) Then
            If NormaliseMonth(CDate(cellValue)) = monthValue Then
                Set FindMonthRow = lr
                Exit Function
            End If
        End If
    Next lr
End Function






