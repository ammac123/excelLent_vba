Attribute VB_Name = "ScenarioTokenLibrary"
Option Explicit

' Workbook-level library of reusable tokens that the token builder UI reads and
' writes. One row per saved token on a hidden sheet (98_SCN_TokenLibrary), persisted
' with the workbook. This module owns the sheet and the table; nothing else should
' write to tblSavedTokens directly.

Public Type SavedToken
    tokenId As String
    label As String
    kind As String        ' Metric / Current / Delta / DeltaPct / Label / Scenario / ChartNamed / ChartInline
    token As String
    createdOn As Date
    modifiedOn As Date
End Type

Private Const COL_ID As String = "TokenId"
Private Const COL_LABEL As String = "Label"
Private Const COL_KIND As String = "Kind"
Private Const COL_TOKEN As String = "Token"
Private Const COL_CREATED As String = "CreatedOn"
Private Const COL_MODIFIED As String = "ModifiedOn"

Private Const ID_PREFIX As String = "TOK"
Private Const ID_DIGITS As Long = 4

'=======================================================================================
' Bootstrap
'=======================================================================================

' Idempotent. Creates the hidden sheet and empty table if either is missing. Safe to
' call before any other API; every public entry does.
Public Sub EnsureLibrarySheet()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim sheetName As String, tableName As String

    sheetName = TokenLibrarySheetName()
    tableName = SavedTokensTableName()

    Set ws = SheetOrNothing(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
        ws.Visible = xlSheetHidden
    End If

    Set lo = TableOrNothing(tableName)
    If Not lo Is Nothing Then Exit Sub

    ws.Range("A1").value = COL_ID
    ws.Range("B1").value = COL_LABEL
    ws.Range("C1").value = COL_KIND
    ws.Range("D1").value = COL_TOKEN
    ws.Range("E1").value = COL_CREATED
    ws.Range("F1").value = COL_MODIFIED

    ' Pass a header-only range so the table starts at zero data rows.
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F1"), , xlYes)
    lo.Name = tableName
    lo.ListColumns(COL_CREATED).Range.numberFormat = "yyyy-mm-dd hh:mm"
    lo.ListColumns(COL_MODIFIED).Range.numberFormat = "yyyy-mm-dd hh:mm"
End Sub

'=======================================================================================
' Read
'=======================================================================================

' All library rows in table order. The caller must handle an unassigned result
' (returned when the table is empty) via LibraryRowCount.
Public Function LibraryRows() As SavedToken()
    Dim lo As ListObject
    Dim body As Variant
    Dim rows() As SavedToken
    Dim cId As Long, cLabel As Long, cKind As Long, cToken As Long
    Dim cCreated As Long, cModified As Long
    Dim r As Long, n As Long

    EnsureLibrarySheet
    Set lo = RequireTable(SavedTokensTableName())
    n = TableRowCount(lo)
    If n = 0 Then Exit Function

    body = TableBody(lo)
    cId = RequireColumnIndex(lo, COL_ID)
    cLabel = RequireColumnIndex(lo, COL_LABEL)
    cKind = RequireColumnIndex(lo, COL_KIND)
    cToken = RequireColumnIndex(lo, COL_TOKEN)
    cCreated = RequireColumnIndex(lo, COL_CREATED)
    cModified = RequireColumnIndex(lo, COL_MODIFIED)

    ReDim rows(1 To n)
    For r = 1 To n
        rows(r).tokenId = CStr(body(r, cId))
        rows(r).label = CStr(body(r, cLabel))
        rows(r).kind = CStr(body(r, cKind))
        rows(r).token = CStr(body(r, cToken))
        If IsDate(body(r, cCreated)) Then rows(r).createdOn = CDate(body(r, cCreated))
        If IsDate(body(r, cModified)) Then rows(r).modifiedOn = CDate(body(r, cModified))
    Next r

    LibraryRows = rows
End Function

Public Function LibraryRowCount() As Long
    EnsureLibrarySheet
    LibraryRowCount = TableRowCount(RequireTable(SavedTokensTableName()))
End Function

Public Function LibraryRowById(ByVal tokenId As String) As SavedToken
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim created As Variant, modified As Variant

    EnsureLibrarySheet
    Set lo = RequireTable(SavedTokensTableName())
    rowIndex = FindRowIndex(lo, COL_ID, tokenId)
    If rowIndex = 0 Then
        Err.Raise vbObjectError + 760, "ScenarioTokenLibrary", _
                  "No saved token found with id '" & tokenId & "'."
    End If

    LibraryRowById.tokenId = tokenId
    LibraryRowById.label = CStr(CellInRow(lo, rowIndex, COL_LABEL).value)
    LibraryRowById.kind = CStr(CellInRow(lo, rowIndex, COL_KIND).value)
    LibraryRowById.token = CStr(CellInRow(lo, rowIndex, COL_TOKEN).value)
    created = CellInRow(lo, rowIndex, COL_CREATED).value
    modified = CellInRow(lo, rowIndex, COL_MODIFIED).value
    If IsDate(created) Then LibraryRowById.createdOn = CDate(created)
    If IsDate(modified) Then LibraryRowById.modifiedOn = CDate(modified)
End Function

'=======================================================================================
' Write
'=======================================================================================

' Updates an existing row (matched by tokenId when non-empty) or appends a new row.
' Returns the resulting row's tokenId so the caller can rebind the form's editing
' state to the newly-created row.
Public Function UpsertToken(ByVal tokenId As String, label As String, kind As String, _
                            token As String) As String
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim newRow As ListRow
    Dim prevCreated As Variant

    EnsureLibrarySheet
    Set lo = RequireTable(SavedTokensTableName())

    If Len(tokenId) > 0 Then
        rowIndex = FindRowIndex(lo, COL_ID, tokenId)
        If rowIndex > 0 Then
            prevCreated = CellInRow(lo, rowIndex, COL_CREATED).value
            WriteRow lo, rowIndex, tokenId, label, kind, token, prevCreated, Now
            UpsertToken = tokenId
            Exit Function
        End If
    End If

    tokenId = NewTokenId()
    Set newRow = lo.ListRows.Add
    WriteRow lo, newRow.Index, tokenId, label, kind, token, Now, Now
    UpsertToken = tokenId
End Function

Public Sub DeleteToken(ByVal tokenId As String)
    Dim lo As ListObject
    Dim rowIndex As Long

    If Len(tokenId) = 0 Then Exit Sub

    EnsureLibrarySheet
    Set lo = RequireTable(SavedTokensTableName())
    rowIndex = FindRowIndex(lo, COL_ID, tokenId)
    If rowIndex = 0 Then Exit Sub

    lo.ListRows(rowIndex).Delete
End Sub

' Returns a token id that is not already in the table. Scans existing ids to find the
' highest numeric suffix and increments, so deletions do not reuse ids.
Public Function NewTokenId() As String
    Dim lo As ListObject
    Dim body As Variant
    Dim r As Long, cId As Long
    Dim id As String, suffix As Long, highest As Long

    EnsureLibrarySheet
    Set lo = RequireTable(SavedTokensTableName())
    cId = RequireColumnIndex(lo, COL_ID)

    If TableRowCount(lo) > 0 Then
        body = TableBody(lo)
        For r = 1 To UBound(body, 1)
            id = CStr(body(r, cId))
            If Len(id) > Len(ID_PREFIX) Then
                suffix = Val(Mid$(id, Len(ID_PREFIX) + 1))
                If suffix > highest Then highest = suffix
            End If
        Next r
    End If

    NewTokenId = ID_PREFIX & Format$(highest + 1, String$(ID_DIGITS, "0"))
End Function

Private Sub WriteRow(lo As ListObject, ByVal rowIndex As Long, _
                     tokenId As String, label As String, kind As String, _
                     token As String, createdOn As Variant, modifiedOn As Variant)
    CellInRow(lo, rowIndex, COL_ID).value = tokenId
    CellInRow(lo, rowIndex, COL_LABEL).value = label
    CellInRow(lo, rowIndex, COL_KIND).value = kind
    CellInRow(lo, rowIndex, COL_TOKEN).value = token
    CellInRow(lo, rowIndex, COL_CREATED).value = createdOn
    CellInRow(lo, rowIndex, COL_MODIFIED).value = modifiedOn
End Sub
