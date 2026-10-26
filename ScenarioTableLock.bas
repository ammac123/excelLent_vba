Attribute VB_Name = "ScenarioTableLock"
Option Explicit

' Keeps tblScenarios read-only to the user. Every change to a scenario's metadata goes
' through frmScenarioEditor (the Edit Scenario and New Scenario buttons), which validates
' it first; typing, pasting or deleting rows in the table directly is blocked.
'
' The lock is ordinary Excel sheet protection with only the table's cells locked, so the
' rest of the sheet (picker, buttons, other tables) stays editable. Selecting a cell in
' the table shows an input message pointing at the Edit Scenario button.
'
' Protection stops macros resizing tables on the sheet too, even with UserInterfaceOnly,
' so BeginBatch lifts it and EndBatch puts it back. Writers that run outside a batch
' call UnlockScenarioSheetIfLocked / RelockScenarioSheet themselves.
'
' Setup: run PrepareScenarioTableLock once. Then hook it up in ThisWorkbook, because
' UserInterfaceOnly does not survive closing the file:
'   Private Sub Workbook_Open()
'       WriteModelInfo
'       LockScenarioTable
'   End Sub

Private Const SHEET_PASSWORD As String = ""

Private Const MESSAGE_TITLE As String = "Read-only"
Private Const MESSAGE_TEXT As String = "Use the Edit Scenario button to change a scenario, " & _
                                       "or double-click its row."

' One-time setup: unlocks every cell on the sheet, then locks just the table. Run it
' again after moving the table or adding inputs elsewhere on the sheet.
Public Sub PrepareScenarioTableLock()
    Dim lo As ListObject

    Set lo = RequireTable(TBL_SCENARIOS)
    lo.Parent.Unprotect SHEET_PASSWORD
    lo.Parent.Cells.Locked = False
    LockScenarioTable

    MsgBox TBL_SCENARIOS & " on " & lo.Parent.Name & " is now read-only. Everything else " & _
           "on the sheet can still be edited.", vbInformation
End Sub

' Locks the table and protects its sheet. Safe to call from Workbook_Open: a failure is
' reported to the Immediate window rather than stopping the open.
Public Sub LockScenarioTable()
    Dim lo As ListObject

    On Error GoTo Failed
    Set lo = TableOrNothing(TBL_SCENARIOS)
    If lo Is Nothing Then Exit Sub

    ProtectTableSheet lo
    Exit Sub

Failed:
    Debug.Print "LockScenarioTable: " & Err.description
End Sub

' The table's sheet, unprotected, when it was protected; Nothing when there was nothing
' to lift. Pass the result to RelockScenarioSheet afterwards.
Public Function UnlockScenarioSheetIfLocked() As Worksheet
    Dim lo As ListObject

    Set lo = TableOrNothing(TBL_SCENARIOS)
    If lo Is Nothing Then Exit Function
    If Not lo.Parent.ProtectContents Then Exit Function

    ' A wrong password leaves the sheet locked; the write that follows then fails with
    ' Excel's own error, which says more than anything raised here would.
    On Error Resume Next
    lo.Parent.Unprotect SHEET_PASSWORD
    If Err.Number = 0 Then Set UnlockScenarioSheetIfLocked = lo.Parent
    On Error GoTo 0
End Function

' Puts back the protection UnlockScenarioSheetIfLocked lifted. Nothing does nothing.
' No On Error here on purpose: callers run this from their error handlers and then
' re-raise Err, which an On Error statement in here would clear.
Public Sub RelockScenarioSheet(ws As Worksheet)
    Dim lo As ListObject

    If ws Is Nothing Then Exit Sub

    Set lo = TableOrNothing(TBL_SCENARIOS)
    If lo Is Nothing Then Exit Sub
    ProtectTableSheet lo
End Sub

' Rows added since the last lock pick up the lock and the message here too. Unprotects
' first: on open the sheet is still protected from last time, without UserInterfaceOnly.
Private Sub ProtectTableSheet(lo As ListObject)
    lo.Parent.Unprotect SHEET_PASSWORD
    lo.Range.Locked = True

    With lo.Range.Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .InputTitle = MESSAGE_TITLE
        .InputMessage = MESSAGE_TEXT
        .ShowInput = True
    End With

    lo.Parent.Protect Password:=SHEET_PASSWORD, DrawingObjects:=False, Contents:=True, _
                      UserInterfaceOnly:=True, AllowFormattingColumns:=True, _
                      AllowFormattingRows:=True, AllowFiltering:=True
End Sub
