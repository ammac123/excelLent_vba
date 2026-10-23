Attribute VB_Name = "SheetRecalc"
Option Explicit

' Generic engine for recalculating a single sheet on demand and recording when it ran.
' One instance of this module serves every sheet; each sheet supplies its own RecalcSpec.
'
' Calculation is left on manual afterwards so ordinary edits do not recalculate the sheet.

Public Type RecalcSpec
    key As String              ' slot in the AppState store, e.g. "CurrentStateApps"
    targetSheet As String
    statusCell As String       ' optional, leave empty to skip the status display
    label As String            ' optional, used in messages instead of the sheet name
End Type

Public Function NewRecalcSpec(key As String, targetSheet As String, _
                              Optional statusCell As String = "", _
                              Optional label As String = "") As RecalcSpec
    With NewRecalcSpec
        .key = key
        .targetSheet = targetSheet
        .statusCell = statusCell
        If Len(label) = 0 Then .label = targetSheet Else .label = label
    End With
End Function

Public Sub RunRecalc(spec As RecalcSpec)
    Dim ws As Worksheet
    Dim priorCalculation As XlCalculation
    Dim stampedOn As Date

    On Error GoTo Failed

    Set ws = SheetOrNothing(spec.targetSheet)
    If ws Is Nothing Then
        MsgBox "Sheet '" & spec.targetSheet & "' was not found in this workbook.", vbCritical
        Exit Sub
    End If

    priorCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.StatusBar = "Recalculating " & spec.label & "..."

    ' Manual first so nothing else in the workbook recalculates while this runs.
    ws.EnableCalculation = True
    Application.Calculation = xlCalculationManual
    MarkDirty ws
    

    stampedOn = Now
    LastCalculatedOn(spec.key) = stampedOn
    RefreshRecalcStatus spec

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = priorCalculation

    MsgBox spec.label & " recalculated at " & Format$(stampedOn, "dd mmm yyyy hh:nn") & "." & vbNewLine & _
           "Worksheet automatic calculation is now off.", vbInformation
    ws.EnableCalculation = False
    Exit Sub

Failed:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    ' Restore whatever the workbook was using before the failure.
    If priorCalculation <> 0 Then Application.Calculation = priorCalculation
    MsgBox "The recalculation of " & spec.label & " did not complete." & vbNewLine & _
           Err.description, vbCritical
End Sub

' Writes the recorded timestamp into the spec's status cell. Safe to call from Workbook_Open.
Public Sub RefreshRecalcStatus(spec As RecalcSpec)
    Dim ws As Worksheet

    If Len(spec.statusCell) = 0 Then Exit Sub

    Set ws = SheetOrNothing(spec.targetSheet)
    If ws Is Nothing Then Exit Sub

    ws.Range(spec.statusCell).value = LastCalculatedOnText(spec.key, ifNever:="Never recalculated")
End Sub

' Dirty forces every formula on the sheet to recalculate, not just the stale ones.
Private Sub MarkDirty(ws As Worksheet)
    On Error Resume Next
    ws.UsedRange.Dirty
    On Error GoTo 0
End Sub

Private Function SheetOrNothing(sheetName As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function


