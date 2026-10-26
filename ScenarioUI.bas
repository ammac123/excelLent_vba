Attribute VB_Name = "ScenarioUI"
Option Explicit

' Button handlers for the Scenario Manager sheet. Everything here is prompts, guards and
' messages; the actual work lives in ScenarioStore.
'
' Assign one macro per button. The picker is the cell named scenarioPicker, validated
' against the Name column of tblScenarios.

Private Const MANAGER_SHEET As String = "04_SCN_Manager"
Private Const NAME_PICKER As String = "scenarioPicker"
Private Const NAME_STATUS As String = "scenarioStatus"
Private Const NAME_ACTIVE_NAME As String = "activeScenarioName"

Private Const EDITOR_FORM As String = "frmScenarioEditor"

' Opens frmScenarioEditor in new mode. Creating loads the new scenario over the active
' tables, so unsaved edits are confirmed away before the form opens.
Public Sub NewScenarioButton()
    On Error GoTo Failed
    If Not ConfirmDiscardIfDirty("Create a new scenario") Then Exit Sub

    OpenScenarioEditor newMode:=True
    Exit Sub

Failed:
    ReportFailure "create the scenario"
End Sub

' Opens frmScenarioEditor on the scenario in the picker, else the loaded one.
Public Sub EditScenarioButton()
    On Error GoTo Failed
    OpenScenarioEditor DefaultEditScenarioId()
    Exit Sub

Failed:
    ReportFailure "open the scenario editor"
End Sub

' Kept so a Rename shape assigned before the editor existed still does something useful.
Public Sub RenameScenarioButton()
    EditScenarioButton
End Sub

' Shows the editor modally. scenarioId is the one to open on; "" or an unknown ID opens
' on the first scenario. Late bound so the rest of the project compiles without the form.
Public Sub OpenScenarioEditor(Optional scenarioId As String = "", _
                              Optional newMode As Boolean = False)
    Dim frm As Object
    Dim createdId As String

    On Error GoTo FormMissing
    Set frm = UserForms.Add(EDITOR_FORM)
    On Error GoTo 0

    If newMode Then
        frm.StartNew
    Else
        frm.Preselect scenarioId
    End If
    frm.Show vbModal

    ' The form hides rather than unloads, so its outcome can still be read here.
    createdId = frm.CreatedScenarioId
    Unload frm

    If Len(createdId) > 0 Then
        MsgBox "Created '" & ScenarioName(createdId) & "' (" & createdId & ")." & vbNewLine & _
               vbNewLine & "It starts from the blank default. Enter the ATL adjustments " & _
               "and the BTL discretions, then save.", vbInformation
    End If
    Exit Sub

FormMissing:
    MsgBox EDITOR_FORM & " is not in this workbook yet." & vbNewLine & vbNewLine & _
           "See " & EDITOR_FORM & "_SETUP.md in the project folder for how to add it. " & _
           "Once it's in place, save the workbook and this button will open it.", _
           vbInformation, "Scenario editor"
End Sub

' For Worksheet_BeforeDoubleClick on the manager sheet. A double-click on a scenario row
' opens the editor on it; anywhere else in tblScenarios explains the table is read-only.
' Returns True when the click was handled, so the sheet can cancel Excel's own response.
Public Function HandleScenarioTableDoubleClick(target As Range) As Boolean
    Dim lo As ListObject
    Dim scenarioId As String

    On Error GoTo Failed
    Set lo = TableOrNothing(TBL_SCENARIOS)
    If lo Is Nothing Then Exit Function
    If Not target.Worksheet Is lo.Parent Then Exit Function
    If Application.Intersect(target.Cells(1), lo.Range) Is Nothing Then Exit Function

    HandleScenarioTableDoubleClick = True

    If lo.DataBodyRange Is Nothing Then
        MsgBox TBL_SCENARIOS & " is read-only. Use the New Scenario button to add one.", _
               vbInformation
        Exit Function
    End If
    If Application.Intersect(target.Cells(1), lo.DataBodyRange) Is Nothing Then
        MsgBox TBL_SCENARIOS & " is read-only. Use the Edit Scenario button, or " & _
               "double-click a scenario's row.", vbInformation
        Exit Function
    End If

    scenarioId = Trim$(CStr(CellInRow(lo, target.Row - lo.DataBodyRange.Row + 1, _
                                      COL_SCENARIO_ID).value))
    If IsCurrentStateId(scenarioId) Then
        MsgBox "The current state isn't an editable scenario. Use Run Current State to " & _
               "refresh it.", vbInformation
        Exit Function
    End If

    OpenScenarioEditor scenarioId
    Exit Function

Failed:
    ReportFailure "open the scenario editor"
End Function

' What a new scenario starts with in the editor.
Public Function NewScenarioDefaults() As ScenarioDetails
    With NewScenarioDefaults
        .displayName = SuggestedName()
        .effectiveDate = Date
        .includeInRun = True
        .includeInCompare = True
    End With
End Function

' Called by the editor's Create button once the details have passed
' ScenarioDetailsProblem. Returns the new ID.
Public Function CreateScenarioFromEditor(details As ScenarioDetails) As String
    CreateScenarioFromEditor = NewScenario(Trim$(details.displayName), _
                                           Trim$(details.description), _
                                           CDate(details.effectiveDate), _
                                           details.includeInRun, details.includeInCompare)
    RefreshScenarioManager
End Function

' Called by the editor's Save button. Returns a warning for the user, or "".
Public Function SaveScenarioFromEditor(details As ScenarioDetails) As String
    Dim before As ScenarioDetails

    before = ReadScenarioDetails(details.scenarioId)
    UpdateScenarioDetails details

    RepointPicker NAME_PICKER, before.displayName, Trim$(details.displayName)
    RepointPicker OutputScenarioPickerName(), before.displayName, Trim$(details.displayName)
    RefreshScenarioManager

    If Not SameDate(before.effectiveDate, details.effectiveDate) Then
        If HasStoredResults(details.scenarioId) Then
            SaveScenarioFromEditor = "The stored results for '" & Trim$(details.displayName) & _
                                     "' were run with the old effective date. Run it " & _
                                     "again to refresh them."
        End If
    End If
End Function

Public Sub LoadScenarioButton()
    Dim scenarioId As String

    On Error GoTo Failed
    scenarioId = PickedScenarioId()
    If Len(scenarioId) = 0 Then Exit Sub

    If StrComp(scenarioId, ActiveScenarioIdVal(), vbTextCompare) = 0 And Not ActiveIsDirty() Then
        MsgBox "'" & ScenarioName(scenarioId) & "' is already loaded.", vbInformation
        Exit Sub
    End If

    If Not ConfirmDiscardIfDirty("Load '" & ScenarioName(scenarioId) & "'") Then Exit Sub

    LoadScenario scenarioId
    Application.CalculateFull
    RefreshScenarioManager

    MsgBox "Loaded '" & ScenarioName(scenarioId) & "'.", vbInformation
    Exit Sub

Failed:
    ReportFailure "load the scenario"
End Sub

Public Sub SaveScenarioButton()
    Dim scenarioId As String
    Dim problem As String

    On Error GoTo Failed
    scenarioId = ActiveScenarioIdVal()
    If Len(scenarioId) = 0 Then
        MsgBox "No scenario is loaded. Create one or load one first.", vbExclamation
        Exit Sub
    End If

    If Not ActiveScenarioIsValid(problem) Then
        MsgBox "This scenario cannot be saved yet:" & vbNewLine & vbNewLine & problem, vbExclamation
        Exit Sub
    End If

    SaveActiveScenario
    LastValidatedOn(scenarioId) = Now
    RefreshScenarioManager

    MsgBox "Saved '" & ScenarioName(scenarioId) & "'.", vbInformation
    Exit Sub

Failed:
    ReportFailure "save the scenario"
End Sub

Public Sub SaveScenarioAsButton()
    Dim displayName As String
    Dim scenarioId As String
    Dim problem As String

    On Error GoTo Failed
    If Not ActiveScenarioIsValid(problem) Then
        MsgBox "This scenario cannot be saved yet:" & vbNewLine & vbNewLine & problem, vbExclamation
        Exit Sub
    End If

    displayName = Trim$(InputBox("Name for the copy:", "Save Scenario As", _
                                 ActiveScenarioNameVal() & " (copy)"))
    If Len(displayName) = 0 Then Exit Sub

    If Not NameIsAvailable(displayName) Then
        MsgBox "A scenario named '" & displayName & "' already exists.", vbExclamation
        Exit Sub
    End If

    scenarioId = SaveActiveScenarioAs(displayName)
    LastValidatedOn(scenarioId) = Now
    RefreshScenarioManager

    MsgBox "Saved as '" & displayName & "' (" & scenarioId & "), which is now the loaded " & _
           "scenario.", vbInformation
    Exit Sub

Failed:
    ReportFailure "save a copy of the scenario"
End Sub

Public Sub DeleteScenarioButton()
    Dim scenarioId As String
    Dim displayName As String

    On Error GoTo Failed
    scenarioId = PickedScenarioId()
    If Len(scenarioId) = 0 Then Exit Sub

    displayName = ScenarioName(scenarioId)

    If MsgBox("Delete '" & displayName & "' (" & scenarioId & ")?" & vbNewLine & vbNewLine & _
              "Its inputs and its stored results will be removed. This cannot be undone.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Delete Scenario") = vbNo Then Exit Sub

    DeleteScenario scenarioId
    RefreshScenarioManager

    If Len(ActiveScenarioIdVal()) = 0 Then
        MsgBox "Deleted '" & displayName & "'. It was the loaded scenario, so the input " & _
               "sheets still hold its values. Load another scenario before recalculating.", _
               vbInformation
    Else
        MsgBox "Deleted '" & displayName & "'.", vbInformation
    End If
    Exit Sub

Failed:
    ReportFailure "delete the scenario"
End Sub

' Throws away unsaved edits by reloading the active scenario from the store.
Public Sub RevertScenarioButton()
    Dim scenarioId As String

    On Error GoTo Failed
    scenarioId = ActiveScenarioIdVal()
    If Len(scenarioId) = 0 Then
        MsgBox "No scenario is loaded, so there is nothing to revert to.", vbExclamation
        Exit Sub
    End If

    If Not ActiveIsDirty() Then
        MsgBox "'" & ScenarioName(scenarioId) & "' has no unsaved changes.", vbInformation
        Exit Sub
    End If

    If MsgBox("Discard all unsaved changes and reload '" & ScenarioName(scenarioId) & "'?", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Revert Scenario") = vbNo Then Exit Sub

    LoadScenario scenarioId
    RefreshScenarioManager
    Exit Sub

Failed:
    ReportFailure "revert the scenario"
End Sub

Public Sub ValidateScenarioButton()
    Dim problem As String

    On Error GoTo Failed
    If ActiveScenarioIsValid(problem) Then
        MsgBox "The loaded inputs are valid and ready to run.", vbInformation
    Else
        MsgBox problem, vbExclamation
    End If
    Exit Sub

Failed:
    ReportFailure "validate the scenario"
End Sub

'=======================================================================================
' Metrics and comparison
'=======================================================================================

' Rescans the source tables and refreshes the catalogue, keeping the ticks already set.
Public Sub SyncMetricsButton()
    Dim total As Long, added As Long, removed As Long

    On Error GoTo Failed
    total = SyncMetricCatalogue(added, removed)

    MsgBox "The metric catalogue now lists " & total & " metric(s)." & vbNewLine & _
           added & " new, " & removed & " no longer in the source tables." & vbNewLine & _
           vbNewLine & "Tick the ones you want captured, then run the scenarios.", _
           vbInformation
    Exit Sub

Failed:
    ReportFailure "sync the metric catalogue"
End Sub

Public Sub IncludeSelectedMetricsButton()
    SetMetricSelection True, "captured"
End Sub

Public Sub ExcludeSelectedMetricsButton()
    SetMetricSelection False, "ignored"
End Sub

Public Sub IncludeAllMetricsButton()
    On Error GoTo Failed
    MsgBox SetAllIncluded(True) & " metric(s) switched on.", vbInformation
    Exit Sub

Failed:
    ReportFailure "include every metric"
End Sub

Public Sub ExcludeAllMetricsButton()
    On Error GoTo Failed
    MsgBox SetAllIncluded(False) & " metric(s) switched off.", vbInformation
    Exit Sub

Failed:
    ReportFailure "exclude every metric"
End Sub

Public Sub RunScenarioButton()
    RunActiveScenario
End Sub

Public Sub RunAllScenariosButton()
    RunIncludedScenarios
End Sub

' Refreshes the baseline results independently of any loaded scenario. Pair with
' tokens like {Current.X}, {Delta.X} and {DeltaPct.X} in compare templates.
Public Sub RunCurrentStateButton()
    RunCurrentState
End Sub

Public Sub BuildComparisonButton()
    BuildScenarioComparison
End Sub

' Renders the picked scenario's detail block on the output sheet, including any charts
' declared via {Chart:<name>} tokens in the chosen template.
Public Sub BuildScenarioOutputButton()
    BuildScenarioOutput
End Sub

' Opens the modeless token builder form. Assign this macro to a Shape button on any
' sheet where composing tokens is useful (output sheet, comparison sheet, manager
' sheet) — the form doesn't touch the active sheet or selection, so the user can
' Alt+Tab back to their template and paste.
Public Sub OpenTokenBuilderButton()
    On Error GoTo Failed
    OpenTokenBuilder
    Exit Sub
Failed:
    ReportFailure "open the token builder"
End Sub

' Prompts for the new range for the output template picked in cmbRedefineOutputTemplate
' and repoints its defined name. A missing name is created on the fly.
Public Sub RedefineOutputTemplateRangeButton()
    Dim chosen As CompareTemplateSpec
    Dim target As Range
    Dim prompt As String
    Dim defaultAddress As String

    On Error GoTo Failed
    RefreshOutputRedefineCombo

    chosen = SelectedOutputRedefineTemplate()

    prompt = "Select the new range for the '" & chosen.label & "' output template, then click OK." & _
             vbNewLine & _
             "The label will be written in the row above the top-left cell."

    If TypeName(Selection) = "Range" Then defaultAddress = Selection.Address

    On Error Resume Next
    Set target = Application.InputBox(prompt:=prompt, Title:="Redefine output template range", _
                                      Default:=defaultAddress, Type:=8)
    On Error GoTo Failed

    If target Is Nothing Then Exit Sub

    RedefineSelectedOutputTemplateRange target

    MsgBox "Redefined '" & chosen.label & "' to " & _
           target.Address(False, False) & ".", vbInformation
    Exit Sub

Failed:
    ReportFailure "redefine the output template range"
End Sub

' Re-frames every summary and output template where it stands. Run once to move templates
' marked the old way (border on the range itself) over to the outside highlight frame.
Public Sub RefreshTemplateFramesButton()
    Dim skipped As String

    On Error GoTo Failed
    skipped = RefreshTemplateFrames()

    If Len(skipped) = 0 Then
        MsgBox "Every template range is framed.", vbInformation
    Else
        MsgBox "Framed the template ranges, except:" & vbNewLine & vbNewLine & skipped, _
               vbExclamation
    End If
    Exit Sub

Failed:
    ReportFailure "refresh the template frames"
End Sub

' Prompts for the new range for the template picked in the cmbRedefineTemplate dropdown,
' moves the highlight frame and the header label with it. A missing named range is created
' on the fly so a deleted template range can be reinstated by picking it and selecting a
' rectangle.
Public Sub RedefineTemplateRangeButton()
    Dim chosen As CompareTemplateSpec
    Dim target As Range
    Dim prompt As String
    Dim defaultAddress As String
    Dim picked As Variant

    On Error GoTo Failed
    RefreshRedefineTemplatePicker

    chosen = SelectedRedefineTemplate()

    prompt = "Select the new range for the '" & chosen.label & "' template, then click OK." & _
             vbNewLine & _
             "The label will be written in the row above the top-left cell."

    If TypeName(Selection) = "Range" Then defaultAddress = Selection.Address

    ' InputBox returns False on Cancel, which can't Set into a Range — suppress the
    ' type-mismatch so we can detect cancel cleanly.
    On Error Resume Next
    Set target = Application.InputBox(prompt:=prompt, Title:="Redefine template range", _
                                      Default:=defaultAddress, Type:=8)
    On Error GoTo Failed

    If target Is Nothing Then Exit Sub

    RedefineSelectedTemplateRange target

    MsgBox "Redefined '" & chosen.label & "' to " & _
           target.Address(False, False) & ".", vbInformation
    Exit Sub

Failed:
    ReportFailure "redefine the template range"
End Sub

Private Sub SetMetricSelection(ByVal flag As Boolean, verb As String)
    Dim changed As Long

    On Error GoTo Failed
    changed = SetSelectionIncluded(flag)

    If changed < 0 Then
        MsgBox "Select the rows you want in " & MetricCatalogueTableName() & " first.", _
               vbExclamation
    Else
        MsgBox changed & " metric(s) will now be " & verb & ".", vbInformation
    End If
    Exit Sub

Failed:
    ReportFailure "change the selected metrics"
End Sub

' Writes the summary block on the manager sheet. Safe to call from Workbook_Open.
Public Sub RefreshScenarioManager()
    Dim cell As Range
    Dim scenarioId As String

    scenarioId = ActiveScenarioIdVal()

    Set cell = NamedRangeOrNothing(NAME_ACTIVE_NAME)
    If Not cell Is Nothing Then
        cell.value = IIf(Len(scenarioId) = 0, "No scenario loaded", ScenarioName(scenarioId))
    End If

    Set cell = NamedRangeOrNothing(NAME_STATUS)
    If Not cell Is Nothing Then cell.value = StatusText(scenarioId)

    RefreshOutputScenarioPicker
End Sub

Private Function StatusText(scenarioId As String) As String
    If Len(scenarioId) = 0 Then
        StatusText = "No scenario loaded"
    ElseIf ActiveIsDirty() Then
        StatusText = "Unsaved changes"
    Else
        StatusText = "Saved " & LastValidatedOnText(scenarioId, ifNever:="but never validated")
    End If
End Function

' The scenario currently chosen in the picker. Empty when nothing usable is selected.
Private Function PickedScenarioId() As String
    Dim cell As Range
    Dim picked As String

    Set cell = NamedRangeOrNothing(NAME_PICKER)
    If cell Is Nothing Then
        MsgBox "The named range '" & NAME_PICKER & "' is missing from the " & MANAGER_SHEET & _
               " sheet.", vbCritical
        Exit Function
    End If

    picked = Trim$(CStr(cell.value))
    If Len(picked) = 0 Then
        MsgBox "Choose a scenario in the picker first.", vbExclamation
        Exit Function
    End If

    PickedScenarioId = ScenarioIdForName(picked)
    If Len(PickedScenarioId) = 0 Then
        MsgBox "'" & picked & "' is not one of the saved scenarios.", vbExclamation
    End If
End Function

Private Function ConfirmDiscardIfDirty(action As String) As Boolean
    If Not ActiveIsDirty() Then
        ConfirmDiscardIfDirty = True
        Exit Function
    End If

    ConfirmDiscardIfDirty = (MsgBox("The loaded scenario has unsaved changes." & vbNewLine & _
                                    action & " and discard them?", _
                                    vbYesNo + vbExclamation + vbDefaultButton2, _
                                    "Unsaved Changes") = vbYes)
End Function

' "Scenario 26 Oct 2026", or "... (2)" and so on when that name is already taken.
Private Function SuggestedName() As String
    Dim stem As String
    Dim n As Long

    stem = "Scenario " & Format$(Date, "dd mmm yyyy")
    SuggestedName = stem
    n = 1
    Do While Not NameIsAvailable(SuggestedName)
        n = n + 1
        SuggestedName = stem & " (" & n & ")"
    Loop
End Function

' The scenario the editor opens on: the picker's, else the loaded one, never the current
' state. "" lets the form fall back to the first scenario.
Private Function DefaultEditScenarioId() As String
    Dim cell As Range
    Dim scenarioId As String

    Set cell = NamedRangeOrNothing(NAME_PICKER)
    If Not cell Is Nothing Then
        If Len(Trim$(CStr(cell.value))) > 0 Then scenarioId = ScenarioIdForName(CStr(cell.value))
    End If
    If Len(scenarioId) = 0 Or IsCurrentStateId(scenarioId) Then scenarioId = ActiveScenarioIdVal()
    If IsCurrentStateId(scenarioId) Then scenarioId = ""

    DefaultEditScenarioId = scenarioId
End Function

' Pickers hold scenario names, so after a rename one still showing the old name would
' point at nothing. Events are held off so the write doesn't count as the user picking.
Private Sub RepointPicker(nameRef As String, oldName As String, newName As String)
    Dim cell As Range
    Dim priorEvents As Boolean

    If StrComp(oldName, newName, vbBinaryCompare) = 0 Then Exit Sub

    Set cell = NamedRangeOrNothing(nameRef)
    If cell Is Nothing Then Exit Sub
    If StrComp(Trim$(CStr(cell.value)), oldName, vbTextCompare) <> 0 Then Exit Sub

    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    cell.value = newName
    Application.EnableEvents = priorEvents
End Sub

Private Function SameDate(a As Variant, b As Variant) As Boolean
    If IsDate(a) And IsDate(b) Then
        SameDate = (CDate(a) = CDate(b))
    Else
        SameDate = (CStr(a) = CStr(b))
    End If
End Function

Private Sub ReportFailure(action As String)
    MsgBox "Could not " & action & "." & vbNewLine & vbNewLine & Err.description, vbCritical
End Sub

