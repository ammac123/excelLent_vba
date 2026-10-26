VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmScenarioEditor
   Caption         =   "Edit Scenario"
   ClientHeight    =   4200
   ClientLeft      =   72
   ClientTop       =   240
   ClientWidth     =   5640
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "frmScenarioEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Edits a scenario's metadata (name, description, effective date, run and compare flags)
' and, in new mode, creates a scenario. tblScenarios is read-only to the user, so this is
' the only way those change. The checks and writes live in ScenarioStore and ScenarioUI;
' this module is controls and view state.
'
' Every control is built in code when the form opens, so the form itself is a blank
' UserForm named frmScenarioEditor and nothing lives in a .frx. Opened through
' ScenarioUI.OpenScenarioEditor, which calls Preselect or StartNew, shows the form
' modally, reads CreatedScenarioId once it hides, and unloads it.

Private Const MARGIN As Single = 12
Private Const LABEL_WIDTH As Single = 84
Private Const FIELD_LEFT As Single = 100
Private Const FIELD_WIDTH As Single = 264
Private Const FIELD_HEIGHT As Single = 18
Private Const ROW_STEP As Single = 24
Private Const AUDIT_HEIGHT As Single = 22
Private Const DESCRIPTION_HEIGHT As Single = 54
Private Const DATE_WIDTH As Single = 90
Private Const BANNER_HEIGHT As Single = 36
Private Const BUTTON_WIDTH As Single = 72
Private Const BUTTON_HEIGHT As Single = 24
Private Const BUTTON_GAP As Single = 6

Private Const DATE_FORMAT As String = "dd-mmm-yyyy"

' BGR, as MSForms wants.
Private Const COLOUR_PROBLEM As Long = &HC0FFFF
Private Const COLOUR_NOTE As Long = &HF0F0F0
Private Const COLOUR_QUIET As Long = &H808080

Private WithEvents cboScenario As MSForms.ComboBox
Private WithEvents txtName As MSForms.TextBox
Private WithEvents txtDescription As MSForms.TextBox
Private WithEvents txtEffectiveDate As MSForms.TextBox
Private WithEvents chkIncludeRun As MSForms.CheckBox
Private WithEvents chkIncludeCompare As MSForms.CheckBox
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnRevert As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton
Private lblScenario As MSForms.Label
Private lblAudit As MSForms.Label
Private lblName As MSForms.Label
Private lblDescription As MSForms.Label
Private lblEffectiveDate As MSForms.Label
Private lblDateHint As MSForms.Label
Private lblBanner As MSForms.Label

' The combo shows names; comboIds holds the ScenarioID behind each entry.
Private comboIds() As String
Private comboCount As Long

Private newMode As Boolean
Private shownId As String     ' scenario the fields were loaded from; "" in new mode or when there is none
Private isDirty As Boolean
Private loading As Boolean    ' held while code fills the controls, so their events aren't edits
Private createdId As String

' The scenario Create made, or "" when the form closed without creating one.
Public Property Get CreatedScenarioId() As String
    CreatedScenarioId = createdId
End Property

'=======================================================================================
' Entry points
'=======================================================================================

' Edit mode, opened on scenarioId. "" or an ID not in the list opens on the first scenario.
Public Sub Preselect(ByVal scenarioId As String)
    newMode = False
    Me.Caption = "Edit Scenario"
    btnSave.Caption = "Save"
    ArrangeControls
    FillScenarioCombo scenarioId
    ShowPicked
End Sub

Public Sub StartNew()
    Dim defaults As ScenarioDetails

    newMode = True
    Me.Caption = "New Scenario"
    btnSave.Caption = "Create"
    ArrangeControls
    shownId = ""
    defaults = NewScenarioDefaults()
    ShowDetails defaults
End Sub

'=======================================================================================
' Building the form
'=======================================================================================

Private Sub UserForm_Initialize()
    ' Set at runtime because the VBE designer rescales sizes under Windows display
    ' scaling. ArrangeControls sets the height once it knows which rows are showing.
    Me.Width = FIELD_LEFT + FIELD_WIDTH + MARGIN + (Me.Width - Me.InsideWidth)
    BuildControls
End Sub

' Added in reading order, which is also the tab order.
Private Sub BuildControls()
    Set lblScenario = AddLabel("lblScenario", "Scenario:")
    Set cboScenario = AddField("Forms.ComboBox.1", "cboScenario", FIELD_HEIGHT)
    cboScenario.Style = fmStyleDropDownList

    Set lblAudit = AddLabel("lblAudit", "")
    lblAudit.Left = FIELD_LEFT
    lblAudit.Width = FIELD_WIDTH
    lblAudit.Height = AUDIT_HEIGHT
    lblAudit.WordWrap = True
    lblAudit.ForeColor = COLOUR_QUIET

    Set lblName = AddLabel("lblName", "Name:")
    Set txtName = AddField("Forms.TextBox.1", "txtName", FIELD_HEIGHT)

    Set lblDescription = AddLabel("lblDescription", "Description:")
    Set txtDescription = AddField("Forms.TextBox.1", "txtDescription", DESCRIPTION_HEIGHT)
    txtDescription.MultiLine = True
    txtDescription.WordWrap = True
    txtDescription.EnterKeyBehavior = True
    txtDescription.ScrollBars = fmScrollBarsVertical

    Set lblEffectiveDate = AddLabel("lblEffectiveDate", "Effective date:")
    Set txtEffectiveDate = AddField("Forms.TextBox.1", "txtEffectiveDate", FIELD_HEIGHT)
    txtEffectiveDate.Width = DATE_WIDTH
    Set lblDateHint = AddLabel("lblDateHint", "e.g. " & Format$(Date, DATE_FORMAT))
    lblDateHint.Left = FIELD_LEFT + DATE_WIDTH + 8
    lblDateHint.Width = FIELD_WIDTH - DATE_WIDTH - 8
    lblDateHint.ForeColor = COLOUR_QUIET

    Set chkIncludeRun = AddField("Forms.CheckBox.1", "chkIncludeRun", FIELD_HEIGHT)
    chkIncludeRun.Caption = "Include in run"
    chkIncludeRun.Width = 120
    chkIncludeRun.ControlTipText = "Run this scenario with Run All Scenarios"
    Set chkIncludeCompare = AddField("Forms.CheckBox.1", "chkIncludeCompare", FIELD_HEIGHT)
    chkIncludeCompare.Caption = "Include in compare"
    chkIncludeCompare.Left = FIELD_LEFT + 130
    chkIncludeCompare.Width = FIELD_WIDTH - 130
    chkIncludeCompare.ControlTipText = "Show this scenario in the scenario comparison"

    Set lblBanner = AddLabel("lblBanner", "")
    lblBanner.Width = FIELD_LEFT + FIELD_WIDTH - MARGIN
    lblBanner.Height = BANNER_HEIGHT
    lblBanner.WordWrap = True

    Set btnSave = AddButton("btnSave", "Save", 3)
    Set btnRevert = AddButton("btnRevert", "Revert", 2)
    btnRevert.ControlTipText = "Throw away the edits and show the saved values again"
    Set btnClose = AddButton("btnClose", "Close", 1)
    btnClose.Cancel = True
End Sub

' Rows are stacked top to bottom; new mode drops the scenario picker and audit line.
Private Sub ArrangeControls()
    Dim y As Single

    lblScenario.Visible = Not newMode
    cboScenario.Visible = Not newMode
    lblAudit.Visible = Not newMode
    btnRevert.Visible = Not newMode

    y = MARGIN
    If Not newMode Then
        PlaceRow lblScenario, cboScenario, y
        y = y + ROW_STEP
        lblAudit.Top = y
        y = y + AUDIT_HEIGHT + 4
    End If

    PlaceRow lblName, txtName, y
    y = y + ROW_STEP
    PlaceRow lblDescription, txtDescription, y
    y = y + DESCRIPTION_HEIGHT + 6
    PlaceRow lblEffectiveDate, txtEffectiveDate, y
    lblDateHint.Top = lblEffectiveDate.Top
    y = y + ROW_STEP
    chkIncludeRun.Top = y
    chkIncludeCompare.Top = y
    y = y + ROW_STEP

    lblBanner.Top = y
    y = y + BANNER_HEIGHT + 8
    btnSave.Top = y
    btnRevert.Top = y
    btnClose.Top = y
    y = y + BUTTON_HEIGHT + MARGIN

    Me.Height = y + (Me.Height - Me.InsideHeight)
    CentreFormOnExcel Me
End Sub

' Labels sit a little lower than their field so the text lines up with the field's.
Private Sub PlaceRow(lbl As MSForms.Label, field As Object, ByVal rowTop As Single)
    field.Top = rowTop
    lbl.Top = rowTop + 3
End Sub

Private Function AddLabel(ctlName As String, labelCaption As String) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", ctlName)
    AddLabel.Left = MARGIN
    AddLabel.Width = LABEL_WIDTH
    AddLabel.Height = FIELD_HEIGHT
    AddLabel.Caption = labelCaption
End Function

Private Function AddField(progId As String, ctlName As String, ByVal ctlHeight As Single) As Object
    Set AddField = Me.Controls.Add(progId, ctlName)
    AddField.Left = FIELD_LEFT
    AddField.Width = FIELD_WIDTH
    AddField.Height = ctlHeight
End Function

' slot counts from the right: 1 is the rightmost button.
Private Function AddButton(ctlName As String, buttonCaption As String, _
                           ByVal slot As Long) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", ctlName)
    AddButton.Left = FIELD_LEFT + FIELD_WIDTH - slot * BUTTON_WIDTH - (slot - 1) * BUTTON_GAP
    AddButton.Width = BUTTON_WIDTH
    AddButton.Height = BUTTON_HEIGHT
    AddButton.Caption = buttonCaption
End Function

'=======================================================================================
' Showing a scenario
'=======================================================================================

' Every scenario except the current state, which is not editable here.
Private Sub FillScenarioCombo(ByVal selectId As String)
    Dim ids As Variant
    Dim i As Long
    Dim pick As Long

    loading = True
    cboScenario.Clear
    comboCount = 0
    pick = -1

    ids = ScenarioIds()
    For i = LBound(ids) To UBound(ids)
        If Not IsCurrentStateId(CStr(ids(i))) Then
            ReDim Preserve comboIds(0 To comboCount)
            comboIds(comboCount) = CStr(ids(i))
            cboScenario.AddItem ScenarioName(comboIds(comboCount))
            If StrComp(comboIds(comboCount), selectId, vbTextCompare) = 0 Then pick = comboCount
            comboCount = comboCount + 1
        End If
    Next i

    If comboCount > 0 Then cboScenario.ListIndex = IIf(pick >= 0, pick, 0)
    loading = False
End Sub

Private Function PickedId() As String
    If cboScenario.ListIndex >= 0 Then PickedId = comboIds(cboScenario.ListIndex)
End Function

' Moves the combo without it counting as the user picking.
Private Sub SelectScenario(scenarioId As String)
    Dim i As Long

    loading = True
    For i = 0 To comboCount - 1
        If StrComp(comboIds(i), scenarioId, vbTextCompare) = 0 Then cboScenario.ListIndex = i
    Next i
    loading = False
End Sub

' Loads the fields from the stored values of whatever the combo is on.
Private Sub ShowPicked()
    Dim details As ScenarioDetails

    shownId = PickedId()
    SetFieldsEnabled (Len(shownId) > 0)
    If Len(shownId) > 0 Then details = ReadScenarioDetails(shownId)
    ShowDetails details
End Sub

Private Sub ShowDetails(details As ScenarioDetails)
    loading = True
    txtName.Value = details.displayName
    txtDescription.Value = details.description
    txtEffectiveDate.Value = DateText(details.effectiveDate)
    chkIncludeRun.Value = details.includeInRun
    chkIncludeCompare.Value = details.includeInCompare
    lblAudit.Caption = AuditText(details)
    loading = False

    SetEdited False
    ShowStandingNote
End Sub

' What the fields hold now, ready for ScenarioDetailsProblem. A date that doesn't parse
' is passed through as text so the check can say so.
Private Function FieldDetails() As ScenarioDetails
    Dim typed As String

    typed = Trim$(txtEffectiveDate.Text)

    With FieldDetails
        .scenarioId = shownId
        .displayName = Trim$(txtName.Text)
        .description = Trim$(txtDescription.Text)
        If IsDate(typed) Then
            .effectiveDate = CDate(typed)
        Else
            .effectiveDate = typed
        End If
        .includeInRun = CBool(chkIncludeRun.Value)
        .includeInCompare = CBool(chkIncludeCompare.Value)
    End With
End Function

Private Sub SetFieldsEnabled(ByVal enabled As Boolean)
    txtName.Enabled = enabled
    txtDescription.Enabled = enabled
    txtEffectiveDate.Enabled = enabled
    chkIncludeRun.Enabled = enabled
    chkIncludeCompare.Enabled = enabled
End Sub

Private Sub SetEdited(ByVal edited As Boolean)
    isDirty = edited
    If newMode Then
        btnSave.Enabled = True
    Else
        btnSave.Enabled = edited
        btnRevert.Enabled = edited
    End If
End Sub

' The banner's resting message for the current mode and scenario.
Private Sub ShowStandingNote()
    If newMode Then
        ShowBanner "Create loads the new scenario into the input tables, starting from the " & _
                   "blank default."
    ElseIf Len(shownId) = 0 Then
        ShowBanner "There are no scenarios to edit yet. Use New Scenario to create one."
    ElseIf StrComp(shownId, ActiveScenarioIdVal(), vbTextCompare) = 0 Then
        ShowBanner "This is the loaded scenario, so saving also sets the model's effective " & _
                   "date to the one shown here."
    Else
        ShowBanner ""
    End If
End Sub

Private Sub ShowBanner(text As String, Optional ByVal isProblem As Boolean = False)
    lblBanner.Caption = text
    If Len(text) = 0 Then
        lblBanner.BackStyle = fmBackStyleTransparent
        lblBanner.BorderStyle = fmBorderStyleNone
    Else
        lblBanner.BackStyle = fmBackStyleOpaque
        lblBanner.BackColor = IIf(isProblem, COLOUR_PROBLEM, COLOUR_NOTE)
        lblBanner.BorderStyle = fmBorderStyleSingle
    End If
End Sub

Private Function DateText(value As Variant) As String
    If IsDate(value) Then
        DateText = Format$(CDate(value), DATE_FORMAT)
    Else
        DateText = CStr(value)
    End If
End Function

' "SCN0003. Created 12-Oct-2026 by A Macleod. Last changed 26-Oct-2026 by A Macleod."
Private Function AuditText(details As ScenarioDetails) As String
    If Len(details.scenarioId) = 0 Then Exit Function
    AuditText = details.scenarioId & ". Created " & _
                AuditStamp(details.createdOn, details.createdBy) & ". Last changed " & _
                AuditStamp(details.modifiedOn, details.modifiedBy) & "."
End Function

Private Function AuditStamp(stamp As Variant, who As String) As String
    If IsDate(stamp) Then
        AuditStamp = Format$(CDate(stamp), DATE_FORMAT)
    Else
        AuditStamp = "on an unknown date"
    End If
    If Len(who) > 0 Then AuditStamp = AuditStamp & " by " & who
End Function

'=======================================================================================
' Saving and creating
'=======================================================================================

' Saves the fields over shownId, then shows showAfter ("" shows the same scenario again).
' False when nothing was saved; the banner says why.
Private Function SaveChanges(Optional ByVal showAfter As String = "") As Boolean
    Dim details As ScenarioDetails
    Dim problem As String
    Dim warning As String

    details = FieldDetails()
    problem = ScenarioDetailsProblem(details)
    If Len(problem) > 0 Then
        ShowBanner problem, isProblem:=True
        Exit Function
    End If

    On Error GoTo Failed
    warning = SaveScenarioFromEditor(details)
    On Error GoTo 0

    If Len(showAfter) = 0 Then showAfter = details.scenarioId
    FillScenarioCombo showAfter
    ShowPicked

    If Len(warning) > 0 Then MsgBox warning, vbExclamation, Me.Caption
    SaveChanges = True
    Exit Function

Failed:
    ShowBanner "Could not save: " & Err.description, isProblem:=True
End Function

Private Sub CreateScenario()
    Dim details As ScenarioDetails
    Dim problem As String

    details = FieldDetails()
    problem = ScenarioDetailsProblem(details)
    If Len(problem) > 0 Then
        ShowBanner problem, isProblem:=True
        Exit Sub
    End If

    On Error GoTo Failed
    createdId = CreateScenarioFromEditor(details)
    Me.Hide
    Exit Sub

Failed:
    ShowBanner "Could not create the scenario: " & Err.description, isProblem:=True
End Sub

' Unsaved edits are offered for saving first. New mode just closes: nothing is created
' until Create is pressed.
Private Sub CloseForm()
    If isDirty And Not newMode Then
        Select Case MsgBox("Save your changes to '" & ScenarioName(shownId) & "'?", _
                           vbYesNoCancel + vbQuestion, Me.Caption)
            Case vbYes
                If Not SaveChanges() Then Exit Sub
            Case vbCancel
                Exit Sub
        End Select
    End If
    Me.Hide
End Sub

'=======================================================================================
' Events
'=======================================================================================

Private Sub UserForm_Activate()
    ' The suggested name is there to be typed over.
    If newMode Then
        txtName.SetFocus
        txtName.SelStart = 0
        txtName.SelLength = Len(txtName.Text)
    End If
End Sub

' The close box hides like the Close button does, so OpenScenarioEditor can still read
' the outcome. Unload from code goes through untouched.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> vbFormControlMenu Then Exit Sub
    Cancel = True
    CloseForm
End Sub

Private Sub cboScenario_Change()
    Dim pickedNow As String

    If loading Then Exit Sub
    pickedNow = PickedId()
    If StrComp(pickedNow, shownId, vbTextCompare) = 0 Then Exit Sub

    If isDirty Then
        Select Case MsgBox("Save your changes to '" & ScenarioName(shownId) & "' first?", _
                           vbYesNoCancel + vbQuestion, Me.Caption)
            Case vbYes
                If Not SaveChanges(pickedNow) Then SelectScenario shownId
                Exit Sub
            Case vbCancel
                SelectScenario shownId
                Exit Sub
        End Select
    End If

    ShowPicked
End Sub

Private Sub txtName_Change()
    FieldEdited
End Sub

Private Sub txtDescription_Change()
    FieldEdited
End Sub

Private Sub txtEffectiveDate_Change()
    FieldEdited
End Sub

Private Sub chkIncludeRun_Click()
    FieldEdited
End Sub

Private Sub chkIncludeCompare_Click()
    FieldEdited
End Sub

Private Sub FieldEdited()
    If loading Then Exit Sub
    SetEdited True
    ShowStandingNote
End Sub

Private Sub btnSave_Click()
    If newMode Then
        CreateScenario
    Else
        SaveChanges
    End If
End Sub

Private Sub btnRevert_Click()
    ShowPicked
End Sub

Private Sub btnClose_Click()
    CloseForm
End Sub
