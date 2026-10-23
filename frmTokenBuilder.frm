VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTokenBuilder 
   Caption         =   "Token builder"
   ClientHeight    =   7572
   ClientLeft      =   72
   ClientTop       =   240
   ClientWidth     =   5412
   OleObjectBlob   =   "frmTokenBuilder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTokenBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Event handlers and view state for the token builder form. All compose /
' parse / validate / lookup logic lives in ScenarioTokenBuilder.

Private editingTokenId As String
Private suppressRefresh As Boolean

' The segment attribute panel is built in code when the form opens rather than in the
' designer, so adding it needed no change to the .frx. It copies the scenario panel's
' geometry and font: one row per pick, then the live value, then the format row.
Private Const SEGMENT_ROW_STEP As Single = 22

' MSForms list boxes drift when their contents are rebuilt: IntegralHeight snaps the
' height to whole rows and display scaling compounds it, so the series lists jump
' each time a source or baseline table is picked. The designed geometry is captured
' on open and put back after every repopulate.
Private seriesListBox As Variant
Private baseSeriesListBox As Variant

Private WithEvents cboSegmentId As MSForms.ComboBox
Private WithEvents cboSegmentCol As MSForms.ComboBox
Private lblSegmentId As MSForms.Label
Private lblSegmentCol As MSForms.Label
Private lblSegmentValue As MSForms.Label
Private lblSegmentValueVal As MSForms.Label

Private Sub UserForm_Initialize()
    ' Set at runtime because the VBE designer rescales these under Windows
    ' display scaling (e.g. 125%) and won't keep typed values.
    Me.Width = 540
    Me.Height = 620
    lstInlineSeries.IntegralHeight = False
    lstInlineBaseSeries.IntegralHeight = False
    seriesListBox = CaptureBox(lstInlineSeries)
    baseSeriesListBox = CaptureBox(lstInlineBaseSeries)
    BuildSegmentPanel
    ApplyTabOrder
    PopulateKindCombo
    PopulateMetricSourceCombo
    PopulateScenarioCombo
    PopulateSegmentCombos
    PopulateValueFormatCombo
    PopulateNamedChartCombo
    PopulateInlineKindCombo
    PopulateSourceCombos
    optMetric.Value = True
    cboTokenKind.ListIndex = 0
    RefreshSavedTokens
    ApplyKindVisibility
    btnCancel.Enabled = False
    RebuildPreview
End Sub

' Tab order follows reading order. Set in code because TabIndex lives in the
' .frx binary. Hidden panels are skipped automatically, so all four panels can
' share the sequence.
Private Sub ApplyTabOrder()
    Dim order As Variant, i As Long
    order = Array(cboTokenKind, _
                  cboMetricSource, cboMetricSegment, cboMetricType, _
                  optMetric, optCurrent, optDelta, optDeltaPct, optLabel, _
                  cboScenarioCol, _
                  cboSegmentId, cboSegmentCol, _
                  cboValueFormat, _
                  cboChartNamed, _
                  cboInlineKind, cboInlineSource, cboInlineX, lstInlineSeries, _
                  cboInlineBase, lstInlineBaseSeries, txtInlineTitle, _
                  txtInlineFormat, chkInlineLegend, txtInlineHeight, txtInlineWidth, _
                  txtPreview, btnCopy, btnSave, btnCancel, btnClose, _
                  lstSaved, btnLibCopy, btnLibEdit, btnLibDelete)
    ' Assigning ascending indexes leaves the earlier ones in place.
    For i = LBound(order) To UBound(order)
        order(i).TabIndex = i - LBound(order)
    Next i
End Sub

Private Sub PopulateKindCombo()
    cboTokenKind.Clear
    cboTokenKind.AddItem "Metric value"
    cboTokenKind.AddItem "Scenario field"
    cboTokenKind.AddItem "Chart (named)"
    cboTokenKind.AddItem "Chart (inline)"
    cboTokenKind.AddItem "Segment attribute"
End Sub

' Metric picker: Source -> Segment -> Metric. Each Fill* rebuilds the list below
' it and cascades down; a list with one option selects it automatically. The
' Fill* subs hold suppressRefresh while they change lists, so the combos' own
' Change handlers don't re-enter.
Private Sub PopulateMetricSourceCombo()
    FillCombo cboMetricSource, MetricSourceChoices()
End Sub

Private Sub FillMetricSegments()
    Dim prior As Boolean
    prior = suppressRefresh
    suppressRefresh = True
    FillCombo cboMetricSegment, MetricSegmentChoices(cboMetricSource.Value)
    If cboMetricSegment.ListCount = 1 Then cboMetricSegment.ListIndex = 0
    suppressRefresh = prior
    FillMetricTypes
End Sub

Private Sub FillMetricTypes()
    Dim prior As Boolean
    prior = suppressRefresh
    suppressRefresh = True
    FillCombo cboMetricType, MetricTypeChoices(cboMetricSource.Value, cboMetricSegment.Value)
    If cboMetricType.ListCount = 1 Then cboMetricType.ListIndex = 0
    suppressRefresh = prior
End Sub

Private Sub FillCombo(cbo As MSForms.ComboBox, items() As String)
    Dim i As Long
    cbo.Clear
    For i = LBound(items) To UBound(items)
        cbo.AddItem items(i)
    Next i
End Sub

Private Function CurrentMetricId() As String
    CurrentMetricId = MetricIdFor(cboMetricSource.Value, cboMetricSegment.Value, _
                                  cboMetricType.Value)
End Function

' Points all three combos at an existing metric (used by Edit). A metric that has
' left the catalogue can't be split back into its parts, so the combos are
' cleared and the banner explains why.
Private Sub SelectMetricById(metricId As String)
    Dim m As MetricSpec
    m = MetricById(metricId)
    If Len(m.metric) = 0 Then
        cboMetricSource.ListIndex = -1
        cboMetricSegment.Clear
        cboMetricType.Clear
        Exit Sub
    End If
    SelectComboValue cboMetricSource, MetricSourceKey(m)
    FillMetricSegments
    SelectComboValue cboMetricSegment, MetricSegmentKey(m)
    FillMetricTypes
    SelectComboValue cboMetricType, MetricTypeKey(m)
End Sub

' Segment picker: ID and Column are independent, so neither list depends on the other.
Private Sub PopulateSegmentCombos()
    FillCombo cboSegmentId, SegmentIdChoices()
    FillCombo cboSegmentCol, SegmentColumnChoices()
End Sub

Private Sub BuildSegmentPanel()
    Dim rowTop As Single
    rowTop = cboScenarioCol.Top
    Set lblSegmentId = AddPanelLabel("lblSegmentId", "Segment ID:", rowTop)
    Set cboSegmentId = AddPanelCombo("cboSegmentId", rowTop)
    Set lblSegmentCol = AddPanelLabel("lblSegmentCol", "Column:", rowTop + SEGMENT_ROW_STEP)
    Set cboSegmentCol = AddPanelCombo("cboSegmentCol", rowTop + SEGMENT_ROW_STEP)
    Set lblSegmentValue = AddPanelLabel("lblSegmentValue", "Value now:", _
                                        rowTop + 2 * SEGMENT_ROW_STEP)
    Set lblSegmentValueVal = AddPanelLabel("lblSegmentValueVal", "", _
                                           rowTop + 2 * SEGMENT_ROW_STEP)
    lblSegmentValueVal.Left = cboScenarioCol.Left
    lblSegmentValueVal.Width = Me.InsideWidth - cboScenarioCol.Left - lblScenarioCol.Left
End Sub

Private Function AddPanelLabel(name As String, caption As String, _
                               ByVal rowTop As Single) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", name)
    lbl.Left = lblScenarioCol.Left
    lbl.Top = rowTop + (lblScenarioCol.Top - cboScenarioCol.Top)
    lbl.Width = lblScenarioCol.Width
    lbl.Height = lblScenarioCol.Height
    lbl.BackStyle = lblScenarioCol.BackStyle
    lbl.Font.Name = lblScenarioCol.Font.Name
    lbl.Font.Size = lblScenarioCol.Font.Size
    lbl.Caption = caption
    Set AddPanelLabel = lbl
End Function

Private Function AddPanelCombo(name As String, ByVal rowTop As Single) As MSForms.ComboBox
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls.Add("Forms.ComboBox.1", name)
    cbo.Left = cboScenarioCol.Left
    cbo.Top = rowTop
    cbo.Width = cboScenarioCol.Width
    cbo.Height = cboScenarioCol.Height
    cbo.Style = cboScenarioCol.Style
    cbo.Font.Name = cboScenarioCol.Font.Name
    cbo.Font.Size = cboScenarioCol.Font.Size
    Set AddPanelCombo = cbo
End Function

Private Sub PopulateScenarioCombo()
    Dim names() As String, i As Long
    cboScenarioCol.Clear
    names = ScenarioFieldChoices()
    For i = LBound(names) To UBound(names)
        cboScenarioCol.AddItem names(i)
    Next i
End Sub

' Presets for the common cases. The box stays editable, so any Excel number format
' code can be typed in; blank leaves the token unformatted.
Private Sub PopulateValueFormatCombo()
    Dim presets As Variant, i As Long
    presets = Array("0.00%", "0.0%", "#,##0", "#,##0.00", "$#,##0", _
                    "#,##0;(#,##0);-", "0.00", "dd-mmm-yy", "mmm-yy")
    cboValueFormat.Style = fmStyleDropDownCombo
    cboValueFormat.Clear
    For i = LBound(presets) To UBound(presets)
        cboValueFormat.AddItem presets(i)
    Next i
End Sub

Private Sub PopulateNamedChartCombo()
    Dim specs() As ChartSpec, i As Long
    cboChartNamed.Clear
    specs = NamedChartChoices()
    For i = LBound(specs) To UBound(specs)
        cboChartNamed.AddItem specs(i).label
    Next i
End Sub

Private Sub PopulateInlineKindCombo()
    cboInlineKind.Clear
    cboInlineKind.AddItem "Line"
    cboInlineKind.AddItem "Column"
    cboInlineKind.AddItem "Bar"
    cboInlineKind.AddItem "Combo"
    cboInlineKind.ListIndex = 0
End Sub

Private Sub PopulateSourceCombos()
    Dim names() As String, i As Long
    cboInlineSource.Clear
    cboInlineBase.Clear
    cboInlineBase.AddItem ""
    names = SourceTableChoices()
    If (LBound(names) <= UBound(names)) Then
        For i = LBound(names) To UBound(names)
            cboInlineSource.AddItem names(i)
            cboInlineBase.AddItem names(i)
        Next i
    End If
End Sub

Private Sub ApplyKindVisibility()
    Dim m As Boolean, s As Boolean, cn As Boolean, ci As Boolean, sg As Boolean
    Select Case cboTokenKind.ListIndex
        Case 0: m = True
        Case 1: s = True
        Case 2: cn = True
        Case 3: ci = True
        Case 4: sg = True
    End Select
    SetGroup Array(lblMetric, cboMetricSource, lblMetricSegment, cboMetricSegment, _
                   lblMetricType, cboMetricType, lblFlavour, optMetric, optCurrent, _
                   optDelta, optDeltaPct, optLabel, lblMetricLabel, _
                   lblMetricLabelVal, lblMetricGroup, lblMetricGroupVal, _
                   lblMetricFormat, lblMetricFormatVal, lblMetricSource, _
                   lblMetricSourceVal, lblMetricKey, lblMetricKeyVal), m
    SetGroup Array(lblScenarioCol, cboScenarioCol), s
    SetGroup Array(lblSegmentId, cboSegmentId, lblSegmentCol, cboSegmentCol, _
                   lblSegmentValue, lblSegmentValueVal), sg
    ' One format row serves every value panel, so it sits under whichever is showing.
    SetGroup Array(lblValueFormat, cboValueFormat, lblValueFormatHint), m Or s Or sg
    If s Then
        PlaceValueFormatRow 56
    ElseIf sg Then
        PlaceValueFormatRow cboSegmentId.Top + 3 * SEGMENT_ROW_STEP
    Else
        PlaceValueFormatRow 212
    End If
    SetGroup Array(lblChartNamed, cboChartNamed, lblChartNamedKind, _
                   lblChartNamedKindVal, lblChartNamedSource, _
                   lblChartNamedSourceVal, lblChartNamedSeries, _
                   lblChartNamedSeriesVal), cn
    SetGroup Array(lblInlineKind, cboInlineKind, lblInlineSource, _
                   cboInlineSource, lblInlineX, cboInlineX, lblInlineSeries, _
                   lstInlineSeries, lblInlineBase, cboInlineBase, _
                   lblInlineBaseSeries, lstInlineBaseSeries, lblInlineTitle, _
                   txtInlineTitle, lblInlineFormat, txtInlineFormat, _
                   chkInlineLegend, lblInlineHeight, txtInlineHeight, _
                   lblInlineWidth, txtInlineWidth), ci
End Sub

Private Sub PlaceValueFormatRow(ByVal rowTop As Single)
    lblValueFormat.Top = rowTop + 2
    cboValueFormat.Top = rowTop
    lblValueFormatHint.Top = rowTop + 2
End Sub

Private Sub SetGroup(controls As Variant, visible As Boolean)
    Dim i As Long
    For i = LBound(controls) To UBound(controls)
        controls(i).Visible = visible
    Next i
End Sub

Private Sub cboTokenKind_Change()
    ApplyKindVisibility
    RebuildPreview
End Sub

Private Sub cboMetricSource_Change()
    If suppressRefresh Then Exit Sub
    FillMetricSegments
    RefreshMetricReadout
    RebuildPreview
End Sub

Private Sub cboMetricSegment_Change()
    If suppressRefresh Then Exit Sub
    FillMetricTypes
    RefreshMetricReadout
    RebuildPreview
End Sub

Private Sub cboMetricType_Change()
    If suppressRefresh Then Exit Sub
    RefreshMetricReadout
    RebuildPreview
End Sub

' Blank until all three picks resolve to a catalogue entry.
Private Sub RefreshMetricReadout()
    Dim m As MetricSpec
    m = MetricById(CurrentMetricId())
    lblMetricLabelVal.Caption = m.label
    lblMetricGroupVal.Caption = m.groupName
    lblMetricFormatVal.Caption = m.numberFormat
    lblMetricSourceVal.Caption = m.metric
    lblMetricKeyVal.Caption = m.labelColumn
End Sub

Private Sub optMetric_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub optCurrent_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub optDelta_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub optDeltaPct_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub optLabel_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub

Private Sub cboScenarioCol_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub

Private Sub cboSegmentId_Change()
    If suppressRefresh Then Exit Sub
    RefreshSegmentReadout
    RebuildPreview
End Sub

Private Sub cboSegmentCol_Change()
    If suppressRefresh Then Exit Sub
    RefreshSegmentReadout
    RebuildPreview
End Sub

' What the token would show right now, read live from the segment map.
Private Sub RefreshSegmentReadout()
    lblSegmentValueVal.Caption = SegmentPreview(cboSegmentId.Value, cboSegmentCol.Value)
End Sub

Private Sub cboValueFormat_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub

Private Sub cboChartNamed_Change()
    If suppressRefresh Then Exit Sub
    RefreshNamedChartReadout
    RebuildPreview
End Sub

Private Sub RefreshNamedChartReadout()
    Dim specs() As ChartSpec, i As Long
    specs = NamedChartChoices()
    For i = LBound(specs) To UBound(specs)
        If StrComp(specs(i).label, cboChartNamed.Value, vbTextCompare) = 0 Then
            lblChartNamedKindVal.Caption = specs(i).kind
            lblChartNamedSourceVal.Caption = specs(i).sourceTable
            lblChartNamedSeriesVal.Caption = specs(i).seriesColumns
            Exit Sub
        End If
    Next i
    lblChartNamedKindVal.Caption = ""
    lblChartNamedSourceVal.Caption = ""
    lblChartNamedSeriesVal.Caption = ""
End Sub

Private Sub cboInlineSource_Change()
    RefreshInlineColumnLists False
    RebuildPreview
End Sub

Private Sub cboInlineBase_Change()
    RefreshInlineColumnLists True
    RebuildPreview
End Sub

Private Sub cboInlineKind_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub cboInlineX_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub lstInlineSeries_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub lstInlineBaseSeries_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub txtInlineTitle_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub txtInlineFormat_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub chkInlineLegend_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub txtInlineHeight_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub
Private Sub txtInlineWidth_Change()
    If Not suppressRefresh Then RebuildPreview
End Sub

Private Sub RefreshInlineColumnLists(baselineSide As Boolean)
    Dim names() As String, i As Long
    If Not baselineSide Then
        cboInlineX.Clear
        lstInlineSeries.Clear
        If Len(cboInlineSource.Value) > 0 Then
            names = TableColumnChoices(cboInlineSource.Value)
            For i = LBound(names) To UBound(names)
                cboInlineX.AddItem names(i)
                lstInlineSeries.AddItem names(i)
            Next i
        End If
    End If
    lstInlineBaseSeries.Clear
    If Len(cboInlineBase.Value) > 0 Then
        names = TableColumnChoices(cboInlineBase.Value)
        For i = LBound(names) To UBound(names)
            lstInlineBaseSeries.AddItem names(i)
        Next i
    End If
    PinInlineLists
End Sub

Private Sub PinInlineLists()
    RestoreBox lstInlineSeries, seriesListBox
    RestoreBox lstInlineBaseSeries, baseSeriesListBox
End Sub

Private Function CaptureBox(ctl As Object) As Variant
    CaptureBox = Array(ctl.Left, ctl.Top, ctl.Width, ctl.Height)
End Function

Private Sub RestoreBox(ctl As Object, box As Variant)
    ctl.Left = box(0)
    ctl.Top = box(1)
    ctl.Width = box(2)
    ctl.Height = box(3)
End Sub

Private Sub RebuildPreview()
    Dim token As String
    token = ComposeCurrentToken()
    txtPreview.text = token
    btnCopy.Caption = "Copy to clipboard"
    btnSave.Caption = "Save to library"
    btnSave.Enabled = (Len(token) > 0)
    btnCopy.Enabled = (Len(token) > 0)
    If Len(token) > 0 Then
        lblBanner.Caption = ValidateSavedToken(token)
    Else
        lblBanner.Caption = ""
    End If
    lblBanner.Visible = (Len(lblBanner.Caption) > 0)
End Sub

Private Function ComposeCurrentToken() As String
    Dim spec As ChartSpec
    Select Case cboTokenKind.ListIndex
        Case 0:  ComposeCurrentToken = ComposeMetricToken(SelectedFlavour, CurrentMetricId(), _
                                                          cboValueFormat.text)
        Case 1:  ComposeCurrentToken = ComposeScenarioToken(cboScenarioCol.Value, _
                                                            cboValueFormat.text)
        Case 2:  ComposeCurrentToken = ComposeChartNamedToken(cboChartNamed.Value)
        Case 3
            spec = CurrentInlineChartSpec()
            ComposeCurrentToken = ComposeChartInlineToken(spec)
        Case 4:  ComposeCurrentToken = ComposeSegmentToken(cboSegmentId.Value, _
                                                           cboSegmentCol.Value, _
                                                           cboValueFormat.text)
    End Select
End Function

Private Function SelectedFlavour() As String
    If optCurrent.Value Then SelectedFlavour = KIND_CURRENT: Exit Function
    If optDelta.Value Then SelectedFlavour = KIND_DELTA: Exit Function
    If optDeltaPct.Value Then SelectedFlavour = KIND_DELTA_PCT: Exit Function
    If optLabel.Value Then SelectedFlavour = KIND_LABEL: Exit Function
    SelectedFlavour = KIND_METRIC
End Function

Private Function CurrentInlineChartSpec() As ChartSpec
    Dim spec As ChartSpec
    spec.kind = cboInlineKind.Value
    spec.sourceTable = cboInlineSource.Value
    spec.xColumn = cboInlineX.Value
    spec.seriesColumns = JoinSelected(lstInlineSeries)
    spec.baselineTable = cboInlineBase.Value
    spec.baselineColumns = JoinSelected(lstInlineBaseSeries)
    spec.title = txtInlineTitle.text
    spec.numberFormat = txtInlineFormat.text
    spec.showLegend = chkInlineLegend.Value
    spec.heightRows = CLng(Val("0" & txtInlineHeight.text))
    spec.widthCols = CLng(Val("0" & txtInlineWidth.text))
    CurrentInlineChartSpec = spec
End Function

Private Function JoinSelected(lst As MSForms.ListBox) As String
    Dim i As Long, out As String
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If Len(out) > 0 Then out = out & ","
            out = out & lst.List(i)
        End If
    Next i
    JoinSelected = out
End Function

Private Sub btnCopy_Click()
    Dim ok As Boolean
    If Len(txtPreview.text) = 0 Then Exit Sub
    ok = PushClipboardText(txtPreview.text)
    If ok Then
        btnCopy.Caption = "Copied " & ChrW$(10003)
    Else
        btnCopy.Caption = "Clipboard failed"
    End If
    Application.OnTime Now + TimeValue("00:00:02"), "ResetCopyCaption"
End Sub

Public Sub ResetCopyCaption()
    btnCopy.Caption = "Copy to clipboard"
End Sub

Private Function PushClipboardText(text As String) As Boolean
    Dim dobj As Object
    On Error GoTo Failed
    Set dobj = New MSForms.DataObject
    dobj.SetText text
    dobj.PutInClipboard
    PushClipboardText = True
    Exit Function
Failed:
    PushClipboardText = False
End Function

Private Sub btnSave_Click()
    Dim label As String, kind As String
    If Len(txtPreview.text) = 0 Then Exit Sub
    label = InputBox("Name this token:", "Save token", DefaultLabel())
    If Len(Trim$(label)) = 0 Then Exit Sub
    kind = CurrentKindForStorage()
    editingTokenId = UpsertToken(editingTokenId, label, kind, txtPreview.text)
    btnSave.Caption = "Saved " & ChrW$(10003)
    Application.OnTime Now + TimeValue("00:00:02"), "ResetSaveCaption"
    RefreshSavedTokens
    btnCancel.Enabled = True
End Sub

Public Sub ResetSaveCaption()
    btnSave.Caption = "Save to library"
End Sub

Private Function CurrentKindForStorage() As String
    Select Case cboTokenKind.ListIndex
        Case 0: CurrentKindForStorage = SelectedFlavour
        Case 1: CurrentKindForStorage = KIND_SCENARIO
        Case 2: CurrentKindForStorage = KIND_CHART_NAMED
        Case 3: CurrentKindForStorage = KIND_CHART_INLINE
        Case 4: CurrentKindForStorage = KIND_SEGMENT
    End Select
End Function

Private Function DefaultLabel() As String
    Select Case cboTokenKind.ListIndex
        Case 0: DefaultLabel = MetricById(CurrentMetricId()).label & " (" & SelectedFlavour & ")"
        Case 1: DefaultLabel = "Scenario " & cboScenarioCol.Value
        Case 2: DefaultLabel = cboChartNamed.Value
        Case 3: DefaultLabel = cboInlineSource.Value & " chart"
        Case 4: DefaultLabel = "Segment " & cboSegmentId.Value & " " & cboSegmentCol.Value
    End Select
End Function

Private Sub btnCancel_Click()
    editingTokenId = ""
    btnCancel.Enabled = False
    ClearInputs
    RebuildPreview
End Sub

Private Sub btnClose_Click()
    editingTokenId = ""
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    StopWheelScroll
End Sub

' Mouse-wheel routing (see MouseWheel). Hovering a list or combo makes it the scroll
' target; hovering blank form space clears the target so the wheel does nothing there.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll Nothing, Me
End Sub
Private Sub cboTokenKind_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboTokenKind, Me
End Sub
Private Sub cboMetricSource_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboMetricSource, Me
End Sub
Private Sub cboMetricSegment_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboMetricSegment, Me
End Sub
Private Sub cboMetricType_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboMetricType, Me
End Sub
Private Sub cboScenarioCol_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboScenarioCol, Me
End Sub
Private Sub cboSegmentId_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboSegmentId, Me
End Sub
Private Sub cboSegmentCol_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboSegmentCol, Me
End Sub
Private Sub cboValueFormat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboValueFormat, Me
End Sub
Private Sub cboChartNamed_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboChartNamed, Me
End Sub
Private Sub cboInlineKind_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboInlineKind, Me
End Sub
Private Sub cboInlineSource_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboInlineSource, Me
End Sub
Private Sub cboInlineX_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboInlineX, Me
End Sub
Private Sub lstInlineSeries_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll lstInlineSeries, Me
End Sub
Private Sub cboInlineBase_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll cboInlineBase, Me
End Sub
Private Sub lstInlineBaseSeries_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll lstInlineBaseSeries, Me
End Sub
Private Sub lstSaved_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WheelScroll lstSaved, Me
End Sub

Private Sub ClearInputs()
    suppressRefresh = True
    cboMetricSource.ListIndex = -1
    cboMetricSegment.Clear
    cboMetricType.Clear
    cboScenarioCol.ListIndex = -1
    cboSegmentId.ListIndex = -1
    cboSegmentCol.ListIndex = -1
    cboValueFormat.text = ""
    cboChartNamed.ListIndex = -1
    cboInlineSource.ListIndex = -1
    cboInlineX.ListIndex = -1
    lstInlineSeries.Clear
    cboInlineBase.ListIndex = -1
    lstInlineBaseSeries.Clear
    PinInlineLists
    txtInlineTitle.text = ""
    txtInlineFormat.text = ""
    chkInlineLegend.Value = True
    txtInlineHeight.text = ""
    txtInlineWidth.text = ""
    RefreshMetricReadout
    RefreshNamedChartReadout
    RefreshSegmentReadout
    suppressRefresh = False
End Sub

Private Sub RefreshSavedTokens()
    Dim rows() As SavedToken, i As Long
    lstSaved.Clear
    If LibraryRowCount() = 0 Then Exit Sub
    rows = LibraryRows()
    lstSaved.ColumnCount = 3
    lstSaved.ColumnWidths = "150;80;240"
    For i = LBound(rows) To UBound(rows)
        lstSaved.AddItem rows(i).label
        lstSaved.List(lstSaved.ListCount - 1, 1) = rows(i).kind
        lstSaved.List(lstSaved.ListCount - 1, 2) = rows(i).token
    Next i
End Sub

Private Sub btnLibCopy_Click()
    DoLibCopy
End Sub

Private Sub lstSaved_DblClick(ByVal c As MSForms.ReturnBoolean)
    DoLibCopy
End Sub

Private Sub DoLibCopy()
    Dim id As String, row As SavedToken
    id = SelectedLibraryId()
    If Len(id) = 0 Then Exit Sub
    row = LibraryRowById(id)
    PushClipboardText row.token
    btnLibCopy.Caption = "Copied " & ChrW$(10003)
    Application.OnTime Now + TimeValue("00:00:02"), "ResetLibCopyCaption"
End Sub

Public Sub ResetLibCopyCaption()
    btnLibCopy.Caption = "Copy"
End Sub

Private Sub btnLibEdit_Click()
    Dim id As String, row As SavedToken, parsed As ParsedToken
    id = SelectedLibraryId()
    If Len(id) = 0 Then Exit Sub
    row = LibraryRowById(id)
    parsed = ParseTokenForEdit(row.token)
    LoadParsedIntoControls parsed
    editingTokenId = id
    btnCancel.Enabled = True
    lblBanner.Caption = ValidateSavedToken(row.token)
    lblBanner.Visible = (Len(lblBanner.Caption) > 0)
End Sub

Private Sub btnLibDelete_Click()
    Dim id As String
    id = SelectedLibraryId()
    If Len(id) = 0 Then Exit Sub
    If MsgBox("Delete the selected token?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    DeleteToken id
    RefreshSavedTokens
End Sub

Private Function SelectedLibraryId() As String
    Dim i As Long, rows() As SavedToken
    If lstSaved.ListIndex < 0 Then Exit Function
    If LibraryRowCount() = 0 Then Exit Function
    rows = LibraryRows()
    i = LBound(rows) + lstSaved.ListIndex
    If i >= LBound(rows) And i <= UBound(rows) Then SelectedLibraryId = rows(i).tokenId
End Function

Private Sub LoadParsedIntoControls(parsed As ParsedToken)
    suppressRefresh = True
    cboValueFormat.text = parsed.numberFormat
    Select Case parsed.kind
        Case KIND_METRIC, KIND_CURRENT, KIND_DELTA, KIND_DELTA_PCT, KIND_LABEL
            cboTokenKind.ListIndex = 0
            SelectMetricById parsed.metricId
            optMetric.Value = (parsed.kind = KIND_METRIC)
            optCurrent.Value = (parsed.kind = KIND_CURRENT)
            optDelta.Value = (parsed.kind = KIND_DELTA)
            optDeltaPct.Value = (parsed.kind = KIND_DELTA_PCT)
            optLabel.Value = (parsed.kind = KIND_LABEL)
        Case KIND_SCENARIO
            cboTokenKind.ListIndex = 1
            SelectComboValue cboScenarioCol, parsed.column
        Case KIND_SEGMENT
            cboTokenKind.ListIndex = 4
            SelectComboValue cboSegmentId, parsed.segmentId
            SelectComboValue cboSegmentCol, parsed.column
        Case KIND_CHART_NAMED
            cboTokenKind.ListIndex = 2
            SelectComboValue cboChartNamed, parsed.chartLabel
        Case KIND_CHART_INLINE
            cboTokenKind.ListIndex = 3
            SelectComboValue cboInlineKind, parsed.chartSpec.kind
            SelectComboValue cboInlineSource, parsed.chartSpec.sourceTable
            RefreshInlineColumnLists False
            SelectComboValue cboInlineX, parsed.chartSpec.xColumn
            SelectListBoxValues lstInlineSeries, parsed.chartSpec.seriesColumns
            SelectComboValue cboInlineBase, parsed.chartSpec.baselineTable
            RefreshInlineColumnLists True
            SelectListBoxValues lstInlineBaseSeries, parsed.chartSpec.baselineColumns
            txtInlineTitle.text = parsed.chartSpec.title
            txtInlineFormat.text = parsed.chartSpec.numberFormat
            chkInlineLegend.Value = parsed.chartSpec.showLegend
            txtInlineHeight.text = IIf(parsed.chartSpec.heightRows > 0, CStr(parsed.chartSpec.heightRows), "")
            txtInlineWidth.text = IIf(parsed.chartSpec.widthCols > 0, CStr(parsed.chartSpec.widthCols), "")
    End Select
    ApplyKindVisibility
    RefreshMetricReadout
    RefreshNamedChartReadout
    RefreshSegmentReadout
    suppressRefresh = False
    RebuildPreview
End Sub

Private Sub SelectComboValue(cbo As MSForms.ComboBox, value As String)
    Dim i As Long
    cbo.ListIndex = -1
    If Len(value) = 0 Then Exit Sub
    For i = 0 To cbo.ListCount - 1
        If StrComp(cbo.List(i), value, vbTextCompare) = 0 Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
    ' Value not in list — add it so the user can see what's broken.
    cbo.AddItem value
    cbo.ListIndex = cbo.ListCount - 1
End Sub

Private Sub SelectListBoxValues(lst As MSForms.ListBox, csv As String)
    Dim parts() As String, i As Long, j As Long
    If Len(csv) = 0 Then Exit Sub
    parts = Split(csv, ",")
    For j = LBound(parts) To UBound(parts)
        For i = 0 To lst.ListCount - 1
            If StrComp(lst.List(i), Trim$(parts(j)), vbTextCompare) = 0 Then
                lst.Selected(i) = True
                Exit For
            End If
        Next i
    Next j
End Sub
