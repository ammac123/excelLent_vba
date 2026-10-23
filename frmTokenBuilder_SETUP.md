# frmTokenBuilder — manual setup

The `.bas` modules (`ScenarioTokenBuilder`, `ScenarioTokenLibrary`,
`ScenarioOutput`, `ScenarioTokens`, etc.) already carry every bit of logic the
token builder needs. The only thing left is the `UserForm` itself, which the
code refers to as **`frmTokenBuilder`**.

This workbook can't materialise it programmatically because the "Trust access
to the VBA project object model" setting is locked off. So the form is built
once by hand in VBE following the steps below. After it exists and you save
the workbook, it ships with every `extract_bas.sh pack` run (both `.frm` and
`.frx` are now in the pack filter).

## Prerequisites

- All of the `.bas` files in this folder have been imported into the VBE
  project. In particular: `ScenarioTokenBuilder.bas`, `ScenarioTokenLibrary.bas`,
  `ScenarioTokens.bas`, `ScenarioOutput.bas`, `ScenarioStore.bas`,
  `ScenarioConfig.bas`, `ScenarioMetrics.bas`, `TableUtils.bas`.
- At least one Shape button somewhere in the workbook will be assigned to
  `OpenTokenBuilderButton`. You can add that before or after the form exists.

> **Already built the form before the value format box existed?** You don't
> need to rebuild it. Follow
> [Adding the value format row to an existing form](#adding-the-value-format-row-to-an-existing-form)
> instead.

## Layout at a glance

The form is a single page about 540×620 points. Panels overlap vertically —
only one panel is visible at a time, picked by the `Token kind:` combo at the
top. In VBE's designer you'll see every control stacked, which is fine; runtime
code toggles their `Visible` property.

```
+-----------------------------------------------------------------+
| Token kind: [ Metric value              v ]                     |
|                                                                 |
| [ One of four panels — see below ]                              |
|                                                                 |
| [ yellow banner, visible on validation issues ]                 |
|                                                                 |
| Preview:  [ composed token appears here       ]                 |
| [ Copy to clipboard ] [ Save to library ] [ Cancel ]  [ Close ] |
|                                                                 |
| Saved tokens:                                                   |
| [ listbox: Label | Kind | Token                               ] |
| [ Copy ] [ Edit ] [ Delete ]                                    |
+-----------------------------------------------------------------+
```

## Step 1 — Open VBE and insert the form

1. In Excel, press **Alt+F11** to open the Visual Basic Editor.
2. In the Project Explorer (left-hand tree), right-click the workbook's VBAProject
   node and choose **Insert → UserForm**.
3. Open the Properties window (**F4** if not visible) and set:
   - `(Name)` → `frmTokenBuilder`
   - `Caption` → `Token builder`
   - `Width` → `540`
   - `Height` → `620`

   If the designer snaps these to other values (e.g. 495.6), that's Windows
   display scaling. Ignore it — `UserForm_Initialize` sets the real size at
   runtime.

## Step 2 — Add the controls

Open the Toolbox (**View → Toolbox** in VBE) and drag the right type of control
onto the form for each row below. After placing one, click it, open the
Properties window, and set the properties listed. Positions are in points.

### Common pattern
- `(Name)` is the key field — the handler code references controls by this
  exact name. Case doesn't matter in VBA but keep it tidy.
- Set `Left`, `Top`, `Width`, `Height` to the four numbers in order.
- Set `Caption` when the row shows one. Blank captions are literally blank.

### Row 1 — Token kind picker (always visible)

| (Name)         | Type        | Left | Top | Width | Height | Caption          | Extra               |
|----------------|-------------|------|-----|-------|--------|------------------|---------------------|
| `lblKind`      | Label       | 10   | 10  | 70    | 15     | `Token kind:`    |                     |
| `cboTokenKind` | ComboBox    | 85   | 8   | 160   | 18     |                  |                     |

### Metric panel (shown for Metric / Current / Delta / DeltaPct / Label)

| (Name)                | Type         | Left | Top | Width | Height | Caption     | Extra                   |
|-----------------------|--------------|------|-----|-------|--------|-------------|-------------------------|
| `lblMetric`           | Label        | 10   | 36  | 70    | 15     | `Source:`      |                         |
| `cboMetricSource`     | ComboBox     | 85   | 34  | 240   | 18     |                | `Style` = `2 - fmStyleDropDownList` |
| `lblMetricSegment`    | Label        | 10   | 58  | 70    | 15     | `Segment:`     |                         |
| `cboMetricSegment`    | ComboBox     | 85   | 56  | 240   | 18     |                | `Style` = `2 - fmStyleDropDownList` |
| `lblMetricType`       | Label        | 10   | 80  | 70    | 15     | `Metric:`      |                         |
| `cboMetricType`       | ComboBox     | 85   | 78  | 240   | 18     |                | `Style` = `2 - fmStyleDropDownList` |
| `lblFlavour`          | Label        | 10   | 104 | 70    | 15     | `Flavour:`     |                         |
| `optMetric`           | OptionButton | 85   | 104 | 55    | 15     | `Metric`       | `GroupName` = `flav`    |
| `optCurrent`          | OptionButton | 140  | 104 | 65    | 15     | `Current`      | `GroupName` = `flav`    |
| `optDelta`            | OptionButton | 205  | 104 | 55    | 15     | `Delta`        | `GroupName` = `flav`    |
| `optDeltaPct`         | OptionButton | 260  | 104 | 70    | 15     | `DeltaPct`     | `GroupName` = `flav`    |
| `optLabel`            | OptionButton | 330  | 104 | 55    | 15     | `Label`        | `GroupName` = `flav`    |
| `lblMetricLabel`      | Label        | 10   | 126 | 70    | 15     | `Label:`       |                         |
| `lblMetricLabelVal`   | Label        | 85   | 126 | 440   | 15     |                |                         |
| `lblMetricGroup`      | Label        | 10   | 142 | 70    | 15     | `Group:`       |                         |
| `lblMetricGroupVal`   | Label        | 85   | 142 | 440   | 15     |                |                         |
| `lblMetricFormat`     | Label        | 10   | 158 | 70    | 15     | `Format:`      |                         |
| `lblMetricFormatVal`  | Label        | 85   | 158 | 440   | 15     |                |                         |
| `lblMetricSource`     | Label        | 10   | 174 | 70    | 15     | `Metric ID:`   |                         |
| `lblMetricSourceVal`  | Label        | 85   | 174 | 440   | 15     |                |                         |
| `lblMetricKey`        | Label        | 10   | 190 | 70    | 15     | `Segment col:` |                         |
| `lblMetricKeyVal`     | Label        | 85   | 190 | 440   | 15     |                |                         |

The metric is picked in three steps: **Source** (the table it comes from),
**Segment** (the row, e.g. a pricing segment ID) and **Metric** (the column,
e.g. NIM). Each list only shows what exists under the choice above it, and a
list with a single option fills itself in. Metrics that come from a named range
rather than a table appear under the source `(named ranges)`, and single-row
tables use the segment `(no segment)`.

The `GroupName = flav` setting makes those five option buttons mutually
exclusive. Make sure every one of them has that string in `GroupName`.

### Scenario-field panel (shown for Scenario field)

| (Name)            | Type     | Left | Top | Width | Height | Caption    |
|-------------------|----------|------|-----|-------|--------|------------|
| `lblScenarioCol`  | Label    | 10   | 36  | 70    | 15     | `Column:`  |
| `cboScenarioCol`  | ComboBox | 85   | 34  | 240   | 18     |            |

### Value format row (shown for Metric value and Scenario field)

These three controls are shared by both panels. The code moves them at runtime
so they sit directly under whichever panel is showing, so the `Top` values below
only matter in the designer.

| (Name)               | Type     | Left | Top | Width | Height | Caption                                | Extra |
|----------------------|----------|------|-----|-------|--------|----------------------------------------|-------|
| `lblValueFormat`     | Label    | 10   | 214 | 70    | 15     | `Value format:`                        |       |
| `cboValueFormat`     | ComboBox | 85   | 212 | 140   | 18     |                                        | `Style` = `0 - fmStyleDropDownCombo` (the default; it must stay editable) |
| `lblValueFormatHint` | Label    | 230  | 214 | 295   | 15     | `Blank = template / catalogue format`  |       |

### Chart (named) panel (shown for Chart (named))

| (Name)                     | Type     | Left | Top | Width | Height | Caption   |
|----------------------------|----------|------|-----|-------|--------|-----------|
| `lblChartNamed`            | Label    | 10   | 36  | 70    | 15     | `Chart:`  |
| `cboChartNamed`            | ComboBox | 85   | 34  | 240   | 18     |           |
| `lblChartNamedKind`        | Label    | 10   | 58  | 70    | 15     | `Kind:`   |
| `lblChartNamedKindVal`     | Label    | 85   | 58  | 240   | 15     |           |
| `lblChartNamedSource`      | Label    | 10   | 74  | 70    | 15     | `Source:` |
| `lblChartNamedSourceVal`   | Label    | 85   | 74  | 440   | 15     |           |
| `lblChartNamedSeries`      | Label    | 10   | 90  | 70    | 15     | `Series:` |
| `lblChartNamedSeriesVal`   | Label    | 85   | 90  | 440   | 15     |           |

### Chart (inline) panel (shown for Chart (inline))

| (Name)                 | Type     | Left | Top | Width | Height | Caption         | Extra                              |
|------------------------|----------|------|-----|-------|--------|-----------------|------------------------------------|
| `lblInlineKind`        | Label    | 10   | 36  | 70    | 15     | `Kind:`         |                                    |
| `cboInlineKind`        | ComboBox | 85   | 34  | 100   | 18     |                 |                                    |
| `lblInlineSource`      | Label    | 10   | 58  | 70    | 15     | `Source:`       |                                    |
| `cboInlineSource`      | ComboBox | 85   | 56  | 240   | 18     |                 |                                    |
| `lblInlineX`           | Label    | 10   | 80  | 70    | 15     | `X column:`     |                                    |
| `cboInlineX`           | ComboBox | 85   | 78  | 240   | 18     |                 |                                    |
| `lblInlineSeries`      | Label    | 10   | 102 | 70    | 15     | `Series:`       |                                    |
| `lstInlineSeries`      | ListBox  | 85   | 100 | 240   | 54     |                 | `MultiSelect` = `1 - fmMultiSelectMulti` |
| `lblInlineBase`        | Label    | 10   | 160 | 70    | 15     | `Baseline tbl:` |                                    |
| `cboInlineBase`        | ComboBox | 85   | 158 | 240   | 18     |                 |                                    |
| `lblInlineBaseSeries`  | Label    | 10   | 182 | 70    | 15     | `B. series:`    |                                    |
| `lstInlineBaseSeries`  | ListBox  | 85   | 180 | 240   | 54     |                 | `MultiSelect` = `1 - fmMultiSelectMulti` |
| `lblInlineTitle`       | Label    | 10   | 240 | 70    | 15     | `Title:`        |                                    |
| `txtInlineTitle`       | TextBox  | 85   | 238 | 240   | 18     |                 |                                    |
| `lblInlineFormat`      | Label    | 10   | 262 | 70    | 15     | `Format:`       |                                    |
| `txtInlineFormat`      | TextBox  | 85   | 260 | 100   | 18     |                 |                                    |
| `chkInlineLegend`      | CheckBox | 195  | 262 | 70    | 15     | `Legend`        | `Value` = `True`                   |
| `lblInlineHeight`      | Label    | 10   | 284 | 70    | 15     | `Height:`       |                                    |
| `txtInlineHeight`      | TextBox  | 85   | 282 | 50    | 18     |                 |                                    |
| `lblInlineWidth`       | Label    | 150  | 284 | 50    | 15     | `Width:`        |                                    |
| `txtInlineWidth`       | TextBox  | 205  | 282 | 50    | 18     |                 |                                    |

### Banner, preview, action buttons (always visible)

| (Name)        | Type           | Left | Top | Width | Height | Caption               | Extra                   |
|---------------|----------------|------|-----|-------|--------|-----------------------|-------------------------|
| `lblBanner`   | Label          | 10   | 340 | 515   | 24     |                       | `BackColor` = `&H00C0FFFF` (pale yellow), `Visible` = `False` |
| `lblPreview`  | Label          | 10   | 372 | 70    | 15     | `Preview:`            |                         |
| `txtPreview`  | TextBox        | 85   | 370 | 440   | 18     |                       | `Locked` = `True` (optional, prevents accidental typing) |
| `btnCopy`     | CommandButton  | 10   | 400 | 110   | 24     | `Copy to clipboard`   |                         |
| `btnSave`     | CommandButton  | 125  | 400 | 100   | 24     | `Save to library`     |                         |
| `btnCancel`   | CommandButton  | 230  | 400 | 70    | 24     | `Cancel`              | `Enabled` = `False`     |
| `btnClose`    | CommandButton  | 455  | 400 | 70    | 24     | `Close`               |                         |

### Saved tokens section (always visible)

| (Name)          | Type           | Left | Top | Width | Height | Caption          |
|-----------------|----------------|------|-----|-------|--------|------------------|
| `lblSaved`      | Label          | 10   | 432 | 100   | 15     | `Saved tokens:`  |
| `lstSaved`      | ListBox        | 10   | 450 | 515   | 90     |                  |
| `btnLibCopy`    | CommandButton  | 10   | 548 | 70    | 24     | `Copy`           |
| `btnLibEdit`    | CommandButton  | 85   | 548 | 70    | 24     | `Edit`           |
| `btnLibDelete`  | CommandButton  | 160  | 548 | 70    | 24     | `Delete`         |

## Step 3 — Paste the code module

1. In the Project Explorer, double-click `frmTokenBuilder` under *Forms* to open
   its code module (or press **F7** while the form is selected).
2. Delete anything VBE pre-populated.
3. Paste the whole block below.

```vba
Option Explicit

' Event handlers and view state for the token builder form. All compose /
' parse / validate / lookup logic lives in ScenarioTokenBuilder.

Private editingTokenId As String
Private suppressRefresh As Boolean

Private Sub UserForm_Initialize()
    ' Set at runtime because the VBE designer rescales these under Windows
    ' display scaling (e.g. 125%) and won't keep typed values.
    Me.Width = 540
    Me.Height = 620
    ApplyTabOrder
    PopulateKindCombo
    PopulateMetricSourceCombo
    PopulateScenarioCombo
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
    Dim m As Boolean, s As Boolean, cn As Boolean, ci As Boolean
    Select Case cboTokenKind.ListIndex
        Case 0: m = True
        Case 1: s = True
        Case 2: cn = True
        Case 3: ci = True
    End Select
    SetGroup Array(lblMetric, cboMetricSource, lblMetricSegment, cboMetricSegment, _
                   lblMetricType, cboMetricType, lblFlavour, optMetric, optCurrent, _
                   optDelta, optDeltaPct, optLabel, lblMetricLabel, _
                   lblMetricLabelVal, lblMetricGroup, lblMetricGroupVal, _
                   lblMetricFormat, lblMetricFormatVal, lblMetricSource, _
                   lblMetricSourceVal, lblMetricKey, lblMetricKeyVal), m
    SetGroup Array(lblScenarioCol, cboScenarioCol), s
    ' One format row serves both value panels, so it sits under whichever is showing.
    SetGroup Array(lblValueFormat, cboValueFormat, lblValueFormatHint), m Or s
    If s Then PlaceValueFormatRow 56 Else PlaceValueFormatRow 212
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
    End Select
End Function

Private Function DefaultLabel() As String
    Select Case cboTokenKind.ListIndex
        Case 0: DefaultLabel = MetricById(CurrentMetricId()).label & " (" & SelectedFlavour & ")"
        Case 1: DefaultLabel = "Scenario " & cboScenarioCol.Value
        Case 2: DefaultLabel = cboChartNamed.Value
        Case 3: DefaultLabel = cboInlineSource.Value & " chart"
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

Private Sub ClearInputs()
    suppressRefresh = True
    cboMetricSource.ListIndex = -1
    cboMetricSegment.Clear
    cboMetricType.Clear
    cboScenarioCol.ListIndex = -1
    cboValueFormat.text = ""
    cboChartNamed.ListIndex = -1
    cboInlineSource.ListIndex = -1
    cboInlineX.ListIndex = -1
    lstInlineSeries.Clear
    cboInlineBase.ListIndex = -1
    lstInlineBaseSeries.Clear
    txtInlineTitle.text = ""
    txtInlineFormat.text = ""
    chkInlineLegend.Value = True
    txtInlineHeight.text = ""
    txtInlineWidth.text = ""
    RefreshMetricReadout
    RefreshNamedChartReadout
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
```

## Step 4 — Save and test

1. Press **Ctrl+S** in VBE (or save from Excel). The form is now part of the
   workbook.
2. In Excel, add a Shape button on any sheet and assign the macro
   `OpenTokenBuilderButton`. Right-click the shape → Assign Macro → pick it.
3. Click the button. The form should appear modelessly. Compose a metric
   token, click **Copy to clipboard**, confirm the caption flashes `Copied ✓`
   for two seconds, paste into a template cell on another sheet to verify.
4. **File → Export File** on `frmTokenBuilder` in VBE and commit the generated
   `frmTokenBuilder.frm` + `frmTokenBuilder.frx` so future users get the form
   for free when they pull the latest pack.

## Adding the value format row to an existing form

Use this if `frmTokenBuilder` already exists in your workbook from an earlier
version of this guide. You only add three controls and swap the code; nothing
else on the form changes.

1. Import the updated `ScenarioTokens.bas` and `ScenarioTokenBuilder.bas`
   (remove the old modules first: right-click → **Remove**, answer **No** to
   exporting, then **File → Import File**).
2. In the Project Explorer, double-click `frmTokenBuilder` under *Forms* to
   open the designer.
3. From the Toolbox, add the three controls in
   [Value format row](#value-format-row-shown-for-metric-value-and-scenario-field)
   and set their properties exactly as listed. The names matter; the positions
   are only a starting point because the code moves the row at runtime.
   - If the Metric panel's controls overlap where the row goes, that's fine —
     panels are stacked in the designer anyway.
4. Press **F7** to open the form's code, select everything, delete it, and paste
   the code block from [Step 3](#step-3--paste-the-code-module).
5. **Debug → Compile VBAProject.** It should compile without errors. If it
   stops on `cboValueFormat`, `lblValueFormat` or `lblValueFormatHint`, that
   control's `(Name)` is misspelled (see Troubleshooting).
6. Save, then **File → Export File** on `frmTokenBuilder` so the new `.frm` /
   `.frx` pair goes into the pack.

## Using the value format box

The **Value format** box adds an Excel number format to the token, e.g.
`{Metric.MetricsOutSegment_42_NIM|0.00%}`. It appears for **Metric value**
(all five flavours) and **Scenario field** tokens; charts keep their own
`Format:` box on the inline panel.

- **Pick a preset or type your own.** The list holds common formats (`0.00%`,
  `#,##0`, `#,##0;(#,##0);-`, `dd-mmm-yy`, …). Anything you could type into
  Excel's *Format Cells → Custom* box works, including `;` sections and quoted
  text like `0.0"x"`. The one character you can't use is `}`, since it ends
  the token. The yellow banner warns if you try.
- **Leave it blank** for the original behaviour: the template cell's own format
  wins, and the catalogue's `NumberFormat` only fills in cells left as General.
- **The `Format:` line in the Metric readout** shows the catalogue format for
  the selected metric. That's what you get with the box blank, so use it as a
  starting point when you want something slightly different.
- **Editing a saved token** puts its format back in the box, and **Save to
  library** stores the token with its format.

Where the token lands decides how the format is used:

| Template cell contains                  | Result |
|-----------------------------------------|--------|
| Only the token: `{Metric.X\|0.00%}`     | The cell gets the format; the value stays a number. Overrides the template cell's own format and the catalogue. |
| A formula: `={Metric.X\|0.00%}*1`       | The value stays a number so the formula still works; the cell gets the format. If several tokens in one formula have formats, the first one wins. |
| Text: `NIM: {Metric.X\|0.00%}`          | The value is written as text through Excel's `TEXT()`, e.g. `NIM: 2.35%`. |

If Excel rejects a format, the token behaves as though the box were blank.

## Troubleshooting

- **Compile error "Variable not defined" on `cboValueFormat`,
  `lblValueFormat` or `lblValueFormatHint`.** The form code is the new version
  but the controls haven't been added, or a `(Name)` is misspelled. Add them as
  described in
  [Adding the value format row to an existing form](#adding-the-value-format-row-to-an-existing-form).
- **Compile error "Wrong number of arguments" on `ComposeMetricToken` or
  `ComposeScenarioToken`.** The form code is new but `ScenarioTokenBuilder.bas`
  is the old one. Re-import it.
- **The value format box won't let you type.** Its `Style` was set to
  `2 - fmStyleDropDownList`. The code resets it to an editable combo on open,
  but set it back to `0 - fmStyleDropDownCombo` in the designer too.

- **Compile error on `MSForms.DataObject` or `MSForms.ListBox`.** Your project
  may not have a reference to Microsoft Forms 2.0. In VBE: Tools → References,
  tick **Microsoft Forms 2.0 Object Library**. Normally it's added
  automatically when you insert a UserForm, but if not, adding it manually
  fixes compile errors in the clipboard and multi-select code.
- **"Can't find project or library" on `CatalogueChoices` / `MetricById` /
  `ComposeMetricToken` etc.** The `.bas` modules weren't imported. See
  prerequisites.
- **Form opens but combos are empty.** The catalogue table / tblScenarios /
  tblMetricsOut* may not exist in your workbook yet. Nothing crashes — the
  combos just show empty lists.
- **Copy button flashes "Clipboard failed".** You're on Mac Excel or an
  environment where `MSForms.DataObject` isn't usable. The composed token is
  still visible in the Preview textbox; select and Ctrl+C it manually.
- **Yellow banner appears on a token you expect to be valid.** The banner
  text names what's missing (`"Metric 'XYZ' is no longer in the catalogue."`
  etc.). Either the catalogue entry was removed or renamed, or the referenced
  table/column was. The composed token still copies either way — the banner
  is informational.
- **Someone else has "Trust access to the VBA project object model" enabled.**
  They can build the form in VBE on their machine the normal way, then
  **File → Export File** on `frmTokenBuilder` and commit the resulting
  `frmTokenBuilder.frm` + `frmTokenBuilder.frx`. Anyone pulling the latest
  pack after that gets the form for free and doesn't have to repeat these
  steps.
