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
| [ One of five panels — see below ]                              |
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

A table with row keys also lists `(all segments)` at the top of the Segment
list. That composes a wildcard token such as `{Metric.MetricsOutSegment_*_NIM}`:
the template row holding it repeats once per segment when the sheet is built
(see the "Wildcard rows" notes at the top of `ScenarioTokens.bas`). Patterns
that cover only part of the key, like `t1*`, are typed into the cell by hand;
Edit on a saved one shows the pattern in the Segment list. No designer change
is needed for any of this.

The `GroupName = flav` setting makes those five option buttons mutually
exclusive. Make sure every one of them has that string in `GroupName`.

### Scenario-field panel (shown for Scenario field)

| (Name)            | Type     | Left | Top | Width | Height | Caption    |
|-------------------|----------|------|-----|-------|--------|------------|
| `lblScenarioCol`  | Label    | 10   | 36  | 70    | 15     | `Column:`  |
| `cboScenarioCol`  | ComboBox | 85   | 34  | 240   | 18     |            |

### Segment attribute panel (shown for Segment attribute)

Nothing to add in the designer. The form builds these controls itself when it
opens (`BuildSegmentPanel`), copying the position, size and font of
`lblScenarioCol` / `cboScenarioCol`, so they line up with the Scenario-field
panel:

| Created in code      | Type     | Row                    | Caption        |
|----------------------|----------|------------------------|----------------|
| `lblSegmentId`       | Label    | same row as `Column:`  | `Segment ID:`  |
| `cboSegmentId`       | ComboBox | same row as `Column:`  |                |
| `lblSegmentCol`      | Label    | one row down           | `Column:`      |
| `cboSegmentCol`      | ComboBox | one row down           |                |
| `lblSegmentValue`    | Label    | two rows down          | `Value now:`   |
| `lblSegmentValueVal` | Label    | two rows down          | live value     |

Don't add controls with these names in the designer; `Controls.Add` would fail
on the duplicate name.

`cboSegmentId` starts with `(all segments)`, which composes
`{Segment.*.<Column>}`. A row holding it repeats once per segment, in the same
order as a `(all segments)` metric row, so names and metrics line up. The
`Value now:` readout then shows the first row's value and how many rows the
token fills.

### Value format row (shown for Metric value, Scenario field and Segment attribute)

These three controls are shared by all three value panels. The code moves them at runtime
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
| `lblInlineBase`        | Label    | 10   | 160 | 70    | 15     | `Baseline tbl:` | Hidden at runtime (see below)      |
| `cboInlineBase`        | ComboBox | 85   | 158 | 240   | 18     |                 | Hidden at runtime (see below)      |
| `lblInlineBaseSeries`  | Label    | 10   | 182 | 70    | 15     | `B. series:`    | Renamed `Vs current:` at runtime   |
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

Charts plot stored scenario results, and their baseline is always the current
state's stored values for the same table. So the code hides the baseline-table
combo when the form opens and puts the include row (below) in its place. Keep
`lblInlineBase` and `cboInlineBase` on the form, because the code refers to them.
The two series lists get a second column, set in code, showing each column's
include status.

### Include row (Metric value and Chart (inline) panels)

Nothing to add in the designer. The form builds these when it opens
(`BuildIncludeRow`) and moves them under whichever of the two panels is showing:
row 208 for Metric value, where the value format row then drops to 232, and
row 158 for Chart (inline).

| Created in code       | Type          | Caption    |
|-----------------------|---------------|------------|
| `lblIncludeStatus`    | Label         | `Status:`  |
| `lblIncludeStatusVal` | Label         | the pick's include status |
| `btnInclude`          | CommandButton | `Include`  |

Don't add controls with these names in the designer.

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

' The include row is built in code too, and shared by the metric and inline chart panels:
' where the pick stands (ticked Include? run?) and a button that ticks what's missing.
Private Const INCLUDE_ROW_METRIC As Single = 208
Private Const INCLUDE_ROW_CHART As Single = 158
Private WithEvents btnInclude As MSForms.CommandButton
Private lblIncludeStatus As MSForms.Label
Private lblIncludeStatusVal As MSForms.Label

Private Sub UserForm_Initialize()
    ' Set at runtime because the VBE designer rescales these under Windows
    ' display scaling (e.g. 125%) and won't keep typed values.
    Me.Width = 540
    Me.Height = 620
    seriesListBox = CaptureBox(lstInlineSeries)
    baseSeriesListBox = CaptureBox(lstInlineBaseSeries)
    ForgetStoredMetrics
    BuildSegmentPanel
    BuildIncludeRow
    PrepareInlineChartPanel
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
                  btnInclude, lstInlineBaseSeries, txtInlineTitle, _
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

Private Sub BuildIncludeRow()
    Set lblIncludeStatus = AddPanelLabel("lblIncludeStatus", "Status:", 0)
    Set lblIncludeStatusVal = AddPanelLabel("lblIncludeStatusVal", "", 0)
    lblIncludeStatusVal.Left = cboScenarioCol.Left
    lblIncludeStatusVal.Width = 150
    Set btnInclude = Me.Controls.Add("Forms.CommandButton.1", "btnInclude")
    btnInclude.Left = cboScenarioCol.Left + 160
    btnInclude.Width = 80
    btnInclude.Height = cboScenarioCol.Height
    btnInclude.Font.Name = cboScenarioCol.Font.Name
    btnInclude.Font.Size = cboScenarioCol.Font.Size
    btnInclude.Caption = "Include"
    btnInclude.ControlTipText = "Tick Include in the metric catalogue for everything this pick needs"
End Sub

Private Sub PlaceIncludeRow(ByVal rowTop As Single)
    lblIncludeStatus.Top = rowTop + (lblScenarioCol.Top - cboScenarioCol.Top)
    lblIncludeStatusVal.Top = lblIncludeStatus.Top
    btnInclude.Top = rowTop
End Sub

' Charts plot stored results, whose baseline is always the current state's values for
' the same source, so the designer's baseline-table combo is retired and its series
' list becomes "compare with the current state" over the same columns. The include row
' takes the combo's place. Both series lists show each column's include status.
Private Sub PrepareInlineChartPanel()
    lblInlineBase.Visible = False
    cboInlineBase.Visible = False
    lblInlineBaseSeries.Caption = "Vs current:"
    lstInlineSeries.ColumnCount = 2
    lstInlineSeries.ColumnWidths = "120;110"
    lstInlineBaseSeries.ColumnCount = 2
    lstInlineBaseSeries.ColumnWidths = "120;110"
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

' Only tables the metric catalogue scans: nothing else has stored results to plot.
Private Sub PopulateSourceCombos()
    FillCombo cboInlineSource, ChartSourceChoices()
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
        PlaceValueFormatRow INCLUDE_ROW_METRIC + SEGMENT_ROW_STEP + 2
    End If
    SetGroup Array(lblIncludeStatus, lblIncludeStatusVal, btnInclude), m Or ci
    If ci Then PlaceIncludeRow INCLUDE_ROW_CHART Else PlaceIncludeRow INCLUDE_ROW_METRIC
    SetGroup Array(lblChartNamed, cboChartNamed, lblChartNamedKind, _
                   lblChartNamedKindVal, lblChartNamedSource, _
                   lblChartNamedSourceVal, lblChartNamedSeries, _
                   lblChartNamedSeriesVal), cn
    SetGroup Array(lblInlineKind, cboInlineKind, lblInlineSource, _
                   cboInlineSource, lblInlineX, cboInlineX, lblInlineSeries, _
                   lstInlineSeries, _
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

' Switching kind or source re-reads tblScenarioResults, so a run made while the form
' was open shows up in the include status.
Private Sub cboTokenKind_Change()
    ForgetStoredMetrics
    ApplyKindVisibility
    RebuildPreview
End Sub

Private Sub cboMetricSource_Change()
    If suppressRefresh Then Exit Sub
    ForgetStoredMetrics
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
    ForgetStoredMetrics
    RefreshInlineColumnLists
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

' X defaults to the source's row-key column. Both series lists carry the same catalogued
' columns, each with its include status alongside.
Private Sub RefreshInlineColumnLists()
    Dim prior As Boolean
    Dim names() As String, statuses() As String, i As Long
    prior = suppressRefresh
    suppressRefresh = True
    cboInlineX.Clear
    lstInlineSeries.Clear
    lstInlineBaseSeries.Clear
    If Len(cboInlineSource.Value) > 0 Then
        FillCombo cboInlineX, ChartXChoices(cboInlineSource.Value)
        SelectComboValue cboInlineX, ChartRowKeyColumn(cboInlineSource.Value)
        names = ChartSeriesChoices(cboInlineSource.Value)
        statuses = ChartColumnStatuses(cboInlineSource.Value)
        For i = LBound(names) To UBound(names)
            lstInlineSeries.AddItem names(i)
            lstInlineSeries.List(lstInlineSeries.ListCount - 1, 1) = statuses(i)
            lstInlineBaseSeries.AddItem names(i)
            lstInlineBaseSeries.List(lstInlineBaseSeries.ListCount - 1, 1) = statuses(i)
        Next i
    End If
    PinInlineLists
    suppressRefresh = prior
End Sub

' Rewrites the status column in place, so the user's ticks survive an Include.
Private Sub RefreshInlineColumnStatuses()
    Dim statuses() As String, i As Long
    If Len(cboInlineSource.Value) = 0 Then Exit Sub
    statuses = ChartColumnStatuses(cboInlineSource.Value)
    For i = 0 To lstInlineSeries.ListCount - 1
        If i + 1 > UBound(statuses) Then Exit For
        lstInlineSeries.List(i, 1) = statuses(i + 1)
        lstInlineBaseSeries.List(i, 1) = statuses(i + 1)
    Next i
End Sub

' Status of whatever the visible panel points at. Blank for a Label token, which reads
' the catalogue rather than a stored result.
Private Sub RefreshIncludeStatus()
    Dim status As String
    Select Case cboTokenKind.ListIndex
        Case 0
            If SelectedFlavour <> KIND_LABEL Then status = MetricIncludeStatus(CurrentMetricId())
        Case 3
            status = ChartIncludeStatus(cboInlineSource.Value, InlineColumnsPicked())
    End Select
    lblIncludeStatusVal.Caption = status
    btnInclude.Enabled = StatusNeedsInclude(status)
End Sub

' Every column ticked in either series list.
Private Function InlineColumnsPicked() As String
    InlineColumnsPicked = JoinSelected(lstInlineSeries) & "," & JoinSelected(lstInlineBaseSeries)
End Function

Private Sub btnInclude_Click()
    Dim changed As Long
    On Error GoTo Failed
    Select Case cboTokenKind.ListIndex
        Case 0: changed = IncludeMetricPick(CurrentMetricId())
        Case 3: changed = IncludeChartPick(cboInlineSource.Value, InlineColumnsPicked())
    End Select
    RefreshInlineColumnStatuses
    RebuildPreview
    MsgBox changed & " metric(s) ticked Include in the catalogue." & vbNewLine & vbNewLine & _
           "Run the scenarios (and the current state, for Current / Delta and " & _
           "vs-current series) to store their values.", vbInformation
    Exit Sub
Failed:
    MsgBox "Couldn't tick the metrics in the catalogue." & vbNewLine & vbNewLine & _
           Err.Description, vbExclamation
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
    RefreshIncludeStatus
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
            out = out & lst.List(i, 0)
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
    cboInlineX.Clear
    lstInlineSeries.Clear
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
            RefreshInlineColumnLists
            SelectComboValue cboInlineX, parsed.chartSpec.xColumn
            SelectListBoxValues lstInlineSeries, parsed.chartSpec.seriesColumns
            ' An older token's baselineSource is dropped: stored charts never read it.
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
            If StrComp(lst.List(i, 0), Trim$(parts(j)), vbTextCompare) = 0 Then
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
   [Value format row](#value-format-row-shown-for-metric-value-scenario-field-and-segment-attribute)
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
(all five flavours), **Scenario field** and **Segment attribute** tokens;
charts keep their own `Format:` box on the inline panel.

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

## Building tables and charts from scenario metrics

Everything a template shows for a scenario comes from `tblScenarioResults`, not
from the live model tables. A number only gets there when all three of these are true:

1. **It's in the catalogue.** `tblScenarioMetrics` has one row per (table, row,
   column) cell of each table listed in `MetricSourceTables()` in `ScenarioConfig`.
   If a table or column is missing, add it there and press **Sync metrics**.
2. **It's ticked Include** in `tblScenarioMetrics`.
3. **The scenario has been run since it was ticked.** A run stores only included
   metrics. `{Current.X}`, `{Delta.X}` and a chart's vs-current series also need
   the current state to have been run.

The token builder shows where each pick stands:

| Status              | Meaning                                         | What to do |
|---------------------|-------------------------------------------------|------------|
| `ready`             | Included and stored                             | Nothing    |
| `not included`      | Nothing behind the pick is ticked               | Press **Include**, then run the scenarios |
| `3/12 included`     | Only some rows are ticked                       | Press **Include**, then run the scenarios |
| `included, not run` | Ticked, but no run has stored it yet            | Run the scenarios |
| `included, 5/12 run`| Some ticked rows were added after the last run  | Run the scenarios |

**Include** ticks exactly the catalogue rows behind the current pick. It doesn't
run anything, so run the scenarios afterwards. The yellow banner repeats the
advice in words. "Run" counts a metric that has been stored for any scenario,
so a scenario that hasn't been run since still shows gaps.

### A table

There's no separate table token. A table is metric tokens laid out across
template rows.

1. **Token kind:** `Metric value`. Pick **Source**, then **Segment**, then **Metric**.
2. For one row per segment, choose `(all segments)` as the segment. That gives
   a wildcard token like `{Metric.MetricsOutSegment_*_NIM}`, and the template
   row holding it repeats once per segment.
3. Build one token per column (NIM, Fees, ...) and put them side by side in the
   same template row. A `Segment attribute` token with `(all segments)` in the
   first column labels each row.
4. Check each token shows `ready` before you paste it.

### A chart

1. **Token kind:** `Chart (inline)`.
2. **Source** only lists tables the catalogue scans, because nothing else has
   stored data to plot.
3. **X column** defaults to the table's row-key column (one point per row).
4. Tick one or more **Series** columns. The status beside each column says
   whether it's ready.
5. Optionally tick columns in **Vs current** to draw the current state's
   values alongside.
6. If the status isn't `ready`, press **Include** and run the scenarios.
7. Copy the token into a cell of the output or comparison template. Merge the
   area you want the chart to fill.

The chart redraws from the picked scenario's stored results every time the
output sheet (or each comparison block) is built.

## Segment attribute tokens

**Segment attribute** in the `Token kind:` list builds a token that reads one
cell of the segment map, `tblProductMap`, keyed by `Pricing_Segment_ID`:

```
{Segment.42.Segment Description}
{Segment.42.LVR_Band|0%}
```

- **Segment ID** lists every ID in the table, in the table's own row order.
- **Column** lists every other column of the table. Each column is an attribute
  of the segment.
- **Value now** shows what the token would render right now.
- **Value format** works the same as for metric and scenario tokens.

The token isn't a metric. It's never stored in `tblScenarioResults`, it's the
same in every scenario's block, and it's read from the table each time a
template renders, so edits to `tblProductMap` show up on the next build. A
render reads the table once per block, however many segment tokens the block
holds.

The table and key column are set in `ScenarioConfig`
(`SegmentMapTableName`, `SegmentMapKeyColumn`). An ID or column that isn't in
the table renders as `n/a` (`NA()` inside a formula), and the yellow banner
names what's missing.

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
