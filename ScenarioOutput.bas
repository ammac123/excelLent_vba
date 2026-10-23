Attribute VB_Name = "ScenarioOutput"
Option Explicit

' Detailed single-scenario output sheet. Same template-and-token mechanism as
' ScenarioCompare, but renders one block for one picked scenario instead of a row of
' blocks, and also understands chart tokens so a template can declare charts in cells
' without any VBA edit.
'
'=======================================================================================
' What the user controls on the sheet
'=======================================================================================
'
'   ScenarioOutputTemplates()  Named blocks declared in ScenarioConfig, same contract as
'                              CompareTemplates(). Each is a defined name over a block
'                              laid out and formatted by hand.
'   outputTemplatePicker       Cell holding the chosen template's label.
'   outputScenarioPicker       Cell holding the scenario's display name. Empty, "Loaded
'                              (<name>)" or an unknown name falls back to the currently loaded
'                              scenario, so a user who hasn't thought about it still
'                              gets what the model is calculating.
'   outputAnchor               Top-left cell the block lands on.
'   cmbRedefineOutputTemplate  Form-control combo box that pairs with the Redefine
'                              Output Template Range button.
'
'=======================================================================================
' Token grammar
'=======================================================================================
'
' Everything the comparison supports works here: {Scenario.X}, {Metric.X}, {Current.X},
' {Delta.X}, {DeltaPct.X}, {Label.X}. See ScenarioTokens for the full list. Because
' this sheet renders one scenario, {Scenario.Index} is always 1.
'
' In addition, two chart tokens are recognised in cells whose entire contents are the
' token (not embedded in text, not inside formulas):
'
'   {Chart:<Name>}        Drops a chart from ScenarioOutputCharts() where <Name> is the
'                         spec's label. Case-insensitive.
'   {Chart:k=v;k=v;...}   Drops a chart from an inline spec. Recognised keys:
'                           kind              Line, Column, Bar, Combo (default Line)
'                           source            table name (required)
'                           x                 x-axis column (required)
'                           series            comma-separated series columns (required)
'                           baselineSource    optional parallel table
'                           baselineSeries    comma-separated series from baselineSource
'                           title             chart title (default = spec label or
'                                             "<source>")
'                           format            value-axis number format, e.g. 0.0%
'                           legend            Y/N (default Y)
'                           height            chart height in rows at the anchor cell
'                           width             chart width in columns at the anchor cell
'                         Semicolons separate pairs, equals separates key and value, and
'                         commas separate items within a list. Keys are case-insensitive.
'
' Chart size: height/width when the spec sets them, otherwise the token cell's merged
' area, otherwise 15 rows by 8 columns. On this sheet the chart is drawn at that size
' even when it runs past the template's edge. The comparison sheet renders the same
' tokens but shrinks each chart to fit inside its scenario block, and plots stored
' results instead of the live tables (see ScenarioCompare).
'
' A cell holding a chart token is cleared before the chart is anchored, so the text
' never shows through. Any ChartObject this module creates is named "scnOut_<spec>_<n>"
' so a subsequent build can find and delete it cleanly. A chart that cannot be built is
' skipped and listed in the closing message; the rest of the output still renders.
'
' Merged cells are safe to use for chart areas, including one that runs past the
' template's defined range: the template is widened to take in every merged area it
' touches before it is copied, because copying or clearing part of a merged area fails.
'
'=======================================================================================
' Rebuild semantics
'=======================================================================================
'
' Everything this module writes is disposable. On rebuild it:
'   1. Deletes every ChartObject on the sheet whose name starts with "scnOut_".
'   2. Clears the previously recorded rectangle on the sheet.
'   3. Copies the picked template to the anchor.
'   4. Fills tokens for the picked scenario.
'   5. Walks the block for chart tokens and creates ChartObjects.
'   6. Records the new extent.
'
' So the output can never drift out of step with tblScenarioResults or the live
' tblMetricsOut* tables that charts bind to.

Private Const EXTENT_KEY As String = "OutputExtent"
Private Const CHART_NAME_PREFIX As String = "scnOut_"
Private Const CHART_TOKEN_PREFIX As String = "Chart:"
Private Const LOADED_SENTINEL As String = "Loaded"
Private Const STAGING_SHEET As String = "ChartStaging"

'=======================================================================================
' Entry point
'=======================================================================================

Public Sub BuildScenarioOutput()
    Dim ws As Worksheet
    Dim chosen As CompareTemplateSpec
    Dim template As Range, anchor As Range, dest As Range
    Dim scenarioId As String
    Dim results As Object, labels As Object, formats As Object, groups As Object
    Dim guard As CalcGuard
    Dim chartProblems As String

    On Error GoTo Failed

    RefreshOutputScenarioPicker
    RefreshOutputTemplatePicker

    scenarioId = ResolveOutputScenarioId()
    If Len(scenarioId) = 0 Then
        MsgBox "No scenario to render. Load one, or pick a saved scenario in '" & _
               OutputScenarioPickerName() & "'.", vbExclamation
        Exit Sub
    End If

    chosen = SelectedOutputTemplate()
    Set template = WholeMergeFootprint(RequireNamedRange(chosen.nameRef))
    Set anchor = RequireNamedRange(OutputAnchorName())
    Set ws = anchor.Worksheet

    Set results = ResultsLookup()
    LoadCatalogueMaps labels, formats, groups

    guard = BeginBatch("Building the scenario output...")

    ClearPreviousOutput ws

    Set dest = anchor.Cells(1, 1).Resize(template.Rows.count, template.Columns.count)
    ' A merged area straddling the landing area, from a hand edit or an older build,
    ' would make the paste fail.
    WholeMergeFootprint(dest).UnMerge
    template.Copy dest.Cells(1, 1)
    CopyColumnWidthsAndRowHeights template, dest
    Application.CutCopyMode = False

    FillTokens dest, scenarioId, 1, results, labels, formats

    chartProblems = RenderChartTokens(ws, dest, CHART_NAME_PREFIX, False)

    ' Template formulas are written while calculation is manual; settle them here so the
    ' sheet never shows stale results between the write and the next recalc.
    ws.Calculate

    RecordExtent dest

    EndBatch guard

    ws.Activate
    anchor.Select
    If Len(chartProblems) = 0 Then
        MsgBox "Rendered scenario '" & ScenarioName(scenarioId) & "' using the '" & _
               chosen.label & "' template.", vbInformation
    Else
        MsgBox "Rendered scenario '" & ScenarioName(scenarioId) & "' using the '" & _
               chosen.label & "' template, but some charts were skipped:" & vbNewLine & _
               vbNewLine & chartProblems, vbExclamation
    End If
    Exit Sub

Failed:
    EndBatch guard
    MsgBox "The scenario output could not be built." & vbNewLine & vbNewLine & _
           Err.description, vbCritical
End Sub

'=======================================================================================
' Scenario picker
'=======================================================================================

' Scenario ID the output should render. Order:
'   1. The picker cell, if it names a saved scenario.
'   2. The "Loaded (...)" entry, or an empty cell, falls back to the active scenario.
'   3. Fails silently (returns "") so the caller can message instead of crashing.
Public Function ResolveOutputScenarioId() As String
    Dim cell As Range
    Dim picked As String
    Dim id As String

    Set cell = NamedRangeOrNothing(OutputScenarioPickerName())
    If Not cell Is Nothing Then picked = Trim$(CStr(cell.value))

    If Len(picked) > 0 And Not IsLoadedLabel(picked) Then
        id = ScenarioIdForName(picked)
        If Len(id) > 0 Then
            ResolveOutputScenarioId = id
            Exit Function
        End If
    End If

    ResolveOutputScenarioId = ActiveScenarioIdVal()
End Function

' Rebuilds the picker's dropdown from the saved scenarios plus a "Loaded (<name>)" entry
' that means "whichever scenario is currently loaded." Called from the build button and
' whenever the manager refreshes, so loading, adding or renaming a scenario is picked up
' automatically.
Public Sub RefreshOutputScenarioPicker()
    Dim cell As Range
    Dim ids As Variant
    Dim list As String
    Dim loadedLabel As String
    Dim i As Long

    Set cell = NamedRangeOrNothing(OutputScenarioPickerName())
    If cell Is Nothing Then Exit Sub

    loadedLabel = CurrentLoadedLabel()
    list = loadedLabel

    ids = ScenarioIds()
    If UBound(ids) >= LBound(ids) Then
        For i = LBound(ids) To UBound(ids)
            list = list & "," & ScenarioName(CStr(ids(i)))
        Next i
    End If

    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=list
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' An empty cell, or a "Loaded (...)" label naming whatever was loaded last time, is
    ' brought up to date so the cell always names the scenario it will render.
    If Len(Trim$(CStr(cell.value))) = 0 Or IsLoadedLabel(CStr(cell.value)) Then
        cell.value = loadedLabel
    End If
End Sub

' "Loaded (Base Case)", or plain "Loaded" when nothing is loaded.
Private Function CurrentLoadedLabel() As String
    Dim scenarioId As String

    scenarioId = ActiveScenarioIdVal()
    If Len(scenarioId) = 0 Then
        CurrentLoadedLabel = LOADED_SENTINEL
    Else
        CurrentLoadedLabel = LOADED_SENTINEL & " (" & ScenarioName(scenarioId) & ")"
    End If
End Function

' True for the bare sentinel and for any "Loaded (...)" label, including a stale one
' that names a scenario which has since been unloaded or renamed.
Private Function IsLoadedLabel(ByVal text As String) As Boolean
    text = Trim$(text)
    If StrComp(text, LOADED_SENTINEL, vbTextCompare) = 0 Then
        IsLoadedLabel = True
    Else
        IsLoadedLabel = (LCase$(text) Like LCase$(LOADED_SENTINEL) & " (*)")
    End If
End Function

'=======================================================================================
' Template picker
'=======================================================================================

Public Function SelectedOutputTemplate() As CompareTemplateSpec
    Dim templates() As CompareTemplateSpec
    Dim cell As Range
    Dim picked As String
    Dim i As Long

    templates = ScenarioOutputTemplates()
    If OutputTemplateCount(templates) = 0 Then
        Err.Raise vbObjectError + 730, "ScenarioOutput", _
                  "No output templates are declared in ScenarioOutputTemplates()."
    End If

    Set cell = NamedRangeOrNothing(OutputTemplatePickerName())
    If Not cell Is Nothing Then picked = Trim$(CStr(cell.value))

    For i = LBound(templates) To UBound(templates)
        If StrComp(templates(i).label, picked, vbTextCompare) = 0 Then
            SelectedOutputTemplate = templates(i)
            Exit Function
        End If
    Next i

    SelectedOutputTemplate = templates(LBound(templates))
End Function

Public Sub RefreshOutputTemplatePicker()
    Dim templates() As CompareTemplateSpec
    Dim cell As Range
    Dim list As String
    Dim i As Long

    Set cell = NamedRangeOrNothing(OutputTemplatePickerName())
    If cell Is Nothing Then Exit Sub

    templates = ScenarioOutputTemplates()
    For i = LBound(templates) To UBound(templates)
        If Len(list) > 0 Then list = list & ","
        list = list & templates(i).label
    Next i

    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=list
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    If Len(Trim$(CStr(cell.value))) = 0 Then cell.value = templates(LBound(templates)).label
End Sub

Private Function OutputTemplateCount(templates() As CompareTemplateSpec) As Long
    On Error Resume Next
    OutputTemplateCount = UBound(templates) - LBound(templates) + 1
    If Err.Number <> 0 Then OutputTemplateCount = 0
    Err.Clear
End Function

'=======================================================================================
' Template range redefinition
'=======================================================================================

' Mirrors the compare-sheet redefine flow. The output template design area has its own
' combo box (cmbRedefineOutputTemplate) and its own Redefine button.
Public Sub RefreshOutputRedefineCombo()
    Dim templates() As CompareTemplateSpec
    Dim shp As Shape
    Dim ctrl As ControlFormat
    Dim i As Long
    Dim previous As String
    Dim restoreTo As Long

    Set shp = FindShapeAnywhere(OutputRedefineComboName())
    If shp Is Nothing Then Exit Sub

    Set ctrl = shp.ControlFormat
    If ctrl.ListIndex >= 1 Then previous = CStr(ctrl.list(ctrl.ListIndex))

    ctrl.ListFillRange = ""
    ctrl.RemoveAllItems

    templates = ScenarioOutputTemplates()
    For i = LBound(templates) To UBound(templates)
        ctrl.AddItem templates(i).label
        If Len(previous) > 0 And StrComp(templates(i).label, previous, vbTextCompare) = 0 Then
            restoreTo = i - LBound(templates) + 1
        End If
    Next i

    If restoreTo >= 1 Then
        ctrl.ListIndex = restoreTo
    ElseIf ctrl.ListIndex < 1 Then
        ctrl.ListIndex = 1
    End If
End Sub

Public Function SelectedOutputRedefineTemplate() As CompareTemplateSpec
    Dim templates() As CompareTemplateSpec
    Dim shp As Shape
    Dim ctrl As ControlFormat
    Dim picked As String
    Dim i As Long

    Set shp = FindShapeAnywhere(OutputRedefineComboName())
    If shp Is Nothing Then
        Err.Raise vbObjectError + 740, "ScenarioOutput", _
                  "The combo box '" & OutputRedefineComboName() & "' is missing. " & _
                  "Add it to the output template design area, or rename the existing one."
    End If

    Set ctrl = shp.ControlFormat
    If ctrl.ListIndex < 1 Then
        Err.Raise vbObjectError + 741, "ScenarioOutput", _
                  "Choose a template in '" & OutputRedefineComboName() & "' first."
    End If

    picked = CStr(ctrl.list(ctrl.ListIndex))

    templates = ScenarioOutputTemplates()
    For i = LBound(templates) To UBound(templates)
        If StrComp(templates(i).label, picked, vbTextCompare) = 0 Then
            SelectedOutputRedefineTemplate = templates(i)
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 742, "ScenarioOutput", _
              "'" & picked & "' is not one of the configured output templates."
End Function

' Repoints the picked output template's named range at newRange and moves the highlight
' frame and header label with it (see ScenarioCompare's template range redefinition). A
' missing named range is created on the fly.
Public Sub RedefineSelectedOutputTemplateRange(ByVal newRange As Range)
    Dim chosen As CompareTemplateSpec
    Dim oldRange As Range

    If newRange Is Nothing Then
        Err.Raise vbObjectError + 743, "ScenarioOutput", _
                  "Select the new area for the template first."
    End If
    If newRange.Areas.count > 1 Then
        Err.Raise vbObjectError + 744, "ScenarioOutput", _
                  "The template range must be a single rectangle, not a multi-area selection."
    End If
    If newRange.Row = 1 Then
        Err.Raise vbObjectError + 745, "ScenarioOutput", _
                  "The template range cannot start on row 1 because the label needs the row above it."
    End If
    If Not TemplateFrameFits(newRange) Then
        Err.Raise vbObjectError + 746, "ScenarioOutput", _
                  "The template range needs a free cell on every side for its highlight " & _
                  "frame, so it cannot start in column A or touch the sheet's last row or column."
    End If

    chosen = SelectedOutputRedefineTemplate()

    Set oldRange = NamedRangeOrNothing(chosen.nameRef)
    If Not oldRange Is Nothing Then ClearTemplateRangeMarkers oldRange, chosen.label

    newRange.Name = chosen.nameRef
    MarkTemplateRange newRange, chosen.label
End Sub

Private Function FindShapeAnywhere(shapeName As String) As Shape
    Dim ws As Worksheet
    Dim shp As Shape

    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            If StrComp(shp.Name, shapeName, vbTextCompare) = 0 Then
                Set FindShapeAnywhere = shp
                Exit Function
            End If
        Next shp
    Next ws
End Function

'=======================================================================================
' Chart tokens
'=======================================================================================

' Walks block for cells whose entire contents are a chart token, clears each, and
' anchors a ChartObject there named chartPrefix & <spec> & "_" & <n>.
'
'   fitToBlock         shrink each chart so it stays inside block
'   storedScenarioId   plot that scenario's stored results (with the current state as
'                      the baseline) instead of binding to the live tables; needs results
'
' A chart that fails is skipped rather than aborting the caller's build. Returns one
' line per skipped chart, or "" when every chart was placed.
Public Function RenderChartTokens(ws As Worksheet, block As Range, chartPrefix As String, _
                                  ByVal fitToBlock As Boolean, _
                                  Optional storedScenarioId As String = "", _
                                  Optional results As Object = Nothing) As String
    Dim cell As Range
    Dim text As String
    Dim body As String
    Dim spec As ChartSpec
    Dim serial As Long
    Dim failures As String

    For Each cell In block.Cells
        If cell.HasFormula Then GoTo NextCell
        text = Trim$(CStr(cell.value))
        If Not IsChartTokenCell(text) Then GoTo NextCell

        On Error GoTo ChartFailed

        ' The whole merged area: clearing just its top-left cell raises. Cleared before
        ' the spec is resolved so a failed chart never leaves raw token text behind.
        cell.MergeArea.ClearContents
        serial = serial + 1

        body = Mid$(Mid$(text, 2, Len(text) - 2), Len(CHART_TOKEN_PREFIX) + 1)
        spec = ResolveChartSpec(body)
        PlaceChart ws, cell, block, spec, chartPrefix & SafeName(spec.label) & "_" & serial, _
                   fitToBlock, storedScenarioId, results

        On Error GoTo 0
NextCell:
    Next cell

    RenderChartTokens = failures
    Exit Function

ChartFailed:
    failures = failures & "- " & cell.Address(False, False) & ": " & Err.description & vbNewLine
    Resume NextCell
End Function

' Deletes every ChartObject on ws whose name starts with prefix.
Public Sub DeleteRenderedCharts(ws As Worksheet, prefix As String)
    Dim i As Long

    For i = ws.ChartObjects.count To 1 Step -1
        If StartsWithIgnoreCase(ws.ChartObjects(i).Name, prefix) Then ws.ChartObjects(i).Delete
    Next i
End Sub

Private Function IsChartTokenCell(text As String) As Boolean
    Dim inner As String

    If Len(text) < Len(CHART_TOKEN_PREFIX) + 2 Then Exit Function
    If Left$(text, 1) <> TokenOpen() Or Right$(text, 1) <> TokenClose() Then Exit Function
    If InStr(text, TokenClose()) <> Len(text) Then Exit Function

    inner = Trim$(Mid$(text, 2, Len(text) - 2))
    IsChartTokenCell = (StrComp(Left$(inner, Len(CHART_TOKEN_PREFIX)), _
                                CHART_TOKEN_PREFIX, vbTextCompare) = 0)
End Function

' The content after "Chart:" is either the label of a named spec or an inline
' key=value;... spec. The two forms are told apart by the presence of "=".
Private Function ResolveChartSpec(body As String) As ChartSpec
    Dim specs() As ChartSpec
    Dim i As Long
    Dim trimmed As String

    trimmed = Trim$(body)

    If InStr(trimmed, "=") = 0 Then
        specs = ScenarioOutputCharts()
        For i = LBound(specs) To UBound(specs)
            If StrComp(specs(i).label, trimmed, vbTextCompare) = 0 Then
                ResolveChartSpec = specs(i)
                Exit Function
            End If
        Next i
        Err.Raise vbObjectError + 750, "ScenarioOutput", _
                  "No chart named '" & trimmed & "' is declared in ScenarioOutputCharts()."
    End If

    ResolveChartSpec = ParseInlineChartSpec(trimmed)
End Function

Public Function ParseInlineChartSpec(body As String) As ChartSpec
    Dim parts() As String
    Dim pair() As String
    Dim i As Long
    Dim key As String, value As String

    ParseInlineChartSpec.kind = "Line"
    ParseInlineChartSpec.showLegend = True

    parts = Split(body, ";")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) = 0 Then GoTo NextPart

        pair = Split(parts(i), "=", 2)
        If UBound(pair) < 1 Then GoTo NextPart

        key = LCase$(Trim$(pair(0)))
        value = Trim$(pair(1))

        Select Case key
            Case "kind":            ParseInlineChartSpec.kind = value
            Case "source":          ParseInlineChartSpec.sourceTable = value
            Case "x":               ParseInlineChartSpec.xColumn = value
            Case "series":          ParseInlineChartSpec.seriesColumns = value
            Case "baselinesource":  ParseInlineChartSpec.baselineTable = value
            Case "baselineseries":  ParseInlineChartSpec.baselineColumns = value
            Case "title":           ParseInlineChartSpec.title = value
            Case "format":          ParseInlineChartSpec.numberFormat = value
            Case "legend":          ParseInlineChartSpec.showLegend = IsYes(value)
            Case "height":          ParseInlineChartSpec.heightRows = CLng(Val(value))
            Case "width":           ParseInlineChartSpec.widthCols = CLng(Val(value))
            Case "label":           ParseInlineChartSpec.label = value
        End Select

NextPart:
    Next i

    If Len(ParseInlineChartSpec.label) = 0 Then _
        ParseInlineChartSpec.label = ParseInlineChartSpec.sourceTable
End Function

Private Function IsYes(value As String) As Boolean
    Select Case LCase$(Trim$(value))
        Case "y", "yes", "true", "1": IsYes = True
    End Select
End Function

' Anchors a ChartObject at cell and binds its series, either to the configured table
' columns or, given storedScenarioId, to that scenario's stored results staged on a
' hidden sheet. A chart that fails part-way is deleted before the error is passed on, so
' a skipped chart never leaves an empty frame behind.
Private Sub PlaceChart(ws As Worksheet, anchorCell As Range, block As Range, _
                       spec As ChartSpec, chartName As String, ByVal fitToBlock As Boolean, _
                       storedScenarioId As String, results As Object)
    Dim source As ListObject, baseline As ListObject
    Dim chartObj As ChartObject
    Dim chart As chart
    Dim staged As Range
    Dim leftPx As Double, topPx As Double
    Dim widthPx As Double, heightPx As Double
    Dim seriesList() As String, baseList() As String
    Dim hasBaseline As Boolean
    Dim errNumber As Long, errText As String
    Dim i As Long

    If Len(spec.sourceTable) = 0 Then
        Err.Raise vbObjectError + 751, "ScenarioOutput", _
                  "Chart '" & spec.label & "' has no source table."
    End If
    If Len(spec.xColumn) = 0 Then
        Err.Raise vbObjectError + 752, "ScenarioOutput", _
                  "Chart '" & spec.label & "' has no x-axis column."
    End If
    If Len(spec.seriesColumns) = 0 Then
        Err.Raise vbObjectError + 753, "ScenarioOutput", _
                  "Chart '" & spec.label & "' has no series columns."
    End If

    seriesList = SplitAndTrim(spec.seriesColumns)
    If Len(spec.baselineColumns) > 0 Then
        baseList = SplitAndTrim(spec.baselineColumns)
        hasBaseline = True
    End If

    If Len(storedScenarioId) = 0 Then
        Set source = RequireTable(spec.sourceTable)
        If Len(spec.baselineTable) > 0 Then Set baseline = TableOrNothing(spec.baselineTable)
        hasBaseline = hasBaseline And Not baseline Is Nothing
    Else
        Set staged = StageStoredSeries(spec, storedScenarioId, results, seriesList, _
                                       baseList, hasBaseline)
    End If

    ChartFrame anchorCell, block, spec, fitToBlock, leftPx, topPx, widthPx, heightPx

    On Error GoTo Failed

    Set chartObj = ws.ChartObjects.Add(leftPx, topPx, widthPx, heightPx)
    chartObj.Name = chartName

    Set chart = chartObj.chart
    chart.ChartType = MapChartKind(spec.kind)

    ' Remove the autogenerated series so SeriesCollection starts empty.
    Do While chart.SeriesCollection.count > 0
        chart.SeriesCollection(1).Delete
    Loop

    If staged Is Nothing Then
        For i = LBound(seriesList) To UBound(seriesList)
            AddSeriesFromTable chart, source, spec.xColumn, seriesList(i), seriesList(i)
        Next i
        If hasBaseline Then
            For i = LBound(baseList) To UBound(baseList)
                AddSeriesFromTable chart, baseline, spec.xColumn, baseList(i), _
                                   baseList(i) & " (baseline)"
            Next i
        End If
    Else
        ' Staged layout: header row, then x in the first column and one column per series.
        For i = 2 To staged.Columns.count
            AddSeries chart, staged.Cells(2, 1).Resize(staged.Rows.count - 1, 1), _
                      staged.Cells(2, i).Resize(staged.Rows.count - 1, 1), _
                      CStr(staged.Cells(1, i).value)
        Next i
    End If

    chart.HasTitle = True
    chart.ChartTitle.text = IIf(Len(spec.title) > 0, spec.title, spec.label)
    chart.HasLegend = spec.showLegend

    If Len(spec.numberFormat) > 0 Then
        On Error Resume Next
        chart.Axes(xlValue).TickLabels.numberFormat = spec.numberFormat
        On Error GoTo 0
    End If
    Exit Sub

Failed:
    errNumber = Err.Number
    errText = Err.description
    If Not chartObj Is Nothing Then chartObj.Delete
    Err.Raise errNumber, "ScenarioOutput", errText
End Sub

' Where the chart goes and how big it is. The intended size is the spec's height/width
' when set, else the anchor's merged area when it has one, else 15 rows by 8 columns.
' fitToBlock then trims that so the chart ends at block's right and bottom edges.
Private Sub ChartFrame(anchorCell As Range, block As Range, spec As ChartSpec, _
                       ByVal fitToBlock As Boolean, ByRef leftPx As Double, _
                       ByRef topPx As Double, ByRef widthPx As Double, _
                       ByRef heightPx As Double)
    Dim area As Range

    Set area = anchorCell.MergeArea
    leftPx = anchorCell.Left
    topPx = anchorCell.Top

    If spec.widthCols > 0 Then
        widthPx = anchorCell.Resize(1, spec.widthCols).Width
    ElseIf area.Cells.count > 1 Then
        widthPx = area.Width
    Else
        widthPx = anchorCell.Resize(1, 8).Width
    End If

    If spec.heightRows > 0 Then
        heightPx = anchorCell.Resize(spec.heightRows, 1).Height
    ElseIf area.Cells.count > 1 Then
        heightPx = area.Height
    Else
        heightPx = anchorCell.Resize(15, 1).Height
    End If

    If Not fitToBlock Then Exit Sub
    widthPx = Application.Min(widthPx, block.Left + block.Width - leftPx)
    heightPx = Application.Min(heightPx, block.Top + block.Height - topPx)
End Sub

Private Sub AddSeriesFromTable(chart As chart, lo As ListObject, xColumn As String, _
                               valueColumn As String, seriesName As String)
    Dim xCol As ListColumn, vCol As ListColumn

    Set xCol = lo.ListColumns(xColumn)
    Set vCol = lo.ListColumns(valueColumn)

    If lo.ListRows.count = 0 Then Exit Sub

    AddSeries chart, xCol.DataBodyRange, vCol.DataBodyRange, seriesName
End Sub

Private Sub AddSeries(chart As chart, xRange As Range, vRange As Range, seriesName As String)
    Dim ser As Series

    Set ser = chart.SeriesCollection.NewSeries
    ser.values = vRange
    ser.XValues = xRange
    ser.Name = seriesName
End Sub

'=======================================================================================
' Stored-result chart data
'=======================================================================================
'
' The live tables a chart normally binds to only ever hold the loaded scenario, so a
' chart drawn for any other scenario has to be fed from tblScenarioResults instead. The
' metric catalogue says which stored metric holds each (row, column) cell of a source
' table, so a chart's x/series columns map straight onto metric IDs. The values are
' written to a very hidden sheet, one block of columns per chart, and the series bind to
' those cells: series built from literal arrays hit Excel's series-formula length limit
' on anything longer than a few dozen points.
'
' Only metrics that are ticked Include and have been run are stored, so a chart over
' uncaptured columns plots #N/A gaps. Baseline series read the current-state results for
' the same metrics, the same way {Current.X} does; baselineSource is not consulted.

Private Function StageStoredSeries(spec As ChartSpec, scenarioId As String, _
                                   results As Object, seriesList() As String, _
                                   baseList() As String, ByVal hasBaseline As Boolean) As Range
    Dim specs() As MetricSpec
    Dim rowKeys As Object, cellIds As Object
    Dim keys As Variant
    Dim data() As Variant
    Dim stage As Worksheet
    Dim target As Range
    Dim labelColumn As String, rowKey As String
    Dim nSeries As Long, nBase As Long, nRows As Long, col As Long
    Dim r As Long, i As Long

    If results Is Nothing Then
        Err.Raise vbObjectError + 756, "ScenarioOutput", _
                  "Chart '" & spec.label & "' needs stored results but none were passed in."
    End If

    ' Row keys in catalogue order, which follows the source table's own row order.
    specs = CatalogueSpecs(False)
    Set rowKeys = CreateObject("Scripting.Dictionary")
    Set cellIds = CreateObject("Scripting.Dictionary")
    rowKeys.CompareMode = vbTextCompare
    cellIds.CompareMode = vbTextCompare
    For i = 1 To MetricSpecCount(specs)
        If StrComp(specs(i).sourceTable, spec.sourceTable, vbTextCompare) = 0 Then
            If Not rowKeys.Exists(specs(i).labelValue) Then rowKeys.Add specs(i).labelValue, True
            cellIds(specs(i).labelValue & "|" & specs(i).sourceColumn) = specs(i).metric
            labelColumn = specs(i).labelColumn
        End If
    Next i

    If rowKeys.count = 0 Then
        Err.Raise vbObjectError + 755, "ScenarioOutput", _
                  "Chart '" & spec.label & "' plots stored results, but no metric in " & _
                  MetricCatalogueTableName() & " comes from " & spec.sourceTable & _
                  ". Sync the metrics and run the scenarios."
    End If

    nRows = rowKeys.count
    nSeries = UBound(seriesList) - LBound(seriesList) + 1
    If hasBaseline Then nBase = UBound(baseList) - LBound(baseList) + 1

    ReDim data(1 To nRows + 1, 1 To 1 + nSeries + nBase)
    data(1, 1) = spec.xColumn
    For i = 1 To nSeries
        data(1, 1 + i) = seriesList(LBound(seriesList) + i - 1)
    Next i
    For i = 1 To nBase
        data(1, 1 + nSeries + i) = baseList(LBound(baseList) + i - 1) & " (baseline)"
    Next i

    keys = rowKeys.keys
    For r = 1 To nRows
        rowKey = CStr(keys(r - 1))

        ' The x column is usually the row key itself; anything else is a captured
        ' metric of its own, falling back to the key when it was not captured.
        If StrComp(spec.xColumn, labelColumn, vbTextCompare) = 0 Then
            data(r + 1, 1) = rowKey
        Else
            data(r + 1, 1) = StoredPoint(results, scenarioId, cellIds, rowKey, spec.xColumn, False)
            If IsError(data(r + 1, 1)) Then data(r + 1, 1) = rowKey
        End If

        For i = 1 To nSeries
            data(r + 1, 1 + i) = StoredPoint(results, scenarioId, cellIds, rowKey, _
                                             seriesList(LBound(seriesList) + i - 1), True)
        Next i
        For i = 1 To nBase
            data(r + 1, 1 + nSeries + i) = StoredPoint(results, CurrentStateScenarioId(), _
                                                       cellIds, rowKey, _
                                                       baseList(LBound(baseList) + i - 1), True)
        Next i
    Next r

    Set stage = ChartStagingSheet()
    If IsEmpty(stage.Cells(1, 1).value) Then
        col = 1
    Else
        col = stage.Cells(1, stage.Columns.count).End(xlToLeft).column + 2
    End If

    Set target = stage.Cells(1, col).Resize(nRows + 1, 1 + nSeries + nBase)
    target.value = data
    Set StageStoredSeries = target
End Function

' One stored cell of a source table, or #N/A when it was never captured. #N/A rather than
' blank or zero so a line chart shows a gap instead of a plunge to zero. numericOnly also
' turns the run's text markers (#ROW MISSING etc.) into #N/A, since text plots as zero.
Private Function StoredPoint(results As Object, scenarioId As String, cellIds As Object, _
                             rowKey As String, column As String, _
                             ByVal numericOnly As Boolean) As Variant
    Dim key As String
    Dim value As Variant

    StoredPoint = CVErr(xlErrNA)

    key = rowKey & "|" & column
    If Not cellIds.Exists(key) Then Exit Function

    value = MetricValue(results, scenarioId, CStr(cellIds(key)), Empty)
    If IsEmpty(value) Then Exit Function
    If numericOnly And Not IsNumeric(value) Then Exit Function

    StoredPoint = value
End Function

' Created on first use and kept very hidden, so it never shows in the sheet tabs or the
' Unhide dialog.
Private Function ChartStagingSheet() As Worksheet
    Dim stage As Worksheet

    Set stage = SheetOrNothing(STAGING_SHEET)
    If stage Is Nothing Then
        Set stage = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        stage.Name = STAGING_SHEET
        stage.Visible = xlSheetVeryHidden
    End If

    Set ChartStagingSheet = stage
End Function

' Empties the staging sheet. Call only after deleting the charts that read from it.
Public Sub ClearStagedChartData()
    Dim stage As Worksheet

    Set stage = SheetOrNothing(STAGING_SHEET)
    If Not stage Is Nothing Then stage.Cells.Clear
End Sub

Private Function MapChartKind(kind As String) As XlChartType
    Select Case LCase$(Trim$(kind))
        Case "column":                  MapChartKind = xlColumnClustered
        Case "bar":                     MapChartKind = xlBarClustered
        Case "combo", "comboclustered": MapChartKind = xlColumnClustered
        Case Else:                      MapChartKind = xlLine
    End Select
End Function

Private Function SplitAndTrim(csv As String) As String()
    Dim raw() As String
    Dim out() As String
    Dim i As Long, n As Long

    raw = Split(csv, ",")
    ReDim out(LBound(raw) To UBound(raw))
    For i = LBound(raw) To UBound(raw)
        If Len(Trim$(raw(i))) > 0 Then
            out(n + LBound(raw)) = Trim$(raw(i))
            n = n + 1
        End If
    Next i
    If n = 0 Then
        Err.Raise vbObjectError + 754, "ScenarioOutput", _
                  "Expected one or more comma-separated names, got '" & csv & "'."
    End If
    ReDim Preserve out(LBound(raw) To LBound(raw) + n - 1)
    SplitAndTrim = out
End Function

Private Function SafeName(text As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Then
            out = out & ch
        End If
    Next i
    If Len(out) = 0 Then out = "chart"
    SafeName = out
End Function

'=======================================================================================
' Generated area
'=======================================================================================

Private Sub ClearPreviousOutput(ws As Worksheet)
    Dim target As Range
    Dim extent As String

    ' Delete charts first so the subsequent Clear is not fighting their anchor cells.
    DeleteRenderedCharts ws, CHART_NAME_PREFIX

    extent = SettingText(EXTENT_KEY)
    If Len(extent) = 0 Then Exit Sub

    On Error Resume Next
    Set target = ws.Range(extent)
    On Error GoTo 0
    If target Is Nothing Then
        SettingText(EXTENT_KEY) = ""
        Exit Sub
    End If

    Set target = WholeMergeFootprint(target)
    target.Clear
    target.ClearFormats

    SettingText(EXTENT_KEY) = ""
End Sub

Private Sub RecordExtent(covered As Range)
    SettingText(EXTENT_KEY) = covered.Address
End Sub

' rng widened until no merged area straddles its edge. Copying, clearing or pasting over
' part of a merged area raises "We can't do that to a merged cell", which is exactly what
' a chart area merged past a template's defined range runs into.
Public Function WholeMergeFootprint(rng As Range) As Range
    Dim grown As Range
    Dim cell As Range
    Dim area As Range
    Dim changed As Boolean

    Set grown = rng
    Do
        ' MergeCells is False when nothing in the range is merged: the common case, and
        ' far cheaper than visiting every cell.
        If Not IsNull(grown.MergeCells) Then
            If Not grown.MergeCells Then Exit Do
        End If

        changed = False
        For Each cell In grown.Cells
            If cell.MergeCells Then
                Set area = cell.MergeArea
                If Application.Intersect(area, grown).Cells.count < area.Cells.count Then
                    Set grown = BoundingBox(grown, area)
                    changed = True
                    Exit For
                End If
            End If
        Next cell
    Loop While changed

    Set WholeMergeFootprint = grown
End Function

Private Function BoundingBox(a As Range, b As Range) As Range
    Dim ws As Worksheet

    Set ws = a.Worksheet
    Set BoundingBox = ws.Range( _
        ws.Cells(Application.Min(a.Row, b.Row), Application.Min(a.column, b.column)), _
        ws.Cells(Application.Max(a.Row + a.Rows.count, b.Row + b.Rows.count) - 1, _
                 Application.Max(a.column + a.Columns.count, b.column + b.Columns.count) - 1))
End Function

Private Sub CopyColumnWidthsAndRowHeights(source As Range, dest As Range)
    Dim i As Long

    For i = 1 To source.Columns.count
        dest.Columns(i).ColumnWidth = source.Columns(i).ColumnWidth
    Next i
    For i = 1 To source.Rows.count
        dest.Rows(i).RowHeight = source.Rows(i).RowHeight
    Next i
End Sub

Private Function StartsWithIgnoreCase(text As String, prefix As String) As Boolean
    If Len(text) < Len(prefix) Then Exit Function
    StartsWithIgnoreCase = (StrComp(Left$(text, Len(prefix)), prefix, vbTextCompare) = 0)
End Function
