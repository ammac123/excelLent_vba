Attribute VB_Name = "ScenarioCompare"
Option Explicit

' Side-by-side comparison of however many scenarios you tick.
'
' The sheet is built from things the user controls directly:
'
'   CompareTemplates()      A list of sample summary blocks, declared in ScenarioConfig.
'                           Each is a defined name covering a block laid out and
'                           formatted by hand. One is picked at a time and copied once
'                           per selected scenario, so changing how a summary looks is a
'                           spreadsheet job rather than a code job.
'   compareTemplatePicker   Cell holding the chosen template's label.
'   compareAnchor           The top-left cell the first copy lands on.
'
' The token grammar and resolver live in ScenarioTokens. See the comment block at the
' top of that module for the full list of {Scenario.X}, {Metric.X}, {Current.X},
' {Delta.X}, {DeltaPct.X} and {Label.X} tokens and how they behave in cells, text and
' formulas.
'
' Ordinary formulas are copied, not flattened, so a template can calculate off its own
' cells. Relative references inside the block translate with the copy as they would if
' you dragged it. References to anything outside the block must be absolute, or they will
' translate too and point at the wrong cells in the second and later blocks.
'
' Underneath the blocks the same metrics are written out again as one grid,
' tblScenarioComparison, with a row per metric and a column per scenario.
'
' Chart tokens ({Chart:<Name>} and inline {Chart:k=v;...}, see ScenarioOutput) render
' once per block, with two differences from the output sheet:
'
'   - Each chart is shrunk to fit inside its own block, so a chart sized for the output
'     sheet can't spill over the next scenario or the grid. Give the chart room by
'     merging the area it should fill, or by making the template bigger.
'   - Each chart plots that block's stored results (baseline series: the current
'     state's), not the live tables, which only ever hold the loaded scenario. A series
'     only has data for metrics that are ticked Include and have been run.
'
' Everything this module writes is disposable: the generated area is cleared and rebuilt
' from tblScenarioResults every time, so the comparison can never drift out of step with
' the stored results.

' The generated area is recorded here so a rebuild only ever clears what a previous
' rebuild put there.
Private Const EXTENT_KEY As String = "CompareExtent"
Private Const CHART_NAME_PREFIX As String = "scnCmp_"

'=======================================================================================
' Entry point
'=======================================================================================

Public Sub BuildScenarioComparison()
    Dim ws As Worksheet
    Dim chosen As CompareTemplateSpec
    Dim template As Range, anchor As Range
    Dim ids() As String
    Dim metrics() As String
    Dim results As Object, labels As Object, formats As Object, groups As Object
    Dim guard As CalcGuard
    Dim count As Long
    Dim chartProblems As String
    Dim summary As String

    On Error GoTo Failed

    ids = ComparisonScenarioIds()
    count = StringArrayCount(ids)
    If count = 0 Then
        MsgBox "No scenarios are ticked " & COL_COMPARE & " in " & TBL_SCENARIOS & ".", _
               vbExclamation
        Exit Sub
    End If

    RefreshCompareTemplatePicker
    chosen = SelectedTemplate()
    ' Widened to whole merged areas: copying part of one fails, and the block size (and so
    ' the spacing between blocks) has to include anything merged past the named range.
    Set template = WholeMergeFootprint(RequireNamedRange(chosen.nameRef))
    Set anchor = RequireNamedRange(CompareAnchorName())
    Set ws = anchor.Worksheet

    metrics = TemplateMetricIds(template)

    Set results = ResultsLookup()
    LoadCatalogueMaps labels, formats, groups

    guard = BeginBatch("Building the scenario comparison...")

    ClearPreviousOutput ws
    chartProblems = PaintBlocks(ws, template, anchor, ids, results, labels, formats)
    BuildComparisonGrid ws, anchor, template, ids, metrics, results, labels, formats, groups
    ' Template formulas are written while calculation is manual, so settle them here
    ' rather than leaving the sheet full of stale results.
    ws.Calculate

    EndBatch guard

    ws.Activate
    anchor.Select
    summary = "Compared " & count & " scenario(s) over " & StringArrayCount(metrics) & _
              " metric(s) using the '" & chosen.label & "' template."
    If Len(chartProblems) = 0 Then
        MsgBox summary, vbInformation
    Else
        MsgBox summary & vbNewLine & vbNewLine & "Some charts were skipped:" & vbNewLine & _
               chartProblems, vbExclamation
    End If
    Exit Sub

Failed:
    EndBatch guard
    MsgBox "The scenario comparison could not be built." & vbNewLine & vbNewLine & _
           Err.description, vbCritical
End Sub

' Scenarios ticked for comparison, in table order. Falls back to the run flag, and then
' to every scenario, so the sheet still works before the column has been added.
Public Function ComparisonScenarioIds() As String()
    Dim lo As ListObject
    Dim ids As Variant
    Dim picked() As String
    Dim column As String
    Dim wanted As Boolean
    Dim i As Long, n As Long

    ids = ScenarioIds()
    If UBound(ids) < LBound(ids) Then Exit Function

    Set lo = RequireTable(TBL_SCENARIOS)
    If ColumnIndexOrZero(lo, COL_COMPARE) > 0 Then
        column = COL_COMPARE
    ElseIf ColumnIndexOrZero(lo, COL_INCLUDE) > 0 Then
        column = COL_INCLUDE
    End If

    ReDim picked(1 To UBound(ids) - LBound(ids) + 1)
    For i = LBound(ids) To UBound(ids)
        ' No flag column at all means compare everything.
        If Len(column) = 0 Then
            wanted = True
        Else
            wanted = IsTruthyValue(MetaValue(CStr(ids(i)), column))
        End If

        If wanted Then
            n = n + 1
            picked(n) = CStr(ids(i))
        End If
    Next i

    If n = 0 Then Exit Function
    ReDim Preserve picked(1 To n)
    ComparisonScenarioIds = picked
End Function

'=======================================================================================
' Template selection
'=======================================================================================

' The template named in the picker. Falls back to the first configured one, so an empty
' or stale picker still produces a comparison rather than an error.
Public Function SelectedTemplate() As CompareTemplateSpec
    Dim templates() As CompareTemplateSpec
    Dim cell As Range
    Dim picked As String
    Dim i As Long

    templates = CompareTemplates()
    If CompareTemplateCount(templates) = 0 Then
        Err.Raise vbObjectError + 630, "ScenarioCompare", _
                  "No summary templates are declared in CompareTemplates()."
    End If

    Set cell = NamedRangeOrNothing(CompareTemplatePickerName())
    If Not cell Is Nothing Then picked = Trim$(CStr(cell.value))

    For i = LBound(templates) To UBound(templates)
        If StrComp(templates(i).label, picked, vbTextCompare) = 0 Then
            SelectedTemplate = templates(i)
            Exit Function
        End If
    Next i

    SelectedTemplate = templates(LBound(templates))
End Function

' Rebuilds the picker's dropdown from the configured templates, so adding one to
' ScenarioConfig is the only change a new template needs.
Public Sub RefreshCompareTemplatePicker()
    Dim templates() As CompareTemplateSpec
    Dim cell As Range
    Dim list As String
    Dim i As Long

    Set cell = NamedRangeOrNothing(CompareTemplatePickerName())
    If cell Is Nothing Then Exit Sub

    templates = CompareTemplates()
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

Private Function CompareTemplateCount(templates() As CompareTemplateSpec) As Long
    On Error Resume Next
    CompareTemplateCount = UBound(templates) - LBound(templates) + 1
    If Err.Number <> 0 Then CompareTemplateCount = 0
    Err.Clear
End Function

'=======================================================================================
' Template range redefinition
'=======================================================================================
'
' Button handler on the template design sheet. The user selects the new rectangle, picks
' which template that rectangle belongs to from the cmbRedefineTemplate combo box, and
' clicks the button. The old range loses its highlight frame and header label; the new
' range gets them.
'
' The frame is drawn on the ring of cells one outside the template range, never on the
' range itself: a fill on the ring plus a thick black border around its outer edge, with
' the label in the ring cell above the top-left. Rendering copies only the range, so the
' marker can't follow a template into the generated blocks. The output sheet's templates
' share the same marker code.

' Rebuilds the redefine combo box's list from the configured templates. Called by the
' button so adding a template is still just an edit to CompareTemplates().
Public Sub RefreshRedefineTemplatePicker()
    Dim templates() As CompareTemplateSpec
    Dim shp As Shape
    Dim ctrl As ControlFormat
    Dim i As Long
    Dim previous As String
    Dim restoreTo As Long

    Set shp = FindShape(RedefineTemplateComboName())
    If shp Is Nothing Then Exit Sub

    Set ctrl = shp.ControlFormat
    ' Remember what was selected so rebuilding the list does not silently redirect the
    ' redefine action to a different template.
    If ctrl.ListIndex >= 1 Then previous = CStr(ctrl.list(ctrl.ListIndex))

    ' A combo bound to a sheet range ignores AddItem, so detach first.
    ctrl.ListFillRange = ""
    ctrl.RemoveAllItems

    templates = CompareTemplates()
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

' The template picked in the redefine combo box. Raises when the combo is missing or
' nothing is selected, so the caller gets a clear error rather than silently redefining
' the wrong range.
Public Function SelectedRedefineTemplate() As CompareTemplateSpec
    Dim templates() As CompareTemplateSpec
    Dim shp As Shape
    Dim ctrl As ControlFormat
    Dim picked As String
    Dim i As Long

    Set shp = FindShape(RedefineTemplateComboName())
    If shp Is Nothing Then
        Err.Raise vbObjectError + 640, "ScenarioCompare", _
                  "The combo box '" & RedefineTemplateComboName() & "' is missing. " & _
                  "Add it to the template design sheet, or rename the existing one."
    End If

    Set ctrl = shp.ControlFormat
    If ctrl.ListIndex < 1 Then
        Err.Raise vbObjectError + 641, "ScenarioCompare", _
                  "Choose a template in '" & RedefineTemplateComboName() & "' first."
    End If

    picked = CStr(ctrl.list(ctrl.ListIndex))

    templates = CompareTemplates()
    For i = LBound(templates) To UBound(templates)
        If StrComp(templates(i).label, picked, vbTextCompare) = 0 Then
            SelectedRedefineTemplate = templates(i)
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 642, "ScenarioCompare", _
              "'" & picked & "' is not one of the configured templates."
End Function

Private Function FindShape(shapeName As String) As Shape
    Dim ws As Worksheet
    Dim shp As Shape

    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            If StrComp(shp.Name, shapeName, vbTextCompare) = 0 Then
                Set FindShape = shp
                Exit Function
            End If
        Next shp
    Next ws
End Function

' Points the picked template's named range at newRange, clears the frame around the
' previous range, and frames the new one.
Public Sub RedefineSelectedTemplateRange(ByVal newRange As Range)
    Dim chosen As CompareTemplateSpec
    Dim oldRange As Range

    If newRange Is Nothing Then
        Err.Raise vbObjectError + 643, "ScenarioCompare", _
                  "Select the new area for the template first."
    End If
    If newRange.Areas.count > 1 Then
        Err.Raise vbObjectError + 644, "ScenarioCompare", _
                  "The template range must be a single rectangle, not a multi-area selection."
    End If
    If newRange.Row = 1 Then
        Err.Raise vbObjectError + 645, "ScenarioCompare", _
                  "The template range cannot start on row 1 because the label needs the row above it."
    End If
    If Not TemplateFrameFits(newRange) Then
        Err.Raise vbObjectError + 646, "ScenarioCompare", _
                  "The template range needs a free cell on every side for its highlight " & _
                  "frame, so it cannot start in column A or touch the sheet's last row or column."
    End If

    chosen = SelectedRedefineTemplate()

    Set oldRange = NamedRangeOrNothing(chosen.nameRef)
    If Not oldRange Is Nothing Then ClearTemplateRangeMarkers oldRange, chosen.label

    newRange.Name = chosen.nameRef
    MarkTemplateRange newRange, chosen.label
End Sub

' Re-frames every summary and output template where it currently stands. Clearing first
' also strips the old-style border drawn on the range itself, so running this once moves
' existing templates over to the outside frame. Returns the labels of any template that
' was skipped, one per line, or "" when all were framed.
Public Function RefreshTemplateFrames() As String
    Dim skipped As String

    skipped = RefreshFramesFor(CompareTemplates())
    skipped = skipped & RefreshFramesFor(ScenarioOutputTemplates())
    RefreshTemplateFrames = skipped
End Function

Private Function RefreshFramesFor(templates() As CompareTemplateSpec) As String
    Dim rng As Range
    Dim i As Long

    For i = LBound(templates) To UBound(templates)
        Set rng = NamedRangeOrNothing(templates(i).nameRef)
        If rng Is Nothing Then
            RefreshFramesFor = RefreshFramesFor & templates(i).label & _
                               " (no range named " & templates(i).nameRef & ")" & vbNewLine
        ElseIf Not TemplateFrameFits(rng) Then
            RefreshFramesFor = RefreshFramesFor & templates(i).label & _
                               " (no room for a frame around " & rng.Address(False, False) & _
                               "; redefine it away from row 1, column A and the sheet edge)" & vbNewLine
        Else
            ClearTemplateRangeMarkers rng, templates(i).label
            MarkTemplateRange rng, templates(i).label
        End If
    Next i
End Function

' True when rng has a cell on every side for the frame to sit in.
Public Function TemplateFrameFits(rng As Range) As Boolean
    Dim ws As Worksheet

    Set ws = rng.Worksheet
    TemplateFrameFits = rng.Row > 1 And rng.Column > 1 And _
                        rng.Row + rng.Rows.count - 1 < ws.Rows.count And _
                        rng.Column + rng.Columns.count - 1 < ws.Columns.count
End Function

' Removes the frame and the header label from a previously marked range. Only cells and
' edges still carrying the marker's own fill and border are cleared, and the label only
' when it still matches the template, so the user's own formatting is left alone.
Public Sub ClearTemplateRangeMarkers(rng As Range, label As String)
    Dim cell As Range
    Dim labelCell As Range

    ' Ranges marked before the frame moved outside carry the border on the range itself;
    ' strip it here so a redefine also cleans up the old style.
    ClearMarkerBorder rng

    If TemplateFrameFits(rng) Then
        ClearMarkerBorder TemplateFrame(rng)
        For Each cell In TemplateFrameRing(rng).Cells
            If cell.Interior.Color = TemplateFrameColor() Then cell.Interior.Pattern = xlNone
        Next cell
    End If

    If rng.Row > 1 Then
        Set labelCell = rng.Cells(1, 1).Offset(-1, 0)
        If StrComp(Trim$(CStr(labelCell.value)), label, vbTextCompare) = 0 Then
            labelCell.ClearContents
        End If
    End If
End Sub

' Callers check TemplateFrameFits first.
Public Sub MarkTemplateRange(rng As Range, label As String)
    TemplateFrameRing(rng).Interior.Color = TemplateFrameColor()
    TemplateFrame(rng).BorderAround LineStyle:=xlContinuous, Weight:=xlThick, Color:=vbBlack
    rng.Cells(1, 1).Offset(-1, 0).value = label
End Sub

' rng grown by one cell on every side.
Private Function TemplateFrame(rng As Range) As Range
    Set TemplateFrame = rng.Offset(-1, -1).Resize(rng.Rows.count + 2, rng.Columns.count + 2)
End Function

' Just the ring of cells between rng and the outer edge of its frame.
Private Function TemplateFrameRing(rng As Range) As Range
    Dim frame As Range

    Set frame = TemplateFrame(rng)
    Set TemplateFrameRing = Union(frame.Rows(1), frame.Rows(frame.Rows.count), _
                                  frame.Columns(1), frame.Columns(frame.Columns.count))
End Function

' Clears each perimeter edge of rng that is the marker's thick black line. A mixed edge
' reads back as Null, which If treats as False, so it is left alone.
Private Sub ClearMarkerBorder(rng As Range)
    Dim edges As Variant
    Dim i As Long

    edges = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
    For i = LBound(edges) To UBound(edges)
        With rng.Borders(edges(i))
            If .LineStyle = xlContinuous And .Weight = xlThick And .Color = vbBlack Then
                .LineStyle = xlNone
            End If
        End With
    Next i
End Sub

'=======================================================================================
' Summary blocks
'=======================================================================================

' Copies the template once per scenario, fills in its tokens and draws its charts.
' Records the rectangle covered so the next rebuild knows exactly what to clear. Returns
' one line per chart that could not be drawn, prefixed with its scenario.
Private Function PaintBlocks(ws As Worksheet, template As Range, anchor As Range, _
                             ids() As String, results As Object, labels As Object, _
                             formats As Object) As String
    Dim dest As Range
    Dim rowStep As Long, colStep As Long
    Dim i As Long, count As Long
    Dim problems As String, allProblems As String

    count = StringArrayCount(ids)

    If CompareLayoutIsAcross() Then
        colStep = template.Columns.count + CompareBlockGap()
    Else
        rowStep = template.Rows.count + CompareBlockGap()
    End If

    For i = 1 To count
        Set dest = anchor.Cells(1, 1).Offset(rowStep * (i - 1), colStep * (i - 1)) _
                         .Resize(template.Rows.count, template.Columns.count)

        ' A merged area straddling the landing area would make the paste fail.
        WholeMergeFootprint(dest).UnMerge
        template.Copy dest.Cells(1, 1)
        CopyColumnWidthsAndRowHeights template, dest
        FillTokens dest, ids(i), i, results, labels, formats

        ' The block index keeps chart names unique across blocks.
        problems = RenderChartTokens(ws, dest, CHART_NAME_PREFIX & i & "_", True, ids(i), results)
        If Len(problems) > 0 Then
            allProblems = allProblems & ScenarioName(ids(i)) & ":" & vbNewLine & problems
        End If
    Next i

    Application.CutCopyMode = False
    RecordExtent anchor.Cells(1, 1).Resize( _
        template.Rows.count + rowStep * (count - 1), _
        template.Columns.count + colStep * (count - 1))

    PaintBlocks = allProblems
End Function

' Column widths and row heights do not travel with Range.Copy, so set them explicitly.
' For an across layout each block has its own columns; for a stacked one each block has
' its own rows. The other dimension is shared and the second assignment is a no-op.
Private Sub CopyColumnWidthsAndRowHeights(source As Range, dest As Range)
    Dim i As Long

    For i = 1 To source.Columns.count
        dest.Columns(i).ColumnWidth = source.Columns(i).ColumnWidth
    Next i
    For i = 1 To source.Rows.count
        dest.Rows(i).RowHeight = source.Rows(i).RowHeight
    Next i
End Sub

'=======================================================================================
' Comparison grid
'=======================================================================================

' Metric by scenario, in the order the template referenced them.
Private Sub BuildComparisonGrid(ws As Worksheet, anchor As Range, template As Range, _
                                ids() As String, metrics() As String, results As Object, _
                                labels As Object, formats As Object, groups As Object)
    Dim grid As Range
    Dim values() As Variant
    Dim rows As Long, cols As Long, scenarios As Long
    Dim topRow As Long
    Dim m As Long, s As Long

    scenarios = StringArrayCount(ids)
    rows = StringArrayCount(metrics)
    cols = scenarios + 3   ' label, scenarios..., group, metric ID

    topRow = BlocksBottomRow(anchor, template, scenarios) + 2

    ReDim values(1 To rows + 1, 1 To cols)
    values(1, 1) = "Metric"
    For s = 1 To scenarios
        values(1, s + 1) = ScenarioName(ids(s))
    Next s
    values(1, cols - 1) = "Group"
    values(1, cols) = "MetricID"

    For m = 1 To rows
        values(m + 1, 1) = LookupOr(labels, metrics(m), metrics(m))
        For s = 1 To scenarios
            ' Blank rather than a marker, so the column stays numeric.
            values(m + 1, s + 1) = MetricValue(results, ids(s), metrics(m), "")
        Next s
        values(m + 1, cols - 1) = LookupOr(groups, metrics(m), "Ungrouped")
        values(m + 1, cols) = metrics(m)
    Next m

    Set grid = ws.Cells(topRow, anchor.column).Resize(rows + 1, cols)
    grid.value = values
    ExtendExtent ws, grid

    FormatGridRows grid, metrics, formats, scenarios
    ReplaceComparisonTable ws, grid
End Sub

Private Sub FormatGridRows(grid As Range, metrics() As String, formats As Object, _
                           ByVal scenarioCount As Long)
    Dim m As Long
    Dim fmt As String

    For m = 1 To StringArrayCount(metrics)
        fmt = LookupOr(formats, metrics(m), "")
        If Len(fmt) > 0 Then
            grid.Cells(m + 1, 2).Resize(1, scenarioCount).numberFormat = fmt
        End If
    Next m
End Sub

' Rebuilds the ListObject over the grid. Dropping and re-adding is simpler than resizing
' and guarantees the table never keeps columns from a comparison with more scenarios.
Private Sub ReplaceComparisonTable(ws As Worksheet, grid As Range)
    Dim lo As ListObject

    Set lo = TableOrNothing(ComparisonTableName())
    If Not lo Is Nothing Then lo.Unlist

    Set lo = ws.ListObjects.Add(xlSrcRange, grid, , xlYes)
    lo.Name = ComparisonTableName()
    lo.ShowAutoFilterDropDown = False
    ApplyComparisonTableStyle lo
End Sub

' Applies the configured custom table style. Falls back to leaving the default style
' when the style has not been defined in the workbook yet, so a missing style only
' affects the look of the grid rather than failing the whole build.
Private Sub ApplyComparisonTableStyle(lo As ListObject)
    Dim styleName As String

    styleName = ComparisonTableStyleName()
    If Len(styleName) = 0 Then Exit Sub

    On Error Resume Next
    lo.TableStyle = styleName
    If Err.Number <> 0 Then
        Err.Clear
        ' Fallback keeps the header readable without the custom style.
        lo.HeaderRowRange.Font.Bold = True
    End If
    On Error GoTo 0
End Sub

'=======================================================================================
' Generated area
'=======================================================================================

' Removes whatever the last build left behind. Only the recorded rectangle is touched, so
' anything else on the sheet survives. The extent is cleared twice: Clear removes
' values and most formatting, ClearFormats then catches banding and fills a dismissed
' table style can otherwise leave on cells beyond the new, narrower rectangle.
Private Sub ClearPreviousOutput(ws As Worksheet)
    Dim lo As ListObject
    Dim extent As String
    Dim target As Range

    ' Charts before the data they plot, and both before the cells they sit over.
    DeleteRenderedCharts ws, CHART_NAME_PREFIX
    ClearStagedChartData

    ' Unlist first so the comparison table's style is no longer "owned" by a live
    ' ListObject; once detached, Clear/ClearFormats can actually reset the cells.
    Set lo = TableOrNothing(ComparisonTableName())
    If Not lo Is Nothing Then lo.Unlist

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

' The whole generated area as one rectangle. A rectangle rather than a union of blocks:
' the address is stored in a defined name, and a union's address grows without limit.
Private Sub RecordExtent(covered As Range)
    SettingText(EXTENT_KEY) = covered.Address
End Sub

Private Sub ExtendExtent(ws As Worksheet, extra As Range)
    Dim prior As String
    Dim covered As Range

    prior = SettingText(EXTENT_KEY)
    If Len(prior) = 0 Then
        SettingText(EXTENT_KEY) = extra.Address
        Exit Sub
    End If

    Set covered = ws.Range(prior)
    SettingText(EXTENT_KEY) = ws.Range( _
        ws.Cells(Application.Min(covered.Row, extra.Row), _
                 Application.Min(covered.column, extra.column)), _
        ws.Cells(Application.Max(covered.Row + covered.Rows.count, _
                                 extra.Row + extra.Rows.count) - 1, _
                 Application.Max(covered.column + covered.Columns.count, _
                                 extra.column + extra.Columns.count) - 1)).Address
End Sub

Private Function BlocksBottomRow(anchor As Range, template As Range, _
                                 ByVal scenarioCount As Long) As Long
    If CompareLayoutIsAcross() Then
        BlocksBottomRow = anchor.Row + template.Rows.count - 1
    Else
        BlocksBottomRow = anchor.Row + _
                          (template.Rows.count + CompareBlockGap()) * scenarioCount - _
                          CompareBlockGap() - 1
    End If
End Function

'=======================================================================================
' Small helpers
'=======================================================================================

Private Function LookupOr(map As Object, key As String, fallback As String) As String
    If map.Exists(key) Then LookupOr = CStr(map(key))
    If Len(LookupOr) = 0 Then LookupOr = fallback
End Function
