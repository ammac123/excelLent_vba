Attribute VB_Name = "ScenarioConfig"
Option Explicit

' Declares what a scenario is made of. ScenarioStore holds all the mechanics, so adding
' a fourth input set later is a new entry in ScenarioParts and nothing else.
'
' The active tables are the model's existing input tables. tblDiscretions in particular
' keeps its name, its sheet and its validation rules; it simply gets its contents from
' whichever scenario is loaded. No formula in the workbook needs to change.

' Table that collects the metrics from each scenario run.
Private Const TBL_RESULTS As String = "tblScenarioResults"

' Master list of every metric a run can capture, with the Include flag the user ticks.
Private Const TBL_METRIC_CATALOGUE As String = "tblScenarioMetrics"

' Where the side-by-side comparison is drawn.
Private Const SHEET_COMPARE As String = "10_SCN_Comparison"
Private Const NAME_COMPARE_ANCHOR As String = "compareAnchor"
Private Const NAME_COMPARE_PICKER As String = "compareTemplatePicker"
Private Const NAME_REDEFINE_COMBO As String = "cmbRedefineTemplate"
Private Const TBL_COMPARISON As String = "tblScenarioComparison"

' Where the single-scenario output is drawn. Same template-and-token model as the
' comparison sheet, but one block for the picked scenario instead of one block per
' scenario; also supports chart tokens that drop ChartObjects onto the sheet.
Private Const SHEET_OUTPUT As String = "09_SCN_Output"
Private Const NAME_OUTPUT_ANCHOR As String = "outputAnchor"
Private Const NAME_OUTPUT_PICKER As String = "outputScenarioPicker"
Private Const NAME_OUTPUT_TEMPLATE_PICKER As String = "outputTemplatePicker"
Private Const NAME_OUTPUT_REDEFINE_COMBO As String = "cmbRedefineOutputTemplate"

' Segment attributes for {Segment.<ID>.<Column>} tokens: one row per segment, keyed by
' the segment ID, with every other column an attribute of that segment.
Private Const TBL_SEGMENT_MAP As String = "tblProductMap"
Private Const COL_SEGMENT_MAP_KEY As String = "Pricing_Segment_ID"

' Hidden sheet that stores reusable tokens for the token builder UI.
Private Const SHEET_TOKEN_LIBRARY As String = "98_SCN_TokenLibrary"
Private Const TBL_SAVED_TOKENS As String = "tblSavedTokens"

' Store tables: long format, one row per scenario per input row.
Private Const TBL_STORE_ATL As String = "tblScenarioATL"
Private Const TBL_STORE_BTL As String = "tblScenarioDiscretions"

' Active tables: what the model calculates from.
Private Const TBL_ACTIVE_ATL As String = "tblActiveATL"
Private Const TBL_ACTIVE_BTL As String = "tblActiveDiscretions"

' Templates: the blank default a new scenario starts from.
Private Const TBL_TEMPLATE_ATL As String = "tblATLTemplate"
Private Const TBL_TEMPLATE_BTL As String = "tblDiscretionsTemplate"

Private Const COL_EFFECTIVE_DATE As String = "EffectiveDate"
Private Const NAME_EFFECTIVE_DATE As String = "activeEffectiveDate"

' Reserved ScenarioID for the "current state" pseudo-scenario. It lives in tblScenarios
' alongside real scenarios so compare, token resolution and the results store all work
' without special cases, but it has its own run path and cannot be loaded/saved/deleted
' through the ordinary scenario flows.
Private Const ID_CURRENT_STATE As String = "CURRENT"
Private Const NAME_CURRENT_STATE As String = "Current State"
Private Const DESC_CURRENT_STATE As String = "Baseline read directly from the current-state " & _
                                              "calc sheets. Not editable as a scenario."

Public Function ResultsTableName() As String
    ResultsTableName = TBL_RESULTS
End Function

Public Function MetricCatalogueTableName() As String
    MetricCatalogueTableName = TBL_METRIC_CATALOGUE
End Function

'=======================================================================================
' Current state
'=======================================================================================

Public Function CurrentStateScenarioId() As String
    CurrentStateScenarioId = ID_CURRENT_STATE
End Function

Public Function CurrentStateScenarioName() As String
    CurrentStateScenarioName = NAME_CURRENT_STATE
End Function

Public Function CurrentStateScenarioDescription() As String
    CurrentStateScenarioDescription = DESC_CURRENT_STATE
End Function

' Sheets to recalculate before harvesting current-state metrics. One RecalcSpec per sheet,
' in the order they should fire. Add or remove entries here as the current-state model
' grows; nothing else in the codebase needs to change.
Public Function CurrentStateRecalcs() As RecalcSpec()
    Dim specs(1 To 1) As RecalcSpec

    specs(1) = NewRecalcSpec(key:="CurrentStateModel", _
                             targetSheet:="33_MDL_CurrentMetrics", _
                             statusCell:="", _
                             label:="Current state metrics")

    CurrentStateRecalcs = specs
End Function

'=======================================================================================
' Metrics
'=======================================================================================

' The output tables a run harvests. Every data row x column cell in these becomes one
' metric, identified by its label value and column header rather than by position.
'
' A trailing "*" in tableName matches every table whose name starts with that stem, which
' is how the generated tblMetricsOut1, tblMetricsOut2, ... family is picked up without
' having to name each one as it appears.
'
' labelColumn names the column that identifies each row. Leave it empty for a table that
' only ever holds a single row of headline numbers.
Public Function MetricSourceTables() As MetricTableSpec()
    Dim specs(1 To 3) As MetricTableSpec

    ' currentStateTable: parallel output table a current-state run reads instead of the
    ' scenario-driven one. If your current-state calc sheets feed into a different set of
    ' output tables, name them here. Leaving the fourth argument out means the current-state
    ' run reads the regular source, which is fine for metrics whose source does not depend
    ' on the loaded scenario.
    specs(1) = MakeMetricSource("tblMetricsOutSegment", "Pricing_Segment_ID", "Segment Description", _
                                "tblMetricsOutSegment_CURRENT")
    specs(2) = MakeMetricSource("tblMetricsOutProduct", "Product", "Product Aggregation", _
                                "tblMetricsOutProduct_CURRENT")
    specs(3) = MakeMetricSource("tblMetricsOutRFCMonthly", "Month", _
                            "Monthly RFC", "tblMetricsOutRFCMonthly_CURRENT")


    MetricSourceTables = specs
End Function

' Columns whose headers start with any of these are never captured. Bookkeeping columns
' and anything prefixed "_EX" are noise in a comparison.
Public Function MetricColumnExcludePrefixes() As String()
    Dim prefixes(1 To 5) As String

    prefixes(1) = "ScenarioID"
    prefixes(2) = "RunOn"
    prefixes(3) = "Notes"
    prefixes(4) = "Comment"
    prefixes(5) = "_EX"

    MetricColumnExcludePrefixes = prefixes
End Function

' Called after the model recalculates and before a run reads its metrics. Anything that
' builds an output table out of the loaded inputs has to be rebuilt here, otherwise the
' run captures the previous scenario's numbers.
'
' The settlement projection is written as values, and the metric tables calculate off
' it, so the run is still in manual calculation with those dependents out of date.
' Application.Calculate settles only what changed (plus volatiles), not the whole model
' again.
Public Sub PrepareScenarioOutputs()
    GenerateDiscretionSummary
    RefreshSettlementProjection
    Application.Calculate
End Sub

'=======================================================================================
' Comparison view
'=======================================================================================

Public Function CompareSheetName() As String
    CompareSheetName = SHEET_COMPARE
End Function

' The sample summary blocks. One is picked at a time and copied once per selected
' scenario, so how a summary looks is a spreadsheet job rather than a code job. Each
' nameRef is a defined name covering the block; the label is what the picker offers.
Public Function CompareTemplates() As CompareTemplateSpec()
    Dim templates(1 To 3) As CompareTemplateSpec

    templates(1) = MakeCompareTemplate("Headline summary", "compareTemplateHeadline")
    templates(2) = MakeCompareTemplate("Discretion grid", "compareTemplateGrid")
    templates(3) = MakeCompareTemplate("Full detail", "compareTemplateDetail")

    CompareTemplates = templates
End Function

' Cell holding the chosen template's label. Its dropdown is rebuilt from CompareTemplates
' on every build, so adding a fourth template here is the only change needed.
Public Function CompareTemplatePickerName() As String
    CompareTemplatePickerName = NAME_COMPARE_PICKER
End Function

' Form-control combo box on the template design sheet whose selection picks which template
' range the "redefine" button acts on. Shares its list with CompareTemplates().
Public Function RedefineTemplateComboName() As String
    RedefineTemplateComboName = NAME_REDEFINE_COMBO
End Function

' Top-left cell of the generated area. The first scenario's block lands here.
Public Function CompareAnchorName() As String
    CompareAnchorName = NAME_COMPARE_ANCHOR
End Function

' Tidy metric-by-scenario grid built underneath the blocks.
Public Function ComparisonTableName() As String
    ComparisonTableName = TBL_COMPARISON
End Function

' Custom table style applied to the generated comparison table. The style has to be
' defined in the workbook (New Table Style... in the Table Design ribbon) before a
' rebuild can pick it up; a missing style is reported and the table keeps the default.
Public Function ComparisonTableStyleName() As String
    ComparisonTableStyleName = "TableStyle Compare"
End Function

' True lays the scenario blocks out left to right, False stacks them top to bottom.
Public Function CompareLayoutIsAcross() As Boolean
    CompareLayoutIsAcross = True
End Function

' Blank rows or columns left between one scenario's block and the next.
Public Function CompareBlockGap() As Long
    CompareBlockGap = 1
End Function

' Fill for the highlight frame drawn on the ring of cells around each summary and output
' template range. The frame sits outside the range, so it never reaches a rendered block.
Public Function TemplateFrameColor() As Long
    TemplateFrameColor = RGB(255, 230, 153)
End Function

'=======================================================================================
' Scenario output view
'=======================================================================================

Public Function OutputSheetName() As String
    OutputSheetName = SHEET_OUTPUT
End Function

' Top-left cell of the rendered block. The picked scenario's block lands here.
Public Function OutputAnchorName() As String
    OutputAnchorName = NAME_OUTPUT_ANCHOR
End Function

' Cell holding the scenario the user wants to see. Empty / unknown falls back to the
' currently loaded scenario, so opening the sheet without touching anything shows what
' the model is already configured for.
Public Function OutputScenarioPickerName() As String
    OutputScenarioPickerName = NAME_OUTPUT_PICKER
End Function

' Cell holding the chosen output-template label. Rebuilt from ScenarioOutputTemplates()
' every build, so adding a template there is the only change needed.
Public Function OutputTemplatePickerName() As String
    OutputTemplatePickerName = NAME_OUTPUT_TEMPLATE_PICKER
End Function

' Form-control combo box on the output template design area whose selection picks which
' template range the "redefine" button acts on. Shares its list with
' ScenarioOutputTemplates().
Public Function OutputRedefineComboName() As String
    OutputRedefineComboName = NAME_OUTPUT_REDEFINE_COMBO
End Function

'=======================================================================================
' Segment attributes
'=======================================================================================

' Table that {Segment.<ID>.<Column>} tokens read from. It is read live at render time,
' never copied into a scenario, so editing it changes every template on the next build.
Public Function SegmentMapTableName() As String
    SegmentMapTableName = TBL_SEGMENT_MAP
End Function

' Column holding the segment ID each row is keyed by. Every other column is offered as
' an attribute.
Public Function SegmentMapKeyColumn() As String
    SegmentMapKeyColumn = COL_SEGMENT_MAP_KEY
End Function

'=======================================================================================
' Token builder library
'=======================================================================================

' Hidden sheet that stores the reusable-token library the token builder UI reads and
' writes. ScenarioTokenLibrary creates and maintains the sheet; nothing else should
' write to it.
Public Function TokenLibrarySheetName() As String
    TokenLibrarySheetName = SHEET_TOKEN_LIBRARY
End Function

Public Function SavedTokensTableName() As String
    SavedTokensTableName = TBL_SAVED_TOKENS
End Function

' Templates available on the output sheet. Each is a defined name covering a block laid
' out and formatted by hand, same contract as CompareTemplates(). One block is rendered
' per build (for the picked scenario), so templates can be larger and richer than the
' comparison ones.
Public Function ScenarioOutputTemplates() As CompareTemplateSpec()
    Dim templates(1 To 2) As CompareTemplateSpec

    templates(1) = MakeCompareTemplate("Scenario detail", "outputTemplateDetail")
    templates(2) = MakeCompareTemplate("Executive summary", "outputTemplateExec")

    ScenarioOutputTemplates = templates
End Function

' Named chart specs addressable from an output template via {Chart:<name>} tokens. The
' name is case-insensitive. An inline chart token
'   {Chart:kind=Line;source=tblMetricsOutRFCMonthly;x=Month;series=Value;baselineSource=tblMetricsOutRFCMonthly_CURRENT;baselineSeries=Value;title=Monthly RFC}
' overrides anything configured here, so a template can declare one-off charts without
' editing code. Series must be real column names of the referenced table; the output
' sheet re-renders each build, so changing a spec here flows through at the next click.
Public Function ScenarioOutputCharts() As ChartSpec()
    Dim specs(1 To 4) As ChartSpec

    specs(1) = MakeChartSpec(label:="Monthly RFC vs baseline", _
                             kind:="Line", _
                             sourceTable:="tblMetricsOutRFCMonthly", _
                             xColumn:="Month", _
                             seriesColumns:="All_App_Volumes", _
                             baselineTable:="tblMetricsOutRFCMonthly_CURRENT", _
                             baselineColumns:="All_App_Volumes", _
                             title:="Monthly RFC Apps: scenario vs current state", _
                             numberFormat:="#,##0", _
                             showLegend:=True, _
                             heightRows:=18, _
                             widthCols:=10)

    specs(2) = MakeChartSpec(label:="Monthly RFC vs baseline - OA", _
                             kind:="Line", _
                             sourceTable:="tblMetricsOutRFCMonthly", _
                             xColumn:="Month", _
                             seriesColumns:="OA_App_Volumes", _
                             baselineTable:="tblMetricsOutRFCMonthly_CURRENT", _
                             baselineColumns:="OA_App_Volumes", _
                             title:="Monthly RFC OA Apps: scenario vs current state", _
                             numberFormat:="#,##0", _
                             showLegend:=True, _
                             heightRows:=18, _
                             widthCols:=10)


    specs(3) = MakeChartSpec(label:="Monthly RFC vs baseline - MS", _
                             kind:="Line", _
                             sourceTable:="tblMetricsOutRFCMonthly", _
                             xColumn:="Month", _
                             seriesColumns:="MS_App_Volumes", _
                             baselineTable:="tblMetricsOutRFCMonthly_CURRENT", _
                             baselineColumns:="MS_App_Volumes", _
                             title:="Monthly RFC MS Apps: scenario vs current state", _
                             numberFormat:="#,##0", _
                             showLegend:=True, _
                             heightRows:=18, _
                             widthCols:=10)


    specs(4) = MakeChartSpec(label:="Segment NIM", _
                             kind:="Column", _
                             sourceTable:="tblMetricsOutSegment", _
                             xColumn:="Segment Description", _
                             seriesColumns:="NIM", _
                             title:="Segment NIM", _
                             heightRows:=14, _
                             widthCols:=10)

    ScenarioOutputCharts = specs
End Function

' Order matters only for the order of progress messages.
Public Function ScenarioParts() As ScenarioPart()
    Dim parts(1 To 2) As ScenarioPart

    parts(1) = MakePart("ATL price adjustments", TBL_STORE_ATL, TBL_ACTIVE_ATL, TBL_TEMPLATE_ATL)
    parts(2) = MakePart("BTL discretions", TBL_STORE_BTL, TBL_ACTIVE_BTL, TBL_TEMPLATE_BTL)

    ScenarioParts = parts
End Function

Public Function ScenarioScalars() As ScenarioScalar()
    Dim scalars(1 To 1) As ScenarioScalar

    scalars(1) = MakeScalar("effective date", COL_EFFECTIVE_DATE, NAME_EFFECTIVE_DATE)

    ScenarioScalars = scalars
End Function

' Output tables whose contents count as part of the active scenario's state. An edit
' inside one of these marks the active scenario dirty, alongside edits to the input
' tables declared in ScenarioParts() and the scalars declared in ScenarioScalars().
' Edits anywhere else on an input or output sheet are treated as navigation or
' scratch work and leave the dirty flag alone.
Public Function ScenarioDirtyOutputTables() As Variant
    ScenarioDirtyOutputTables = Array( _
        "tblMetricsOutSegment", _
        "tblMetricsOutProduct", _
        "tblMetricsOutRFCMonthly", _
        "tblActiveATL", _
        "tblActiveDiscretions", _
        "tblPriceElasticity", _
        "tblCurrentDiscretions")
End Function

' Tables whose filters are cleared whenever code looks them up. These are rewritten as
' whole blocks, and a filter hiding rows during a read, delete or resize loses data.
Public Function UnfilteredTableNames() As Variant
    UnfilteredTableNames = Array( _
        TBL_STORE_ATL, _
        TBL_STORE_BTL, _
        TBL_RESULTS, _
        TBL_METRIC_CATALOGUE)
End Function

' What a brand new scenario starts with when the caller does not supply a value.
Public Function ScenarioScalarDefault(scalar As ScenarioScalar) As Variant
    Select Case scalar.column
        Case COL_EFFECTIVE_DATE
            ScenarioScalarDefault = Date
        Case Else
            ScenarioScalarDefault = Empty
    End Select
End Function

' Gate for Save and Run. Returns False and fills problem with what to fix.
Public Function ActiveScenarioIsValid(ByRef problem As String) As Boolean
    Dim cell As Range
    Dim lo As ListObject
    Dim summary As String

    problem = ""

    Set cell = NamedRangeOrNothing(NAME_EFFECTIVE_DATE)
    If cell Is Nothing Then
        problem = "The named range '" & NAME_EFFECTIVE_DATE & "' is missing from this workbook."
        Exit Function
    End If
    If Not IsDate(cell.value) Then
        problem = "The effective date is blank or is not a date."
        Exit Function
    End If

    Set lo = TableOrNothing(TBL_ACTIVE_ATL)
    If lo Is Nothing Then
        problem = "Table '" & TBL_ACTIVE_ATL & "' is missing from this workbook."
        Exit Function
    End If
    If TableRowCount(lo) = 0 Then
        problem = "There are no ATL price adjustment rows. Load or create a scenario first."
        Exit Function
    End If

    If Not DiscretionsAreValid(TBL_ACTIVE_BTL, summary) Then
        problem = "The BTL discretions do not cover every segment cleanly:" & vbNewLine & summary
        Exit Function
    End If

    ActiveScenarioIsValid = True
End Function

Private Function MakePart(label As String, storeTable As String, activeTable As String, _
                          templateTable As String) As ScenarioPart
    With MakePart
        .label = label
        .storeTable = storeTable
        .activeTable = activeTable
        .templateTable = templateTable
    End With
End Function

Private Function MakeScalar(label As String, column As String, nameRef As String) As ScenarioScalar
    With MakeScalar
        .label = label
        .column = column
        .nameRef = nameRef
    End With
End Function

Private Function MakeMetricSource(tableName As String, labelColumn As String, _
                                  groupName As String, _
                                  Optional currentStateTable As String = "") As MetricTableSpec
    With MakeMetricSource
        .tableName = tableName
        .labelColumn = labelColumn
        .groupName = groupName
        .currentStateTable = currentStateTable
    End With
End Function

Private Function MakeCompareTemplate(label As String, nameRef As String) As CompareTemplateSpec
    With MakeCompareTemplate
        .label = label
        .nameRef = nameRef
    End With
End Function

Private Function MakeChartSpec(label As String, kind As String, sourceTable As String, _
                               xColumn As String, seriesColumns As String, _
                               Optional baselineTable As String = "", _
                               Optional baselineColumns As String = "", _
                               Optional title As String = "", _
                               Optional numberFormat As String = "", _
                               Optional showLegend As Boolean = True, _
                               Optional heightRows As Long = 0, _
                               Optional widthCols As Long = 0) As ChartSpec
    With MakeChartSpec
        .label = label
        .kind = kind
        .sourceTable = sourceTable
        .xColumn = xColumn
        .seriesColumns = seriesColumns
        .baselineTable = baselineTable
        .baselineColumns = baselineColumns
        .title = title
        .numberFormat = numberFormat
        .showLegend = showLegend
        .heightRows = heightRows
        .widthCols = widthCols
    End With
End Function


