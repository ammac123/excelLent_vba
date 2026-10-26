Attribute VB_Name = "ScenarioTokenBuilder"
Option Explicit

' Business logic for the token builder UI. The form (frmTokenBuilder) owns event
' handling and clipboard delivery; this module owns composing, parsing, validating
' and choice lookups so the compose grammar stays testable from .bas-land and never
' drifts from what ScenarioTokens / ScenarioOutput actually accept.
'
' Entry point:
'
'   OpenTokenBuilder()      Loads frmTokenBuilder and shows it modelessly. Multiple
'                           Shape buttons on different sheets can call this; the
'                           form never activates another sheet or alters the
'                           selection, so the user's context is preserved. If the
'                           form is already open, it is brought back to the front
'                           instead of opening a second copy.
'
' Compose: one function per token kind, each returning the exact grammar string.
'
'   ComposeMetricToken(flavour, metricId, [numberFormat])
'   ComposeScenarioToken(column, [numberFormat])
'                                   numberFormat, when set, is appended as
'                                   "|<format>" per the ScenarioTokens grammar.
'   ComposeSegmentToken(segmentId, column, [numberFormat])
'   ComposeChartNamedToken(label)
'   ComposeChartInlineToken(spec)   emits kind=;source=;x=;series=;... matching
'                                   ScenarioOutput.ParseInlineChartSpec.
'
' Reverse direction:
'
'   ParseTokenForEdit(token)        Returns a ParsedToken filled with whatever the
'                                   form needs to repopulate its controls, so the
'                                   Edit flow on the Saved tokens tab round-trips.
'   ValidateSavedToken(token)       Returns "" when the token references nothing
'                                   missing, otherwise a short reason suitable for
'                                   the yellow banner. Never blocks the UI.
'
' Choice lookups: hand the form the lists it needs to populate dropdowns.
'
'   CatalogueChoices()              every live metric in the catalogue
'   ScenarioFieldChoices()          built-in field names + custom columns on
'                                   tblScenarios, with bookkeeping columns hidden
'   NamedChartChoices()             entries from ScenarioOutputCharts()
'   SourceTableChoices()            every ListObject in the workbook
'   TableColumnChoices(tableName)   that table's column names
'   SegmentIdChoices()              segment IDs in the segment map, in table order
'   SegmentColumnChoices()          the segment map's attribute columns
'   SegmentPreview(segmentId, col)  the live value a Segment token would show
'   MetricById(metricId)            the live MetricSpec for a given id, or an
'                                   empty struct if the metric no longer exists

Public Const KIND_METRIC As String = "Metric"
Public Const KIND_CURRENT As String = "Current"
Public Const KIND_DELTA As String = "Delta"
Public Const KIND_DELTA_PCT As String = "DeltaPct"
Public Const KIND_LABEL As String = "Label"
Public Const KIND_SCENARIO As String = "Scenario"
Public Const KIND_SEGMENT As String = "Segment"
Public Const KIND_CHART_NAMED As String = "ChartNamed"
Public Const KIND_CHART_INLINE As String = "ChartInline"
Public Const KIND_UNKNOWN As String = "Unknown"

' Placeholder keys for the metric picker; see "Metric picker" below.
Public Const METRIC_SOURCE_NAMED As String = "(named ranges)"
Public Const METRIC_SEGMENT_NONE As String = "(no segment)"
Public Const METRIC_SEGMENT_ALL As String = "(all segments)"

Public Type ParsedToken
    kind As String          ' one of the KIND_* constants
    metricId As String      ' for Metric / Current / Delta / DeltaPct / Label
    column As String        ' for Scenario, and the attribute column for Segment
    segmentId As String     ' for Segment
    numberFormat As String  ' the "|<format>" suffix on a metric, scenario or segment token
    chartLabel As String    ' for ChartNamed
    chartSpec As ChartSpec  ' for ChartInline
End Type

'=======================================================================================
' Entry
'=======================================================================================

Public Sub OpenTokenBuilder()
    Dim frm As Object

    EnsureLibrarySheet

    ' UserForms.Add always creates a fresh instance, so check for one that's already
    ' open first. Calling Show on a loaded modeless form re-shows it if hidden and
    ' activates it, which brings it to the front with focus.
    Set frm = LoadedTokenBuilder()
    If Not frm Is Nothing Then
        frm.Show vbModeless
        Exit Sub
    End If

    ' A missing form raises; we catch and offer the bootstrap rather than crashing
    ' out of a Shape button click.
    On Error GoTo FormMissing
    UserForms.Add("frmTokenBuilder").Show vbModeless
    Exit Sub

FormMissing:
    MsgBox "frmTokenBuilder is not in this workbook yet." & vbNewLine & vbNewLine & _
           "See frmTokenBuilder_SETUP.md in the project folder for step-by-step " & _
           "instructions on building it in the VBE. Once it's in place, save the " & _
           "workbook and this button will open it.", _
           vbInformation, "Token builder"
End Sub

'=======================================================================================
' Compose
'=======================================================================================

' flavour: one of KIND_METRIC/KIND_CURRENT/KIND_DELTA/KIND_DELTA_PCT/KIND_LABEL.
' Returns "" when the inputs are not enough to form a valid token so the preview
' can stay empty in a half-filled compose state.
Public Function ComposeMetricToken(flavour As String, metricId As String, _
                                   Optional numberFormat As String = "") As String
    Dim prefix As String

    If Len(Trim$(metricId)) = 0 Then Exit Function
    Select Case flavour
        Case KIND_METRIC:    prefix = "Metric."
        Case KIND_CURRENT:   prefix = "Current."
        Case KIND_DELTA:     prefix = "Delta."
        Case KIND_DELTA_PCT: prefix = "DeltaPct."
        Case KIND_LABEL:     prefix = "Label."
        Case Else:           Exit Function
    End Select
    ComposeMetricToken = TokenOpen() & prefix & Trim$(metricId) & _
                         FormatSuffix(numberFormat) & TokenClose()
End Function

Public Function ComposeScenarioToken(column As String, _
                                     Optional numberFormat As String = "") As String
    If Len(Trim$(column)) = 0 Then Exit Function
    ComposeScenarioToken = TokenOpen() & "Scenario." & Trim$(column) & _
                           FormatSuffix(numberFormat) & TokenClose()
End Function

Public Function ComposeSegmentToken(segmentId As String, column As String, _
                                    Optional numberFormat As String = "") As String
    If Len(Trim$(segmentId)) = 0 Or Len(Trim$(column)) = 0 Then Exit Function
    ComposeSegmentToken = TokenOpen() & "Segment." & Trim$(segmentId) & "." & Trim$(column) & _
                          FormatSuffix(numberFormat) & TokenClose()
End Function

Private Function FormatSuffix(numberFormat As String) As String
    If Len(Trim$(numberFormat)) = 0 Then Exit Function
    FormatSuffix = TokenFormatSeparator() & Trim$(numberFormat)
End Function

Public Function ComposeChartNamedToken(label As String) As String
    If Len(Trim$(label)) = 0 Then Exit Function
    ComposeChartNamedToken = TokenOpen() & "Chart:" & Trim$(label) & TokenClose()
End Function

' Only the required three (source/x/series) are mandatory; the rest fold in only when
' set. Grammar matches ScenarioOutput.ParseInlineChartSpec so a composed token
' round-trips cleanly.
Public Function ComposeChartInlineToken(spec As ChartSpec) As String
    Dim parts As String
    Dim kind As String

    If Len(Trim$(spec.sourceTable)) = 0 Then Exit Function
    If Len(Trim$(spec.xColumn)) = 0 Then Exit Function
    If Len(Trim$(spec.seriesColumns)) = 0 Then Exit Function

    kind = Trim$(spec.kind)
    If Len(kind) = 0 Then kind = "Line"

    parts = "kind=" & kind
    parts = parts & ";source=" & Trim$(spec.sourceTable)
    parts = parts & ";x=" & Trim$(spec.xColumn)
    parts = parts & ";series=" & Trim$(spec.seriesColumns)

    If Len(Trim$(spec.baselineTable)) > 0 Then _
        parts = parts & ";baselineSource=" & Trim$(spec.baselineTable)
    If Len(Trim$(spec.baselineColumns)) > 0 Then _
        parts = parts & ";baselineSeries=" & Trim$(spec.baselineColumns)
    If Len(Trim$(spec.title)) > 0 Then parts = parts & ";title=" & Trim$(spec.title)
    If Len(Trim$(spec.numberFormat)) > 0 Then parts = parts & ";format=" & Trim$(spec.numberFormat)
    If Not spec.showLegend Then parts = parts & ";legend=N"
    If spec.heightRows > 0 Then parts = parts & ";height=" & spec.heightRows
    If spec.widthCols > 0 Then parts = parts & ";width=" & spec.widthCols

    ComposeChartInlineToken = TokenOpen() & "Chart:" & parts & TokenClose()
End Function

'=======================================================================================
' Parse
'=======================================================================================

Public Function ParseTokenForEdit(token As String) As ParsedToken
    Dim trimmed As String, inner As String, body As String
    Dim fmt As String
    Dim segmentId As String, column As String

    ParseTokenForEdit.kind = KIND_UNKNOWN
    trimmed = Trim$(token)
    If Len(trimmed) < 3 Then Exit Function
    If Left$(trimmed, 1) <> TokenOpen() Then Exit Function
    If Right$(trimmed, 1) <> TokenClose() Then Exit Function

    inner = Trim$(Mid$(trimmed, 2, Len(trimmed) - 2))

    If StartsWithCI(inner, "Chart:") Then
        body = Mid$(inner, Len("Chart:") + 1)
        If InStr(body, "=") = 0 Then
            ParseTokenForEdit.kind = KIND_CHART_NAMED
            ParseTokenForEdit.chartLabel = Trim$(body)
        Else
            ParseTokenForEdit.kind = KIND_CHART_INLINE
            ParseTokenForEdit.chartSpec = ParseInlineChartSpec(body)
        End If
        Exit Function
    End If

    ' Charts are handled above because an inline spec carries its own format= key and
    ' its title may legitimately contain a "|".
    inner = SplitTokenFormat(inner, fmt)
    ParseTokenForEdit.numberFormat = fmt

    If StartsWithCI(inner, "Scenario.") Then
        ParseTokenForEdit.kind = KIND_SCENARIO
        ParseTokenForEdit.column = Mid$(inner, Len("Scenario.") + 1)
        Exit Function
    End If

    If StartsWithCI(inner, "Segment.") Then
        SplitSegmentBody Mid$(inner, Len("Segment.") + 1), segmentId, column
        ParseTokenForEdit.kind = KIND_SEGMENT
        ParseTokenForEdit.segmentId = segmentId
        ParseTokenForEdit.column = column
        Exit Function
    End If

    ' DeltaPct before Delta.
    If StartsWithCI(inner, "DeltaPct.") Then
        ParseTokenForEdit.kind = KIND_DELTA_PCT
        ParseTokenForEdit.metricId = Mid$(inner, Len("DeltaPct.") + 1)
        Exit Function
    End If
    If StartsWithCI(inner, "Delta.") Then
        ParseTokenForEdit.kind = KIND_DELTA
        ParseTokenForEdit.metricId = Mid$(inner, Len("Delta.") + 1)
        Exit Function
    End If
    If StartsWithCI(inner, "Current.") Then
        ParseTokenForEdit.kind = KIND_CURRENT
        ParseTokenForEdit.metricId = Mid$(inner, Len("Current.") + 1)
        Exit Function
    End If
    If StartsWithCI(inner, "Label.") Then
        ParseTokenForEdit.kind = KIND_LABEL
        ParseTokenForEdit.metricId = Mid$(inner, Len("Label.") + 1)
        Exit Function
    End If
    If StartsWithCI(inner, "Metric.") Then
        ParseTokenForEdit.kind = KIND_METRIC
        ParseTokenForEdit.metricId = Mid$(inner, Len("Metric.") + 1)
        Exit Function
    End If
End Function

'=======================================================================================
' Validate
'=======================================================================================

Public Function ValidateSavedToken(token As String) As String
    Dim parsed As ParsedToken
    Dim built As Boolean
    Dim lo As ListObject
    Dim specs() As ChartSpec
    Dim i As Long

    parsed = ParseTokenForEdit(token)

    ' A "}" ends the token early, so everything after it would be pasted as plain text.
    If InStr(parsed.numberFormat, TokenClose()) > 0 Then
        ValidateSavedToken = "A number format cannot contain '" & TokenClose() & "'."
        Exit Function
    End If

    Select Case parsed.kind
        Case KIND_METRIC, KIND_CURRENT, KIND_DELTA, KIND_DELTA_PCT, KIND_LABEL
            If IsWildcardMetricId(parsed.metricId) Then
                ValidateSavedToken = ValidateWildcardMetric(parsed.metricId)
            ElseIf Len(MetricById(parsed.metricId).metric) = 0 Then
                ValidateSavedToken = "Metric '" & parsed.metricId & _
                                     "' is no longer in the catalogue."
            End If

        Case KIND_SCENARIO
            Select Case UCase$(parsed.column)
                Case "NAME", "ID", "INDEX", "RUNON": built = True
            End Select
            If Not built Then
                Set lo = TableOrNothing(TBL_SCENARIOS)
                If lo Is Nothing Then
                    ValidateSavedToken = "Table '" & TBL_SCENARIOS & "' is missing."
                ElseIf ColumnIndexOrZero(lo, parsed.column) = 0 Then
                    ValidateSavedToken = "Scenario column '" & parsed.column & _
                                         "' is no longer on " & TBL_SCENARIOS & "."
                End If
            End If

        Case KIND_SEGMENT
            ValidateSavedToken = ValidateSegmentToken(parsed.segmentId, parsed.column)

        Case KIND_CHART_NAMED
            specs = ScenarioOutputCharts()
            For i = LBound(specs) To UBound(specs)
                If StrComp(specs(i).label, parsed.chartLabel, vbTextCompare) = 0 Then _
                    Exit Function
            Next
            ValidateSavedToken = "Chart '" & parsed.chartLabel & _
                                 "' is no longer declared in ScenarioOutputCharts()."

        Case KIND_CHART_INLINE
            ValidateSavedToken = ValidateInlineChart(parsed.chartSpec)

        Case Else
            ValidateSavedToken = "This token does not match any supported pattern."
    End Select
End Function

' A wildcard row renders one copy per ticked match, so matches that are all unticked
' would leave a single blank row.
Private Function ValidateWildcardMetric(pattern As String) As String
    Dim specs() As MetricSpec
    Dim matched As Boolean
    Dim i As Long

    specs = CatalogueSpecs(False)
    For i = 1 To MetricSpecCount(specs)
        If MatchesWildcardMetric(pattern, specs(i)) Then
            If specs(i).include Then Exit Function
            matched = True
        End If
    Next i

    If matched Then
        ValidateWildcardMetric = "Every metric matching '" & pattern & "' is unticked in " & _
                                 MetricCatalogueTableName() & ", so its row renders empty."
    Else
        ValidateWildcardMetric = "No metric in the catalogue matches '" & pattern & "'."
    End If
End Function

Private Function ValidateSegmentToken(segmentId As String, column As String) As String
    Dim lo As ListObject

    If Len(segmentId) = 0 Or Len(column) = 0 Then
        ValidateSegmentToken = "A segment token needs both an ID and a column: " & _
                               "{Segment.<ID>.<Column>}."
        Exit Function
    End If

    Set lo = TableOrNothing(SegmentMapTableName())
    If lo Is Nothing Then
        ValidateSegmentToken = "Table '" & SegmentMapTableName() & "' is missing."
    ElseIf ColumnIndexOrZero(lo, SegmentMapKeyColumn()) = 0 Then
        ValidateSegmentToken = "Key column '" & SegmentMapKeyColumn() & "' is missing on " & _
                               SegmentMapTableName() & "."
    ElseIf ColumnIndexOrZero(lo, column) = 0 Then
        ValidateSegmentToken = "Column '" & column & "' is no longer on " & _
                               SegmentMapTableName() & "."
    Else
        ' Re-read so the banner reflects edits made to the table while the form is open.
        ForgetSegmentMap
        If Not SegmentMapHasId(segmentId) Then
            ValidateSegmentToken = "Segment '" & segmentId & "' is no longer in " & _
                                   SegmentMapTableName() & "."
        End If
    End If
End Function

Private Function ValidateInlineChart(spec As ChartSpec) As String
    Dim lo As ListObject
    Dim parts() As String
    Dim i As Long
    Dim col As String

    Set lo = TableOrNothing(spec.sourceTable)
    If lo Is Nothing Then
        ValidateInlineChart = "Chart source table '" & spec.sourceTable & "' is missing."
        Exit Function
    End If
    If ColumnIndexOrZero(lo, spec.xColumn) = 0 Then
        ValidateInlineChart = "X column '" & spec.xColumn & "' is missing on " & _
                              spec.sourceTable & "."
        Exit Function
    End If

    parts = Split(spec.seriesColumns, ",")
    For i = LBound(parts) To UBound(parts)
        col = Trim$(parts(i))
        If Len(col) > 0 And ColumnIndexOrZero(lo, col) = 0 Then
            ValidateInlineChart = "Series column '" & col & "' is missing on " & _
                                  spec.sourceTable & "."
            Exit Function
        End If
    Next

    If Len(spec.baselineTable) = 0 Then Exit Function

    Set lo = TableOrNothing(spec.baselineTable)
    If lo Is Nothing Then
        ValidateInlineChart = "Baseline table '" & spec.baselineTable & "' is missing."
        Exit Function
    End If

    If Len(spec.baselineColumns) = 0 Then Exit Function

    parts = Split(spec.baselineColumns, ",")
    For i = LBound(parts) To UBound(parts)
        col = Trim$(parts(i))
        If Len(col) > 0 And ColumnIndexOrZero(lo, col) = 0 Then
            ValidateInlineChart = "Baseline series '" & col & "' is missing on " & _
                                  spec.baselineTable & "."
            Exit Function
        End If
    Next
End Function

'=======================================================================================
' Choice lookups
'=======================================================================================

Public Function CatalogueChoices() As MetricSpec()
    CatalogueChoices = CatalogueSpecs(False)
End Function

' Returns the live MetricSpec for metricId, or a zero-filled spec when the metric is
' not in the catalogue. Callers test via Len(result.metric).
'
' A wildcard ID (see ScenarioTokens) comes back as its first match with the row key
' swapped for the key pattern, so the picker reads it as (all segments) or "t1*" and
' the readout shows e.g. "* NIM".
Public Function MetricById(metricId As String) As MetricSpec
    Dim specs() As MetricSpec
    Dim keyPattern As String
    Dim i As Long

    specs = CatalogueSpecs(False)
    For i = 1 To MetricSpecCount(specs)
        If IsWildcardMetricId(metricId) Then
            If MatchesWildcardMetric(metricId, specs(i)) Then
                keyPattern = WildcardKeyPattern(metricId, specs(i))
                MetricById = specs(i)
                MetricById.metric = metricId
                MetricById.labelValue = keyPattern
                MetricById.label = keyPattern & " " & specs(i).sourceColumn
                Exit Function
            End If
        ElseIf StrComp(specs(i).metric, metricId, vbTextCompare) = 0 Then
            MetricById = specs(i)
            Exit Function
        End If
    Next i
End Function

'---------------------------------------------------------------------------------------
' Metric picker: Source -> Segment -> Metric type
'---------------------------------------------------------------------------------------
'
' The form picks a metric in three steps instead of one long list of IDs. Each step
' reads the same three catalogue fields a scanned metric is built from:
'
'   Source   MetricSpec.sourceTable   e.g. tblMetricsOutSegment
'   Segment  MetricSpec.labelValue    the row key, e.g. 42 under Pricing_Segment_ID
'   Type     MetricSpec.sourceColumn  the column, e.g. NIM
'
' Metrics that don't fit that shape still get a place: named-range metrics sit under
' a "(named ranges)" source and use their metric ID as the type, and single-row tables
' with no label column use "(no segment)". The three Key functions are the only place
' that mapping lives, so the lists and MetricIdFor can't disagree. The placeholder
' constants (METRIC_SOURCE_NAMED, METRIC_SEGMENT_NONE, METRIC_SEGMENT_ALL) are declared at
' the top of the module, since VBA won't accept module-level declarations after the
' first procedure.
'
' A source with row keys also offers "(all segments)" at the top of the Segment list,
' which composes a wildcard token ({Metric.MetricsOutSegment_*_NIM}) whose template row
' repeats once per segment. A segment value holding a "*", such as "t1*" from editing a
' hand-typed token, is treated as a key pattern the same way.

Public Function MetricSourceKey(m As MetricSpec) As String
    If Len(m.sourceTable) = 0 Then
        MetricSourceKey = METRIC_SOURCE_NAMED
    Else
        MetricSourceKey = m.sourceTable
    End If
End Function

Public Function MetricSegmentKey(m As MetricSpec) As String
    If Len(m.labelValue) = 0 Then
        MetricSegmentKey = METRIC_SEGMENT_NONE
    ElseIf m.labelValue = "*" Then
        MetricSegmentKey = METRIC_SEGMENT_ALL
    Else
        MetricSegmentKey = m.labelValue
    End If
End Function

Public Function MetricTypeKey(m As MetricSpec) As String
    If Len(m.sourceTable) = 0 Or Len(m.sourceColumn) = 0 Then
        MetricTypeKey = m.metric
    Else
        MetricTypeKey = m.sourceColumn
    End If
End Function

' Distinct sources, sorted.
Public Function MetricSourceChoices() As String()
    Dim specs() As MetricSpec
    Dim seen As Object
    Dim i As Long

    Set seen = NewTextSet()
    specs = CatalogueSpecs(False)
    For i = 1 To MetricSpecCount(specs)
        AddDistinct seen, MetricSourceKey(specs(i))
    Next i
    MetricSourceChoices = TextSetToArray(seen, True)
End Function

' Distinct segments for one source, in catalogue order (which follows the source
' table's row order, so segments read in the same order as the sheet).
Public Function MetricSegmentChoices(source As String) As String()
    Dim specs() As MetricSpec
    Dim seen As Object, choices As Object
    Dim key As Variant
    Dim wildcardable As Boolean
    Dim i As Long

    Set seen = NewTextSet()
    specs = CatalogueSpecs(False)
    For i = 1 To MetricSpecCount(specs)
        If KeyMatches(MetricSourceKey(specs(i)), source) Then
            AddDistinct seen, MetricSegmentKey(specs(i))
            If Len(WildcardMetricId(specs(i), "*")) > 0 Then wildcardable = True
        End If
    Next i

    ' "(all segments)" leads, so it sits above a long list of IDs.
    Set choices = NewTextSet()
    If wildcardable Then AddDistinct choices, METRIC_SEGMENT_ALL
    For Each key In seen.keys
        AddDistinct choices, CStr(key)
    Next key
    MetricSegmentChoices = TextSetToArray(choices, False)
End Function

' Distinct metric types for one source + segment, in catalogue (column) order.
Public Function MetricTypeChoices(source As String, segment As String) As String()
    Dim specs() As MetricSpec
    Dim seen As Object
    Dim i As Long

    Set seen = NewTextSet()
    specs = CatalogueSpecs(False)
    For i = 1 To MetricSpecCount(specs)
        If KeyMatches(MetricSourceKey(specs(i)), source) And _
           SegmentMatches(specs(i), segment) Then
            AddDistinct seen, MetricTypeKey(specs(i))
        End If
    Next i
    MetricTypeChoices = TextSetToArray(seen, False)
End Function

' The metric ID the three picks point at, or "" until all three resolve to a
' catalogue entry. A segment pattern gives the wildcard ID instead.
Public Function MetricIdFor(source As String, segment As String, metricType As String) As String
    Dim specs() As MetricSpec
    Dim i As Long

    If Len(source) = 0 Or Len(segment) = 0 Or Len(metricType) = 0 Then Exit Function

    specs = CatalogueSpecs(False)
    For i = 1 To MetricSpecCount(specs)
        If KeyMatches(MetricSourceKey(specs(i)), source) And _
           SegmentMatches(specs(i), segment) And _
           KeyMatches(MetricTypeKey(specs(i)), metricType) Then
            If Len(SegmentPattern(segment)) = 0 Then
                MetricIdFor = specs(i).metric
            Else
                MetricIdFor = WildcardMetricId(specs(i), SegmentPattern(segment))
            End If
            Exit Function
        End If
    Next i
End Function

' The key pattern a Segment pick stands for, or "" for a single segment.
Private Function SegmentPattern(segment As String) As String
    If KeyMatches(segment, METRIC_SEGMENT_ALL) Then
        SegmentPattern = "*"
    ElseIf IsWildcardMetricId(segment) Then
        SegmentPattern = segment
    End If
End Function

Private Function SegmentMatches(m As MetricSpec, segment As String) As Boolean
    Dim pattern As String

    pattern = SegmentPattern(segment)
    If Len(pattern) = 0 Then
        SegmentMatches = KeyMatches(MetricSegmentKey(m), segment)
    Else
        SegmentMatches = MatchesWildcardMetric(WildcardMetricId(m, pattern), m)
    End If
End Function

Private Function KeyMatches(a As String, b As String) As Boolean
    KeyMatches = (StrComp(a, b, vbTextCompare) = 0)
End Function

Private Function NewTextSet() As Object
    Set NewTextSet = CreateObject("Scripting.Dictionary")
    NewTextSet.CompareMode = vbTextCompare
End Function

Private Sub AddDistinct(seen As Object, key As String)
    If Not seen.Exists(key) Then seen.Add key, True
End Sub

' An empty set comes back as Split(vbNullString): a real zero-length array, so the
' form's LBound/UBound loops simply don't run instead of raising.
Private Function TextSetToArray(seen As Object, ByVal sorted As Boolean) As String()
    Dim out() As String
    Dim keys As Variant
    Dim i As Long

    If seen.count = 0 Then
        TextSetToArray = Split(vbNullString)
        Exit Function
    End If

    keys = seen.keys
    ReDim out(1 To seen.count)
    For i = 1 To seen.count
        out(i) = CStr(keys(i - 1))
    Next i
    If sorted Then SortStringArray out
    TextSetToArray = out
End Function

' Built-in Scenario-field names first, then any other tblScenarios columns the user
' has added. Bookkeeping columns (IDs, timestamps, user names) are hidden to keep
' the dropdown short.
Public Function ScenarioFieldChoices() As String()
    Dim lo As ListObject, lc As ListColumn
    Dim names() As String
    Dim n As Long

    ReDim names(1 To 32)
    names(1) = "Name"
    names(2) = "ID"
    names(3) = "Index"
    names(4) = "RunOn"
    n = 4

    Set lo = TableOrNothing(TBL_SCENARIOS)
    If Not lo Is Nothing Then
        For Each lc In lo.ListColumns
            If Not IsBookkeepingColumn(lc.Name) Then
                n = n + 1
                If n > UBound(names) Then ReDim Preserve names(1 To n + 16)
                names(n) = lc.Name
            End If
        Next
    End If

    ReDim Preserve names(1 To n)
    ScenarioFieldChoices = names
End Function

Private Function IsBookkeepingColumn(name As String) As Boolean
    Select Case LCase$(name)
        Case LCase$(COL_SCENARIO_ID), LCase$(COL_NAME), _
             LCase$(COL_CREATED_ON), LCase$(COL_MODIFIED_ON), _
             LCase$(COL_CREATED_BY), LCase$(COL_MODIFIED_BY)
            IsBookkeepingColumn = True
    End Select
End Function

Public Function NamedChartChoices() As ChartSpec()
    NamedChartChoices = ScenarioOutputCharts()
End Function

Public Function SourceTableChoices() As String()
    Dim ws As Worksheet, lo As ListObject
    Dim names() As String
    Dim n As Long

    ReDim names(1 To 128)
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            n = n + 1
            If n > UBound(names) Then ReDim Preserve names(1 To n + 64)
            names(n) = lo.Name
        Next
    Next
    If n = 0 Then
        SourceTableChoices = Split(vbNullString)
        Exit Function
    End If

    ReDim Preserve names(1 To n)
    SortStringArray names
    SourceTableChoices = names
End Function

Public Function TableColumnChoices(tableName As String) As String()
    Dim lo As ListObject
    Dim names() As String
    Dim i As Long, n As Long

    TableColumnChoices = Split(vbNullString)

    Set lo = TableOrNothing(tableName)
    If lo Is Nothing Then Exit Function

    n = lo.ListColumns.count
    If n = 0 Then Exit Function

    ReDim names(1 To n)
    For i = 1 To n
        names(i) = lo.ListColumns(i).Name
    Next
    TableColumnChoices = names
End Function

'---------------------------------------------------------------------------------------
' Segment picker: ID -> Column
'---------------------------------------------------------------------------------------

' Segment IDs in the segment map's own row order.
Public Function SegmentIdChoices() As String()
    ForgetSegmentMap
    SegmentIdChoices = SegmentMapIds()
End Function

' Every column of the segment map except the key, in table order.
Public Function SegmentColumnChoices() As String()
    Dim lo As ListObject
    Dim names() As String
    Dim i As Long, n As Long

    SegmentColumnChoices = Split(vbNullString)

    Set lo = TableOrNothing(SegmentMapTableName())
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.count < 2 Then Exit Function

    ReDim names(1 To lo.ListColumns.count)
    For i = 1 To lo.ListColumns.count
        If StrComp(lo.ListColumns(i).Name, SegmentMapKeyColumn(), vbTextCompare) <> 0 Then
            n = n + 1
            names(n) = lo.ListColumns(i).Name
        End If
    Next i
    If n = 0 Then Exit Function

    ReDim Preserve names(1 To n)
    SegmentColumnChoices = names
End Function

' The live value a Segment token would render, as text for the form's readout. Blank
' until both picks are made; the missing marker when they don't resolve.
Public Function SegmentPreview(segmentId As String, column As String) As String
    If Len(segmentId) = 0 Or Len(column) = 0 Then Exit Function
    ForgetSegmentMap
    SegmentPreview = CStr(SegmentAttribute(segmentId, column, TokenMissingText()))
End Function

'=======================================================================================
' Caption-reset callbacks (Application.OnTime targets)
'=======================================================================================
'
' The form schedules these so the "Copied ✅" / "Saved ✅" caption flashes revert
' automatically. They're public because Application.OnTime cannot target a form's
' private handler. Each sub walks the loaded forms to find frmTokenBuilder and ask
' it to reset; if the form has been unloaded between the schedule and the fire, we
' silently exit.

Public Sub ResetCopyCaption()
    CallFormResetter "ResetCopyCaption"
End Sub

Public Sub ResetSaveCaption()
    CallFormResetter "ResetSaveCaption"
End Sub

Public Sub ResetLibCopyCaption()
    CallFormResetter "ResetLibCopyCaption"
End Sub

Private Sub CallFormResetter(methodName As String)
    Dim frm As Object

    Set frm = LoadedTokenBuilder()
    If frm Is Nothing Then Exit Sub

    On Error Resume Next
    CallByName frm, methodName, VbMethod
    On Error GoTo 0
End Sub

' Returns the loaded frmTokenBuilder instance, or Nothing if it isn't open. Walks
' UserForms by TypeName so this compiles even before the form has been imported.
Private Function LoadedTokenBuilder() As Object
    Dim i As Long

    For i = 0 To UserForms.count - 1
        If StrComp(TypeName(UserForms(i)), "frmTokenBuilder", vbTextCompare) = 0 Then
            Set LoadedTokenBuilder = UserForms(i)
            Exit Function
        End If
    Next i
End Function

'=======================================================================================
' Small helpers
'=======================================================================================

Private Function StartsWithCI(text As String, prefix As String) As Boolean
    If Len(text) < Len(prefix) Then Exit Function
    StartsWithCI = (StrComp(Left$(text, Len(prefix)), prefix, vbTextCompare) = 0)
End Function

' In-place alphabetical sort. Simple insertion sort suits the small arrays the
' choice lookups produce.
Private Sub SortStringArray(ByRef values() As String)
    Dim i As Long, j As Long
    Dim tmp As String

    For i = LBound(values) + 1 To UBound(values)
        tmp = values(i)
        j = i - 1
        Do While j >= LBound(values)
            If StrComp(values(j), tmp, vbTextCompare) <= 0 Then Exit Do
            values(j + 1) = values(j)
            j = j - 1
        Loop
        values(j + 1) = tmp
    Next i
End Sub
