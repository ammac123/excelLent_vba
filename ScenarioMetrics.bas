Attribute VB_Name = "ScenarioMetrics"
Option Explicit

' The metric catalogue: the master list of every number a scenario run captures.
'
' A scenario's outputs are spread over several tables, one per area of the model. The
' catalogue flattens them into a single list of "metric ID -> where to read it", so a
' run becomes "read every included metric and stamp it with the ScenarioID", and the
' comparison view becomes "pick some metrics and pivot them by scenario".
'
' Three tables are involved and they are easy to confuse:
'
'   source tables       The model's own output tables, declared by MetricSourceTables()
'                       in ScenarioConfig. They hold live values for whatever scenario is
'                       loaded right now and carry no ScenarioID. Scanned, never written.
'   tblScenarioMetrics  The catalogue. One row per metric, with an Include flag the user
'                       ticks to choose what gets captured. SyncMetricCatalogue rebuilds
'                       it from the source tables, preserving the flags already set.
'   tblScenarioResults  The long store. One row per scenario per metric, rewritten each
'                       time that scenario runs.
'
' A metric addresses its source by label value and column header rather than by row and
' column number, so inserting or reordering rows in a source table cannot silently
' repoint a metric at somebody else's number.

Public Type MetricSpec
    metric As String         ' unique key; this is what lands in tblScenarioResults
    label As String          ' display text for summaries and chart categories
    groupName As String      ' groups metrics into one summary section and one chart
    nameRef As String        ' defined name source; empty when table sourced
    sourceTable As String    ' table source; empty when name sourced
    labelColumn As String    ' column in sourceTable holding the row key
    labelValue As String     ' row key identifying the row
    sourceColumn As String   ' column header holding the value
    numberFormat As String
    include As Boolean
    origin As String         ' ORIGIN_SCAN or ORIGIN_MANUAL
End Type

' Columns of tblScenarioResults. Public because ScenarioRun writes them and
' ScenarioCompare reads them.
Public Const COL_RESULT_METRIC As String = "Metric"
Public Const COL_RESULT_VALUE As String = "Value"
Public Const COL_RESULT_RUN_ON As String = "RunOn"

' Columns of the catalogue. Any that are missing are added by EnsureCatalogue.
Public Const COL_CAT_METRIC As String = "Metric"
Public Const COL_CAT_LABEL As String = "Label"
Public Const COL_CAT_GROUP As String = "Group"
Public Const COL_CAT_NAME_REF As String = "NameRef"
Public Const COL_CAT_TABLE As String = "SourceTable"
Public Const COL_CAT_LABEL_COL As String = "LabelColumn"
Public Const COL_CAT_LABEL_VAL As String = "LabelValue"
Public Const COL_CAT_VALUE_COL As String = "SourceColumn"
Public Const COL_CAT_FORMAT As String = "NumberFormat"
Public Const COL_CAT_INCLUDE As String = "Include"
Public Const COL_CAT_ORIGIN As String = "Origin"

Public Const ORIGIN_SCAN As String = "Scan"
Public Const ORIGIN_MANUAL As String = "Manual"

' Markers written into the results table in place of a value that could not be read.
' A run stores these rather than failing, so one broken metric cannot lose the rest.
Private Const ERR_NO_TABLE As String = "#TABLE MISSING"
Private Const ERR_NO_COLUMN As String = "#COLUMN MISSING"
Private Const ERR_NO_ROW As String = "#ROW MISSING"
Private Const ERR_NO_NAME As String = "#NAME MISSING"

'=======================================================================================
' Catalogue table
'=======================================================================================

' The catalogue, with every column this module needs present. Columns are added rather
' than demanded, so an older two-column tblScenarioMetrics upgrades itself in place.
Public Function EnsureCatalogue() As ListObject
    Dim lo As ListObject

    Set lo = RequireTable(MetricCatalogueTableName())

    AddColumnIfMissing lo, COL_CAT_METRIC
    AddColumnIfMissing lo, COL_CAT_LABEL
    AddColumnIfMissing lo, COL_CAT_GROUP
    AddColumnIfMissing lo, COL_CAT_NAME_REF
    AddColumnIfMissing lo, COL_CAT_TABLE
    AddColumnIfMissing lo, COL_CAT_LABEL_COL
    AddColumnIfMissing lo, COL_CAT_LABEL_VAL
    AddColumnIfMissing lo, COL_CAT_VALUE_COL
    AddColumnIfMissing lo, COL_CAT_FORMAT
    AddColumnIfMissing lo, COL_CAT_INCLUDE
    AddColumnIfMissing lo, COL_CAT_ORIGIN

    Set EnsureCatalogue = lo
End Function

Private Sub AddColumnIfMissing(lo As ListObject, header As String)
    If ColumnIndexOrZero(lo, header) > 0 Then Exit Sub
    lo.ListColumns.Add.Name = header
End Sub

' Rescans the source tables and rewrites the catalogue. Scanned rows that no longer
' exist are dropped; rows the user typed in by hand are always kept. Returns the number
' of metrics in the catalogue afterwards, and reports how many were added via addedOut.
Public Function SyncMetricCatalogue(Optional ByRef addedOut As Long, _
                                    Optional ByRef removedOut As Long) As Long
    Dim lo As ListObject
    Dim existing() As MetricSpec
    Dim discovered() As MetricSpec
    Dim merged() As MetricSpec
    Dim byMetric As Object
    Dim seen As Object
    Dim eCount As Long, dCount As Long, n As Long, i As Long
    Dim spec As MetricSpec
    Dim guard As CalcGuard

    guard = BeginBatch("Scanning the source tables for metrics...")
    On Error GoTo Failed

    Set lo = EnsureCatalogue()

    existing = CatalogueSpecs(False)
    eCount = MetricSpecCount(existing)

    Set byMetric = CreateObject("Scripting.Dictionary")
    byMetric.CompareMode = 1
    For i = 1 To eCount
        If Not byMetric.Exists(existing(i).metric) Then byMetric.Add existing(i).metric, i
    Next i

    discovered = DiscoverMetricSpecs()
    dCount = MetricSpecCount(discovered)

    ReDim merged(1 To eCount + dCount + 1)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    ' Hand-written rows first, so the catalogue keeps the order the user arranged.
    For i = 1 To eCount
        If StrComp(existing(i).origin, ORIGIN_SCAN, vbTextCompare) <> 0 Then
            If Not seen.Exists(existing(i).metric) Then
                n = n + 1
                merged(n) = existing(i)
                merged(n).origin = ORIGIN_MANUAL
                seen.Add existing(i).metric, True
            End If
        End If
    Next i

    For i = 1 To dCount
        If Not seen.Exists(discovered(i).metric) Then
            spec = discovered(i)
            If byMetric.Exists(spec.metric) Then
                CarryOverUserEdits spec, existing(byMetric(spec.metric))
            Else
                addedOut = addedOut + 1
            End If
            n = n + 1
            merged(n) = spec
            seen.Add spec.metric, True
        End If
    Next i

    For i = 1 To eCount
        If StrComp(existing(i).origin, ORIGIN_SCAN, vbTextCompare) = 0 Then
            If Not seen.Exists(existing(i).metric) Then removedOut = removedOut + 1
        End If
    Next i

    If n = 0 Then
        ResizeTableBody lo, 0
    Else
        ReDim Preserve merged(1 To n)
        WriteCatalogue lo, merged
    End If

    EndBatch guard
    SyncMetricCatalogue = n
    Exit Function

Failed:
    EndBatch guard
    Err.Raise Err.Number, "ScenarioMetrics.SyncMetricCatalogue", Err.description
End Function

' A resync must not undo the choices the user made in the catalogue, only refresh where
' each metric is read from.
Private Sub CarryOverUserEdits(ByRef fresh As MetricSpec, prior As MetricSpec)
    fresh.include = prior.include
    If Len(prior.label) > 0 Then fresh.label = prior.label
    If Len(prior.groupName) > 0 Then fresh.groupName = prior.groupName
    If Len(prior.numberFormat) > 0 Then fresh.numberFormat = prior.numberFormat
End Sub

Private Sub WriteCatalogue(lo As ListObject, specs() As MetricSpec)
    Dim isFormula As Variant
    Dim n As Long

    n = MetricSpecCount(specs)
    ResizeTableBody lo, n
    If n = 0 Then Exit Sub

    ' Read after the resize: Excel refills calculated columns as the table grows.
    isFormula = FormulaColumnFlags(lo)

    WriteCatalogueColumn lo, isFormula, COL_CAT_METRIC, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_LABEL, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_GROUP, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_NAME_REF, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_TABLE, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_LABEL_COL, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_LABEL_VAL, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_VALUE_COL, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_FORMAT, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_INCLUDE, specs
    WriteCatalogueColumn lo, isFormula, COL_CAT_ORIGIN, specs
End Sub

' One block write per column, skipping any the user has turned into a calculated column.
Private Sub WriteCatalogueColumn(lo As ListObject, isFormula As Variant, header As String, _
                                 specs() As MetricSpec)
    Dim values() As Variant
    Dim col As Long, i As Long, n As Long

    col = ColumnIndexOrZero(lo, header)
    If col = 0 Then Exit Sub
    If isFormula(col) Then Exit Sub

    n = MetricSpecCount(specs)
    ReDim values(1 To n, 1 To 1)
    For i = 1 To n
        Select Case header
            Case COL_CAT_METRIC: values(i, 1) = specs(i).metric
            Case COL_CAT_LABEL: values(i, 1) = specs(i).label
            Case COL_CAT_GROUP: values(i, 1) = specs(i).groupName
            Case COL_CAT_NAME_REF: values(i, 1) = specs(i).nameRef
            Case COL_CAT_TABLE: values(i, 1) = specs(i).sourceTable
            Case COL_CAT_LABEL_COL: values(i, 1) = specs(i).labelColumn
            Case COL_CAT_LABEL_VAL: values(i, 1) = specs(i).labelValue
            Case COL_CAT_VALUE_COL: values(i, 1) = specs(i).sourceColumn
            Case COL_CAT_FORMAT: values(i, 1) = specs(i).numberFormat
            Case COL_CAT_INCLUDE: values(i, 1) = specs(i).include
            Case COL_CAT_ORIGIN: values(i, 1) = specs(i).origin
        End Select
    Next i

    lo.ListColumns(col).DataBodyRange.value = values
End Sub

' Every row of the catalogue, or only the ticked ones. Rows with no metric ID are skipped
' so a half-typed row cannot break a run.
Public Function CatalogueSpecs(ByVal includedOnly As Boolean) As MetricSpec()
    Dim lo As ListObject
    Dim body As Variant
    Dim specs() As MetricSpec
    Dim cMetric As Long, cLabel As Long, cGroup As Long, cNameRef As Long
    Dim cTable As Long, cLabelCol As Long, cLabelVal As Long, cValueCol As Long
    Dim cFormat As Long, cInclude As Long, cOrigin As Long
    Dim r As Long, n As Long
    Dim included As Boolean

    Set lo = TableOrNothing(MetricCatalogueTableName())
    If lo Is Nothing Then Exit Function

    body = TableBody(lo)
    If IsEmpty(body) Then Exit Function

    cMetric = ColumnIndexOrZero(lo, COL_CAT_METRIC)
    If cMetric = 0 Then Exit Function

    cLabel = ColumnIndexOrZero(lo, COL_CAT_LABEL)
    cGroup = ColumnIndexOrZero(lo, COL_CAT_GROUP)
    cNameRef = ColumnIndexOrZero(lo, COL_CAT_NAME_REF)
    cTable = ColumnIndexOrZero(lo, COL_CAT_TABLE)
    cLabelCol = ColumnIndexOrZero(lo, COL_CAT_LABEL_COL)
    cLabelVal = ColumnIndexOrZero(lo, COL_CAT_LABEL_VAL)
    cValueCol = ColumnIndexOrZero(lo, COL_CAT_VALUE_COL)
    cFormat = ColumnIndexOrZero(lo, COL_CAT_FORMAT)
    cInclude = ColumnIndexOrZero(lo, COL_CAT_INCLUDE)
    cOrigin = ColumnIndexOrZero(lo, COL_CAT_ORIGIN)

    ReDim specs(1 To UBound(body, 1))
    For r = 1 To UBound(body, 1)
        If Len(Trim$(CStr(body(r, cMetric)))) > 0 Then
            ' An absent Include column means capture everything.
            If cInclude = 0 Then included = True Else included = IsTruthyValue(body(r, cInclude))
            If included Or Not includedOnly Then
                n = n + 1
                With specs(n)
                    .metric = Trim$(CStr(body(r, cMetric)))
                    .label = CellText(body, r, cLabel)
                    .groupName = CellText(body, r, cGroup)
                    .nameRef = CellText(body, r, cNameRef)
                    .sourceTable = CellText(body, r, cTable)
                    .labelColumn = CellText(body, r, cLabelCol)
                    .labelValue = CellText(body, r, cLabelVal)
                    .sourceColumn = CellText(body, r, cValueCol)
                    .numberFormat = CellText(body, r, cFormat)
                    .include = included
                    .origin = CellText(body, r, cOrigin)
                    If Len(.label) = 0 Then .label = .metric
                    If Len(.origin) = 0 Then .origin = ORIGIN_MANUAL
                End With
            End If
        End If
    Next r

    If n = 0 Then Exit Function
    ReDim Preserve specs(1 To n)
    CatalogueSpecs = specs
End Function

' What a run captures. The catalogue is the single source of truth; falling back to a
' raw scan when it is empty keeps the first run working before anyone has synced.
Public Function IncludedMetricSpecs() As MetricSpec()
    Dim specs() As MetricSpec

    specs = CatalogueSpecs(True)
    If MetricSpecCount(specs) > 0 Then
        IncludedMetricSpecs = specs
        Exit Function
    End If

    specs = DiscoverMetricSpecs()
    If MetricSpecCount(specs) = 0 Then
        Err.Raise vbObjectError + 620, "ScenarioMetrics", _
                  "No metrics are configured. Tick some rows in " & _
                  MetricCatalogueTableName() & ", or declare source tables in " & _
                  "MetricSourceTables() and press Sync Metrics."
    End If

    IncludedMetricSpecs = specs
End Function

' Sets the Include flag on every catalogue row at once. Returns how many rows changed.
Public Function SetAllIncluded(ByVal flag As Boolean) As Long
    Dim lo As ListObject
    Dim col As Long, r As Long, n As Long

    Set lo = EnsureCatalogue()
    If TableRowCount(lo) = 0 Then Exit Function

    col = RequireColumnIndex(lo, COL_CAT_INCLUDE)
    For r = 1 To TableRowCount(lo)
        With lo.DataBodyRange.Cells(r, col)
            If IsTruthyValue(.value) <> flag Then
                .value = flag
                n = n + 1
            End If
        End With
    Next r

    SetAllIncluded = n
End Function

' Sets the Include flag on whatever rows of the catalogue the user has selected, so
' picking metrics is "highlight the rows, press the button". Returns how many changed,
' or -1 when the selection does not touch the catalogue at all.
Public Function SetSelectionIncluded(ByVal flag As Boolean) As Long
    Dim lo As ListObject
    Dim overlap As Range
    Dim area As Range
    Dim cell As Range
    Dim col As Long, n As Long

    Set lo = EnsureCatalogue()
    If TableRowCount(lo) = 0 Then Exit Function
    If Not TypeOf Selection Is Range Then
        SetSelectionIncluded = -1
        Exit Function
    End If

    Set overlap = Intersect(Selection, lo.DataBodyRange)
    If overlap Is Nothing Then
        SetSelectionIncluded = -1
        Exit Function
    End If

    col = RequireColumnIndex(lo, COL_CAT_INCLUDE)
    For Each area In overlap.Areas
        For Each cell In Intersect(area.EntireRow, lo.ListColumns(col).DataBodyRange)
            If IsTruthyValue(cell.value) <> flag Then
                cell.value = flag
                n = n + 1
            End If
        Next cell
    Next area

    SetSelectionIncluded = n
End Function

'=======================================================================================
' Discovery
'=======================================================================================

' Walks every table named by MetricSourceTables() and turns each (row, column) cell into
' one metric. The metric ID is "<label value>_<column header>", or just the column header
' when the source declares no label column.
Public Function DiscoverMetricSpecs() As MetricSpec()
    Dim sources() As MetricTableSpec
    Dim excludes() As String
    Dim result() As MetricSpec
    Dim tables As Collection
    Dim lo As ListObject
    Dim source As MetricTableSpec
    Dim body As Variant
    Dim seen As Object
    Dim capacity As Long, labelCol As Long
    Dim t As Long, r As Long, c As Long, n As Long
    Dim header As String, labelValue As String, metricId As String, groupName As String

    sources = MetricSourceTables()
    If MetricTableSpecCount(sources) = 0 Then Exit Function
    excludes = MetricColumnExcludePrefixes()

    Set tables = New Collection
    For t = LBound(sources) To UBound(sources)
        CollectMatchingTables sources(t), tables
    Next t
    If tables.count = 0 Then Exit Function

    For t = 1 To tables.count
        Set lo = TableOrNothing(CStr(tables(t)))
        If Not lo Is Nothing Then capacity = capacity + TableRowCount(lo) * lo.ListColumns.count
    Next t
    If capacity = 0 Then Exit Function

    ReDim result(1 To capacity)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    For t = 1 To tables.count
        Set lo = TableOrNothing(CStr(tables(t)))
        If lo Is Nothing Then GoTo NextTable

        body = TableBody(lo)
        If IsEmpty(body) Then GoTo NextTable

        source = SourceFor(sources, lo.Name)
        groupName = IIf(Len(source.groupName) = 0, lo.Name, source.groupName)

        labelCol = 0
        If Len(source.labelColumn) > 0 Then labelCol = ColumnIndexOrZero(lo, source.labelColumn)

        For r = 1 To UBound(body, 1)
            ' Not IIf: it evaluates both branches, so body(r, 0) would be read.
            labelValue = ""
            If labelCol > 0 Then labelValue = Trim$(CStr(body(r, labelCol)))

            For c = 1 To UBound(body, 2)
                If c = labelCol Then GoTo NextColumn

                header = lo.ListColumns(c).Name
                If HasExcludedPrefix(header, excludes) Then GoTo NextColumn
                If Len(labelValue) = 0 And labelCol > 0 Then GoTo NextColumn

                metricId = BuildMetricId(lo.Name, labelValue, header)
                If seen.Exists(metricId) Then GoTo NextColumn
                seen.Add metricId, True

                n = n + 1
                With result(n)
                    .metric = metricId
                    .label = IIf(Len(labelValue) = 0, header, labelValue & " " & header)
                    .groupName = groupName
                    .sourceTable = lo.Name
                    .labelColumn = source.labelColumn
                    .labelValue = labelValue
                    .sourceColumn = header
                    .numberFormat = CellFormat(lo, r, c)
                    .include = True
                    .origin = ORIGIN_SCAN
                End With

NextColumn:
            Next c
        Next r
NextTable:
    Next t

    If n = 0 Then Exit Function
    ReDim Preserve result(1 To n)
    DiscoverMetricSpecs = result
End Function

' Metric IDs have to stay unique across source tables, so the table name leads. The
' prefix is stripped from the display label, which only needs to read well in a chart.
Private Function BuildMetricId(tableName As String, labelValue As String, header As String) As String
    Dim bare As String

    bare = tableName
    If LCase$(Left$(bare, 3)) = "tbl" Then bare = Mid$(bare, 4)

    If Len(labelValue) = 0 Then
        BuildMetricId = bare & "_" & header
    Else
        BuildMetricId = bare & "_" & labelValue & "_" & header
    End If
End Function

' Resolves one source entry to real table names. A trailing "*" matches every table whose
' name starts with the stem, which is how generated families like tblMetricsOut1,
' tblMetricsOut2 get picked up without naming each one.
Private Sub CollectMatchingTables(source As MetricTableSpec, ByRef tables As Collection)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim stem As String
    Dim name As String

    name = Trim$(source.tableName)
    If Len(name) = 0 Then Exit Sub

    If Right$(name, 1) <> "*" Then
        If Not TableOrNothing(name) Is Nothing Then tables.Add name
        Exit Sub
    End If

    stem = Left$(name, Len(name) - 1)
    If Len(stem) = 0 Then Exit Sub

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(Left$(lo.Name, Len(stem)), stem, vbTextCompare) = 0 Then tables.Add lo.Name
        Next lo
    Next ws
End Sub

Private Function SourceFor(sources() As MetricTableSpec, tableName As String) As MetricTableSpec
    Dim i As Long
    Dim pattern As String

    For i = LBound(sources) To UBound(sources)
        pattern = Trim$(sources(i).tableName)
        If Right$(pattern, 1) = "*" Then
            pattern = Left$(pattern, Len(pattern) - 1)
            If StrComp(Left$(tableName, Len(pattern)), pattern, vbTextCompare) = 0 Then
                SourceFor = sources(i)
                Exit Function
            End If
        ElseIf StrComp(pattern, tableName, vbTextCompare) = 0 Then
            SourceFor = sources(i)
            Exit Function
        End If
    Next i
End Function

Private Function HasExcludedPrefix(header As String, excludes() As String) As Boolean
    Dim i As Long

    If StringArrayCount(excludes) = 0 Then Exit Function

    For i = LBound(excludes) To UBound(excludes)
        If Len(excludes(i)) > 0 Then
            If StrComp(Left$(header, Len(excludes(i))), excludes(i), vbTextCompare) = 0 Then
                HasExcludedPrefix = True
                Exit Function
            End If
        End If
    Next i
End Function

' The source cell's own format, so a percentage stays a percentage in every summary.
Private Function CellFormat(lo As ListObject, ByVal r As Long, ByVal c As Long) As String
    Dim raw As Variant

    raw = lo.DataBodyRange.Cells(r, c).numberFormat
    If VarType(raw) = vbString Then CellFormat = CStr(raw)
End Function

'=======================================================================================
' Reading values
'=======================================================================================

' Metric name in column 1, value in column 2, ready for StoreResults. A source that
' cannot be resolved yields an error marker instead of raising, so one broken metric
' cannot lose the others.
' === PERF PATCH 2026-10-18 #2 Cache ListObject lookups across the run === START
' Was: ReadOne called TableOrNothing for every metric, each scan iterating every sheet
' and ListObject in the workbook. Now the lookup result is cached per-run.
Public Function ReadMetricValues(specs() As MetricSpec) As Variant
    Dim values() As Variant
    Dim rowMaps As Object, colMaps As Object, tableMaps As Object
    Dim n As Long, i As Long

    n = MetricSpecCount(specs)
    If n = 0 Then Exit Function

    Set rowMaps = CreateObject("Scripting.Dictionary")
    Set colMaps = CreateObject("Scripting.Dictionary")
    Set tableMaps = CreateObject("Scripting.Dictionary")

    ReDim values(1 To n, 1 To 2)
    For i = 1 To n
        values(i, 1) = specs(i).metric
        values(i, 2) = ReadOne(specs(i), rowMaps, colMaps, tableMaps)
    Next i

    ReadMetricValues = values
End Function

' Same contract as ReadMetricValues, but each spec's sourceTable is first swapped for its
' parallel current-state table (per MetricSourceTables in ScenarioConfig). A spec whose
' source has no current-state counterpart is read from its original table, which is the
' right behaviour for metrics that do not vary with the loaded scenario.
Public Function ReadMetricValuesForCurrentState(specs() As MetricSpec) As Variant
    Dim swapped() As MetricSpec
    Dim target As String
    Dim n As Long, i As Long

    n = MetricSpecCount(specs)
    If n = 0 Then Exit Function

    ReDim swapped(1 To n)
    For i = 1 To n
        swapped(i) = specs(i)
        target = CurrentStateTableFor(specs(i).sourceTable)
        If Len(target) > 0 Then swapped(i).sourceTable = target
    Next i

    ReadMetricValuesForCurrentState = ReadMetricValues(swapped)
End Function

' Resolves a regular source table name to its parallel current-state table name, or ""
' when none is declared. Handles the trailing-"*" stem form: a concrete discovered name
' is matched against each configured stem, and the current-state name has the matching
' suffix appended.
Private Function CurrentStateTableFor(sourceTable As String) As String
    Dim sources() As MetricTableSpec
    Dim pattern As String, replacement As String
    Dim stem As String
    Dim i As Long

    If Len(sourceTable) = 0 Then Exit Function

    sources = MetricSourceTables()
    For i = LBound(sources) To UBound(sources)
        If Len(sources(i).currentStateTable) > 0 Then
            pattern = sources(i).tableName
            replacement = sources(i).currentStateTable

            If Right$(pattern, 1) = "*" Then
                stem = Left$(pattern, Len(pattern) - 1)
                If StrComp(Left$(sourceTable, Len(stem)), stem, vbTextCompare) = 0 Then
                    ' Carry the discovered suffix across to the current-state family.
                    CurrentStateTableFor = Left$(replacement, Len(replacement) - 1) & _
                                            Mid$(sourceTable, Len(stem) + 1)
                    Exit Function
                End If
            ElseIf StrComp(pattern, sourceTable, vbTextCompare) = 0 Then
                CurrentStateTableFor = replacement
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ReadOne(spec As MetricSpec, rowMaps As Object, colMaps As Object, _
                         tableMaps As Object) As Variant
    Dim lo As ListObject
    Dim cell As Range
    Dim rowIndex As Long, colIndex As Long

    If Len(spec.sourceTable) = 0 Then
        Set cell = NamedRangeOrNothing(spec.nameRef)
        If cell Is Nothing Then ReadOne = ERR_NO_NAME Else ReadOne = cell.Cells(1, 1).value
        Exit Function
    End If

    Set lo = CachedTable(spec.sourceTable, tableMaps)
    If lo Is Nothing Then
        ReadOne = ERR_NO_TABLE
        Exit Function
    End If

    colIndex = CachedColumn(lo, spec.sourceColumn, colMaps)
    If colIndex = 0 Then
        ReadOne = ERR_NO_COLUMN
        Exit Function
    End If

    rowIndex = CachedRow(lo, spec, rowMaps)
    If rowIndex = 0 Then
        ReadOne = ERR_NO_ROW
        Exit Function
    End If

    ReadOne = lo.DataBodyRange.Cells(rowIndex, colIndex).value
End Function

Private Function CachedTable(tableName As String, tableMaps As Object) As ListObject
    Dim lo As ListObject

    If tableMaps.Exists(tableName) Then
        Set CachedTable = tableMaps(tableName)
        Exit Function
    End If

    Set lo = TableOrNothing(tableName)
    If lo Is Nothing Then
        tableMaps.Add tableName, Nothing
    Else
        tableMaps.Add tableName, lo
    End If
    Set CachedTable = lo
End Function
' === PERF PATCH 2026-10-18 #2 Cache ListObject lookups across the run === END

Private Function CachedColumn(lo As ListObject, header As String, colMaps As Object) As Long
    Dim key As String

    key = lo.Name & "|" & header
    If Not colMaps.Exists(key) Then colMaps.Add key, ColumnIndexOrZero(lo, header)
    CachedColumn = colMaps(key)
End Function

' One pass over the label column per table, reused by every metric that reads it.
Private Function CachedRow(lo As ListObject, spec As MetricSpec, rowMaps As Object) As Long
    Dim map As Object
    Dim body As Variant
    Dim key As String, labelValue As String
    Dim labelCol As Long, r As Long

    If Len(spec.labelColumn) = 0 Then
        CachedRow = IIf(TableRowCount(lo) > 0, 1, 0)
        Exit Function
    End If

    key = lo.Name & "|" & spec.labelColumn
    If Not rowMaps.Exists(key) Then
        Set map = CreateObject("Scripting.Dictionary")
        map.CompareMode = 1

        labelCol = ColumnIndexOrZero(lo, spec.labelColumn)
        body = TableBody(lo)
        If labelCol > 0 And Not IsEmpty(body) Then
            For r = 1 To UBound(body, 1)
                labelValue = Trim$(CStr(body(r, labelCol)))
                If Len(labelValue) > 0 And Not map.Exists(labelValue) Then map.Add labelValue, r
            Next r
        End If

        rowMaps.Add key, map
    End If

    Set map = rowMaps(key)
    If map.Exists(spec.labelValue) Then CachedRow = map(spec.labelValue)
End Function

'=======================================================================================
' Small helpers
'=======================================================================================

' UBound on an array that was never dimensioned raises, which is exactly what a function
' returning "no metrics" leaves behind.
Public Function MetricSpecCount(specs() As MetricSpec) As Long
    On Error Resume Next
    MetricSpecCount = UBound(specs) - LBound(specs) + 1
    If Err.Number <> 0 Then MetricSpecCount = 0
    Err.Clear
End Function

Public Function MetricTableSpecCount(specs() As MetricTableSpec) As Long
    On Error Resume Next
    MetricTableSpecCount = UBound(specs) - LBound(specs) + 1
    If Err.Number <> 0 Then MetricTableSpecCount = 0
    Err.Clear
End Function

Public Function StringArrayCount(values() As String) As Long
    On Error Resume Next
    StringArrayCount = UBound(values) - LBound(values) + 1
    If Err.Number <> 0 Then StringArrayCount = 0
    Err.Clear
End Function

Public Function IsTruthyValue(value As Variant) As Boolean
    Select Case UCase$(Trim$(CStr(value)))
        Case "TRUE", "YES", "Y", "1"
            IsTruthyValue = True
    End Select
End Function

Private Function CellText(body As Variant, ByVal r As Long, ByVal c As Long) As String
    If c = 0 Then Exit Function
    CellText = Trim$(CStr(body(r, c)))
End Function
