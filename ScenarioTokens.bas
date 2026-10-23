Attribute VB_Name = "ScenarioTokens"
Option Explicit

' Token substitution engine shared by ScenarioCompare and ScenarioOutput.
'
' A "template range" is a rectangular block laid out and formatted by hand on a worksheet,
' with cells that mix ordinary content and bracketed tokens like {Metric.NIM}. FillTokens
' walks a copy of the block and replaces every recognised token with the value that
' applies to one scenario. Everything else is copied through untouched, so the sheet
' designer stays in charge of layout, fonts, colours and borders.
'
'=======================================================================================
' Token grammar
'=======================================================================================
'
' Tokens are case-insensitive and delimited by { and }. A cell whose entire contents are
' a single token keeps the resolved value's own type (a number stays a number). A token
' embedded in longer text is substituted in place. A token inside a formula is replaced
' with a literal, so "={Metric.A}/{Metric.B}-1" works and a metric with no stored value
' becomes NA() rather than silently reading as zero.
'
' Scalar tokens (resolve to a single value):
'
'   {Scenario.Name}            the scenario's display name
'   {Scenario.ID}              its ScenarioID
'   {Scenario.Index}           1 for the first block painted, 2 for the second, etc.
'                              Always 1 for a single-scenario render.
'   {Scenario.RunOn}           when it last produced results
'   {Scenario.<Column>}        any other column of tblScenarios, e.g. {Scenario.EffectiveDate}
'   {Metric.<MetricID>}        that scenario's stored value for the metric
'   {Label.<MetricID>}         the metric's display label from the catalogue, with
'                              underscores shown as spaces
'   {Current.<MetricID>}       the current-state baseline for the metric
'   {Delta.<MetricID>}         this scenario's value minus the baseline
'   {DeltaPct.<MetricID>}      (scenario - current) / current, as a decimal
'   {Segment.<ID>.<Column>}    an attribute of a segment from the segment map table
'                              (tblProductMap, keyed by Pricing_Segment_ID; see
'                              ScenarioConfig), e.g. {Segment.42.Segment Description}.
'                              The ID runs to the first "." and the column is the rest.
'                              Read live, and the same in every scenario's block. An
'                              unknown ID or column resolves like a missing metric.
'
' Unknown tokens are left exactly as written, so braces used for other purposes survive
' untouched.
'
' Any token can carry an Excel number format after a "|", e.g. {Metric.NIM|0.00%},
' {Delta.Balance|#,##0;(#,##0);-} or {Scenario.EffectiveDate|dd-mmm-yy}. The split is
' on the first "|", so the format may use ";" sections and quoted text freely, but it
' cannot contain "}".
'
' Prefix ordering matters: DeltaPct is tested before Delta because Delta. is a prefix of
' DeltaPct. and the opposite order would silently swallow the DeltaPct case.
'
'=======================================================================================
' Formatting
'=======================================================================================
'
' A format written into the token itself beats everything:
'
'   whole-cell token    the cell gets that number format, the value stays a number
'   token in a formula  the value stays a number so the formula still calculates, and
'                       the cell gets the format (the first formatted token wins)
'   token in text       the value is written through Excel's TEXT(), e.g. "NIM 2.35%"
'
' A format Excel rejects is ignored, falling back to the rules below.
'
' Without one, a template cell's own number format wins. The catalogue format only fills
' in a cell that was left as General, so a designer who typed "0.0%" in the template
' keeps it. Metric, Current and Delta tokens share the metric's unit and inherit its
' format; DeltaPct is a ratio and is left alone so the template can format it however it
' likes (0.0% in one block, +0.0;-0.0 in another).
'
'=======================================================================================
' Writing a new template
'=======================================================================================
'
' 1. Lay out the block on a hidden "templates" area of the sheet, with a label row above.
' 2. Reference metrics by their MetricID (the value from the metric catalogue, not the
'    display label). Mix text, formulas and tokens freely.
' 3. Keep references to anything *outside* the block absolute, otherwise they translate
'    when copies are painted into neighbouring positions.
' 4. Declare the block's defined name in ScenarioConfig (CompareTemplates for compare,
'    ScenarioOutputTemplates for output). Nothing else needs to change.
'
'=======================================================================================
' Public API summary
'=======================================================================================
'
'   FillTokens block, scenarioId, blockIndex, results, labels, formats
'       Rewrites every recognised token inside block.
'   TemplateMetricIds(template)
'       Deduplicated MetricIDs the template references, in reading order. Formulas are
'       scanned too, so a metric only used inside a calculation still appears in the
'       grid under the blocks.
'   ResultsLookup()
'       Dictionary keyed "ScenarioID|MetricID" built from tblScenarioResults.
'   LoadCatalogueMaps labels, formats, groups
'       Three dictionaries keyed by MetricID.
'   MetricValue results, scenarioId, metricId, ifMissing
'       The stored value or ifMissing if the pair is unknown.
'   ResolveToken token, scenarioId, blockIndex, results, labels, missing
'       The resolved value for one token. Exposed for callers (like chart series
'       binding) that already have a token in hand and only want its value.
'   WholeCellToken(text), IsKnownToken(token), MetricIdFromToken(token)
'       Low-level helpers for callers that scan their own cells. All of them accept a
'       token with or without a "|format" suffix.
'   SplitTokenFormat(token, numberFormat)
'       The token without its "|format" suffix, with the format returned via
'       numberFormat.
'   SegmentAttribute(segmentId, column, missing)
'       One cell of the segment map, or missing when the ID or column is unknown.
'   SplitSegmentBody(body, segmentId, column)
'       Splits the "<ID>.<Column>" part of a Segment token.
'   SegmentMapIds(), SegmentMapHasId(segmentId)
'       The IDs in the segment map, and whether one is there.
'   ForgetSegmentMap
'       Drops the cached segment map so the next lookup re-reads the table.

Private Const TOKEN_OPEN As String = "{"
Private Const TOKEN_CLOSE As String = "}"
Private Const FORMAT_SEPARATOR As String = "|"
Private Const PREFIX_SCENARIO As String = "Scenario."
Private Const PREFIX_METRIC As String = "Metric."
Private Const PREFIX_LABEL As String = "Label."
Private Const PREFIX_CURRENT As String = "Current."
Private Const PREFIX_DELTA As String = "Delta."
Private Const PREFIX_DELTA_PCT As String = "DeltaPct."
Private Const PREFIX_SEGMENT As String = "Segment."

Private Const MISSING_VALUE As String = "n/a"

' The segment map, read once and reused for every Segment token until ForgetSegmentMap.
' FillTokens forgets it at the start of each block, so a build always sees the live table
' but never reads it more than once per block.
Private segmentMapLoaded As Boolean
Private segmentMapBody As Variant
Private segmentMapRows As Object     ' segment ID -> row in segmentMapBody
Private segmentMapCols As Object     ' column header -> column in segmentMapBody

'=======================================================================================
' Delimiter accessors
'=======================================================================================

Public Function TokenOpen() As String
    TokenOpen = TOKEN_OPEN
End Function

Public Function TokenClose() As String
    TokenClose = TOKEN_CLOSE
End Function

Public Function TokenFormatSeparator() As String
    TokenFormatSeparator = FORMAT_SEPARATOR
End Function

Public Function TokenMissingText() As String
    TokenMissingText = MISSING_VALUE
End Function

'=======================================================================================
' Rewriting cells in a block
'=======================================================================================

' Replaces the tokens in one block. Formulas are rewritten in place rather than flattened,
' so whatever the template calculates it still calculates here.
Public Sub FillTokens(block As Range, scenarioId As String, ByVal blockIndex As Long, _
                      results As Object, labels As Object, formats As Object)
    Dim cell As Range
    Dim text As String
    Dim token As String
    Dim tokenFormat As String

    ForgetSegmentMap

    For Each cell In block.Cells
        If cell.HasFormula Then
            text = cell.Formula
            If InStr(text, TOKEN_OPEN) > 0 Then
                cell.Formula = SubstituteTokens(text, True, scenarioId, blockIndex, _
                                                results, labels, tokenFormat)
                If Len(tokenFormat) > 0 Then TrySetNumberFormat cell, tokenFormat
            End If
        Else
            text = CStr(cell.value)
            If InStr(text, TOKEN_OPEN) > 0 Then
                token = WholeCellToken(text)
                If Len(token) > 0 Then
                    ' A cell that is nothing but a token keeps the value's own type.
                    cell.value = ResolveToken(token, scenarioId, blockIndex, results, _
                                              labels, MISSING_VALUE)
                    ApplyMetricFormat cell, token, formats
                Else
                    cell.value = SubstituteTokens(text, False, scenarioId, blockIndex, _
                                                  results, labels)
                End If
            End If
        End If
    Next cell
End Sub

' The token when the cell's entire contents are "{...}", otherwise an empty string.
Public Function WholeCellToken(text As String) As String
    Dim trimmed As String

    trimmed = Trim$(text)
    If Len(trimmed) < 3 Then Exit Function
    If Left$(trimmed, 1) <> TOKEN_OPEN Or Right$(trimmed, 1) <> TOKEN_CLOSE Then Exit Function
    If InStr(trimmed, TOKEN_CLOSE) <> Len(trimmed) Then Exit Function

    trimmed = Trim$(Mid$(trimmed, 2, Len(trimmed) - 2))
    If IsKnownToken(trimmed) Then WholeCellToken = trimmed
End Function

' Rewrites every recognised token in a piece of text or formula. Tokens that are not ours
' are left exactly as they were, so braces used for anything else survive untouched.
' In a formula a token's format cannot be applied inline, so the first one found is
' handed back via firstFormat for the caller to put on the cell.
Public Function SubstituteTokens(text As String, ByVal asFormula As Boolean, _
                                 scenarioId As String, ByVal blockIndex As Long, _
                                 results As Object, labels As Object, _
                                 Optional ByRef firstFormat As String) As String
    Dim out As String
    Dim token As String
    Dim tokenFormat As String
    Dim value As Variant
    Dim missing As Variant
    Dim pos As Long, openAt As Long, closeAt As Long

    ' A metric with no stored value has to become NA() in a formula: any text would read
    ' as zero and quietly produce a wrong comparison.
    If asFormula Then missing = Empty Else missing = MISSING_VALUE
    firstFormat = ""

    pos = 1
    Do
        openAt = InStr(pos, text, TOKEN_OPEN)
        If openAt = 0 Then Exit Do
        closeAt = InStr(openAt + 1, text, TOKEN_CLOSE)
        If closeAt = 0 Then Exit Do

        out = out & Mid$(text, pos, openAt - pos)
        token = SplitTokenFormat(Mid$(text, openAt + 1, closeAt - openAt - 1), tokenFormat)

        If IsKnownToken(token) Then
            value = ResolveToken(token, scenarioId, blockIndex, results, labels, missing)
            If asFormula Then
                out = out & FormulaLiteral(value)
                If Len(firstFormat) = 0 Then firstFormat = tokenFormat
            Else
                out = out & FormattedText(value, tokenFormat)
            End If
        Else
            out = out & Mid$(text, openAt, closeAt - openAt + 1)
        End If

        pos = closeAt + 1
    Loop

    SubstituteTokens = out & Mid$(text, pos)
End Function

Public Function IsKnownToken(rawToken As String) As Boolean
    Dim token As String

    token = BareToken(rawToken)

    ' DeltaPct is tested before Delta: Delta is a prefix of DeltaPct, and StartsWith on
    ' "Delta." would otherwise swallow the "DeltaPct." case silently.
    IsKnownToken = StartsWith(token, PREFIX_METRIC) Or _
                   StartsWith(token, PREFIX_LABEL) Or _
                   StartsWith(token, PREFIX_SCENARIO) Or _
                   StartsWith(token, PREFIX_CURRENT) Or _
                   StartsWith(token, PREFIX_DELTA_PCT) Or _
                   StartsWith(token, PREFIX_DELTA) Or _
                   StartsWith(token, PREFIX_SEGMENT)
End Function

' Splits "Metric.NIM|0.00%" into "Metric.NIM" and "0.00%". Only the first "|" splits,
' so a format can carry one inside quoted text.
Public Function SplitTokenFormat(rawToken As String, ByRef numberFormat As String) As String
    Dim cut As Long

    numberFormat = ""
    cut = InStr(rawToken, FORMAT_SEPARATOR)
    If cut = 0 Then
        SplitTokenFormat = Trim$(rawToken)
        Exit Function
    End If

    SplitTokenFormat = Trim$(Left$(rawToken, cut - 1))
    numberFormat = Trim$(Mid$(rawToken, cut + 1))
End Function

Private Function BareToken(rawToken As String) As String
    Dim ignored As String

    BareToken = SplitTokenFormat(rawToken, ignored)
End Function

'=======================================================================================
' Resolving one token to a value
'=======================================================================================

Public Function ResolveToken(rawToken As String, scenarioId As String, ByVal blockIndex As Long, _
                             results As Object, labels As Object, _
                             missing As Variant) As Variant
    Dim token As String
    Dim body As String
    Dim segmentId As String, column As String

    token = BareToken(rawToken)

    If StartsWith(token, PREFIX_METRIC) Then
        body = Mid$(token, Len(PREFIX_METRIC) + 1)
        ResolveToken = MetricValue(results, scenarioId, body, missing)
        Exit Function
    End If

    If StartsWith(token, PREFIX_LABEL) Then
        body = Mid$(token, Len(PREFIX_LABEL) + 1)
        ' Labels are built from raw IDs and headers like "SEG_01 Net_Margin"; underscores
        ' read as noise on a rendered sheet.
        If labels.Exists(body) Then ResolveToken = labels(body) Else ResolveToken = body
        ResolveToken = Replace(CStr(ResolveToken), "_", " ")
        Exit Function
    End If

    If StartsWith(token, PREFIX_CURRENT) Then
        body = Mid$(token, Len(PREFIX_CURRENT) + 1)
        ResolveToken = MetricValue(results, CurrentStateScenarioId(), body, missing)
        Exit Function
    End If

    ' DeltaPct must be checked before Delta.
    If StartsWith(token, PREFIX_DELTA_PCT) Then
        body = Mid$(token, Len(PREFIX_DELTA_PCT) + 1)
        ResolveToken = DeltaValue(results, scenarioId, body, True, missing)
        Exit Function
    End If

    If StartsWith(token, PREFIX_DELTA) Then
        body = Mid$(token, Len(PREFIX_DELTA) + 1)
        ResolveToken = DeltaValue(results, scenarioId, body, False, missing)
        Exit Function
    End If

    If StartsWith(token, PREFIX_SEGMENT) Then
        If SplitSegmentBody(Mid$(token, Len(PREFIX_SEGMENT) + 1), segmentId, column) Then
            ResolveToken = SegmentAttribute(segmentId, column, missing)
        Else
            ResolveToken = missing
        End If
        Exit Function
    End If

    ResolveToken = ScenarioField(Mid$(token, Len(PREFIX_SCENARIO) + 1), scenarioId, blockIndex)
End Function

' Encodes a value for pasting into formula text. Str$ is used for numbers because it
' always emits a period decimal separator, which is what Range.Formula expects whatever
' the regional settings are.
Public Function FormulaLiteral(value As Variant) As String
    If IsEmpty(value) Or IsError(value) Then
        FormulaLiteral = "NA()"
    ElseIf VarType(value) = vbBoolean Then
        FormulaLiteral = IIf(value, "TRUE", "FALSE")
    ElseIf VarType(value) = vbDate Then
        FormulaLiteral = Trim$(Str$(CDbl(value)))
    ElseIf IsNumeric(value) Then
        FormulaLiteral = Trim$(Str$(CDbl(value)))
    Else
        FormulaLiteral = """" & Replace(CStr(value), """", """""") & """"
    End If
End Function

' Returns the scenario-minus-current difference, or (scenario - current)/current when
' asPct is True. Any missing or non-numeric input yields the caller's missing marker so
' a stale metric cannot silently read as zero and warp the delta.
Public Function DeltaValue(results As Object, scenarioId As String, metricId As String, _
                           ByVal asPct As Boolean, missing As Variant) As Variant
    Dim scenVal As Variant, currVal As Variant
    Dim s As Double, c As Double

    scenVal = MetricValue(results, scenarioId, metricId, Empty)
    currVal = MetricValue(results, CurrentStateScenarioId(), metricId, Empty)

    If IsEmpty(scenVal) Or IsEmpty(currVal) Then
        DeltaValue = missing
        Exit Function
    End If
    If VarType(scenVal) = vbString Or VarType(currVal) = vbString Then
        DeltaValue = missing
        Exit Function
    End If
    If Not IsNumeric(scenVal) Or Not IsNumeric(currVal) Then
        DeltaValue = missing
        Exit Function
    End If

    s = CDbl(scenVal)
    c = CDbl(currVal)

    If asPct Then
        If c = 0 Then
            DeltaValue = missing
            Exit Function
        End If
        DeltaValue = (s - c) / c
    Else
        DeltaValue = s - c
    End If
End Function

Public Function ScenarioField(field As String, scenarioId As String, _
                              ByVal blockIndex As Long) As Variant
    Dim lo As ListObject

    Select Case UCase$(field)
        Case "NAME"
            ScenarioField = ScenarioName(scenarioId)
        Case "ID"
            ScenarioField = scenarioId
        Case "INDEX"
            ScenarioField = blockIndex
        Case "RUNON"
            ScenarioField = LastCalculatedOnText(scenarioId, ifNever:="Not yet run")
        Case Else
            Set lo = RequireTable(TBL_SCENARIOS)
            If ColumnIndexOrZero(lo, field) = 0 Then
                ScenarioField = TOKEN_OPEN & PREFIX_SCENARIO & field & TOKEN_CLOSE
            Else
                ScenarioField = MetaValue(scenarioId, field)
            End If
    End Select
End Function

' A format written into the token wins outright. Otherwise the template's own formatting
' wins, and the catalogue format only fills in cells the template left as General.
' Current and Delta tokens share the metric's unit so they inherit its format; DeltaPct
' is a ratio and is left alone for the template to format.
Public Sub ApplyMetricFormat(cell As Range, rawToken As String, formats As Object)
    Dim token As String
    Dim tokenFormat As String
    Dim metricId As String

    token = SplitTokenFormat(rawToken, tokenFormat)
    If Len(tokenFormat) > 0 Then
        If TrySetNumberFormat(cell, tokenFormat) Then Exit Sub
    End If

    If cell.numberFormat <> "General" Then Exit Sub

    If StartsWith(token, PREFIX_METRIC) Then
        metricId = Mid$(token, Len(PREFIX_METRIC) + 1)
    ElseIf StartsWith(token, PREFIX_CURRENT) Then
        metricId = Mid$(token, Len(PREFIX_CURRENT) + 1)
    ElseIf StartsWith(token, PREFIX_DELTA_PCT) Then
        Exit Sub
    ElseIf StartsWith(token, PREFIX_DELTA) Then
        metricId = Mid$(token, Len(PREFIX_DELTA) + 1)
    Else
        Exit Sub
    End If

    If Not formats.Exists(metricId) Then Exit Sub
    If Len(formats(metricId)) = 0 Then Exit Sub

    cell.numberFormat = formats(metricId)
End Sub

' False when Excel rejects the format, leaving the cell's existing format in place.
Private Function TrySetNumberFormat(cell As Range, numberFormat As String) As Boolean
    On Error Resume Next
    cell.numberFormat = numberFormat
    TrySetNumberFormat = (Err.Number = 0)
    On Error GoTo 0
End Function

' Excel's TEXT() rather than VBA's Format$, so a code means exactly what it would in a
' cell's number format. Text values, and any code TEXT rejects, pass through as plain text.
Private Function FormattedText(value As Variant, numberFormat As String) As String
    FormattedText = CStr(value)
    If Len(numberFormat) = 0 Or IsEmpty(value) Then Exit Function
    If Not IsNumeric(value) And VarType(value) <> vbDate Then Exit Function

    On Error Resume Next
    FormattedText = Application.WorksheetFunction.Text(value, numberFormat)
    On Error GoTo 0
End Function

'=======================================================================================
' Scanning templates
'=======================================================================================

' Metric IDs referenced by the template, deduplicated, in reading order. Formulas are
' scanned too, so a metric only ever used inside a calculation still reaches the grid.
' Any of Metric/Current/Delta/DeltaPct counts: the caller should carry a row for a
' metric the template cares about however it addresses it.
Public Function TemplateMetricIds(template As Range) As String()
    Dim cell As Range
    Dim seen As Object
    Dim found() As String
    Dim text As String, token As String, metricId As String
    Dim pos As Long, openAt As Long, closeAt As Long
    Dim n As Long

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    ReDim found(1 To template.Cells.count)

    For Each cell In template.Cells
        If cell.HasFormula Then text = cell.Formula Else text = CStr(cell.value)

        pos = 1
        Do
            openAt = InStr(pos, text, TOKEN_OPEN)
            If openAt = 0 Then Exit Do
            closeAt = InStr(openAt + 1, text, TOKEN_CLOSE)
            If closeAt = 0 Then Exit Do

            token = Trim$(Mid$(text, openAt + 1, closeAt - openAt - 1))
            metricId = MetricIdFromToken(token)
            If Len(metricId) > 0 Then
                If Not seen.Exists(metricId) Then
                    seen.Add metricId, True
                    n = n + 1
                    found(n) = metricId
                End If
            End If

            pos = closeAt + 1
        Loop
    Next cell

    If n = 0 Then Exit Function
    ReDim Preserve found(1 To n)
    TemplateMetricIds = found
End Function

' The metric ID inside a value-bearing token, or "" when the token is not one that
' addresses a metric. DeltaPct must be matched before Delta.
Public Function MetricIdFromToken(rawToken As String) As String
    Dim token As String

    token = BareToken(rawToken)

    If StartsWith(token, PREFIX_METRIC) Then
        MetricIdFromToken = Mid$(token, Len(PREFIX_METRIC) + 1)
    ElseIf StartsWith(token, PREFIX_CURRENT) Then
        MetricIdFromToken = Mid$(token, Len(PREFIX_CURRENT) + 1)
    ElseIf StartsWith(token, PREFIX_DELTA_PCT) Then
        MetricIdFromToken = Mid$(token, Len(PREFIX_DELTA_PCT) + 1)
    ElseIf StartsWith(token, PREFIX_DELTA) Then
        MetricIdFromToken = Mid$(token, Len(PREFIX_DELTA) + 1)
    End If
End Function

'=======================================================================================
' Lookups
'=======================================================================================

' Every stored result keyed "ScenarioID|Metric". A run replaces a scenario's rows
' wholesale, so there is exactly one value per scenario per metric.
Public Function ResultsLookup() As Object
    Dim lo As ListObject
    Dim body As Variant
    Dim map As Object
    Dim cId As Long, cMetric As Long, cValue As Long
    Dim key As String
    Dim r As Long

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = 1
    Set ResultsLookup = map

    Set lo = TableOrNothing(ResultsTableName())
    If lo Is Nothing Then Exit Function

    body = TableBody(lo)
    If IsEmpty(body) Then Exit Function

    cId = RequireColumnIndex(lo, COL_SCENARIO_ID)
    cMetric = RequireColumnIndex(lo, COL_RESULT_METRIC)
    cValue = RequireColumnIndex(lo, COL_RESULT_VALUE)

    For r = 1 To UBound(body, 1)
        key = Trim$(CStr(body(r, cId))) & "|" & Trim$(CStr(body(r, cMetric)))
        map(key) = body(r, cValue)
    Next r
End Function

Public Function MetricValue(results As Object, scenarioId As String, metricId As String, _
                            ifMissing As Variant) As Variant
    Dim key As String

    key = scenarioId & "|" & metricId
    If results.Exists(key) Then MetricValue = results(key) Else MetricValue = ifMissing
End Function

Public Sub LoadCatalogueMaps(ByRef labels As Object, ByRef formats As Object, _
                             ByRef groups As Object)
    Dim specs() As MetricSpec
    Dim i As Long

    Set labels = CreateObject("Scripting.Dictionary")
    Set formats = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    labels.CompareMode = 1
    formats.CompareMode = 1
    groups.CompareMode = 1

    specs = CatalogueSpecs(False)
    For i = 1 To MetricSpecCount(specs)
        labels(specs(i).metric) = specs(i).label
        formats(specs(i).metric) = specs(i).numberFormat
        groups(specs(i).metric) = specs(i).groupName
    Next i
End Sub

'---------------------------------------------------------------------------------------
' Segment map
'---------------------------------------------------------------------------------------

' Splits "42.Segment Description" into "42" and "Segment Description". The ID runs to
' the first ".", so a column header may itself contain dots. False when either part is
' missing.
Public Function SplitSegmentBody(body As String, ByRef segmentId As String, _
                                 ByRef column As String) As Boolean
    Dim cut As Long

    segmentId = ""
    column = ""
    cut = InStr(body, ".")
    If cut = 0 Then Exit Function

    segmentId = Trim$(Left$(body, cut - 1))
    column = Trim$(Mid$(body, cut + 1))
    SplitSegmentBody = (Len(segmentId) > 0 And Len(column) > 0)
End Function

' One attribute of one segment, keeping the cell's own type so a number stays a number.
' A blank cell comes back as "" rather than missing: the segment exists, it just has no
' value there. An error in the cell, an unknown ID or column, or a missing table all
' come back as missing.
Public Function SegmentAttribute(segmentId As String, column As String, _
                                 missing As Variant) As Variant
    Dim value As Variant

    SegmentAttribute = missing
    LoadSegmentMap
    If segmentMapRows.count = 0 Then Exit Function
    If Not segmentMapRows.Exists(segmentId) Then Exit Function
    If Not segmentMapCols.Exists(column) Then Exit Function

    value = segmentMapBody(segmentMapRows(segmentId), segmentMapCols(column))
    If IsError(value) Then Exit Function
    If IsEmpty(value) Then value = ""
    SegmentAttribute = value
End Function

' Segment IDs in table order, skipping blanks, errors and repeats. An empty map comes
' back as Split(vbNullString), a real zero-length array.
Public Function SegmentMapIds() As String()
    Dim ids() As String
    Dim key As Variant
    Dim n As Long

    LoadSegmentMap
    If segmentMapRows.count = 0 Then
        SegmentMapIds = Split(vbNullString)
        Exit Function
    End If

    ReDim ids(1 To segmentMapRows.count)
    For Each key In segmentMapRows.Keys
        n = n + 1
        ids(n) = key
    Next key
    SegmentMapIds = ids
End Function

Public Function SegmentMapHasId(segmentId As String) As Boolean
    LoadSegmentMap
    SegmentMapHasId = segmentMapRows.Exists(segmentId)
End Function

Public Sub ForgetSegmentMap()
    segmentMapLoaded = False
    segmentMapBody = Empty
    Set segmentMapRows = Nothing
    Set segmentMapCols = Nothing
End Sub

' Reads the whole table in one go. A missing table or key column leaves both
' dictionaries empty, so every lookup resolves as missing instead of raising mid-render.
' The first row wins when an ID appears twice.
Private Sub LoadSegmentMap()
    Dim lo As ListObject
    Dim keyCol As Long
    Dim key As String
    Dim r As Long, c As Long

    If segmentMapLoaded Then Exit Sub
    segmentMapLoaded = True

    Set segmentMapRows = CreateObject("Scripting.Dictionary")
    Set segmentMapCols = CreateObject("Scripting.Dictionary")
    segmentMapRows.CompareMode = 1
    segmentMapCols.CompareMode = 1

    Set lo = TableOrNothing(SegmentMapTableName())
    If lo Is Nothing Then Exit Sub
    keyCol = ColumnIndexOrZero(lo, SegmentMapKeyColumn())
    If keyCol = 0 Then Exit Sub

    segmentMapBody = TableBody(lo)
    If IsEmpty(segmentMapBody) Then Exit Sub

    For c = 1 To lo.ListColumns.count
        segmentMapCols(lo.ListColumns(c).Name) = c
    Next c
    For r = 1 To UBound(segmentMapBody, 1)
        If Not IsError(segmentMapBody(r, keyCol)) Then
            key = Trim$(CStr(segmentMapBody(r, keyCol)))
            If Len(key) > 0 And Not segmentMapRows.Exists(key) Then segmentMapRows.Add key, r
        End If
    Next r
End Sub

'=======================================================================================
' Small helpers
'=======================================================================================

Private Function StartsWith(text As String, prefix As String) As Boolean
    If Len(text) < Len(prefix) Then Exit Function
    StartsWith = (StrComp(Left$(text, Len(prefix)), prefix, vbTextCompare) = 0)
End Function

