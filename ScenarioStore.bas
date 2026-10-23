Attribute VB_Name = "ScenarioStore"
Option Explicit
 
' Generic scenario persistence engine. Same shape as TableImport and SheetRecalc:
' this module holds the mechanics, ScenarioConfig declares what a scenario is made of.
'
' A scenario is a set of inputs that the model is run against. Its parts live in two
' places:
'
'   Store tables   long format, one row per scenario per input row, keyed by ScenarioID.
'                  Nothing in the workbook calculates off these.
'   Active tables  the working copy. Every formula and chart in the model reads these
'                  and only these, so switching scenario is a data swap and no formula
'                  anywhere needs to change.
'
' The active tables are the ones the model already had, for example tblDiscretions.
' Loading a scenario overwrites them; saving copies them back under a ScenarioID.
'
' Store tables must not contain calculated columns: they are rewritten wholesale.
 
Public Type ScenarioPart
    label As String          ' used in messages, e.g. "ATL price adjustments"
    storeTable As String     ' long-format table carrying a ScenarioID column
    activeTable As String    ' working table the model formulas read
    templateTable As String  ' blank-default rows for a new scenario, "" for none
End Type
 
Public Type ScenarioScalar
    label As String
    column As String         ' column in tblScenarios holding the value
    nameRef As String        ' defined name the model formulas read
End Type

Public Type MetricTableSpec
    tableName As String      ' table to scan; a trailing "*" matches every table with that stem
    labelColumn As String    ' column whose value identifies each row; "" = one row only
    groupName As String      ' groups the metrics into one summary section and one chart
    currentStateTable As String  ' parallel table a current-state run reads instead of tableName;
                                 ' leave empty to read from tableName itself
End Type

Public Type CompareTemplateSpec
    label As String          ' what the user picks from
    nameRef As String        ' defined name covering the sample summary block
End Type

' A single chart's data binding. ScenarioOutput reads these for a predeclared chart named
' in {Chart:<name>} tokens, and also builds them on the fly from the inline form
' {Chart:kind=Line;source=tblMetricsOutRFCMonthly;x=Month;series=Value,BaselineValue;title=...}.
' Any field left empty takes a sensible default: kind = Line, title = label, height/width
' = the token cell's own merged block size.
Public Type ChartSpec
    label As String          ' used in messages and as the default chart title
    kind As String           ' Line, Column, Bar, Combo. Case-insensitive.
    sourceTable As String    ' table whose columns feed the chart
    xColumn As String        ' category/x-axis column in the source table
    seriesColumns As String  ' comma-separated column names to plot
    baselineTable As String  ' optional parallel table (e.g. the _CURRENT twin) read for
                             ' baselineColumns using the same x values
    baselineColumns As String ' comma-separated series pulled from baselineTable
    title As String          ' overrides label at the top of the chart; "" to use label
    numberFormat As String   ' value-axis format, e.g. "#,##0" or "0.0%"; "" to leave default
    showLegend As Boolean    ' default True
    heightRows As Long       ' chart height in rows at the anchor cell; 0 lets it default
    widthCols As Long        ' chart width in columns; 0 lets it default
End Type
 
' Captured application state for the duration of a scenario swap.
Public Type CalcGuard
    priorCalculation As XlCalculation
    priorEvents As Boolean
    priorScreen As Boolean
    active As Boolean
End Type
 
Public Const TBL_SCENARIOS As String = "tblScenarios"
 
Public Const COL_SCENARIO_ID As String = "ScenarioID"
Public Const COL_NAME As String = "Name"
Public Const COL_DESCRIPTION As String = "Description"
Public Const COL_CREATED_ON As String = "CreatedOn"
Public Const COL_MODIFIED_ON As String = "ModifiedOn"
Public Const COL_CREATED_BY As String = "CreatedBy"
Public Const COL_MODIFIED_BY As String = "ModifiedBy"
Public Const COL_INCLUDE As String = "Include_In_Run"
Public Const COL_COMPARE As String = "Include_In_Compare"
 
Private Const NAME_ACTIVE_ID As String = "activeScenarioId"
Private Const NAME_ACTIVE_NAME As String = "activeScenarioName"
Private Const NAME_ACTIVE_DIRTY As String = "activeScenarioDirty"
 
' No hyphen: IDs double as AppState keys, which become defined names.
Private Const ID_PREFIX As String = "SCN"
Private Const ID_DIGITS As Long = 4
 
'=======================================================================================
' Active scenario identity
'=======================================================================================
 
Public Function ActiveScenarioIdVal() As String
    Dim cell As Range
 
    Set cell = NamedRangeOrNothing(NAME_ACTIVE_ID)
    If cell Is Nothing Then Exit Function
    ActiveScenarioIdVal = Trim$(CStr(cell.value))
End Function
 
Public Function ActiveScenarioNameVal() As String
    ActiveScenarioNameVal = ScenarioName(ActiveScenarioIdVal())
End Function
 
' True when the active tables have been edited since they were last loaded or saved.
Public Function ActiveIsDirty() As Boolean
    Dim cell As Range
 
    Set cell = NamedRangeOrNothing(NAME_ACTIVE_DIRTY)
    If cell Is Nothing Then Exit Function
    ActiveIsDirty = (UCase$(Trim$(CStr(cell.value))) = "YES")
End Function
 
' Flips the flag outright. Callers that already know the edit is relevant (eg. a
' delete that invalidates the active scenario) use this directly; sheet-level
' Worksheet_Change handlers should prefer MarkActiveDirtyIfRelevant so that edits
' outside the configured inputs and outputs do not trip the flag.
Public Sub MarkActiveDirty()
    SetDirty True
End Sub

' Range-aware entry point for Worksheet_Change handlers. Only marks dirty when the
' changed range intersects an active input table, a scenario scalar's named range,
' or one of the output tables listed in ScenarioDirtyOutputTables.
Public Sub MarkActiveDirtyIfRelevant(target As Range)
    If target Is Nothing Then Exit Sub
    If TargetHitsScenarioState(target) Then MarkActiveDirty
End Sub

Public Sub MarkActiveClean()
    SetDirty False
End Sub

Private Function TargetHitsScenarioState(target As Range) As Boolean
    Dim parts() As ScenarioPart
    Dim scalars() As ScenarioScalar
    Dim outputs As Variant
    Dim i As Long

    parts = ScenarioParts()
    For i = LBound(parts) To UBound(parts)
        If RangeHitsTable(target, parts(i).activeTable) Then
            TargetHitsScenarioState = True
            Exit Function
        End If
    Next i

    scalars = ScenarioScalars()
    For i = LBound(scalars) To UBound(scalars)
        If RangeHitsNamedRange(target, scalars(i).nameRef) Then
            TargetHitsScenarioState = True
            Exit Function
        End If
    Next i

    outputs = ScenarioDirtyOutputTables()
    For i = LBound(outputs) To UBound(outputs)
        If RangeHitsTable(target, CStr(outputs(i))) Then
            TargetHitsScenarioState = True
            Exit Function
        End If
    Next i
End Function

Private Function RangeHitsTable(target As Range, tableName As String) As Boolean
    Dim lo As ListObject

    Set lo = TableOrNothing(tableName)
    If lo Is Nothing Then Exit Function
    If lo.Range Is Nothing Then Exit Function
    If Not target.Worksheet Is lo.Parent Then Exit Function
    RangeHitsTable = Not (Application.Intersect(target, lo.Range) Is Nothing)
End Function

Private Function RangeHitsNamedRange(target As Range, nameRef As String) As Boolean
    Dim rng As Range

    Set rng = NamedRangeOrNothing(nameRef)
    If rng Is Nothing Then Exit Function
    If Not target.Worksheet Is rng.Worksheet Then Exit Function
    RangeHitsNamedRange = Not (Application.Intersect(target, rng) Is Nothing)
End Function
 
Private Sub SetDirty(ByVal value As Boolean)
    Dim cell As Range
    Dim newValue As String
    Dim priorEvents As Boolean
 
    Set cell = NamedRangeOrNothing(NAME_ACTIVE_DIRTY)
    If cell Is Nothing Then Exit Sub
 
    newValue = IIf(value, "YES", "NO")
    ' Writing unconditionally would refire Worksheet_Change on the dirty cell's own
    ' sheet, which calls back into MarkActiveDirty and loops.
    If StrComp(CStr(cell.value), newValue, vbTextCompare) = 0 Then Exit Sub
 
    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    cell.value = newValue
    Application.EnableEvents = priorEvents
End Sub
 
Private Sub SetActiveIdentity(scenarioId As String)
    Dim cell As Range
 
    Set cell = NamedRangeOrNothing(NAME_ACTIVE_ID)
    If Not cell Is Nothing Then cell.value = scenarioId
 
    Set cell = NamedRangeOrNothing(NAME_ACTIVE_NAME)
    If Not cell Is Nothing Then cell.value = ScenarioName(scenarioId)
End Sub
 
'=======================================================================================
' Scenario metadata
'=======================================================================================
 
Public Function ScenarioExists(scenarioId As String) As Boolean
    If Len(scenarioId) = 0 Then Exit Function
    ScenarioExists = (FindRowIndex(RequireTable(TBL_SCENARIOS), COL_SCENARIO_ID, scenarioId) > 0)
End Function

' The reserved current-state ID. Treated specially throughout the scenario flows: it
' never loads inputs, never saves inputs, and cannot be renamed or deleted.
Public Function IsCurrentStateId(scenarioId As String) As Boolean
    IsCurrentStateId = (StrComp(Trim$(scenarioId), CurrentStateScenarioId(), vbTextCompare) = 0)
End Function

' Creates the reserved current-state row in tblScenarios if it is missing. Safe to call
' repeatedly; the row's identity never changes.
Public Sub EnsureCurrentStateRow()
    Dim lo As ListObject
    Dim newRow As ListRow
    Dim currentId As String

    currentId = CurrentStateScenarioId()
    If ScenarioExists(currentId) Then Exit Sub

    Set lo = RequireTable(TBL_SCENARIOS)

    Set newRow = lo.ListRows.Add
    newRow.Range.Cells(1, RequireColumnIndex(lo, COL_SCENARIO_ID)).value = currentId
    newRow.Range.Cells(1, RequireColumnIndex(lo, COL_NAME)).value = CurrentStateScenarioName()
    SetIfPresent lo, newRow, COL_DESCRIPTION, CurrentStateScenarioDescription()
    SetIfPresent lo, newRow, COL_CREATED_ON, Now
    SetIfPresent lo, newRow, COL_MODIFIED_ON, Now
    SetIfPresent lo, newRow, COL_CREATED_BY, Application.UserName
    SetIfPresent lo, newRow, COL_MODIFIED_BY, Application.UserName
    ' The current state is not a candidate for the regular run queue; it has its own.
    SetIfPresent lo, newRow, COL_INCLUDE, False
    SetIfPresent lo, newRow, COL_COMPARE, True
End Sub
 
Public Function ScenarioName(scenarioId As String) As String
    Dim value As Variant
 
    value = MetaValue(scenarioId, COL_NAME)
    If Not IsEmpty(value) Then ScenarioName = CStr(value)
End Function
 
' Resolves a display name back to its ID. Empty when the name is unknown.
Public Function ScenarioIdForName(displayName As String) As String
    Dim lo As ListObject
    Dim rowIndex As Long
 
    Set lo = RequireTable(TBL_SCENARIOS)
    rowIndex = FindRowIndex(lo, COL_NAME, Trim$(displayName))
    If rowIndex = 0 Then Exit Function
 
    ScenarioIdForName = CStr(CellInRow(lo, rowIndex, COL_SCENARIO_ID).value)
End Function
 
Public Function MetaValue(scenarioId As String, column As String) As Variant
    Dim lo As ListObject
    Dim rowIndex As Long
 
    Set lo = RequireTable(TBL_SCENARIOS)
    rowIndex = FindRowIndex(lo, COL_SCENARIO_ID, scenarioId)
    If rowIndex = 0 Then Exit Function
 
    MetaValue = CellInRow(lo, rowIndex, column).value
End Function
 
Public Sub SetMetaValue(scenarioId As String, column As String, value As Variant)
    Dim lo As ListObject
    Dim rowIndex As Long
 
    Set lo = RequireTable(TBL_SCENARIOS)
    rowIndex = FindRowIndex(lo, COL_SCENARIO_ID, scenarioId)
    If rowIndex = 0 Then Exit Sub
 
    CellInRow(lo, rowIndex, column).value = value
End Sub
 
' IDs in table order, or an empty array when there are none.
Public Function ScenarioIds() As Variant
    Dim lo As ListObject
    Dim body As Variant
    Dim ids() As String
    Dim col As Long, r As Long
 
    Set lo = RequireTable(TBL_SCENARIOS)
    body = TableBody(lo)
    If IsEmpty(body) Then
        ScenarioIds = Array()
        Exit Function
    End If
 
    col = RequireColumnIndex(lo, COL_SCENARIO_ID)
    ReDim ids(1 To UBound(body, 1))
    For r = 1 To UBound(body, 1)
        ids(r) = CStr(body(r, col))
    Next r
 
    ScenarioIds = ids
End Function
 
' Highest existing number plus one, so deleting a scenario never frees its ID for reuse.
Private Function NextScenarioId() As String
    Dim lo As ListObject
    Dim body As Variant
    Dim col As Long, r As Long
    Dim highest As Long, candidate As Long
    Dim raw As String
 
    Set lo = RequireTable(TBL_SCENARIOS)
    body = TableBody(lo)
 
    If Not IsEmpty(body) Then
        col = RequireColumnIndex(lo, COL_SCENARIO_ID)
        For r = 1 To UBound(body, 1)
            raw = CStr(body(r, col))
            If UCase$(Left$(raw, Len(ID_PREFIX))) = UCase$(ID_PREFIX) Then
                raw = Mid$(raw, Len(ID_PREFIX) + 1)
                If IsNumeric(raw) Then
                    candidate = CLng(raw)
                    If candidate > highest Then highest = candidate
                End If
            End If
        Next r
    End If
 
    NextScenarioId = ID_PREFIX & Format$(highest + 1, String$(ID_DIGITS, "0"))
End Function
 
' Names are what the user picks from, so they have to be unique.
Public Function NameIsAvailable(displayName As String, Optional ignoreId As String = "") As Boolean
    Dim foundId As String
 
    If Len(Trim$(displayName)) = 0 Then Exit Function
 
    foundId = ScenarioIdForName(displayName)
    NameIsAvailable = (Len(foundId) = 0) Or (StrComp(foundId, ignoreId, vbTextCompare) = 0)
End Function
 
'=======================================================================================
' Create, load, save, delete
'=======================================================================================
 
' Creates a scenario seeded from the templates and loads it. Returns the new ID.
Public Function NewScenario(displayName As String, Optional description As String = "", _
                            Optional effectiveDate As Variant = Empty) As String
    Dim lo As ListObject
    Dim newRow As ListRow
    Dim scenarioId As String
    Dim parts() As ScenarioPart
    Dim scalars() As ScenarioScalar
    Dim i As Long
    Dim guard As CalcGuard
 
    If StrComp(Trim$(displayName), CurrentStateScenarioName(), vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 621, "ScenarioStore", _
                  "'" & displayName & "' is reserved for the current-state row."
    End If

    If Not NameIsAvailable(displayName) Then
        Err.Raise vbObjectError + 610, "ScenarioStore", _
                  "A scenario named '" & displayName & "' already exists."
    End If

    scenarioId = NextScenarioId()
    Set lo = RequireTable(TBL_SCENARIOS)

    guard = BeginBatch("Creating " & displayName)
    On Error GoTo Failed
 
    Set newRow = lo.ListRows.Add
    newRow.Range.Cells(1, RequireColumnIndex(lo, COL_SCENARIO_ID)).value = scenarioId
    newRow.Range.Cells(1, RequireColumnIndex(lo, COL_NAME)).value = Trim$(displayName)
    SetIfPresent lo, newRow, COL_DESCRIPTION, description
    SetIfPresent lo, newRow, COL_CREATED_ON, Now
    SetIfPresent lo, newRow, COL_MODIFIED_ON, Now
    SetIfPresent lo, newRow, COL_CREATED_BY, Application.UserName
    SetIfPresent lo, newRow, COL_MODIFIED_BY, Application.UserName
    SetIfPresent lo, newRow, COL_INCLUDE, True
    SetIfPresent lo, newRow, COL_COMPARE, True
 
    scalars = ScenarioScalars()
    For i = LBound(scalars) To UBound(scalars)
        SetIfPresent lo, newRow, scalars(i).column, DefaultScalar(scalars(i), effectiveDate)
    Next i
 
    parts = ScenarioParts()
    For i = LBound(parts) To UBound(parts)
        LoadTemplateInto parts(i)
        SavePart parts(i), scenarioId
    Next i
 
    ApplyScalars scenarioId
    SetActiveIdentity scenarioId
    MarkActiveClean
 
    EndBatch guard
    NewScenario = scenarioId
    Exit Function
 
Failed:
    EndBatch guard
    Err.Raise Err.Number, "ScenarioStore.NewScenario", Err.description
End Function
 
' Replaces the active tables with the stored rows for this scenario.
Public Sub LoadScenario(scenarioId As String)
    Dim parts() As ScenarioPart
    Dim i As Long
    Dim guard As CalcGuard

    If IsCurrentStateId(scenarioId) Then
        Err.Raise vbObjectError + 616, "ScenarioStore", _
                  "The current state is not an editable scenario and cannot be loaded. " & _
                  "Use Run Current State to refresh its results."
    End If

    If Not ScenarioExists(scenarioId) Then
        Err.Raise vbObjectError + 611, "ScenarioStore", _
                  "Scenario '" & scenarioId & "' was not found."
    End If
 
    guard = BeginBatch("Loading " & ScenarioName(scenarioId))
    On Error GoTo Failed
 
    parts = ScenarioParts()
    For i = LBound(parts) To UBound(parts)
        CopyInto parts(i).storeTable, parts(i).activeTable, COL_SCENARIO_ID, scenarioId
    Next i
 
    ApplyScalars scenarioId
    SetActiveIdentity scenarioId
    MarkActiveClean
 
    EndBatch guard
    Exit Sub
 
Failed:
    EndBatch guard
    Err.Raise Err.Number, "ScenarioStore.LoadScenario", Err.description
End Sub
 
' Writes the active tables back over the stored rows for the active scenario.
Public Sub SaveActiveScenario()
    SaveTo ActiveScenarioIdVal()
End Sub
 
' Saves the active tables under a brand new scenario and makes it the active one.
Public Function SaveActiveScenarioAs(displayName As String, _
                                     Optional description As String = "") As String
    Dim lo As ListObject
    Dim newRow As ListRow
    Dim scenarioId As String
 
    If Not NameIsAvailable(displayName) Then
        Err.Raise vbObjectError + 612, "ScenarioStore", _
                  "A scenario named '" & displayName & "' already exists."
    End If
 
    scenarioId = NextScenarioId()
    Set lo = RequireTable(TBL_SCENARIOS)
 
    Set newRow = lo.ListRows.Add
    newRow.Range.Cells(1, RequireColumnIndex(lo, COL_SCENARIO_ID)).value = scenarioId
    newRow.Range.Cells(1, RequireColumnIndex(lo, COL_NAME)).value = Trim$(displayName)
    SetIfPresent lo, newRow, COL_DESCRIPTION, description
    SetIfPresent lo, newRow, COL_CREATED_ON, Now
    SetIfPresent lo, newRow, COL_CREATED_BY, Application.UserName
    SetIfPresent lo, newRow, COL_MODIFIED_BY, Application.UserName
    SetIfPresent lo, newRow, COL_INCLUDE, True
    SetIfPresent lo, newRow, COL_COMPARE, True
 
    ' SaveTo lifts the scalars straight off the model, so the copy keeps the live values.
    SaveTo scenarioId
 
    SaveActiveScenarioAs = scenarioId
End Function
 
Private Sub SaveTo(scenarioId As String)
    Dim parts() As ScenarioPart
    Dim scalars() As ScenarioScalar
    Dim i As Long
    Dim guard As CalcGuard

    If IsCurrentStateId(scenarioId) Then
        Err.Raise vbObjectError + 617, "ScenarioStore", _
                  "The current state is not an editable scenario and cannot be saved over."
    End If

    If Len(scenarioId) = 0 Or Not ScenarioExists(scenarioId) Then
        Err.Raise vbObjectError + 613, "ScenarioStore", _
                  "There is no active scenario to save. Create or load one first."
    End If
 
    guard = BeginBatch("Saving " & ScenarioName(scenarioId))
    On Error GoTo Failed
 
    ' Capture the scalars from the model before the store rows, so the metadata row and
    ' the input rows always describe the same state.
    scalars = ScenarioScalars()
    For i = LBound(scalars) To UBound(scalars)
        SetMetaValue scenarioId, scalars(i).column, ActiveScalarValue(scalars(i))
    Next i
 
    parts = ScenarioParts()
    For i = LBound(parts) To UBound(parts)
        SavePart parts(i), scenarioId
    Next i
 
    SetMetaValue scenarioId, COL_MODIFIED_ON, Now
    SetMetaValue scenarioId, COL_MODIFIED_BY, Application.UserName
    SetActiveIdentity scenarioId
    MarkActiveClean
 
    EndBatch guard
    Exit Sub
 
Failed:
    EndBatch guard
    Err.Raise Err.Number, "ScenarioStore.SaveTo", Err.description
End Sub
 
' Removes the metadata row, every stored input row and every stored result row.
Public Sub DeleteScenario(scenarioId As String)
    Dim lo As ListObject
    Dim parts() As ScenarioPart
    Dim rowIndex As Long
    Dim i As Long
    Dim guard As CalcGuard

    If IsCurrentStateId(scenarioId) Then
        Err.Raise vbObjectError + 618, "ScenarioStore", _
                  "The current state row is reserved and cannot be deleted."
    End If

    If Not ScenarioExists(scenarioId) Then Exit Sub
 
    guard = BeginBatch("Deleting " & ScenarioName(scenarioId))
    On Error GoTo Failed
 
    parts = ScenarioParts()
    For i = LBound(parts) To UBound(parts)
        DeleteStoreRows parts(i).storeTable, scenarioId
    Next i
    DeleteStoreRows ResultsTableName(), scenarioId
 
    Set lo = RequireTable(TBL_SCENARIOS)
    rowIndex = FindRowIndex(lo, COL_SCENARIO_ID, scenarioId)
    If rowIndex > 0 Then lo.ListRows(rowIndex).Delete
 
    ' The active scenario no longer exists, so stop pretending it does.
    If StrComp(ActiveScenarioIdVal(), scenarioId, vbTextCompare) = 0 Then
        SetActiveIdentity ""
        MarkActiveDirty
    End If
 
    EndBatch guard
    Exit Sub
 
Failed:
    EndBatch guard
    Err.Raise Err.Number, "ScenarioStore.DeleteScenario", Err.description
End Sub
 
Public Sub RenameScenario(scenarioId As String, newName As String)
    If IsCurrentStateId(scenarioId) Then
        Err.Raise vbObjectError + 619, "ScenarioStore", _
                  "The current state row is reserved and cannot be renamed."
    End If

    If Not NameIsAvailable(newName, scenarioId) Then
        Err.Raise vbObjectError + 614, "ScenarioStore", _
                  "A scenario named '" & newName & "' already exists."
    End If
 
    SetMetaValue scenarioId, COL_NAME, Trim$(newName)
    SetMetaValue scenarioId, COL_MODIFIED_ON, Now
    SetMetaValue scenarioId, COL_MODIFIED_BY, Application.UserName
    If StrComp(ActiveScenarioIdVal(), scenarioId, vbTextCompare) = 0 Then SetActiveIdentity scenarioId
End Sub
 
'=======================================================================================
' Part copying
'=======================================================================================
 
' Seeds the active table from the part's template, or empties it when there is none.
Private Sub LoadTemplateInto(part As ScenarioPart)
    If Len(part.templateTable) = 0 Then
        ResizeTableBody RequireTable(part.activeTable), 0
    Else
        CopyInto part.templateTable, part.activeTable, "", ""
    End If
End Sub
 
' Copies rows from src into dst, matching columns by header name. Columns dst has but
' src does not are left alone, as are dst's calculated columns. An empty filterColumn
' copies every row.
Private Sub CopyInto(srcName As String, dstName As String, _
                     filterColumn As String, filterValue As String)
    Dim src As ListObject, dst As ListObject
    Dim srcBody As Variant
    Dim isFormula As Variant
    Dim rowMap() As Long
    Dim colMap() As Long
    Dim colValues() As Variant
    Dim filterCol As Long
    Dim n As Long, r As Long, c As Long, i As Long
 
    Set src = RequireTable(srcName)
    Set dst = RequireTable(dstName)
 
    srcBody = TableBody(src)
 
    filterCol = 0
    If Len(filterColumn) > 0 Then filterCol = RequireColumnIndex(src, filterColumn)
 
    If Not IsEmpty(srcBody) Then
        ReDim rowMap(1 To UBound(srcBody, 1))
        For r = 1 To UBound(srcBody, 1)
            If filterCol = 0 Then
                n = n + 1
                rowMap(n) = r
            ElseIf StrComp(CStr(srcBody(r, filterCol)), filterValue, vbTextCompare) = 0 Then
                n = n + 1
                rowMap(n) = r
            End If
        Next r
    End If
 
    ResizeTableBody dst, n
    If n = 0 Then Exit Sub
 
    ' Read after the resize: Excel refills calculated columns as the table grows.
    isFormula = FormulaColumnFlags(dst)
 
    ReDim colMap(1 To dst.ListColumns.count)
    For c = 1 To dst.ListColumns.count
        colMap(c) = ColumnIndexOrZero(src, dst.ListColumns(c).Name)
    Next c
 
    ReDim colValues(1 To n, 1 To 1)
    For c = 1 To dst.ListColumns.count
        If colMap(c) > 0 And Not isFormula(c) Then
            For i = 1 To n
                colValues(i, 1) = srcBody(rowMap(i), colMap(c))
            Next i
            dst.ListColumns(c).DataBodyRange.value = colValues
        End If
    Next c
End Sub
 
' Rewrites the store table as "every row belonging to another scenario, then the active
' rows stamped with this scenario's ID". Rebuilding wholesale keeps it one block write.
Private Sub SavePart(part As ScenarioPart, scenarioId As String)
    Dim store As ListObject, active As ListObject
    Dim storeBody As Variant, activeBody As Variant
    Dim result() As Variant
    Dim srcCol() As Long
    Dim idCol As Long, cols As Long
    Dim keepRows As Long, addRows As Long, outRows As Long
    Dim r As Long, c As Long, o As Long
 
    Set store = RequireTable(part.storeTable)
    Set active = RequireTable(part.activeTable)
 
    idCol = RequireColumnIndex(store, COL_SCENARIO_ID)
    cols = store.ListColumns.count
 
    storeBody = TableBody(store)
    activeBody = TableBody(active)
 
    ReDim srcCol(1 To cols)
    For c = 1 To cols
        srcCol(c) = ColumnIndexOrZero(active, store.ListColumns(c).Name)
    Next c
 
    If Not IsEmpty(storeBody) Then
        For r = 1 To UBound(storeBody, 1)
            If StrComp(CStr(storeBody(r, idCol)), scenarioId, vbTextCompare) <> 0 Then
                keepRows = keepRows + 1
            End If
        Next r
    End If
    If Not IsEmpty(activeBody) Then addRows = UBound(activeBody, 1)
 
    outRows = keepRows + addRows
    If outRows = 0 Then
        ResizeTableBody store, 0
        Exit Sub
    End If
 
    ReDim result(1 To outRows, 1 To cols)
 
    If Not IsEmpty(storeBody) Then
        For r = 1 To UBound(storeBody, 1)
            If StrComp(CStr(storeBody(r, idCol)), scenarioId, vbTextCompare) <> 0 Then
                o = o + 1
                For c = 1 To cols
                    result(o, c) = storeBody(r, c)
                Next c
            End If
        Next r
    End If
 
    If Not IsEmpty(activeBody) Then
        For r = 1 To UBound(activeBody, 1)
            o = o + 1
            For c = 1 To cols
                If c = idCol Then
                    result(o, c) = scenarioId
                ElseIf srcCol(c) > 0 Then
                    result(o, c) = activeBody(r, srcCol(c))
                End If
            Next c
        Next r
    End If
 
    ResizeTableBody store, outRows
    store.DataBodyRange.value = result
End Sub
 
' Drops every row belonging to one scenario from a long-format store table.
Public Sub DeleteStoreRows(storeTableName As String, scenarioId As String)
    Dim lo As ListObject
    Dim body As Variant
    Dim result() As Variant
    Dim idCol As Long, cols As Long
    Dim keepRows As Long
    Dim r As Long, c As Long, o As Long
 
    Set lo = TableOrNothing(storeTableName)
    If lo Is Nothing Then Exit Sub
 
    body = TableBody(lo)
    If IsEmpty(body) Then Exit Sub
 
    idCol = RequireColumnIndex(lo, COL_SCENARIO_ID)
    cols = lo.ListColumns.count
 
    For r = 1 To UBound(body, 1)
        If StrComp(CStr(body(r, idCol)), scenarioId, vbTextCompare) <> 0 Then keepRows = keepRows + 1
    Next r
 
    If keepRows = UBound(body, 1) Then Exit Sub
 
    If keepRows = 0 Then
        ResizeTableBody lo, 0
        Exit Sub
    End If
 
    ReDim result(1 To keepRows, 1 To cols)
    For r = 1 To UBound(body, 1)
        If StrComp(CStr(body(r, idCol)), scenarioId, vbTextCompare) <> 0 Then
            o = o + 1
            For c = 1 To cols
                result(o, c) = body(r, c)
            Next c
        End If
    Next r
 
    ResizeTableBody lo, keepRows
    lo.DataBodyRange.value = result
End Sub
 
'=======================================================================================
' Scalars
'=======================================================================================
 
' Pushes the stored scalars for this scenario into the defined names the model reads.
Private Sub ApplyScalars(scenarioId As String)
    Dim scalars() As ScenarioScalar
    Dim cell As Range
    Dim i As Long
 
    scalars = ScenarioScalars()
    For i = LBound(scalars) To UBound(scalars)
        Set cell = NamedRangeOrNothing(scalars(i).nameRef)
        If cell Is Nothing Then
            Err.Raise vbObjectError + 615, "ScenarioStore", _
                      "Named range '" & scalars(i).nameRef & "' was not found. " & _
                      "The model cannot read the scenario's " & scalars(i).label & "."
        End If
        cell.value = MetaValue(scenarioId, scalars(i).column)
    Next i
End Sub
 
Private Function ActiveScalarValue(scalar As ScenarioScalar) As Variant
    Dim cell As Range
 
    Set cell = NamedRangeOrNothing(scalar.nameRef)
    If cell Is Nothing Then Exit Function
    ActiveScalarValue = cell.value
End Function
 
Private Function DefaultScalar(scalar As ScenarioScalar, seed As Variant) As Variant
    If IsEmpty(seed) Then
        DefaultScalar = ScenarioScalarDefault(scalar)
    Else
        DefaultScalar = seed
    End If
End Function
 
Private Sub SetIfPresent(lo As ListObject, target As ListRow, column As String, value As Variant)
    Dim col As Long
 
    col = ColumnIndexOrZero(lo, column)
    If col = 0 Then Exit Sub
    target.Range.Cells(1, col).value = value
End Sub
 
'=======================================================================================
' Batch guard
'=======================================================================================
 
' Scenario swaps touch thousands of cells; recalculating or firing change events on each
' one would be unusably slow and would trip the dirty flag on every write.
Public Function BeginBatch(Optional status As String = "") As CalcGuard
    With BeginBatch
        .priorCalculation = Application.Calculation
        .priorEvents = Application.EnableEvents
        .priorScreen = Application.ScreenUpdating
        .active = True
    End With
 
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If Len(status) > 0 Then Application.StatusBar = status
End Function
 
Public Sub EndBatch(guard As CalcGuard)
    If Not guard.active Then Exit Sub
 
    Application.StatusBar = False
    Application.EnableEvents = guard.priorEvents
    Application.ScreenUpdating = guard.priorScreen
    Application.Calculation = guard.priorCalculation
    guard.active = False
End Sub


