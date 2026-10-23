Attribute VB_Name = "ScenarioRun"
Option Explicit
 
' Runs scenarios and collects their metrics into tblScenarioResults.
'
' This replaces VBAMacro_ScenarioAnalysis.txt, which could only ever run the four
' scenarios wired into input_scenario_1 .. input_scenario_4. Here the list of scenarios
' and the list of metrics are both data, so neither is capped and neither needs code
' changes to extend.
'
' Because results are stored per ScenarioID, the comparison sheet can show any number of
' scenarios side by side and nothing has to be re-run to redraw a chart.
'
' What counts as a metric is decided entirely by ScenarioMetrics: this module only
' recalculates, reads the included metrics and stores them.
 
' Refreshes the current-state baseline. Independent of whatever scenario is loaded:
' recalculates the current-state calc sheets, reads metrics from their parallel output
' tables (declared in ScenarioConfig's MetricSourceTables) and stores them in
' tblScenarioResults under the reserved CURRENT ScenarioID. Nothing in the active
' scenario is touched.
Public Sub RunCurrentState()
    Dim recalcs() As RecalcSpec
    Dim metrics() As MetricSpec
    Dim currentId As String
    Dim guard As CalcGuard
    Dim i As Long

    On Error GoTo Failed

    currentId = CurrentStateScenarioId()
    EnsureCurrentStateRow

    guard = BeginBatch("Refreshing the current state...")

    recalcs = CurrentStateRecalcs()
    For i = LBound(recalcs) To UBound(recalcs)
        RecalcSheetInline recalcs(i)
    Next i

    Application.CalculateFull
    metrics = IncludedMetricSpecs()
    StoreResults currentId, ReadMetricValuesForCurrentState(metrics)
    LastCalculatedOn(currentId) = Now

    EndBatch guard

    RefreshScenarioManager
    MsgBox "Refreshed the current state and stored its results.", vbInformation
    Exit Sub

Failed:
    EndBatch guard
    MsgBox "The current-state refresh did not complete." & vbNewLine & vbNewLine & _
           Err.description, vbCritical
End Sub

' Recalculates one sheet without the status cell / message-box plumbing of RunRecalc.
' The caller is already inside a BeginBatch guard and will show its own summary.
Private Sub RecalcSheetInline(spec As RecalcSpec)
    Dim ws As Worksheet

    ' An unfilled slot in CurrentStateRecalcs() has no sheet; skip it rather than abort.
    If Len(spec.targetSheet) = 0 Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(spec.targetSheet)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 640, "ScenarioRun.RecalcSheetInline", _
                  "Sheet '" & spec.targetSheet & "' named in CurrentStateRecalcs() " & _
                  "was not found in this workbook."
    End If
    ws.EnableCalculation = True
    ws.UsedRange.Dirty
    ws.Calculate
    LastCalculatedOn(spec.key) = Now
    RefreshRecalcStatus spec
    ws.EnableCalculation = False
End Sub

' Recalculates the loaded scenario and stores its metrics, leaving it loaded.
Public Sub RunActiveScenario()
    Dim scenarioId As String
    Dim problem As String
    Dim metrics() As MetricSpec
    Dim guard As CalcGuard
 
    On Error GoTo Failed
 
    scenarioId = ActiveScenarioIdVal()
    If Len(scenarioId) = 0 Then
        MsgBox "No scenario is loaded. Create one or load one first.", vbExclamation
        Exit Sub
    End If
 
    If ActiveIsDirty() Then
        If MsgBox("The loaded scenario has unsaved changes. Run it anyway?" & vbNewLine & _
                  "The results will be stored against '" & ScenarioName(scenarioId) & "'.", _
                  vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If
 
    If Not ActiveScenarioIsValid(problem) Then
        MsgBox "This scenario cannot be run yet:" & vbNewLine & vbNewLine & problem, vbExclamation
        Exit Sub
    End If
 
    guard = BeginBatch("Running " & ScenarioName(scenarioId) & "...")
    Application.CalculateFull
    PrepareScenarioOutputs
    metrics = IncludedMetricSpecs()
    StoreResults scenarioId, ReadMetricValues(metrics)
    LastCalculatedOn(scenarioId) = Now
    EndBatch guard
 
    RefreshScenarioManager
    MsgBox "Ran '" & ScenarioName(scenarioId) & "' and stored its results.", vbInformation
    Exit Sub
 
Failed:
    EndBatch guard
    MsgBox "The scenario run did not complete." & vbNewLine & vbNewLine & Err.description, vbCritical
End Sub
 
' Runs every scenario flagged Include_In_Run, then restores whatever was loaded before.
Public Sub RunIncludedScenarios()
    Dim ids As Variant
    Dim queue() As String
    Dim metrics() As MetricSpec
    Dim restoreId As String
    Dim guard As CalcGuard
    Dim queued As Long, i As Long
    Dim failures As String
 
    On Error GoTo Failed
 
    If ActiveIsDirty() Then
        If MsgBox("The loaded scenario has unsaved changes and running will discard them." & _
                  vbNewLine & "Continue?", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
    End If
 
    ids = ScenarioIds()
    If UBound(ids) < LBound(ids) Then
        MsgBox "There are no saved scenarios to run.", vbExclamation
        Exit Sub
    End If
 
    ReDim queue(1 To UBound(ids) - LBound(ids) + 1)
    For i = LBound(ids) To UBound(ids)
        If IsIncluded(CStr(ids(i))) Then
            queued = queued + 1
            queue(queued) = CStr(ids(i))
        End If
    Next i
 
    If queued = 0 Then
        MsgBox "No scenarios are flagged " & COL_INCLUDE & " in " & TBL_SCENARIOS & ".", _
               vbExclamation
        Exit Sub
    End If
 
    metrics = IncludedMetricSpecs()
    restoreId = ActiveScenarioIdVal()
 
    guard = BeginBatch()
    For i = 1 To queued
        Application.StatusBar = "Running " & ScenarioName(queue(i)) & _
                                " (" & i & " of " & queued & ")..."
        failures = failures & RunOne(queue(i), metrics)
    Next i
 
    If Len(restoreId) > 0 Then
        If ScenarioExists(restoreId) Then LoadScenario restoreId
    End If
    EndBatch guard
 
    RefreshScenarioManager
    ReportRun queued, failures
    Exit Sub
 
Failed:
    EndBatch guard
    MsgBox "The scenario run did not complete." & vbNewLine & vbNewLine & Err.description, vbCritical
End Sub
 
' Returns an empty string on success, or a line describing why the scenario was skipped.
Private Function RunOne(scenarioId As String, metrics() As MetricSpec) As String
    Dim problem As String
    Dim values As Variant                                                              ' PERF
    Dim t As Double                                                                    ' PERF

    On Error GoTo Failed

    t = Timer: LoadScenario scenarioId                                                 ' PERF
    Debug.Print scenarioId & " Load  " & Format$(Timer - t, "0.000") & "s"             ' PERF

    If Not ActiveScenarioIsValid(problem) Then
        RunOne = "- " & ScenarioName(scenarioId) & ": " & FirstLine(problem) & vbNewLine
        Exit Function
    End If

    t = Timer: Application.CalculateFull                                               ' PERF
    Debug.Print scenarioId & " Calc  " & Format$(Timer - t, "0.000") & "s"             ' PERF
    t = Timer: PrepareScenarioOutputs                                                  ' PERF
    Debug.Print scenarioId & " Prep  " & Format$(Timer - t, "0.000") & "s"             ' PERF
    t = Timer: values = ReadMetricValues(metrics)                                      ' PERF
    Debug.Print scenarioId & " Read  " & Format$(Timer - t, "0.000") & "s"             ' PERF
    t = Timer: StoreResults scenarioId, values                                         ' PERF
    Debug.Print scenarioId & " Store " & Format$(Timer - t, "0.000") & "s"             ' PERF
    LastCalculatedOn(scenarioId) = Now
    Exit Function

Failed:
    RunOne = "- " & ScenarioName(scenarioId) & ": " & Err.description & vbNewLine
End Function
 
'=======================================================================================
' Results
'=======================================================================================
 
' Replaces this scenario's rows in tblScenarioResults with the newly harvested ones.
' === PERF PATCH 2026-10-18 #1 StoreResults bulk write === START
' Was: a per-cell write loop (~3,400 COM round trips for 844 metrics x 4 columns).
' Now: build four column arrays and write each as a single block.
Private Sub StoreResults(scenarioId As String, metrics As Variant)
    Dim lo As ListObject
    Dim startRow As Long, n As Long
    Dim cId As Long, cRun As Long, cMetric As Long, cValue As Long
    Dim idCol() As Variant, runCol() As Variant
    Dim metricCol() As Variant, valueCol() As Variant
    Dim runOn As Date
    Dim i As Long

    DeleteStoreRows ResultsTableName(), scenarioId
    If IsEmpty(metrics) Then Exit Sub

    Set lo = RequireTable(ResultsTableName())
    cId = RequireColumnIndex(lo, COL_SCENARIO_ID)
    cRun = RequireColumnIndex(lo, COL_RESULT_RUN_ON)
    cMetric = RequireColumnIndex(lo, COL_RESULT_METRIC)
    cValue = RequireColumnIndex(lo, COL_RESULT_VALUE)

    n = UBound(metrics, 1)
    startRow = TableRowCount(lo)
    ResizeTableBody lo, startRow + n

    runOn = Now
    ReDim idCol(1 To n, 1 To 1)
    ReDim runCol(1 To n, 1 To 1)
    ReDim metricCol(1 To n, 1 To 1)
    ReDim valueCol(1 To n, 1 To 1)
    For i = 1 To n
        idCol(i, 1) = scenarioId
        runCol(i, 1) = runOn
        metricCol(i, 1) = metrics(i, 1)
        valueCol(i, 1) = metrics(i, 2)
    Next i

    With lo.DataBodyRange
        .Cells(startRow + 1, cId).Resize(n, 1).value = idCol
        .Cells(startRow + 1, cRun).Resize(n, 1).value = runCol
        .Cells(startRow + 1, cMetric).Resize(n, 1).value = metricCol
        .Cells(startRow + 1, cValue).Resize(n, 1).value = valueCol
    End With
End Sub
' === PERF PATCH 2026-10-18 #1 StoreResults bulk write === END
 
Private Sub ReportRun(ByVal queued As Long, ByVal failures As String)
    If Len(failures) = 0 Then
        MsgBox "Ran " & queued & " scenario(s). Results are in " & ResultsTableName() & ".", _
               vbInformation
    Else
        MsgBox "Ran " & queued & " scenario(s). Some did not produce results:" & vbNewLine & _
               vbNewLine & failures, vbExclamation
    End If
End Sub
 
Private Function IsIncluded(scenarioId As String) As Boolean
    ' An absent column means run everything. Check it first: MetaValue raises on a
    ' column it cannot find rather than returning Empty.
    If ColumnIndexOrZero(RequireTable(TBL_SCENARIOS), COL_INCLUDE) = 0 Then
        IsIncluded = True
        Exit Function
    End If
 
    IsIncluded = IsTruthyValue(MetaValue(scenarioId, COL_INCLUDE))
End Function
 
Private Function FirstLine(text As String) As String
    Dim cut As Long
 
    cut = InStr(text, vbNewLine)
    If cut = 0 Then FirstLine = text Else FirstLine = Left$(text, cut - 1)
End Function


