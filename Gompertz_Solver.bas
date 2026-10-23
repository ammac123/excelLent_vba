Attribute VB_Name = "Gompertz_Solver"
Option Explicit

' Requires the Solver add-in reference: Tools > References > Solver in the VBA IDE

' Populated after every successful solver run
Public LastSolverRun As Date

' Sheet that holds all calculations and named ranges.
Private Const CALC_SHEET As String = "27_CALC_App_Flighting"

' Named ranges for the objective, decision variables, and their bounds.
Private Const OBJ_CELL As String = "_weightedMSSEGompertz"

Private Const VAR1 As String = "Gompertz_Var_x0"
Private Const VAR2 As String = "Gompertz_Var_k"
Private Const VAR3 As String = "Gompertz_Var_a"

Private Const VAR1_LB As String = "Gompertz_LBound_x0"
Private Const VAR1_UB As String = "Gompertz_UBound_x0"
Private Const VAR2_LB As String = "Gompertz_LBound_k"
Private Const VAR2_UB As String = "Gompertz_UBound_k"
Private Const VAR3_LB As String = "Gompertz_LBound_a"
Private Const VAR3_UB As String = "Gompertz_UBound_a"

Public Function LastSolverRunDate() As Variant
    LastSolverRunDate = LastSolverRun
End Function


Public Sub RunFlightingSolver()
    Dim ws      As Worksheet
    Dim result  As Long
    
    If Not SolverReady() Then
        If Not SolverReady() Then
            MsgBox "Solver is not enabled. Activate it via:" & vbNewLine & _
                   "File > Options > Add-ins > Manage: Excel Add-ins > Go > tick Solver Add-in.", _
                   vbExclamation, "Solver Not Available"
            Exit Sub
        End If
    End If
    
    
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & CALC_SHEET & "' was not found in this workbook.", vbCritical, "Solver"
        Exit Sub
    End If
    
    SolverOk SetCell:=ws.Range(OBJ_CELL), _
             MaxMinVal:=2, _
             ByChange:=Application.Union(ws.Range(VAR1), ws.Range(VAR2), ws.Range(VAR3)), _
             Engine:=1
    
    SolverAdd CellRef:=ws.Range(VAR1), Relation:=3, FormulaText:=ws.Range(VAR1_LB)
    SolverAdd CellRef:=ws.Range(VAR1), Relation:=1, FormulaText:=ws.Range(VAR1_UB)
    SolverAdd CellRef:=ws.Range(VAR2), Relation:=3, FormulaText:=ws.Range(VAR2_LB)
    SolverAdd CellRef:=ws.Range(VAR2), Relation:=1, FormulaText:=ws.Range(VAR2_UB)
    SolverAdd CellRef:=ws.Range(VAR3), Relation:=3, FormulaText:=ws.Range(VAR3_LB)
    SolverAdd CellRef:=ws.Range(VAR3), Relation:=1, FormulaText:=ws.Range(VAR3_UB)
    
    SolverOptions MaxTime:=30, _
                  Precision:=0.00000001, _
                  Scaling:=True, _
                  Convergence:=0.00000001, _
                  Derivatives:=1, _
                  AssumeNonNeg:=True, _
                  RequireBounds:=True
    
    result = SolverSolve(UserFinish:=True)
    SolverFinish KeepFinal:=1
    
    LastSolverRun = Now
    
    MsgBox SolverResultText(result), vbInformation, "Solver"
End Sub


Private Function SolverReady() As Boolean
    Dim errNum As Long
    On Error Resume Next
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0
    SolverReady = (errNum = 0)
End Function


Private Function SolverResultText(result As Long) As String
    Select Case result
        Case 0: SolverResultText = "Solver found a solution. All constraints are satisfied."
        Case 1: SolverResultText = "Solver converged to the current solution."
        Case 2: SolverResultText = "Solver did not converge."
        Case 3: SolverResultText = "Solver could not find a feasible solution."
        Case 5: SolverResultText = "Solver stopped: maximum time or iteration limit reached."
        Case Else: SolverResultText = "Solver stopped with result code " & result & "."
    End Select
End Function
