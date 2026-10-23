Option Explicit

' Requires the Solver add-in reference: Tools > References > Solver in the VBA IDE.

' Populated after every successful solver run.
Public LastSolverRun As Date

' Sheet that holds all calculations and named ranges. Update to match your workbook.
Private Const CALC_SHEET As String = "Calculations"

' Named ranges for the objective, decision variables, and their bounds.
' Update these strings to match the names defined in your workbook.
Private Const OBJ_CELL As String = "Objective"

Private Const VAR1 As String = "Var1"
Private Const VAR2 As String = "Var2"
Private Const VAR3 As String = "Var3"

Private Const VAR1_LB As String = "Var1_LB"
Private Const VAR1_UB As String = "Var1_UB"
Private Const VAR2_LB As String = "Var2_LB"
Private Const VAR2_UB As String = "Var2_UB"
Private Const VAR3_LB As String = "Var3_LB"
Private Const VAR3_UB As String = "Var3_UB"

Public Sub RunSolver()
    Dim ws     As Worksheet
    Dim result As Long

    If Not SolverReady() Then
        MsgBox "Solver is not enabled. Activate it via:" & vbNewLine & _
               "File > Options > Add-ins > Manage: Excel Add-ins > Go > tick Solver Add-in.", _
               vbExclamation, "Solver Not Available"
        Exit Sub
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

' Attempts SolverReset as a live check that the add-in is loaded and callable.
' Also serves as the reset call before configuring a new run.
Private Function SolverReady() As Boolean
    Dim errNum As Long
    On Error Resume Next
    Application.AddIns("Solver Add-in").Installed = True
    DoEvents
    SolverReset
    errNum = Err.Number
    Err.Clear
    On Error GoTo 0
    SolverReady = (errNum = 0)
End Function

' Shifted Gompertz: L * exp(-b * exp(-k*(x - x0)))
' L  = upper asymptote, b = displacement (b > 0), k = growth rate, x0 = horizontal shift
Public Function ShiftedGompertz(x As Double, L As Double, b As Double, _
                                k As Double, x0 As Double) As Double
Attribute ShiftedGompertz.VB_Description = "Shifted Gompertz curve: L * exp(-b * exp(-k*(x-x0))). L=upper asymptote, b=displacement (b>0), k=growth rate, x0=horizontal shift."
    ShiftedGompertz = L * Exp(-b * Exp(-k * (x - x0)))
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
