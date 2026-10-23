Attribute VB_Name = "zz_QualityOfLIfe"
Public Sub HideWorkbookGridlines()
    Dim ws As Worksheet
    Dim currIndex As Long
    
    Application.ScreenUpdating = False
    currIndex = ActiveWorkbook.ActiveSheet.Index
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    Next ws
    
    ActiveWorkbook.Worksheets(currIndex).Activate
    Application.ScreenUpdating = True
End Sub


Public Function Gompertz(ByVal x As Double, _
                         ByVal L As Double, _
                         ByVal k As Double, _
                         ByVal a As Double) As Double
    Gompertz = L * Exp(-k * Exp(-a * x))
End Function


Function tableName(ref As Range) As Variant
    Dim lo As ListObject
    
    On Error Resume Next
    Set lo = ref.Cells(1, 1).ListObject
    On Error GoTo 0
    
    If lo Is Nothing Then
        tableName = CVErr(xlErrNA)
    Else
        tableName = lo.Name
    End If
End Function


Public Sub RegisterGompertz()
    Application.MacroOptions _
        Macro:="Gompertz", _
        Description:="Gompertz curve: L * EXP(-k * EXP(-a * x)). Returns an S-shaped growth value that approaches L as x increases.", _
        Category:="Statistical", _
        ArgumentDescriptions:=Array( _
            "The input value, such as time or period.", _
            "The upper asymptote, which is the maximum value the curve approaches.", _
            "The displacement along the x-axis. The curve equals L*EXP(-k) at x = 0.", _
            "The growth rate. Higher values give a steeper transition.")
End Sub