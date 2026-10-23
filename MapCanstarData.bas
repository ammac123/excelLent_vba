Attribute VB_Name = "MapCanstarData"
Option Explicit

' Rebuilds tblCompetitorMap as an institution-by-pricing-segment matrix of Canstar rates.
'
' Columns  = Pricing_Segment_ID values from tblProductMap where Included_Flag = "Yes"
'            (unique, in order of first appearance).
' Rows     = unique Institution values from tblCanstar, sorted alphabetically.
' Cells    = the lowest "Calculated Interest Rate" from tblCanstar rows whose
'            ING Product Type / ING Product Purpose / ING Repayment Type and active
'            LVR_FLAG_nn match a tblProductMap row for that segment. Blank if no match.
'
' Everything except the first column header ("Institution") is cleared and rebuilt each run.

Private Const PRODUCT_MAP_TABLE As String = "tblProductMap"
Private Const CANSTAR_TABLE As String = "tblCanstar"
Private Const TARGET_SHEET As String = "20_CALC_CompetitorRateMap"
Private Const TARGET_TABLE As String = "tblCompetitorMap"
Private Const TARGET_TABLE_2 As String = "tblCompetitorMapING"
Private Const OWN_BRAND_NAME As String = "ING"
Private Const FIRST_COLUMN_HEADER As String = "Institution"

Private Const RATE_FORMAT As String = "0.00%"
Private Const TEXT_FORMAT As String = "@"
Private Const LVR_BAND_COUNT As Long = 5

Public Sub BuildCompetitorRateMap()
    Dim mapLo As ListObject, canstarLo As ListObject, targetLo As ListObject
    Dim segments As Variant
    Dim criteria As Object
    Dim institutions As Variant
    Dim matrix As Variant
    Dim filled As Long

    On Error GoTo Failed

    Set mapLo = FindTable(PRODUCT_MAP_TABLE)
    Set canstarLo = FindTable(CANSTAR_TABLE)
    Set targetLo = TargetTable()

    Set criteria = CreateObject("Scripting.Dictionary")
    criteria.CompareMode = 1
    segments = IncludedSegments(mapLo, criteria)
    If IsEmpty(segments) Then
        MsgBox "No rows in " & PRODUCT_MAP_TABLE & " have Included_Flag = Yes.", vbExclamation
        Exit Sub
    End If

    institutions = UniqueInstitutions(canstarLo)
    If IsEmpty(institutions) Then
        MsgBox "No institutions were found in " & CANSTAR_TABLE & ".", vbExclamation
        Exit Sub
    End If

    matrix = BuildMatrix(canstarLo, criteria, segments, institutions, filled)

    Application.ScreenUpdating = False
    WriteMatrix targetLo, segments, institutions, matrix
    Application.ScreenUpdating = True

    MsgBox "Rebuilt " & TARGET_TABLE & " with " & (UBound(institutions) + 1) & " institution(s) x " & _
           (UBound(segments) + 1) & " segment(s)." & vbNewLine & _
           filled & " rate(s) matched.", vbInformation
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Rebuild failed at error " & Err.Number & ":" & vbNewLine & Err.description, vbCritical
End Sub

' Returns the included Pricing_Segment_ID values in order of first appearance and fills
' criteria with "product|purpose|repaymentType|lvrId" -> segment index.
Private Function IncludedSegments(mapLo As ListObject, ByVal criteria As Object) As Variant
    Dim body As Variant
    Dim cSeg As Long, cFlag As Long, cProduct As Long, cPurpose As Long, cRepay As Long, cLvr As Long
    Dim seen As Object
    Dim ordered As Object
    Dim r As Long
    Dim segment As String, key As String

    If mapLo.DataBodyRange Is Nothing Then Exit Function
    body = mapLo.DataBodyRange.value

    cSeg = ColIndex(mapLo, "Pricing_Segment_ID")
    cFlag = ColIndex(mapLo, "Included_Flag")
    cProduct = ColIndex(mapLo, "Product")
    cPurpose = ColIndex(mapLo, "Purpose")
    cRepay = ColIndex(mapLo, "Repayment_Type")
    cLvr = ColIndex(mapLo, "LVR_ID")

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    Set ordered = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(body, 1)
        If StrComp(Trim$(CStr(body(r, cFlag))), "Yes", vbTextCompare) = 0 Then
            segment = Trim$(CStr(body(r, cSeg)))
            If Len(segment) > 0 Then
                If Not seen.Exists(segment) Then
                    seen(segment) = ordered.count
                    ordered(ordered.count) = segment
                End If

                key = CriteriaKey(body(r, cProduct), body(r, cPurpose), body(r, cRepay), body(r, cLvr))
                If Not criteria.Exists(key) Then criteria(key) = seen(segment)
            End If
        End If
    Next r

    If ordered.count = 0 Then Exit Function
    IncludedSegments = ordered.Items
End Function

' Returns the distinct Institution values from tblCanstar, sorted alphabetically.
Private Function UniqueInstitutions(canstarLo As ListObject) As Variant
    Dim body As Variant
    Dim cInst As Long
    Dim seen As Object
    Dim r As Long
    Dim Name As String
    Dim result As Variant

    If canstarLo.DataBodyRange Is Nothing Then Exit Function
    body = canstarLo.DataBodyRange.value
    cInst = ColIndex(canstarLo, "Institution")

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    For r = 1 To UBound(body, 1)
        Name = Trim$(CStr(body(r, cInst)))
        If Len(Name) > 0 Then
            If Not seen.Exists(Name) Then seen(Name) = Name
        End If
    Next r

    If seen.count = 0 Then Exit Function

    result = seen.Items
    SortStrings result
    UniqueInstitutions = result
End Function

' Walks tblCanstar once and keeps the lowest matching rate per institution and segment.
Private Function BuildMatrix(canstarLo As ListObject, ByVal criteria As Object, _
                             segments As Variant, institutions As Variant, _
                             ByRef filled As Long) As Variant
    Dim body As Variant
    Dim cInst As Long, cProduct As Long, cPurpose As Long, cRepay As Long, cRate As Long
    Dim flagCols() As Long
    Dim rowIndex As Object
    Dim result As Variant
    Dim r As Long, b As Long, i As Long
    Dim rowNo As Long, colNo As Long
    Dim key As String
    Dim rate As Double

    body = canstarLo.DataBodyRange.value

    cInst = ColIndex(canstarLo, "Institution")
    cProduct = ColIndex(canstarLo, "ING Product Type")
    cPurpose = ColIndex(canstarLo, "ING Product Purpose")
    cRepay = ColIndex(canstarLo, "ING Repayment Type")
    cRate = ColIndex(canstarLo, "Calculated Interest Rate")

    ReDim flagCols(1 To LVR_BAND_COUNT)
    For b = 1 To LVR_BAND_COUNT
        flagCols(b) = ColIndex(canstarLo, "LVR_FLAG_" & Format$(b, "00"))
    Next b

    Set rowIndex = CreateObject("Scripting.Dictionary")
    rowIndex.CompareMode = 1
    For i = LBound(institutions) To UBound(institutions)
        rowIndex(CStr(institutions(i))) = i - LBound(institutions) + 1
    Next i

    ReDim result(1 To UBound(institutions) - LBound(institutions) + 1, _
                 1 To UBound(segments) - LBound(segments) + 1)

    filled = 0

    For r = 1 To UBound(body, 1)
        If Not IsNumeric(body(r, cRate)) Then GoTo nextRow
        If Not rowIndex.Exists(Trim$(CStr(body(r, cInst)))) Then GoTo nextRow

        rowNo = rowIndex(Trim$(CStr(body(r, cInst))))
        rate = CDbl(body(r, cRate))

        For b = 1 To LVR_BAND_COUNT
            If IsFlagTrue(body(r, flagCols(b))) Then
                key = CriteriaKey(body(r, cProduct), body(r, cPurpose), body(r, cRepay), _
                                  "LVR" & Format$(b, "00"))
                If criteria.Exists(key) Then
                    colNo = CLng(criteria(key)) + 1
                    If IsEmpty(result(rowNo, colNo)) Then
                        result(rowNo, colNo) = rate
                        filled = filled + 1
                    ElseIf rate < CDbl(result(rowNo, colNo)) Then
                        result(rowNo, colNo) = rate
                    End If
                End If
            End If
        Next b
nextRow:
    Next r

    BuildMatrix = result
End Function

' Clears the table back to its first header, then resizes and refills it.
Private Sub WriteMatrix(targetLo As ListObject, segments As Variant, _
                        institutions As Variant, matrix As Variant)
    Dim topLeft As Range
    Dim oldCols As Long
    Dim rowCount As Long, colCount As Long
    Dim names As Variant
    Dim i As Long

    rowCount = UBound(institutions) - LBound(institutions) + 1
    colCount = UBound(segments) - LBound(segments) + 2

    oldCols = targetLo.ListColumns.count
    Set topLeft = targetLo.Range.Cells(1, 1)

    If Not targetLo.DataBodyRange Is Nothing Then targetLo.DataBodyRange.Delete

    ' shrink to one column first so stale headers to the right can be wiped
    targetLo.Resize topLeft.Resize(2, 1)
    topLeft.Offset(0, 1).Resize(2, WorksheetFunction.Max(oldCols, colCount)).Clear

    targetLo.Resize topLeft.Resize(rowCount + 1, colCount)

    topLeft.value = FIRST_COLUMN_HEADER
    For i = LBound(segments) To UBound(segments)
        topLeft.Offset(0, i - LBound(segments) + 1).value = CStr(segments(i))
    Next i

    ReDim names(1 To rowCount, 1 To 1)
    For i = LBound(institutions) To UBound(institutions)
        names(i - LBound(institutions) + 1, 1) = CStr(institutions(i))
    Next i

    With targetLo.DataBodyRange
        .Columns(1).NumberFormat = TEXT_FORMAT
        .Columns(1).value = names
        .Offset(0, 1).Resize(rowCount, colCount - 1).NumberFormat = RATE_FORMAT
        .Offset(0, 1).Resize(rowCount, colCount - 1).value = matrix
    End With
End Sub

Private Function CriteriaKey(ByVal product As Variant, ByVal purpose As Variant, _
                             ByVal repayment As Variant, ByVal lvrId As Variant) As String
    CriteriaKey = UCase$(Trim$(CStr(product))) & "|" & _
                  UCase$(Trim$(CStr(purpose))) & "|" & _
                  UCase$(Trim$(CStr(repayment))) & "|" & _
                  UCase$(Trim$(CStr(lvrId)))
End Function

' Accepts the TRUE boolean, the text "TRUE"/"Y"/"YES", and the number 1.
Private Function IsFlagTrue(ByVal value As Variant) As Boolean
    Dim text As String

    If IsError(value) Then Exit Function
    If VarType(value) = vbBoolean Then
        IsFlagTrue = CBool(value)
        Exit Function
    End If

    text = UCase$(Trim$(CStr(value)))
    IsFlagTrue = (text = "TRUE" Or text = "YES" Or text = "Y" Or text = "1")
End Function

Private Function ColIndex(lo As ListObject, header As String) As Long
    Dim lc As ListColumn

    On Error Resume Next
    Set lc = lo.ListColumns(header)
    On Error GoTo 0

    If lc Is Nothing Then
        Err.Raise vbObjectError + 513, "CompetitorRateMap", _
                  "Column '" & header & "' was not found in table '" & lo.Name & "'."
    End If

    ColIndex = lc.Index
End Function

Private Function FindTable(tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws

    Err.Raise vbObjectError + 514, "CompetitorRateMap", _
              "Table '" & tableName & "' was not found in this workbook."
End Function

Private Function TargetTable() As ListObject
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Err.Raise vbObjectError + 515, "CompetitorRateMap", _
                  "Sheet '" & TARGET_SHEET & "' was not found in this workbook."
    End If

    On Error Resume Next
    Set TargetTable = ws.ListObjects(TARGET_TABLE)
    On Error GoTo 0

    If TargetTable Is Nothing Then
        Err.Raise vbObjectError + 516, "CompetitorRateMap", _
                  "Table '" & TARGET_TABLE & "' was not found on sheet '" & TARGET_SHEET & "'."
    End If
End Function

Private Sub SortStrings(ByRef values As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant

    For i = LBound(values) + 1 To UBound(values)
        temp = values(i)
        j = i - 1
        Do While j >= LBound(values)
            If StrComp(CStr(values(j)), CStr(temp), vbTextCompare) <= 0 Then Exit Do
            values(j + 1) = values(j)
            j = j - 1
        Loop
        values(j + 1) = temp
    Next i
End Sub


