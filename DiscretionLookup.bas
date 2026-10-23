Attribute VB_Name = "DiscretionLookup"
Option Explicit

' Writes Discretion_Current and Discretion_Scenario into tblDiscretionMapping as static
' values, replacing the per-row FILTER formula with a single VBA pass.
'
' Each row's (Product, Purpose, Repayment_Type) picks a bucket; the first band in that
' bucket whose LVR and Loan_Size windows contain the application's LVR and Amount wins.
' Bands are inclusive on both ends. Rows with no matching band get 0, matching the
' if_empty fallback on the original FILTER formula.
'
' Discretion_Current reads from tblCurrentDiscretions and is refreshed when imports run
' (CopyApplicationsToDiscretionMapping calls RefreshDiscretionCurrent) or on demand.
' Discretion_Scenario reads from tblActiveDiscretions and is meant to be called from the
' scenario-run entry point so it reflects whichever scenario is loaded.

Private Const MAPPING_TABLE As String = "tblDiscretionMapping"

Private Const SRC_CURRENT As String = "tblCurrentDiscretions"
Private Const SRC_SCENARIO As String = "tblActiveDiscretions"

Private Const DEST_COL_CURRENT As String = "Discretion_Current"
Private Const DEST_COL_SCENARIO As String = "Discretion_Scenario"

Private Const MAP_PRODUCT As String = "Product"
Private Const MAP_PURPOSE As String = "Purpose"
Private Const MAP_REPAY As String = "Repayment_Type"
Private Const MAP_LVR As String = "Application_LVR"
Private Const MAP_AMOUNT As String = "Application_Amount"

Private Const SRC_PRODUCT As String = "Product"
Private Const SRC_PURPOSE As String = "Purpose"
Private Const SRC_REPAY As String = "Repayment_Type"
Private Const SRC_LVR_MIN As String = "LVR_Min"
Private Const SRC_LVR_MAX As String = "LVR_Max"
Private Const SRC_LOAN_MIN As String = "Loan_Size_Min"
Private Const SRC_LOAN_MAX As String = "Loan_Size_Max"
Private Const SRC_DISCRETION As String = "Discretion"

' Returned when no band matches. Matches the ",0" if_empty on the original FILTER.
Private Const NO_MATCH As Double = 0

Public Sub RefreshDiscretionCurrent()
    WriteDiscretionLookup SRC_CURRENT, DEST_COL_CURRENT
End Sub

Public Sub RefreshDiscretionScenario()
    WriteDiscretionLookup SRC_SCENARIO, DEST_COL_SCENARIO
End Sub

Private Sub WriteDiscretionLookup(ByVal srcTableName As String, ByVal destColName As String)
    Dim loMap As ListObject, loSrc As ListObject
    Dim buckets As Object
    Dim mapBody As Variant
    Dim results() As Variant
    Dim productCol As Long, purposeCol As Long, repayCol As Long
    Dim lvrCol As Long, amountCol As Long, destCol As Long
    Dim r As Long, n As Long
    Dim prevScreen As Boolean
    Dim prevCalc As XlCalculation

    Set loMap = RequireTable(MAPPING_TABLE)
    Set loSrc = RequireTable(srcTableName)

    n = TableRowCount(loMap)
    If n = 0 Then Exit Sub

    productCol = RequireColumnIndex(loMap, MAP_PRODUCT)
    purposeCol = RequireColumnIndex(loMap, MAP_PURPOSE)
    repayCol = RequireColumnIndex(loMap, MAP_REPAY)
    lvrCol = RequireColumnIndex(loMap, MAP_LVR)
    amountCol = RequireColumnIndex(loMap, MAP_AMOUNT)
    destCol = RequireColumnIndex(loMap, destColName)

    Set buckets = LoadBuckets(loSrc)
    mapBody = TableBody(loMap)

    ReDim results(1 To n, 1 To 1)
    For r = 1 To n
        results(r, 1) = MatchDiscretion(buckets, _
            CStr(mapBody(r, productCol)), _
            CStr(mapBody(r, purposeCol)), _
            CStr(mapBody(r, repayCol)), _
            mapBody(r, lvrCol), _
            mapBody(r, amountCol))
    Next r

    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    loMap.ListColumns(destCol).DataBodyRange.value = results

    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
End Sub

' One Collection of band tuples per Product|Purpose|Repayment_Type, so each application
' row only scans the handful of bands that could possibly match it.
Private Function LoadBuckets(loSrc As ListObject) As Object
    Dim buckets As Object
    Dim body As Variant
    Dim cProduct As Long, cPurpose As Long, cRepay As Long
    Dim cLvrMin As Long, cLvrMax As Long, cLoanMin As Long, cLoanMax As Long, cDisc As Long
    Dim r As Long, n As Long
    Dim key As String
    Dim band As Variant

    Set buckets = CreateObject("Scripting.Dictionary")
    buckets.CompareMode = 1
    Set LoadBuckets = buckets

    n = TableRowCount(loSrc)
    If n = 0 Then Exit Function

    body = TableBody(loSrc)
    cProduct = RequireColumnIndex(loSrc, SRC_PRODUCT)
    cPurpose = RequireColumnIndex(loSrc, SRC_PURPOSE)
    cRepay = RequireColumnIndex(loSrc, SRC_REPAY)
    cLvrMin = RequireColumnIndex(loSrc, SRC_LVR_MIN)
    cLvrMax = RequireColumnIndex(loSrc, SRC_LVR_MAX)
    cLoanMin = RequireColumnIndex(loSrc, SRC_LOAN_MIN)
    cLoanMax = RequireColumnIndex(loSrc, SRC_LOAN_MAX)
    cDisc = RequireColumnIndex(loSrc, SRC_DISCRETION)

    For r = 1 To n
        key = GroupKey(CStr(body(r, cProduct)), CStr(body(r, cPurpose)), CStr(body(r, cRepay)))
        If Not buckets.Exists(key) Then buckets.Add key, New Collection
        band = Array( _
            SafeDouble(body(r, cLvrMin)), _
            SafeDouble(body(r, cLvrMax)), _
            SafeDouble(body(r, cLoanMin)), _
            SafeDouble(body(r, cLoanMax)), _
            body(r, cDisc))
        buckets(key).Add band
    Next r
End Function

Private Function MatchDiscretion(ByVal buckets As Object, _
                                 ByVal product As String, _
                                 ByVal purpose As String, _
                                 ByVal repay As String, _
                                 ByVal lvr As Variant, _
                                 ByVal amount As Variant) As Variant
    Dim key As String
    Dim bands As Collection
    Dim band As Variant
    Dim lvrNum As Double, amountNum As Double

    key = GroupKey(product, purpose, repay)
    If Not buckets.Exists(key) Then
        MatchDiscretion = NO_MATCH
        Exit Function
    End If
    If Not (IsNumeric(lvr) And IsNumeric(amount)) Then
        MatchDiscretion = NO_MATCH
        Exit Function
    End If

    lvrNum = CDbl(lvr)
    amountNum = CDbl(amount)
    Set bands = buckets(key)

    For Each band In bands
        If band(0) <= lvrNum And band(1) >= lvrNum _
           And band(2) <= amountNum And band(3) >= amountNum Then
            MatchDiscretion = band(4)
            Exit Function
        End If
    Next band

    MatchDiscretion = NO_MATCH
End Function

Private Function GroupKey(ByVal product As String, ByVal purpose As String, ByVal repay As String) As String
    GroupKey = UCase$(Trim$(product)) & "|" & UCase$(Trim$(purpose)) & "|" & UCase$(Trim$(repay))
End Function

Private Function SafeDouble(ByVal v As Variant) As Double
    If IsNumeric(v) Then SafeDouble = CDbl(v)
End Function
