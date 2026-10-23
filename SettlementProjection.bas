Attribute VB_Name = "SettlementProjection"
Option Explicit

'==============================================================================
' BuildSettlementProjection writes two stacked tables on SHEET_OUT:
'   TBL_OUT       - projection from TBL_IN (full uplift) cohorts
'   TBL_OUT_BASE  - same math on (TBL_IN - TBL_IN_INCREMENTAL) cohorts
'
' For each cohort row c contributing to settlement month m:
'   app_anchor_c = mid of c.Month (uniform-arrival mean; avoids day-1 bias
'                                  that overstates near-term flow)
'   marginal_c_m = G(end_of_m - app_anchor_c)
'                  - G(max(0, end_prev_m - app_anchor_c))
'   $_c_m_seg    = cohort_$ * marginal_c_m * conv_seg
' conv_seg is read from TBL_CONVERSION keyed by Pricing_Segment_ID.
'
' G(days) is the Gompertz.AppFlighting LAMBDA. It is sheet-scoped on
' SHEET_FLIGHT because a workbook-level `Gompertz*` name shadows a workbook-
' scoped version. VBA cannot invoke a LAMBDA via Application.Run, so each
' call goes through SHEET_FLIGHT.Evaluate(...) and is cached by day distance.
'
' FitConversionRates backs out conv_seg from tblRFC actuals and writes them
' into tblConversionMapping. Only rows at or after NAME_RFC_REF_MONTH count
' as settlement rows; earlier rows still contribute cohorts, weighted by the
' share landing in the window:
'   row_share_c = G(end_last - anchor_c) - G(max(0, end_pre_ref - anchor_c))
'   rate_g      = sum(settlements_g over ref window) /
'                 sum(apps_c_g * row_share_c over all rows)
'==============================================================================

Private Const SHEET_IN As String = "31_MDL_ApplicationProjection"
Private Const TBL_IN As String = "tblProjAppUplift"
Private Const TBL_IN_INCREMENTAL As String = "tblProjAppIncrementalUplift"

Private Const SHEET_OUT As String = "32_MDL_SettlementProjection"
Private Const TBL_OUT As String = "tblSettlementProjection"
Private Const TBL_OUT_BASE As String = "tblSettlementProjection_Current"
Private Const ANCHOR_OUT As String = "A101"
Private Const GAP_ROWS As Long = 5
Private Const STATUS_CELL As String = ""
Private Const TITLE_CELL As String = "A100"
Private Const TABLE_STYLE As String = "TableStyle Modelling"
Private Const RECALC_KEY As String = "SettlementProjection"

Private Const SHEET_FLIGHT As String = "27_CALC_App_Flighting"
Private Const GOMPERTZ_LAMBDA As String = "Gompertz.AppFlighting"

Private Const TBL_CONVERSION As String = "tblConversionMapping"
Private Const COL_CONV_SEGMENT As String = "Pricing_Segment_ID"
Private Const COL_CONV_RATE As String = "Conversion_Rate"
Private Const COL_CONV_RFC As String = "RFC_Segment_ID"

Private Const TBL_RFC As String = "tblRFC"
Private Const COL_RFC_MONTH As String = "Month"
Private Const SUFFIX_APPS As String = "_App$"
Private Const SUFFIX_SETT As String = "_NS"

Private Const NAME_RFC_REF_MONTH As String = "dateRFCRefMonth"

Private Const COL_MONTH As String = "Month"
Private Const COL_DAYS As String = "Days_Post_Eff_Date"
Private Const COL_IS_CURR As String = "IsCurrentFY"
Private Const COL_IS_NEXT As String = "IsNextFY"

Private mGompertzCache As Object

' Button entry: rebuilds both tables and reports the outcome in a message box.
Public Sub BuildSettlementProjection()
    Dim priorCalc As XlCalculation
    Dim nRows As Long, nSeg As Long
    Dim stampedOn As Date

    On Error GoTo Failed

    Application.ScreenUpdating = False
    priorCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Projecting settlements..."

    stampedOn = ProjectSettlements(nRows, nSeg)

    Application.StatusBar = False
    Application.Calculation = priorCalc
    Application.ScreenUpdating = True

    MsgBox TBL_OUT & " and " & TBL_OUT_BASE & " rebuilt at " & _
           Format$(stampedOn, "dd mmm yyyy hh:nn") & "." & vbNewLine & _
           nRows & " settlement months x " & nSeg & " segments.", vbInformation
    Exit Sub

Failed:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If priorCalc <> 0 Then Application.Calculation = priorCalc
    MsgBox "Settlement projection failed:" & vbNewLine & Err.description, vbCritical
End Sub

' Scenario-run entry, called from ScenarioConfig.PrepareScenarioOutputs. Silent, and it
' leaves calculation mode, screen updating and the status bar alone because the run's
' BeginBatch guard owns them. Raises on failure so the run reports it against the
' scenario instead of storing results from a stale projection.
Public Sub RefreshSettlementProjection()
    Dim nRows As Long, nSeg As Long

    ProjectSettlements nRows, nSeg
End Sub

' Rebuilds both tables from tblProjAppUplift as it currently stands, so the caller must
' have calculated the loaded inputs first. Returns the time stamped on the status, and
' the table shape via nRows and nSeg.
Private Function ProjectSettlements(ByRef nRows As Long, ByRef nSeg As Long) As Date
    Dim loIn As ListObject
    Dim data As Variant
    Dim nCols As Long
    Dim colMonth As Long, colDays As Long, colCur As Long, colNext As Long
    Dim segCols() As Long, segHeaders() As String
    Dim appAnchor() As Date
    Dim cohortFull() As Double, cohortBase() As Double
    Dim outFull As Variant, outBase As Variant
    Dim r As Long, m As Long, s As Long
    Dim monthStart As Date, endOfM As Date, endOfPrev As Date
    Dim dEnd As Double, dStart As Double, marginal As Double
    Dim segRates() As Double
    Dim stampedOn As Date

    On Error GoTo Failed

    Set loIn = RequireTable(TBL_IN)
    data = TableBody(loIn)
    If IsEmpty(data) Then
        Err.Raise vbObjectError + 701, "SettlementProjection", _
                  "'" & TBL_IN & "' has no rows to project."
    End If

    nRows = UBound(data, 1)
    nCols = UBound(data, 2)

    colMonth = RequireColumnIndex(loIn, COL_MONTH)
    colDays = RequireColumnIndex(loIn, COL_DAYS)
    colCur = RequireColumnIndex(loIn, COL_IS_CURR)
    colNext = RequireColumnIndex(loIn, COL_IS_NEXT)

    CollectSegmentColumns loIn, colMonth, colDays, colCur, colNext, _
                          segCols, segHeaders, nSeg
    If nSeg = 0 Then
        Err.Raise vbObjectError + 702, "SettlementProjection", _
                  "'" & TBL_IN & "' has no segment columns beyond the description columns."
    End If

    segRates = LoadSegmentConversionRates(segHeaders)

    BuildCohortMatrices data, colMonth, segCols, segHeaders, nRows, nSeg, _
                        cohortFull, cohortBase

    ReDim appAnchor(1 To nRows)
    For r = 1 To nRows
        appAnchor(r) = MidOfMonth(CDate(data(r, colMonth)))
    Next r

    ReDim outFull(1 To nRows, 1 To 4 + nSeg)
    ReDim outBase(1 To nRows, 1 To 4 + nSeg)

    Set mGompertzCache = CreateObject("Scripting.Dictionary")

    For m = 1 To nRows
        monthStart = FirstOfMonth(CDate(data(m, colMonth)))
        endOfM = EndOfMonth(monthStart)
        endOfPrev = DateAdd("d", -1, monthStart)

        outFull(m, 1) = data(m, colMonth): outBase(m, 1) = data(m, colMonth)
        outFull(m, 2) = data(m, colDays): outBase(m, 2) = data(m, colDays)
        outFull(m, 3) = data(m, colCur): outBase(m, 3) = data(m, colCur)
        outFull(m, 4) = data(m, colNext): outBase(m, 4) = data(m, colNext)
        For s = 1 To nSeg
            outFull(m, 4 + s) = 0#
            outBase(m, 4 + s) = 0#
        Next s

        For r = 1 To m
            dEnd = CDbl(endOfM) - CDbl(appAnchor(r))
            dStart = CDbl(endOfPrev) - CDbl(appAnchor(r))
            If dStart < 0 Then dStart = 0
            If dEnd < dStart Then dEnd = dStart

            marginal = GompertzCached(dEnd) - GompertzCached(dStart)
            If marginal <> 0 Then
                For s = 1 To nSeg
                    outFull(m, 4 + s) = outFull(m, 4 + s) + _
                                        cohortFull(r, s) * marginal * segRates(s)
                    outBase(m, 4 + s) = outBase(m, 4 + s) + _
                                        cohortBase(r, s) * marginal * segRates(s)
                Next s
            End If
        Next r
    Next m

    WriteBothOutputs outFull, outBase, nRows, nSeg, segHeaders

    stampedOn = Now
    LastCalculatedOn(RECALC_KEY) = stampedOn
    WriteStatus stampedOn

    Set mGompertzCache = Nothing
    ProjectSettlements = stampedOn
    Exit Function

Failed:
    Set mGompertzCache = Nothing
    Err.Raise Err.Number, Err.Source, Err.description
End Function

' Rows are matched between the two tables by Month normalised to the first of
' month; segments by column header name, so column order can differ. A row or
' segment missing from the incremental table is treated as zero incremental.
Private Sub BuildCohortMatrices(ByRef upliftData As Variant, _
                                ByVal upliftColMonth As Long, _
                                ByRef segCols() As Long, _
                                ByRef segHeaders() As String, _
                                ByVal nRows As Long, ByVal nSeg As Long, _
                                ByRef cohortFull() As Double, _
                                ByRef cohortBase() As Double)
    Dim loInc As ListObject
    Dim dataInc As Variant
    Dim incColMonth As Long
    Dim incColBySeg() As Long
    Dim incRowByMonth As Object
    Dim r As Long, s As Long, ri As Long
    Dim incRow As Long
    Dim monthKey As Long
    Dim incVal As Double

    ReDim cohortFull(1 To nRows, 1 To nSeg)
    ReDim cohortBase(1 To nRows, 1 To nSeg)

    For r = 1 To nRows
        For s = 1 To nSeg
            cohortFull(r, s) = NumericOrZero(upliftData(r, segCols(s)))
        Next s
    Next r

    Set loInc = RequireTable(TBL_IN_INCREMENTAL)
    dataInc = TableBody(loInc)
    If IsEmpty(dataInc) Then
        For r = 1 To nRows
            For s = 1 To nSeg
                cohortBase(r, s) = cohortFull(r, s)
            Next s
        Next r
        Exit Sub
    End If

    incColMonth = RequireColumnIndex(loInc, COL_MONTH)

    Set incRowByMonth = CreateObject("Scripting.Dictionary")
    For ri = 1 To UBound(dataInc, 1)
        monthKey = CLng(CDbl(FirstOfMonth(CDate(dataInc(ri, incColMonth)))))
        incRowByMonth(monthKey) = ri
    Next ri

    ReDim incColBySeg(1 To nSeg)
    For s = 1 To nSeg
        incColBySeg(s) = ColumnIndexOrZero(loInc, segHeaders(s))
    Next s

    For r = 1 To nRows
        monthKey = CLng(CDbl(FirstOfMonth(CDate(upliftData(r, upliftColMonth)))))
        If incRowByMonth.Exists(monthKey) Then
            incRow = CLng(incRowByMonth(monthKey))
        Else
            incRow = 0
        End If

        For s = 1 To nSeg
            incVal = 0#
            If incRow > 0 And incColBySeg(s) > 0 Then
                incVal = NumericOrZero(dataInc(incRow, incColBySeg(s)))
            End If
            cohortBase(r, s) = cohortFull(r, s) - incVal
        Next s
    Next r
End Sub

' Rows in TBL_CONVERSION whose RFC_Segment has no matching column pair in
' tblRFC are left unchanged, so a partial fit doesn't blank out other rates.
Public Sub FitConversionRates()
    Dim loRfc As ListObject
    Dim data As Variant
    Dim nRows As Long, r As Long, g As Long
    Dim colMonth As Long
    Dim rfcSegs() As String, appCols() As Long, settCols() As Long, nRfc As Long
    Dim appAnchor() As Date
    Dim endOfLast As Date, refMonthStart As Date, endPreRef As Date
    Dim totalSett() As Double, totalExpApps() As Double, rates() As Double
    Dim rowShare As Double, prePart As Double
    Dim rowMonthStart As Date
    Dim isSettlementRow As Boolean
    Dim priorCalc As XlCalculation
    Dim updated As Long, nSett As Long

    On Error GoTo Failed

    Application.ScreenUpdating = False
    priorCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Fitting conversion rates..."

    Set loRfc = RequireTable(TBL_RFC)
    data = TableBody(loRfc)
    If IsEmpty(data) Then
        Err.Raise vbObjectError + 710, "SettlementProjection", _
                  "'" & TBL_RFC & "' has no rows."
    End If

    nRows = UBound(data, 1)
    colMonth = RequireColumnIndex(loRfc, COL_RFC_MONTH)

    DiscoverRfcSegmentColumns loRfc, appCols, settCols, rfcSegs, nRfc
    If nRfc = 0 Then
        Err.Raise vbObjectError + 711, "SettlementProjection", _
                  "'" & TBL_RFC & "' has no " & SUFFIX_APPS & " / " & _
                  SUFFIX_SETT & " column pairs."
    End If

    refMonthStart = FirstOfMonth(CDate(RequireNamedRange(NAME_RFC_REF_MONTH).value))
    endPreRef = DateAdd("d", -1, refMonthStart)

    ReDim appAnchor(1 To nRows)
    For r = 1 To nRows
        appAnchor(r) = MidOfMonth(CDate(data(r, colMonth)))
    Next r
    endOfLast = EndOfMonth(CDate(data(nRows, colMonth)))

    If endOfLast < refMonthStart Then
        Err.Raise vbObjectError + 713, "SettlementProjection", _
                  "'" & NAME_RFC_REF_MONTH & "' (" & _
                  Format$(refMonthStart, "mmm yyyy") & _
                  ") is after the last month in '" & TBL_RFC & "' (" & _
                  Format$(endOfLast, "mmm yyyy") & _
                  "). No settlement rows to fit against."
    End If

    Set mGompertzCache = CreateObject("Scripting.Dictionary")

    ReDim totalSett(1 To nRfc)
    ReDim totalExpApps(1 To nRfc)
    nSett = 0
    For r = 1 To nRows
        rowMonthStart = FirstOfMonth(CDate(data(r, colMonth)))
        isSettlementRow = (rowMonthStart >= refMonthStart)

        prePart = CDbl(endPreRef) - CDbl(appAnchor(r))
        If prePart < 0 Then prePart = 0
        rowShare = GompertzCached(CDbl(endOfLast) - CDbl(appAnchor(r))) - _
                   GompertzCached(prePart)
        If rowShare < 0 Then rowShare = 0

        For g = 1 To nRfc
            If isSettlementRow Then
                totalSett(g) = totalSett(g) + NumericOrZero(data(r, settCols(g)))
            End If
            totalExpApps(g) = totalExpApps(g) + _
                              NumericOrZero(data(r, appCols(g))) * rowShare
        Next g
        If isSettlementRow Then nSett = nSett + 1
    Next r

    Set mGompertzCache = Nothing

    ReDim rates(1 To nRfc)
    For g = 1 To nRfc
        If totalExpApps(g) > 0 Then
            rates(g) = totalSett(g) / totalExpApps(g)
        End If
    Next g

    updated = WriteRatesToMapping(rfcSegs, rates)

    Application.StatusBar = False
    Application.Calculation = priorCalc
    Application.ScreenUpdating = True

    MsgBox "Fitted " & nRfc & " RFC segment rate(s) across " & nSett & _
           " settlement month(s) from " & _
           Format$(refMonthStart, "mmm yyyy") & "; " & updated & _
           " row(s) in '" & TBL_CONVERSION & "' updated.", vbInformation
    Exit Sub

Failed:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If priorCalc <> 0 Then Application.Calculation = priorCalc
    Set mGompertzCache = Nothing
    MsgBox "FitConversionRates failed:" & vbNewLine & Err.description, vbCritical
End Sub

' Prefixes with only one of the two suffix columns are skipped silently so
' partial segments don't fit from thin air.
Private Sub DiscoverRfcSegmentColumns(ByVal lo As ListObject, _
                                     ByRef appCols() As Long, _
                                     ByRef settCols() As Long, _
                                     ByRef rfcSegs() As String, _
                                     ByRef nRfc As Long)
    Dim nCols As Long, c As Long
    Dim name As String, suffix As Integer
    Dim prefix As String, pairCol As Long
    Dim appByPrefix As Object, settByPrefix As Object
    Dim keyVar As Variant

    Set appByPrefix = CreateObject("Scripting.Dictionary")
    appByPrefix.CompareMode = vbTextCompare
    Set settByPrefix = CreateObject("Scripting.Dictionary")
    settByPrefix.CompareMode = vbTextCompare

    nCols = lo.ListColumns.count
    For c = 1 To nCols
        name = lo.ListColumns(c).Name
        If EndsWith(name, SUFFIX_APPS) Then
            prefix = Left$(name, Len(name) - Len(SUFFIX_APPS))
            appByPrefix(prefix) = c
        ElseIf EndsWith(name, SUFFIX_SETT) Then
            prefix = Left$(name, Len(name) - Len(SUFFIX_SETT))
            settByPrefix(prefix) = c
        End If
    Next c

    ReDim rfcSegs(1 To appByPrefix.count + 1)
    ReDim appCols(1 To appByPrefix.count + 1)
    ReDim settCols(1 To appByPrefix.count + 1)
    nRfc = 0

    For Each keyVar In appByPrefix.keys
        If settByPrefix.Exists(keyVar) Then
            nRfc = nRfc + 1
            rfcSegs(nRfc) = CStr(keyVar)
            appCols(nRfc) = CLng(appByPrefix(keyVar))
            settCols(nRfc) = CLng(settByPrefix(keyVar))
        End If
    Next keyVar

    If nRfc = 0 Then Exit Sub
    ReDim Preserve rfcSegs(1 To nRfc)
    ReDim Preserve appCols(1 To nRfc)
    ReDim Preserve settCols(1 To nRfc)
End Sub

Private Function EndsWith(ByVal source As String, ByVal tail As String) As Boolean
    If Len(source) < Len(tail) Then Exit Function
    EndsWith = (StrComp(Right$(source, Len(tail)), tail, vbTextCompare) = 0)
End Function

' Rows whose RFC_Segment isn't in rfcSegs are left untouched so a partial fit
' doesn't blank out other rates. Returns the row count actually written.
Private Function WriteRatesToMapping(ByRef rfcSegs() As String, _
                                     ByRef rates() As Double) As Long
    Dim lo As ListObject
    Dim body As Variant
    Dim colRfc As Long, colRate As Long
    Dim i As Long, g As Long
    Dim rateByRfc As Object
    Dim key As String
    Dim written As Long

    Set lo = RequireTable(TBL_CONVERSION)
    body = TableBody(lo)
    If IsEmpty(body) Then
        Err.Raise vbObjectError + 712, "SettlementProjection", _
                  "'" & TBL_CONVERSION & "' has no rows."
    End If

    colRfc = RequireColumnIndex(lo, COL_CONV_RFC)
    colRate = RequireColumnIndex(lo, COL_CONV_RATE)

    Set rateByRfc = CreateObject("Scripting.Dictionary")
    rateByRfc.CompareMode = vbTextCompare
    For g = 1 To UBound(rfcSegs)
        rateByRfc(rfcSegs(g)) = rates(g)
    Next g

    For i = 1 To UBound(body, 1)
        key = CStr(body(i, colRfc))
        If rateByRfc.Exists(key) Then
            lo.DataBodyRange.Cells(i, colRate).value = rateByRfc(key)
            written = written + 1
        End If
    Next i

    WriteRatesToMapping = written
End Function

Private Sub CollectSegmentColumns(ByVal lo As ListObject, _
                                  ByVal colMonth As Long, ByVal colDays As Long, _
                                  ByVal colCur As Long, ByVal colNext As Long, _
                                  ByRef segCols() As Long, _
                                  ByRef segHeaders() As String, _
                                  ByRef nSeg As Long)
    Dim nCols As Long, c As Long

    nCols = lo.ListColumns.count
    ReDim segCols(1 To nCols)
    ReDim segHeaders(1 To nCols)
    nSeg = 0

    For c = 1 To nCols
        If c <> colMonth And c <> colDays And c <> colCur And c <> colNext Then
            nSeg = nSeg + 1
            segCols(nSeg) = c
            segHeaders(nSeg) = lo.ListColumns(c).Name
        End If
    Next c

    If nSeg = 0 Then Exit Sub
    ReDim Preserve segCols(1 To nSeg)
    ReDim Preserve segHeaders(1 To nSeg)
End Sub

Private Function NumericOrZero(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumericOrZero = CDbl(v)
End Function

' Raises on missing segment or non-numeric rate; a silent 0 would quietly
' zero out that segment's whole projection.
Private Function LoadSegmentConversionRates(ByRef segHeaders() As String) As Double()
    Dim lo As ListObject
    Dim body As Variant
    Dim colSeg As Long, colRate As Long
    Dim lookup As Object
    Dim r As Long, s As Long
    Dim key As String
    Dim rates() As Double

    Set lo = RequireTable(TBL_CONVERSION)
    body = TableBody(lo)
    If IsEmpty(body) Then
        Err.Raise vbObjectError + 704, "SettlementProjection", _
                  "'" & TBL_CONVERSION & "' has no rows."
    End If

    colSeg = RequireColumnIndex(lo, COL_CONV_SEGMENT)
    colRate = RequireColumnIndex(lo, COL_CONV_RATE)

    Set lookup = CreateObject("Scripting.Dictionary")
    lookup.CompareMode = vbTextCompare
    For r = 1 To UBound(body, 1)
        key = CStr(body(r, colSeg))
        If Len(key) = 0 Then
            Err.Raise vbObjectError + 706, "SettlementProjection", _
                      "'" & TBL_CONVERSION & "' row " & r & " has a blank " & _
                      COL_CONV_SEGMENT & "."
        End If
        If Not IsNumeric(body(r, colRate)) Then
            Err.Raise vbObjectError + 707, "SettlementProjection", _
                      "'" & TBL_CONVERSION & "' row for segment '" & key & _
                      "' has a non-numeric " & COL_CONV_RATE & "."
        End If
        lookup(key) = CDbl(body(r, colRate))
    Next r

    ReDim rates(1 To UBound(segHeaders))
    For s = 1 To UBound(segHeaders)
        If Not lookup.Exists(segHeaders(s)) Then
            Err.Raise vbObjectError + 708, "SettlementProjection", _
                      "Segment '" & segHeaders(s) & "' from '" & TBL_IN & _
                      "' has no conversion rate in '" & TBL_CONVERSION & "'."
        End If
        rates(s) = lookup(segHeaders(s))
    Next s

    LoadSegmentConversionRates = rates
End Function

Private Function GompertzCached(ByVal days As Double) As Double
    Dim key As String

    key = CStr(days)
    If Not mGompertzCache.Exists(key) Then
        mGompertzCache(key) = EvalGompertzLambda(days)
    End If
    GompertzCached = mGompertzCache(key)
End Function

' Evaluate must run on SHEET_FLIGHT so the sheet-scoped LAMBDA resolves ahead
' of the workbook-level `Gompertz*` name that would otherwise shadow it. Str$
' forces a period decimal separator independent of locale.
Private Function EvalGompertzLambda(ByVal days As Double) As Double
    Dim expr As String
    Dim host As Worksheet
    Dim result As Variant

    expr = GOMPERTZ_LAMBDA & "(" & Trim$(Str$(days)) & ")"
    Set host = SheetOrNothing(SHEET_FLIGHT)
    If host Is Nothing Then
        Err.Raise vbObjectError + 705, "SettlementProjection", _
                  "Sheet '" & SHEET_FLIGHT & "' was not found. The '" & _
                  GOMPERTZ_LAMBDA & "' LAMBDA is scoped to that sheet, so " & _
                  "Evaluate must run in its context."
    End If

    result = host.Evaluate(expr)
    If IsError(result) Then
        Err.Raise vbObjectError + 703, "SettlementProjection", _
                  "LAMBDA '" & GOMPERTZ_LAMBDA & _
                  "' returned an error for days=" & days & _
                  ". Check it is defined at sheet scope on '" & _
                  SHEET_FLIGHT & "'."
    End If

    EvalGompertzLambda = CDbl(result)
End Function

Private Function FirstOfMonth(ByVal d As Date) As Date
    FirstOfMonth = DateSerial(Year(d), Month(d), 1)
End Function

Private Function MidOfMonth(ByVal d As Date) As Date
    MidOfMonth = DateSerial(Year(d), Month(d), 15)
End Function

Private Function EndOfMonth(ByVal d As Date) As Date
    EndOfMonth = DateAdd("d", -1, DateAdd("m", 1, FirstOfMonth(d)))
End Function

' When both output tables already exist we resize and overwrite in place so
' external formulas (and structured references into these tables) keep
' resolving. A delete+recreate would turn every dependent cell into #REF!.
' If either table is missing we fall back to the original full rebuild and
' accept the broken refs — that path is only hit on first run or after a
' deliberate manual delete.
Private Sub WriteBothOutputs(ByVal outFull As Variant, ByVal outBase As Variant, _
                             ByVal nRows As Long, ByVal nSeg As Long, _
                             ByRef segHeaders() As String)
    Dim ws As Worksheet
    Dim primaryAnchor As Range, baseAnchor As Range
    Dim primaryTitle As Range, baseTitle As Range
    Dim primaryLo As ListObject, baseLo As ListObject
    Dim rowShift As Long
    Dim inPlace As Boolean

    Set ws = EnsureOutputSheet()
    Set primaryLo = ListObjectOrNothing(ws, TBL_OUT)
    Set baseLo = ListObjectOrNothing(ws, TBL_OUT_BASE)
    inPlace = (Not primaryLo Is Nothing) And (Not baseLo Is Nothing)

    If Not inPlace Then
        If Not primaryLo Is Nothing Then primaryLo.Delete
        If Not baseLo Is Nothing Then baseLo.Delete
        ClearOutputRegion ws
        Set primaryLo = Nothing
        Set baseLo = Nothing
    End If

    Set primaryAnchor = ws.Range(ANCHOR_OUT)
    If inPlace Then
        ' Keep the base table where it is so its external refs stay valid even
        ' if the primary row count changed.
        Set baseAnchor = baseLo.Range.Cells(1, 1)
        If primaryAnchor.Row + nRows >= baseAnchor.Row Then
            Err.Raise vbObjectError + 720, "SettlementProjection", _
                      "'" & TBL_OUT & "' resized to " & nRows & _
                      " rows would overlap '" & TBL_OUT_BASE & "' at " & _
                      baseAnchor.Address & ". Move '" & TBL_OUT_BASE & _
                      "' further down, or delete both tables to rebuild " & _
                      "from scratch (dependent formulas will become #REF!)."
        End If
    Else
        rowShift = 1 + nRows + GAP_ROWS   ' +1 for header row
        Set baseAnchor = primaryAnchor.Offset(rowShift, 0)
    End If

    If Len(TITLE_CELL) > 0 Then
        Set primaryTitle = ws.Range(TITLE_CELL)
        Set baseTitle = baseAnchor.Offset(-1, 0)
    End If

    WriteOutputAt ws, outFull, nRows, nSeg, segHeaders, _
                  TBL_OUT, primaryAnchor, primaryTitle, primaryLo
    WriteOutputAt ws, outBase, nRows, nSeg, segHeaders, _
                  TBL_OUT_BASE, baseAnchor, baseTitle, baseLo
End Sub

Private Sub WriteOutputAt(ByVal ws As Worksheet, ByVal out As Variant, _
                          ByVal nRows As Long, ByVal nSeg As Long, _
                          ByRef segHeaders() As String, _
                          ByVal tblName As String, _
                          ByVal anchor As Range, ByVal titleCell As Range, _
                          ByVal existingLo As ListObject)
    Dim lo As ListObject
    Dim headers() As Variant
    Dim i As Long

    ReDim headers(1 To 1, 1 To 4 + nSeg)
    headers(1, 1) = COL_MONTH
    headers(1, 2) = COL_DAYS
    headers(1, 3) = COL_IS_CURR
    headers(1, 4) = COL_IS_NEXT
    For i = 1 To nSeg
        headers(1, 4 + i) = segHeaders(i)
    Next i

    If existingLo Is Nothing Then
        anchor.Resize(1, 4 + nSeg).value = headers
        anchor.Offset(1, 0).Resize(nRows, 4 + nSeg).value = out
        Set lo = ws.ListObjects.Add(xlSrcRange, _
                                    anchor.Resize(nRows + 1, 4 + nSeg), , xlYes)
        lo.Name = tblName
    Else
        Set lo = existingLo
        ResizeAndOverwrite lo, anchor, headers, out, nRows, nSeg
    End If

    ApplyTableStyle lo
    FormatOutputTable lo

    If Not titleCell Is Nothing Then WriteTitleAt titleCell, anchor
End Sub

' Resize the ListObject to the new footprint and overwrite header + body
' values. Cells that were part of the old table but fall outside the new
' footprint are cleared so stale values don't bleed through. The ListObject
' itself is kept, so structured references (tblSettlementProjection[...])
' and A1 refs into cells still inside the new footprint keep resolving.
Private Sub ResizeAndOverwrite(ByVal lo As ListObject, ByVal anchor As Range, _
                               ByRef headers As Variant, ByVal out As Variant, _
                               ByVal nRows As Long, ByVal nSeg As Long)
    Dim oldRange As Range
    Dim oldRows As Long, oldCols As Long
    Dim newRows As Long, newCols As Long

    Set oldRange = lo.Range
    oldRows = oldRange.Rows.count
    oldCols = oldRange.Columns.count
    newRows = 1 + nRows
    newCols = 4 + nSeg

    If oldRange.Cells(1, 1).Address <> anchor.Address Then
        Err.Raise vbObjectError + 721, "SettlementProjection", _
                  "'" & lo.Name & "' is at " & oldRange.Cells(1, 1).Address & _
                  " but was expected at " & anchor.Address & _
                  ". Delete it and rerun to re-anchor."
    End If

    If oldRows > newRows Then
        anchor.Offset(newRows, 0).Resize(oldRows - newRows, oldCols).Clear
    End If
    If oldCols > newCols Then
        anchor.Offset(0, newCols).Resize(oldRows, oldCols - newCols).Clear
    End If

    lo.Resize anchor.Resize(newRows, newCols)

    lo.HeaderRowRange.value = headers
    If nRows > 0 And Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.value = out
    End If
End Sub

Private Function ListObjectOrNothing(ByVal ws As Worksheet, _
                                     ByVal tblName As String) As ListObject
    On Error Resume Next
    Set ListObjectOrNothing = ws.ListObjects(tblName)
    On Error GoTo 0
End Function

' =tableName(<anchor>) is the UDF in zz_QualityOfLIfe that resolves a cell
' reference to its enclosing ListObject's name. Only .Formula is written so
' the cell keeps its existing formatting.
Private Sub WriteTitleAt(ByVal titleCell As Range, ByVal anchor As Range)
    Dim anchorRef As String

    anchorRef = "'" & anchor.Worksheet.Name & "'!" & anchor.Address(True, True)
    titleCell.Formula = "=tableName(" & anchorRef & ")"
End Sub

Private Function EnsureOutputSheet() As Worksheet
    Dim ws As Worksheet

    Set ws = SheetOrNothing(SHEET_OUT)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = SHEET_OUT
        If Len(STATUS_CELL) > 0 Then
            ws.Range(STATUS_CELL).value = "Never recalculated"
        End If
    End If
    Set EnsureOutputSheet = ws
End Function

' Clears everything from ANCHOR_OUT to the end of the used range so stale
' rows or columns from a longer prior run do not bleed through.
Private Sub ClearOutputRegion(ws As Worksheet)
    Dim anchor As Range
    Dim lastRow As Long, lastCol As Long

    Set anchor = ws.Range(ANCHOR_OUT)
    With ws.UsedRange
        lastRow = .Row + .Rows.count - 1
        lastCol = .column + .Columns.count - 1
    End With

    If lastRow >= anchor.Row And lastCol >= anchor.column Then
        ws.Range(anchor, ws.Cells(lastRow, lastCol)).Clear
    End If
End Sub

Private Sub ApplyTableStyle(lo As ListObject)
    On Error Resume Next
    lo.TableStyle = TABLE_STYLE
    On Error GoTo 0
End Sub

Private Sub FormatOutputTable(lo As ListObject)
    Dim c As Long

    If lo.DataBodyRange Is Nothing Then Exit Sub

    lo.ListColumns(COL_MONTH).DataBodyRange.NumberFormat = "yyyy-mm-dd"
    lo.ListColumns(COL_DAYS).DataBodyRange.NumberFormat = "0"
    lo.ListColumns(COL_IS_CURR).DataBodyRange.HorizontalAlignment = xlCenter
    lo.ListColumns(COL_IS_NEXT).DataBodyRange.HorizontalAlignment = xlCenter

    For c = 5 To lo.ListColumns.count
        lo.ListColumns(c).DataBodyRange.NumberFormat = "#,##0;(#,##0);-"
    Next c
End Sub

Private Sub WriteStatus(ByVal stampedOn As Date)
    Dim ws As Worksheet

    If Len(STATUS_CELL) = 0 Then Exit Sub

    Set ws = SheetOrNothing(SHEET_OUT)
    If ws Is Nothing Then Exit Sub
    ws.Range(STATUS_CELL).value = _
        "Last projected " & Format$(stampedOn, "dd mmm yyyy hh:nn")
End Sub

Public Sub RefreshSettlementProjectionStatus()
    Dim ws As Worksheet

    If Len(STATUS_CELL) = 0 Then Exit Sub

    Set ws = SheetOrNothing(SHEET_OUT)
    If ws Is Nothing Then Exit Sub
    ws.Range(STATUS_CELL).value = LastCalculatedOnText(RECALC_KEY, _
                                                      ifNever:="Never recalculated")
End Sub
