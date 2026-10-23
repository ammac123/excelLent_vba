Attribute VB_Name = "DiscretionSummary"
Option Explicit

'==============================================================================
' Builds summary tables of a discretions table on a configured output sheet.
' One table per Purpose x Repayment_Type; P&I and IO are merged into a single
' "P&I/IO" table when their LVR / Loan_Size / Discretion grids are identical.
'
' Call GenerateDiscretionSummaryFor with a DiscretionsSpec, or the no-arg
' GenerateDiscretionSummary which uses the Active spec.
'==============================================================================

Private Const TABLE_STYLE As String = "TableStyle SRC"
Private Const NAME_MAX_LVR As String = "validationMaxLVRValue"
Private Const NAME_MAX_LOAN As String = "validationMaxLoanSize"
Private Const TITLE_SEP As String = " - "      ' e.g. "Owner Occupied - P&I"
Private Const GAP_ROWS As Long = 2             ' blank rows between tables
Private Const BOUND_TOL As Double = 0.005      ' tolerance for 0 / max boundary checks
Private Const EQ_TOL As Double = 0.000001      ' tolerance for value equality

' Module-level source data and column indices
Private mData As Variant
Private cPurpose As Long, cRepay As Long
Private cLvrMin As Long, cLvrMax As Long
Private cLsMin As Long, cLsMax As Long, cDisc As Long
Private mMaxLVR As Double, mMaxLoan As Double
Private mStartCol As Long

' Set per run from the active spec. Private helpers read these instead of module
' constants so one module can serve multiple discretions tables.
Private mSrcTable As String
Private mOutSheet As String
Private mOutStartCell As String
Private mTablePrefix As String

'------------------------------------------------------------------------------
' Entry points
'------------------------------------------------------------------------------
' Legacy no-arg entry: delegates to the Active spec. ScenarioConfig.PrepareScenarioOutputs
' calls this so scenario runs keep behaving as they did.
Public Sub GenerateDiscretionSummary()
    GenerateDiscretionSummaryFor ActiveDiscretionsSpec()
End Sub

' Generates the summary for whichever discretions table the spec names, into the
' spec's summarySheet at summaryAnchor. Output tables are named summaryTablePrefix &
' <ordinal>, so distinct specs can coexist in the same workbook.
Public Sub GenerateDiscretionSummaryFor(ByRef spec As DiscretionsSpec)
    Dim loSrc As ListObject, ws As Worksheet
    Dim nextRow As Long, tblNum As Long
    Dim purposes As Variant, p As Variant
    Dim rowsPI As Collection, rowsIO As Collection

    On Error GoTo ErrH

    If Len(spec.summarySheet) = 0 Or Len(spec.summaryAnchor) = 0 Then
        Err.Raise vbObjectError + 3, , _
                  "Summary not configured for " & spec.displayName & _
                  ": set summarySheet and summaryAnchor on the spec."
    End If

    mSrcTable = spec.tableName
    mOutSheet = spec.summarySheet
    mOutStartCell = spec.summaryAnchor
    mTablePrefix = spec.summaryTablePrefix
    If Len(mTablePrefix) = 0 Then mTablePrefix = "tblMetricsOut"

    Set loSrc = FindTable(mSrcTable)
    If loSrc Is Nothing Then Err.Raise vbObjectError + 1, , "Table '" & mSrcTable & "' not found."
    If loSrc.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 2, , "'" & mSrcTable & "' has no data."

    mData = loSrc.DataBodyRange.value
    cPurpose = loSrc.ListColumns("Purpose").Index
    cRepay = loSrc.ListColumns("Repayment_Type").Index
    cLvrMin = loSrc.ListColumns("LVR_Min").Index
    cLvrMax = loSrc.ListColumns("LVR_Max").Index
    cLsMin = loSrc.ListColumns("Loan_Size_Min").Index
    cLsMax = loSrc.ListColumns("Loan_Size_Max").Index
    cDisc = loSrc.ListColumns("Discretion").Index

    mMaxLVR = CDbl(ThisWorkbook.names(NAME_MAX_LVR).RefersToRange.value)
    mMaxLoan = CDbl(ThisWorkbook.names(NAME_MAX_LOAN).RefersToRange.value)

    Application.ScreenUpdating = False
    Set ws = GetOutputSheet(loSrc.Parent.Parent)

    nextRow = ws.Range(mOutStartCell).Row
    mStartCol = ws.Range(mOutStartCell).column
    purposes = Array("Owner Occupied", "Investment")

    For Each p In purposes
        Set rowsPI = RowsFor(CStr(p), "P&I")
        Set rowsIO = RowsFor(CStr(p), "IO")

        If rowsPI.count > 0 And rowsIO.count > 0 And SameGrid(rowsPI, rowsIO) Then
            BuildSummary ws, nextRow, tblNum, p & TITLE_SEP & "P&I/IO", rowsPI
        Else
            If rowsPI.count > 0 Then BuildSummary ws, nextRow, tblNum, p & TITLE_SEP & "P&I", rowsPI
            If rowsIO.count > 0 Then BuildSummary ws, nextRow, tblNum, p & TITLE_SEP & "IO", rowsIO
        End If
    Next p

    If tblNum = 0 Then MsgBox "No matching rows found - no summary tables created.", vbExclamation

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
ErrH:
    MsgBox "Discretion summary failed:" & vbCrLf & Err.description, vbCritical
    Resume Cleanup
End Sub

'------------------------------------------------------------------------------
' Build one titled summary table starting at nextRow
'------------------------------------------------------------------------------
Private Sub BuildSummary(ws As Worksheet, ByRef nextRow As Long, ByRef tblNum As Long, _
                         ByVal title As String, rowIdx As Collection)
    Dim lvrMin() As Double, lvrMax() As Double, nL As Long
    Dim lsMin() As Double, lsMax() As Double, nS As Long
    Dim grid() As Variant, r As Variant
    Dim i As Long, j As Long
    Dim rng As Range, lo As ListObject

    ' Unique bands present in this group
    For Each r In rowIdx
        AddBand lvrMin, lvrMax, nL, CDbl(mData(r, cLvrMin)), CDbl(mData(r, cLvrMax))
        AddBand lsMin, lsMax, nS, CDbl(mData(r, cLsMin)), CDbl(mData(r, cLsMax))
    Next r
    SortBands lvrMin, lvrMax, nL
    SortBands lsMin, lsMax, nS

    ' Grid: row 1 = headers, column 1 = loan size labels
    ReDim grid(1 To nS + 1, 1 To nL + 1)
    grid(1, 1) = " "   ' tables can't have a truly blank header
    For j = 1 To nL
        grid(1, j + 1) = LvrLabel(lvrMin(j), lvrMax(j))
    Next j
    For i = 1 To nS
        grid(i + 1, 1) = LoanLabel(lsMin(i), lsMax(i))
    Next i
    For Each r In rowIdx
        i = FindBand(lsMin, lsMax, nS, CDbl(mData(r, cLsMin)), CDbl(mData(r, cLsMax)))
        j = FindBand(lvrMin, lvrMax, nL, CDbl(mData(r, cLvrMin)), CDbl(mData(r, cLvrMax)))
        grid(i + 1, j + 1) = CDbl(mData(r, cDisc))
    Next r

    ' Title
    With ws.Cells(nextRow, mStartCol)
        .value = title
        .Font.Bold = True
    End With

    ' Table
    Set rng = ws.Cells(nextRow + 1, mStartCol).Resize(nS + 1, nL + 1)
    rng.value = grid

    tblNum = tblNum + 1
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = mTablePrefix & tblNum
    lo.TableStyle = TABLE_STYLE
    lo.ShowHeaders = True
    lo.ShowTotals = False
    lo.ShowTableStyleFirstColumn = True
    lo.ShowTableStyleLastColumn = False
    lo.ShowTableStyleRowStripes = False
    lo.ShowTableStyleColumnStripes = False
    lo.ShowAutoFilterDropDown = False
    lo.HeaderRowRange.Cells(1, 1).value = " "

    lo.DataBodyRange.Offset(0, 1).Resize(, nL).NumberFormat = "0.00%"
    lo.Range.HorizontalAlignment = xlCenter

    nextRow = nextRow + 1 + (nS + 1) + GAP_ROWS
End Sub

'------------------------------------------------------------------------------
' Labels
'------------------------------------------------------------------------------
Private Function LvrLabel(ByVal mn As Double, ByVal mx As Double) As String
    Dim atLow As Boolean, atHigh As Boolean
    atLow = (mn <= BOUND_TOL)
    atHigh = (mx >= mMaxLVR - BOUND_TOL)

    If atLow And atHigh Then
        LvrLabel = "All LVR Bands"
    ElseIf atLow Then
        LvrLabel = "<" & CStr(Round(mx, 0)) & "% LVR"
    ElseIf atHigh Then
        LvrLabel = ">" & CStr(Round(mn, 0)) & "% LVR"
    Else
        LvrLabel = CStr(Round(mn, 0)) & "-" & CStr(Round(mx, 0)) & "% LVR"
    End If
End Function

Private Function LoanLabel(ByVal mn As Double, ByVal mx As Double) As String
    Dim atLow As Boolean, atHigh As Boolean
    atLow = (mn <= BOUND_TOL)
    atHigh = (mx >= mMaxLoan - BOUND_TOL)

    If atLow And atHigh Then
        LoanLabel = "All Loan Sizes"
    ElseIf atLow Then
        LoanLabel = "<" & Money(mx)
    ElseIf atHigh Then
        LoanLabel = ">" & Money(mn)
    Else
        LoanLabel = Money(mn) & "-" & Money(mx)
    End If
End Function

' 499,999.99 -> $500k ; 999,999.99 -> $1m ; 1,250,000 -> $1.25m
Private Function Money(ByVal x As Double) As String
    Dim v As Double
    v = Round(x / 1000, 0) * 1000
    If v >= 1000000 Then
        Money = "$" & CStr(Round(v / 1000000, 2)) & "m"
    Else
        Money = "$" & CStr(Round(v / 1000, 0)) & "k"
    End If
End Function

'------------------------------------------------------------------------------
' Data helpers
'------------------------------------------------------------------------------
Private Function RowsFor(ByVal purpose As String, ByVal repay As String) As Collection
    Dim c As New Collection, r As Long
    For r = 1 To UBound(mData, 1)
        If StrComp(Trim$(CStr(mData(r, cPurpose))), purpose, vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(mData(r, cRepay))), repay, vbTextCompare) = 0 Then
            c.Add r
        End If
    Next r
    Set RowsFor = c
End Function

' True if both row sets contain exactly the same LVR/Loan/Discretion combos
Private Function SameGrid(a As Collection, b As Collection) As Boolean
    Dim ra As Variant, rb As Variant, found As Boolean
    If a.count <> b.count Then Exit Function
    For Each ra In a
        found = False
        For Each rb In b
            If RowsMatch(CLng(ra), CLng(rb)) Then found = True: Exit For
        Next rb
        If Not found Then Exit Function
    Next ra
    SameGrid = True
End Function

Private Function RowsMatch(ByVal r1 As Long, ByVal r2 As Long) As Boolean
    RowsMatch = Near(mData(r1, cLvrMin), mData(r2, cLvrMin)) And _
                Near(mData(r1, cLvrMax), mData(r2, cLvrMax)) And _
                Near(mData(r1, cLsMin), mData(r2, cLsMin)) And _
                Near(mData(r1, cLsMax), mData(r2, cLsMax)) And _
                Near(mData(r1, cDisc), mData(r2, cDisc))
End Function

Private Function Near(ByVal a As Variant, ByVal b As Variant) As Boolean
    Near = (Abs(CDbl(a) - CDbl(b)) < EQ_TOL)
End Function

Private Sub AddBand(mins() As Double, maxs() As Double, ByRef n As Long, _
                    ByVal mn As Double, ByVal mx As Double)
    If FindBand(mins, maxs, n, mn, mx) > 0 Then Exit Sub
    n = n + 1
    ReDim Preserve mins(1 To n)
    ReDim Preserve maxs(1 To n)
    mins(n) = mn
    maxs(n) = mx
End Sub

Private Function FindBand(mins() As Double, maxs() As Double, ByVal n As Long, _
                          ByVal mn As Double, ByVal mx As Double) As Long
    Dim k As Long
    For k = 1 To n
        If Abs(mins(k) - mn) < EQ_TOL And Abs(maxs(k) - mx) < EQ_TOL Then
            FindBand = k
            Exit Function
        End If
    Next k
End Function

' Insertion sort by min, then max
Private Sub SortBands(mins() As Double, maxs() As Double, ByVal n As Long)
    Dim i As Long, j As Long, tMin As Double, tMax As Double
    For i = 2 To n
        tMin = mins(i): tMax = maxs(i)
        j = i - 1
        Do While j >= 1
            If mins(j) > tMin Or (mins(j) = tMin And maxs(j) > tMax) Then
                mins(j + 1) = mins(j): maxs(j + 1) = maxs(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        mins(j + 1) = tMin: maxs(j + 1) = tMax
    Next i
End Sub

'------------------------------------------------------------------------------
' Workbook helpers
'------------------------------------------------------------------------------
Private Function FindTable(ByVal tblName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set FindTable = ws.ListObjects(tblName)
        If Not FindTable Is Nothing Then Exit Function
    Next ws
End Function

' Returns a cleared output sheet, creating it if needed
Private Function GetOutputSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet, k As Long
    On Error Resume Next
    Set ws = wb.Worksheets(mOutSheet)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = mOutSheet
    Else
        ClearOutputRegion ws
    End If
    Set GetOutputSheet = ws
End Function

' Clears the area the previous summary occupied: from mOutStartCell to the furthest
' row and column reached by its tables (mTablePrefix & <ordinal>). Anything left over
' inside that footprint goes too, but nothing below or beyond it is touched, so input
' tables such as tblActiveATL and tblActiveDiscretions on the same sheet survive.
Private Sub ClearOutputRegion(ws As Worksheet)
    Dim startCell As Range, footprint As Range, k As Long, lo As ListObject
    Dim lastRow As Long, lastCol As Long

    Set startCell = ws.Range(mOutStartCell)

    For k = ws.ListObjects.count To 1 Step -1
        Set lo = ws.ListObjects(k)
        If IsSummaryTableName(lo.Name) Then
            With lo.Range
                If .Row + .Rows.count - 1 > lastRow Then lastRow = .Row + .Rows.count - 1
                If .column + .Columns.count - 1 > lastCol Then lastCol = .column + .Columns.count - 1
            End With
            lo.Delete
        End If
    Next k

    If lastRow < startCell.Row Or lastCol < startCell.column Then Exit Sub
    Set footprint = ws.Range(startCell, ws.Cells(lastRow, lastCol))

    ' Never clear into a table this module did not build.
    For Each lo In ws.ListObjects
        If Not Intersect(lo.Range, footprint) Is Nothing Then
            Err.Raise vbObjectError + 4, , _
                      "Table '" & lo.Name & "' overlaps the summary area at " & mOutStartCell & _
                      " on '" & ws.Name & "'. Move it or change the summary anchor."
        End If
    Next lo

    footprint.Clear
End Sub

' True for names this module generates: the prefix followed only by digits. Keeps
' look-alikes such as tblMetricsOutSegment out of the cleanup.
Private Function IsSummaryTableName(ByVal tblName As String) As Boolean
    Dim suffix As String

    If StrComp(Left$(tblName, Len(mTablePrefix)), mTablePrefix, vbTextCompare) <> 0 Then Exit Function
    suffix = Mid$(tblName, Len(mTablePrefix) + 1)
    IsSummaryTableName = Len(suffix) > 0 And Not suffix Like "*[!0-9]*"
End Function

