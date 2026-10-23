Attribute VB_Name = "ValidateDiscretionsTable"
Option Explicit

' Validates tblDiscretions for complete, non-overlapping coverage.
'
' For every Purpose x Repayment_Type combination the rows must tile the rectangle
'   LVR        0 .. [Validation_Max_LVR_Value]
'   Loan size  0 .. [Validation_Max_Loan_Size]
' exactly once: no gaps, no overlaps, nothing outside the domain.
'
' Bands are inclusive. Two edges within TOL of each other are treated as the same
' boundary, so 0-999,999.99 followed by 1,000,000-1,999,999.99 is contiguous rather
' than a gap, and 0-80 followed by 80-90 is contiguous rather than an overlap.

Private Const NAME_MAX_LVR As String = "validationMaxLVRValue"
Private Const NAME_MAX_LOAN As String = "validationMaxLoanSize"

Private Const COL_PURPOSE As String = "Purpose"
Private Const COL_REPAYMENT As String = "Repayment_Type"
Private Const COL_LVR_MIN As String = "LVR_Min"
Private Const COL_LVR_MAX As String = "LVR_Max"
Private Const COL_LOAN_MIN As String = "Loan_Size_Min"
Private Const COL_LOAN_MAX As String = "Loan_Size_Max"

Private Const TOL As Double = 0.01
Private Const MAX_ISSUES_PER_GROUP As Long = 12
Private Const MAX_ISSUES_TOTAL As Long = 60
Private Const MAX_ISSUES_SUMMARY As Long = 8

' Preserved for backwards compatibility; delegates to the Active spec.
Public Sub ValidateDiscretions()
    ValidateDiscretionsFor ActiveDiscretionsSpec()
End Sub

' Preserved for backwards compatibility; delegates to the Active spec.
Public Sub ValidateDiscretionsAndGenerateSummary()
    ValidateAndSummarizeFor ActiveDiscretionsSpec()
End Sub

' Validate the table named in the spec, show the result in a MsgBox, and stamp
' LastValidatedOn with the spec's state key.
Public Sub ValidateDiscretionsFor(ByRef spec As DiscretionsSpec)
    RunDiscretionChecks spec
End Sub

' Shared workhorse: runs the checks, reports them, and returns the collection so
' chained flows (e.g. ValidateAndSummarizeFor) can gate on the pass/fail result.
' Returns Nothing when the run itself errored, so callers can distinguish "failed to
' run" from "ran and found zero issues".
Public Function RunDiscretionChecks(ByRef spec As DiscretionsSpec) As Collection
    On Error GoTo Failed

    Set RunDiscretionChecks = CollectIssues(spec.tableName)
    ReportIssuesFor RunDiscretionChecks, spec
    Exit Function

Failed:
    MsgBox "Validation of " & spec.displayName & " failed at error " & Err.Number & ":" & _
           vbNewLine & Err.description, vbCritical
    Set RunDiscretionChecks = Nothing
End Function

' Silent pass/fail for callers that need to gate on coverage rather than report it,
' such as saving or running a scenario. Never shows a message box.
Public Function DiscretionsAreValid(Optional ByVal tableName As String = "tblActiveDiscretions", _
                                    Optional ByRef summary As String) As Boolean
    Dim issues As Collection
    Dim i As Long

    summary = ""

    On Error GoTo Failed
    Set issues = CollectIssues(tableName)
    On Error GoTo 0

    If issues.count = 0 Then
        DiscretionsAreValid = True
        Exit Function
    End If

    For i = 1 To issues.count
        If i > MAX_ISSUES_SUMMARY Then
            summary = summary & "- ... and " & (issues.count - MAX_ISSUES_SUMMARY) & " more."
            Exit For
        End If
        summary = summary & "- " & issues(i) & vbNewLine
    Next i
    Exit Function

Failed:
    summary = Err.description
End Function

' The date the named spec's table last passed validation. Empty if it never has.
' Defaults to the Active spec's state key so legacy callers keep working.
Public Function DiscretionsLastValidated(Optional ByVal stateKey As String = "Discretions") As Variant
    DiscretionsLastValidated = LastValidatedOn(stateKey)
End Function

' Every coverage problem found in the named table, newest check wins.
Private Function CollectIssues(ByVal tableName As String) As Collection
    Dim lo As ListObject
    Dim maxLvr As Double, maxLoan As Double
    Dim purposes As Variant, repayments As Variant
    Dim groups As Object
    Dim issues As Collection
    Dim p As Long, r As Long
    Dim key As String

    Set lo = FindTable(tableName)
    maxLvr = NamedNumber(NAME_MAX_LVR)
    maxLoan = NamedNumber(NAME_MAX_LOAN)

    purposes = Array("Owner Occupied", "Investment")
    repayments = Array("P&I", "IO")

    Set issues = New Collection
    Set CollectIssues = issues
    Set groups = LoadGroups(lo, maxLvr, maxLoan, issues)

    For p = LBound(purposes) To UBound(purposes)
        For r = LBound(repayments) To UBound(repayments)
            key = GroupKey(CStr(purposes(p)), CStr(repayments(r)))
            If Not groups.Exists(key) Then
                Add issues, CStr(purposes(p)) & " / " & CStr(repayments(r)) & ": no rows found."
            Else
                CheckGroup groups(key), CStr(purposes(p)) & " / " & CStr(repayments(r)), _
                           maxLvr, maxLoan, issues
            End If
        Next r
    Next p
End Function

' Reads the table into one collection of bands per Purpose/Repayment_Type combination,
' logging any row that is malformed or reaches outside the validation domain.
Private Function LoadGroups(lo As ListObject, ByVal maxLvr As Double, ByVal maxLoan As Double, _
                            ByVal issues As Collection) As Object
    Dim groups As Object
    Dim body As Variant
    Dim cPurpose As Long, cRepay As Long
    Dim cLvrMin As Long, cLvrMax As Long, cLoanMin As Long, cLoanMax As Long
    Dim rowNo As Long, excelRow As Long
    Dim band(0 To 4) As Double
    Dim key As String, label As String

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = 1
    Set LoadGroups = groups

    If lo.DataBodyRange Is Nothing Then
        Add issues, lo.Name & " is empty."
        Exit Function
    End If

    body = lo.DataBodyRange.value
    cPurpose = ColIndex(lo, COL_PURPOSE)
    cRepay = ColIndex(lo, COL_REPAYMENT)
    cLvrMin = ColIndex(lo, COL_LVR_MIN)
    cLvrMax = ColIndex(lo, COL_LVR_MAX)
    cLoanMin = ColIndex(lo, COL_LOAN_MIN)
    cLoanMax = ColIndex(lo, COL_LOAN_MAX)

    For rowNo = 1 To UBound(body, 1)
        excelRow = lo.DataBodyRange.Row + rowNo - 1
        label = "Row " & excelRow

        If Not IsNumeric(body(rowNo, cLvrMin)) Or Not IsNumeric(body(rowNo, cLvrMax)) _
           Or Not IsNumeric(body(rowNo, cLoanMin)) Or Not IsNumeric(body(rowNo, cLoanMax)) Then
            Add issues, label & ": LVR or loan size bounds are not numeric."
            GoTo nextRow
        End If

        band(0) = CDbl(body(rowNo, cLvrMin))
        band(1) = CDbl(body(rowNo, cLvrMax))
        band(2) = CDbl(body(rowNo, cLoanMin))
        band(3) = CDbl(body(rowNo, cLoanMax))
        band(4) = excelRow

        If band(1) - band(0) <= TOL Then
            Add issues, label & ": LVR band is empty or inverted (" & _
                        FmtLvr(band(0)) & " to " & FmtLvr(band(1)) & ")."
            GoTo nextRow
        End If
        If band(3) - band(2) <= TOL Then
            Add issues, label & ": loan size band is empty or inverted (" & _
                        FmtLoan(band(2)) & " to " & FmtLoan(band(3)) & ")."
            GoTo nextRow
        End If
        If band(0) < -TOL Or band(1) > maxLvr + TOL Then
            Add issues, label & ": LVR band " & FmtLvr(band(0)) & " to " & FmtLvr(band(1)) & _
                        " falls outside 0 to " & FmtLvr(maxLvr) & "."
        End If
        If band(2) < -TOL Or band(3) > maxLoan + TOL Then
            Add issues, label & ": loan size band " & FmtLoan(band(2)) & " to " & FmtLoan(band(3)) & _
                        " falls outside " & FmtLoan(0) & " to " & FmtLoan(maxLoan) & "."
        End If

        key = GroupKey(CStr(body(rowNo, cPurpose)), CStr(body(rowNo, cRepay)))
        If Not groups.Exists(key) Then groups.Add key, New Collection
        groups(key).Add band

nextRow:
    Next rowNo
End Function

' Snaps near-identical band edges onto shared boundaries, then coordinate-compresses the
' group into a grid and counts how many bands cover each cell: zero is a gap, more than
' one is an overlap.
Private Sub CheckGroup(ByVal bands As Collection, ByVal groupLabel As String, _
                       ByVal maxLvr As Double, ByVal maxLoan As Double, _
                       ByVal issues As Collection)
    Dim rects() As Double
    Dim xs() As Double, ys() As Double
    Dim i As Long, j As Long, k As Long
    Dim xLo As Double, xHi As Double, yLo As Double, yHi As Double
    Dim covering As Long
    Dim found As Long

    rects = ToArray(bands, maxLvr, maxLoan)
    xs = SnapAxis(rects, 0, 1, maxLvr)
    ys = SnapAxis(rects, 2, 3, maxLoan)

    For i = LBound(xs) To UBound(xs) - 1
        xLo = xs(i): xHi = xs(i + 1)
        For j = LBound(ys) To UBound(ys) - 1
            yLo = ys(j): yHi = ys(j + 1)

            covering = 0
            For k = 1 To UBound(rects, 1)
                If rects(k, 0) <= xLo And rects(k, 1) >= xHi _
                   And rects(k, 2) <= yLo And rects(k, 3) >= yHi Then
                    covering = covering + 1
                End If
            Next k

            If covering <> 1 Then
                found = found + 1
                If found <= MAX_ISSUES_PER_GROUP Then
                    Add issues, groupLabel & ": " & IIf(covering = 0, "gap", "overlap (" & covering & " rows)") & _
                                " at LVR " & FmtLvr(xLo) & "-" & FmtLvr(xHi) & _
                                ", loan size " & FmtLoan(yLo) & "-" & FmtLoan(yHi) & "."
                End If
            End If
        Next j
    Next i

    If found > MAX_ISSUES_PER_GROUP Then
        Add issues, groupLabel & ": " & (found - MAX_ISSUES_PER_GROUP) & " further gap/overlap cell(s) not listed."
    End If
End Sub

' Copies the group's bands into rows of (lvrMin, lvrMax, loanMin, loanMax, excelRow),
' clamped to the validation domain.
Private Function ToArray(ByVal bands As Collection, ByVal maxLvr As Double, _
                         ByVal maxLoan As Double) As Double()
    Dim rects() As Double
    Dim band As Variant
    Dim i As Long

    ReDim rects(1 To bands.count, 0 To 4)

    For Each band In bands
        i = i + 1
        rects(i, 0) = Clamp(band(0), 0, maxLvr)
        rects(i, 1) = Clamp(band(1), 0, maxLvr)
        rects(i, 2) = Clamp(band(2), 0, maxLoan)
        rects(i, 3) = Clamp(band(3), 0, maxLoan)
        rects(i, 4) = band(4)
    Next band

    ToArray = rects
End Function

' Clusters one axis of edges into distinct boundaries, rewrites each band's edges onto its
' boundary, and returns the sorted boundaries. Snapping is what lets a band ending at
' 999,999.99 meet the next one starting at 1,000,000 without leaving a sliver behind.
Private Function SnapAxis(ByRef rects() As Double, ByVal loIdx As Long, ByVal hiIdx As Long, _
                          ByVal domainMax As Double) As Double()
    Dim raw() As Double, reps() As Double
    Dim n As Long, i As Long, kept As Long

    n = UBound(rects, 1)

    ReDim raw(0 To n * 2 + 1)
    raw(0) = 0
    raw(1) = domainMax
    For i = 1 To n
        raw(i * 2) = rects(i, loIdx)
        raw(i * 2 + 1) = rects(i, hiIdx)
    Next i
    SortDoubles raw

    ReDim reps(0 To UBound(raw))
    For i = 0 To UBound(raw)
        If kept = 0 Then
            reps(0) = raw(i)
            kept = 1
        ElseIf Not SameEdge(raw(i), reps(kept - 1)) Then
            reps(kept) = raw(i)
            kept = kept + 1
        End If
    Next i
    ReDim Preserve reps(0 To kept - 1)

    For i = 1 To n
        rects(i, loIdx) = SnapTo(rects(i, loIdx), reps)
        rects(i, hiIdx) = SnapTo(rects(i, hiIdx), reps)
    Next i

    SnapAxis = reps
End Function

Private Function SnapTo(ByVal value As Double, ByRef reps() As Double) As Double
    Dim i As Long, best As Long

    best = LBound(reps)
    For i = LBound(reps) + 1 To UBound(reps)
        If Abs(reps(i) - value) < Abs(reps(best) - value) Then best = i
    Next i

    SnapTo = reps(best)
End Function

' The relative term absorbs the floating-point error in values the size of a loan balance,
' where 1,000,000 - 999,999.99 does not evaluate to exactly 0.01.
Private Function SameEdge(ByVal a As Double, ByVal b As Double) As Boolean
    SameEdge = (Abs(a - b) <= TOL + Abs(b) * 0.000000000001)
End Function

Private Sub ReportIssuesFor(ByVal issues As Collection, ByRef spec As DiscretionsSpec)
    Dim msg As String
    Dim i As Long

    If issues.count = 0 Then
        LastValidatedOn(spec.stateKey) = Date
        MsgBox spec.displayName & " is valid: every Purpose and Repayment_Type combination " & _
               "is fully covered with no overlaps.", vbInformation
        Exit Sub
    End If

    ' A failed run invalidates any earlier pass.
    LastValidatedOn(spec.stateKey) = Empty

    msg = issues.count & " issue(s) found in " & spec.displayName & " (" & spec.tableName & "):" & _
          vbNewLine & vbNewLine
    For i = 1 To issues.count
        If i > MAX_ISSUES_TOTAL Then
            msg = msg & "... and " & (issues.count - MAX_ISSUES_TOTAL) & " more."
            Exit For
        End If
        msg = msg & "- " & issues(i) & vbNewLine
    Next i

    MsgBox msg, vbExclamation
End Sub

Private Sub Add(ByVal issues As Collection, ByVal text As String)
    issues.Add text
End Sub

Private Function GroupKey(ByVal purpose As String, ByVal repayment As String) As String
    GroupKey = UCase$(Trim$(purpose)) & "|" & UCase$(Trim$(repayment))
End Function

Private Function Clamp(ByVal value As Double, ByVal lowest As Double, ByVal highest As Double) As Double
    If value < lowest Then
        Clamp = lowest
    ElseIf value > highest Then
        Clamp = highest
    Else
        Clamp = value
    End If
End Function

Private Function FmtLvr(ByVal value As Double) As String
    If value <= 1.5 Then
        FmtLvr = Format$(value, "0.##%")
    Else
        FmtLvr = Format$(value, "0.##") & "%"
    End If
End Function

Private Function FmtLoan(ByVal value As Double) As String
    FmtLoan = Format$(value, "$#,##0")
End Function

Private Function NamedNumber(nameRef As String) As Double
    Dim nm As Name
    Dim value As Variant

    On Error Resume Next
    Set nm = ThisWorkbook.names(nameRef)
    On Error GoTo 0

    If nm Is Nothing Then
        Err.Raise vbObjectError + 520, "ValidateDiscretions", _
                  "Named reference '" & nameRef & "' was not found in this workbook."
    End If

    On Error Resume Next
    value = nm.RefersToRange.value
    On Error GoTo 0

    If Not IsNumeric(value) Then
        Err.Raise vbObjectError + 521, "ValidateDiscretions", _
                  "Named reference '" & nameRef & "' does not contain a number."
    End If

    NamedNumber = CDbl(value)
End Function

Private Function ColIndex(lo As ListObject, header As String) As Long
    Dim lc As ListColumn

    On Error Resume Next
    Set lc = lo.ListColumns(header)
    On Error GoTo 0

    If lc Is Nothing Then
        Err.Raise vbObjectError + 522, "ValidateDiscretions", _
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

    Err.Raise vbObjectError + 523, "ValidateDiscretions", _
              "Table '" & tableName & "' was not found in this workbook."
End Function

Private Sub SortDoubles(ByRef values() As Double)
    Dim i As Long, j As Long
    Dim temp As Double

    For i = LBound(values) + 1 To UBound(values)
        temp = values(i)
        j = i - 1
        Do While j >= LBound(values)
            If values(j) <= temp Then Exit Do
            values(j + 1) = values(j)
            j = j - 1
        Loop
        values(j + 1) = temp
    Next i
End Sub




