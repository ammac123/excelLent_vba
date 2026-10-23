Attribute VB_Name = "CopyApplicationsToDiscretion"
Option Explicit

' Copies a configurable set of columns from tblApplications into
' tblDiscretionMapping, row for row. The destination is resized to the exact
' source row count on every run, so this fully replaces the destination body.
'
' Matching is by header name. A column named in ColumnsToCopy that is missing
' from either table is skipped silently, so the config list can be a superset
' while the schemas settle.

Private Const SRC_TABLE As String = "tblApplications"
Private Const DEST_TABLE As String = "tblDiscretionMapping"

' Edit this list to control which columns move across. Headers must match the
' column names in both tables exactly (case-insensitive).
Private Function ColumnsToCopy() As Variant
    ColumnsToCopy = Array( _
        "Pricing_Segment_ID", _
        "Application_Submitted_Date", _
        "Application_LVR", _
        "Application_Amount")
End Function

' Entry point. Assign to a button, or call from another macro.
Public Sub CopyApplicationsToDiscretionMapping()
    Dim loSrc As ListObject
    Dim loDest As ListObject
    Dim cols As Variant
    Dim rowCount As Long

    Set loSrc = RequireTable(SRC_TABLE)
    Set loDest = RequireTable(DEST_TABLE)
    cols = ColumnsToCopy

    rowCount = TableRowCount(loSrc)

    Application.ScreenUpdating = False
    ResizeTableBody loDest, rowCount
    If rowCount > 0 Then CopyConfiguredColumns loSrc, loDest, cols, rowCount
    Application.ScreenUpdating = True

    ' Freeze the "current" discretion lookup into the destination now that the inputs
    ' have been copied. The scenario column is refreshed separately on each scenario run.
    RefreshDiscretionCurrent
End Sub

Private Sub CopyConfiguredColumns(loSrc As ListObject, loDest As ListObject, _
                                  cols As Variant, ByVal rowCount As Long)
    Dim i As Long
    Dim header As String
    Dim srcIdx As Long
    Dim destIdx As Long

    For i = LBound(cols) To UBound(cols)
        header = CStr(cols(i))
        srcIdx = ColumnIndexOrZero(loSrc, header)
        destIdx = ColumnIndexOrZero(loDest, header)
        If srcIdx > 0 And destIdx > 0 Then
            loDest.ListColumns(destIdx).DataBodyRange.value = _
                loSrc.ListColumns(srcIdx).DataBodyRange.value
        End If
    Next i
End Sub
