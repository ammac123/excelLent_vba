Attribute VB_Name = "ModelInfo"
Option Explicit

' Version and governance record for this workbook. The constants below are the
' source of truth: WriteModelInfo rebuilds sheet 00_Governance from them every
' time the workbook opens, so anything typed over the sheet is gone by the next
' open. To release a new version, bump MODEL_VERSION, update the review fields
' and add a line to the top of ChangeLog.
'
' Each value cell is given a workbook name (mi_Version, mi_Owner, ...) so other
' sheets can show it, e.g. ="v" & mi_Version in a header or print footer cell.
'
' Hook it up in ThisWorkbook:
'   Private Sub Workbook_Open()
'       WriteModelInfo
'   End Sub

Private Const INFO_SHEET As String = "00_Governance"
Private Const NAME_PREFIX As String = "mi_"
Private Const SHEET_PASSWORD As String = ""

' ---- Edit these per release -------------------------------------------------
Private Const MODEL_NAME As String = "Mortgages Pricing Tool"
Private Const MODEL_VERSION As String = "0.1.0"
Private Const MODEL_PURPOSE As String = "Scenario projection of mortgages pricing changes."

Private Const AUTHOR As String = ""
Private Const MODEL_OWNER As String = ""
Private Const REVIEWED_BY As String = ""
Private Const APPROVED_BY As String = ""

' yyyy-mm-dd, read with IsoDate. A Date literal would be shown as #m/d/yyyy# by the VBE.
Private Const LAST_REVIEWED As String = "2026-10-26"
Private Const REVIEW_MONTHS As Long = 12
Private Const DUE_SOON_DAYS As Long = 30
' -----------------------------------------------------------------------------

' Newest first. Each entry: version | date (yyyy-mm-dd) | author | change.
Private Function ChangeLog() As Variant
    ChangeLog = Array( _
        "0.1.0|2026-10-26|Andrew Macleod|Initial pre-release." _
    )
End Function

' For code that wants to stamp outputs with the model version.
Public Function ModelVersionText() As String
    ModelVersionText = MODEL_VERSION
End Function

' Rebuilds the info sheet and its names. Safe to call from Workbook_Open: any
' failure is reported to the Immediate window rather than stopping the open.
Public Sub WriteModelInfo()
    Dim ws As Worksheet
    Dim wasSaved As Boolean
    Dim r As Long
    Dim nextReview As Date

    On Error GoTo Failed

    ' Rewriting the sheet dirties the workbook; restore the flag so opening and
    ' closing without changes does not prompt to save.
    wasSaved = ThisWorkbook.Saved
    Application.ScreenUpdating = False

    Set ws = InfoSheet()
    ws.Unprotect SHEET_PASSWORD
    ws.Cells.Clear

    ws.Range("A4").value = MODEL_NAME
    ws.Range("A4").Font.Bold = True
    ws.Range("A4").Font.Size = 14
    ws.Range("A5").value = "Generated from the ModelInfo module on open. Edit the module, not this sheet."
    ws.Range("A5").Font.Italic = True

    nextReview = DateAdd("m", REVIEW_MONTHS, IsoDate(LAST_REVIEWED))

    r = 7
    WriteHeading ws, r, "Model"
    WriteField ws, r, "Version", "Version", MODEL_VERSION
    WriteField ws, r, "Purpose", "Purpose", MODEL_PURPOSE

    r = r + 1
    WriteHeading ws, r, "Governance"
    WriteField ws, r, "Author", "Author", AUTHOR
    WriteField ws, r, "Owner", "Owner", MODEL_OWNER
    WriteField ws, r, "Reviewed by", "ReviewedBy", REVIEWED_BY
    WriteField ws, r, "Approved by", "ApprovedBy", APPROVED_BY
    WriteField ws, r, "Last reviewed", "LastReviewed", IsoDate(LAST_REVIEWED)
    WriteField ws, r, "Next review due", "NextReviewDue", nextReview
    WriteField ws, r, "Review status", "ReviewStatus", ReviewStatus(nextReview)

    r = r + 1
    WriteHeading ws, r, "This session"
    WriteField ws, r, "Opened by", "OpenedBy", Application.UserName
    WriteField ws, r, "Opened at", "OpenedAt", Now
    WriteField ws, r, "File", "File", ThisWorkbook.FullName
    WriteField ws, r, "Last saved by", "LastSavedBy", DocProperty("Last author")
    WriteField ws, r, "Last saved at", "LastSavedAt", DocProperty("Last save time")
    WriteField ws, r, "Excel", "Excel", Application.Version & " (" & Application.OperatingSystem & ")"

    r = r + 1
    WriteChangeLog ws, r

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 60
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 60
    ws.Protect SHEET_PASSWORD

    Application.ScreenUpdating = True
    ThisWorkbook.Saved = wasSaved
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    Debug.Print "WriteModelInfo failed: " & Err.Number & " " & Err.Description
End Sub

Private Function InfoSheet() As Worksheet
    Set InfoSheet = SheetOrNothing(INFO_SHEET)
    If Not InfoSheet Is Nothing Then Exit Function

    Set InfoSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    InfoSheet.Name = INFO_SHEET
End Function

Private Sub WriteHeading(ws As Worksheet, ByRef r As Long, text As String)
    ws.Cells(r, 1).value = text
    ws.Cells(r, 1).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    r = r + 1
End Sub

' Writes label | value | name on row r, points the workbook name at the value
' cell and moves r to the next row.
Private Sub WriteField(ws As Worksheet, ByRef r As Long, label As String, _
                       nameSuffix As String, value As Variant)
    Dim cell As Range

    Set cell = ws.Cells(r, 2)

    ' Text is forced to text so "1.10" stays "1.10" rather than becoming 1.1.
    If VarType(value) = vbDate Then
        cell.NumberFormat = IIf(value = Int(value), "yyyy-mm-dd", "yyyy-mm-dd hh:mm")
    Else
        cell.NumberFormat = "@"
    End If
    cell.value = value
    cell.HorizontalAlignment = xlLeft

    ws.Cells(r, 1).value = label
    ws.Cells(r, 3).value = NAME_PREFIX & nameSuffix
    ws.Cells(r, 3).Font.Color = RGB(128, 128, 128)

    ThisWorkbook.names.Add Name:=NAME_PREFIX & nameSuffix, _
                           RefersTo:="='" & ws.Name & "'!" & cell.Address

    r = r + 1
End Sub

Private Sub WriteChangeLog(ws As Worksheet, ByRef r As Long)
    Dim entries As Variant
    Dim parts As Variant
    Dim i As Long

    WriteHeading ws, r, "Change log"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).value = Array("Version", "Date", "Author", "Change")
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Font.Bold = True
    r = r + 1

    entries = ChangeLog()
    For i = LBound(entries) To UBound(entries)
        parts = Split(entries(i), "|")
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).value = Trim$(parts(0))
        ws.Cells(r, 2).value = IsoDate(Trim$(parts(1)))
        ws.Cells(r, 2).NumberFormat = "yyyy-mm-dd"
        ws.Cells(r, 2).HorizontalAlignment = xlLeft
        ws.Cells(r, 3).value = Trim$(parts(2))
        ws.Cells(r, 4).value = Trim$(parts(3))
        ws.Cells(r, 4).WrapText = True
        r = r + 1
    Next i
End Sub

Private Function ReviewStatus(dueOn As Date) As String
    If Date > dueOn Then
        ReviewStatus = "Overdue"
    ElseIf Date > dueOn - DUE_SOON_DAYS Then
        ReviewStatus = "Due within " & DUE_SOON_DAYS & " days"
    Else
        ReviewStatus = "Current"
    End If
End Function

' yyyy-mm-dd parsed by position so regional date settings cannot swap day and month.
Private Function IsoDate(text As String) As Date
    IsoDate = DateSerial(CInt(Left$(text, 4)), CInt(Mid$(text, 6, 2)), CInt(Right$(text, 2)))
End Function

' Built-in properties such as "Last save time" raise an error on a workbook that
' has never been saved; return "" in that case.
Private Function DocProperty(propName As String) As Variant
    DocProperty = ""
    On Error Resume Next
    DocProperty = ThisWorkbook.BuiltinDocumentProperties(propName).value
    On Error GoTo 0
End Function
