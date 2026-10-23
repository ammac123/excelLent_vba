Attribute VB_Name = "ImportDataFromTable"
Option Explicit
 
' Generic engine for copying a block from a source workbook into a ListObject.
' One instance of this module serves every sheet; each sheet supplies its own ImportSpec.
 
Public Type ImportSpec
    key As String
    sourceSheet As String      ' leave empty for a CSV or any single-sheet workbook
    sourceHeaderCell As String
    targetSheet As String
    TargetTable As String
    statusCell As String       ' optional, leave empty to skip the status display
End Type


' ?????????????????????????????????????????????????????????????????????????????
' Table-sourced import: per-column metadata in this workbook drives a
' month-by-month lookup in the source workbook and fills a ListObject.
'
' Config block layout (rows are relative to configTopRow, default 16):
'   Row +0  (A = "date_row")    B, C, … = row# in source sheet where dates live
'   Row +1  (A = "description") B, C, … = personal reference label
'   Row +2  (A = "sheet")       B, C, … = source sheet name
'   Row +3  (A = "row")         B, C, … = row# in source sheet where values live
'   Row +4  (A = "header")      B, C, … = column header to write into the target table
'
' The target table must have "Month" as its first column; each subsequent column
' is assigned positionally (entry 1 ? col 2, entry 2 ? col 3, …) and its header
' is set from the "header" row of the config block.
' ?????????????????????????????????????????????????????????????????????????????

Public Type TableImportSpec
    key           As String
    configSheet   As String   ' sheet in ThisWorkbook holding the 5-row metadata block
    configTopRow  As Long     ' row of the date_row label; default 16
    startDateName As String   ' named range: first month-ending date to import
    endDateName   As String   ' named range: last month-ending date to import
    targetSheet   As String
    TargetTable   As String
    statusCell    As String
End Type
 
Public Function NewImportSpec(key As String, sourceSheet As String, sourceHeaderCell As String, _
                              targetSheet As String, TargetTable As String, _
                              Optional statusCell As String = "") As ImportSpec
    With NewImportSpec
        .key = key
        .sourceSheet = sourceSheet
        .sourceHeaderCell = sourceHeaderCell
        .targetSheet = targetSheet
        .TargetTable = TargetTable
        .statusCell = statusCell
    End With
End Function
 
Public Sub SelectSourceFile(spec As ImportSpec)
    Dim path As String
 
    path = SelectSourceFileFor(spec.key, "Select the source workbook for " & spec.TargetTable)
    If Len(path) = 0 Then Exit Sub
 
    RefreshStatus spec
    MsgBox spec.TargetTable & " source file set to:" & vbNewLine & path, vbInformation
End Sub
 
' Writes the selected file name into the spec's status cell. Safe to call from Workbook_Open.
Public Sub RefreshStatus(spec As ImportSpec)
    Dim ws As Worksheet
    Dim path As String
 
    If Len(spec.statusCell) = 0 Then Exit Sub
 
    Set ws = SheetOrNothing(spec.targetSheet)
    If ws Is Nothing Then Exit Sub
 
    path = SourceFilePath(spec.key)
 
    If Len(path) = 0 Then
        ws.Range(spec.statusCell).value = "No file selected"
    Else
        ws.Range(spec.statusCell).value = Mid$(path, InStrRev(path, "\") + 1)
    End If
End Sub
 
Public Sub RunImport(spec As ImportSpec)
    Dim srcWb As Workbook
    Dim wasOpen As Boolean
    Dim srcWs As Worksheet
    Dim lo As ListObject
    Dim headerRng As Range, dataRng As Range
    Dim values As Variant
    Dim formulas As Variant
    Dim numberFormats As Variant
    Dim rowCount As Long, srcCols As Long
 
    On Error GoTo Failed
 
    Set lo = GetTargetTable(spec)
    If lo Is Nothing Then Exit Sub
 
    Set srcWb = GetSourceWorkbook(spec, wasOpen)
    RefreshStatus spec
    If srcWb Is Nothing Then Exit Sub
 
    Set srcWs = WorkbookSheetOrNothing(srcWb, spec.sourceSheet)
    If srcWs Is Nothing Then
        MsgBox MissingSheetMessage(srcWb, spec.sourceSheet), vbCritical
        CloseIfOpened srcWb, wasOpen
        Exit Sub
    End If
 
    srcCols = SourceColumnCount(srcWs, spec.sourceHeaderCell)
    If srcCols = 0 Then
        MsgBox "No headers were found at " & spec.sourceHeaderCell & " in '" & srcWb.Name & "'.", vbCritical
        CloseIfOpened srcWb, wasOpen
        Exit Sub
    End If
 
    If srcCols > lo.ListColumns.count Then
        MsgBox "The source has " & srcCols & " columns but " & spec.TargetTable & " has only " & _
               lo.ListColumns.count & ".", vbCritical
        CloseIfOpened srcWb, wasOpen
        Exit Sub
    End If
 
    Set headerRng = srcWs.Range(spec.sourceHeaderCell).Resize(1, srcCols)
    If Not HeadersMatch(headerRng, lo) Then
        CloseIfOpened srcWb, wasOpen
        Exit Sub
    End If
 
    Set dataRng = SourceDataRange(headerRng)
    If dataRng Is Nothing Then
        MsgBox "No data rows were found below the header row in '" & spec.sourceSheet & "'.", vbExclamation
        CloseIfOpened srcWb, wasOpen
        Exit Sub
    End If
 
    values = dataRng.value
    rowCount = dataRng.Rows.count
    CloseIfOpened srcWb, wasOpen
 
    Application.ScreenUpdating = False
    formulas = CapturedFormulas(lo, srcCols)
    numberFormats = CapturedNumberFormats(lo, srcCols)
    ClearTableBody lo
    lo.HeaderRowRange.Offset(1, 0).Resize(rowCount, srcCols).value = values
    lo.Resize lo.HeaderRowRange.Resize(rowCount + 1, lo.ListColumns.count)
    RestoreFormulas lo, srcCols, formulas
    RestoreNumberFormats lo, srcCols, numberFormats
    Application.ScreenUpdating = True
 
    SourceLastImportedOn(spec.key) = Now
    UpdateModelCheckForKey spec.key

    MsgBox "Imported " & rowCount & " row(s) into " & spec.TargetTable & ".", vbInformation
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Import failed at error " & Err.Number & ":" & vbNewLine & Err.description, vbCritical
End Sub
 
Private Function GetTargetTable(spec As ImportSpec) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
 
    Set ws = SheetOrNothing(spec.targetSheet)
    If ws Is Nothing Then
        MsgBox "Sheet '" & spec.targetSheet & "' was not found in this workbook.", vbCritical
        Exit Function
    End If
 
    On Error Resume Next
    Set lo = ws.ListObjects(spec.TargetTable)
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Table '" & spec.TargetTable & "' was not found on sheet '" & ws.Name & "'.", vbCritical
        Exit Function
    End If
 
    Set GetTargetTable = lo
End Function
 
' Reuses the workbook if it is already open, otherwise opens it read-only.
Private Function GetSourceWorkbook(spec As ImportSpec, ByRef wasOpen As Boolean) As Workbook
    Dim result As Workbook
    Dim path As String
 
    path = ResolveSourcePath(spec)
    If Len(path) = 0 Then
        MsgBox "No source file was selected, so nothing was imported.", vbInformation
        Exit Function
    End If
 
    Set result = AlreadyOpenWorkbook(path)
    If Not result Is Nothing Then
        wasOpen = True
        Set GetSourceWorkbook = result
        Exit Function
    End If
 
    wasOpen = False
    Set GetSourceWorkbook = OpenReadOnly(path)
End Function
 
' Returns the stored path, prompting the user if it is missing or no longer on disk.
Private Function ResolveSourcePath(spec As ImportSpec) As String
    Dim path As String
 
    path = SourceFilePath(spec.key)
 
    If Len(path) > 0 Then
        If SourceFileExists(spec.key) Then
            ResolveSourcePath = path
            Exit Function
        End If
        MsgBox "The saved source workbook no longer exists:" & vbNewLine & path & _
               vbNewLine & vbNewLine & "Please select a replacement file.", vbExclamation
    End If
 
    ResolveSourcePath = SelectSourceFileFor(spec.key, _
                            "Select the source workbook for " & spec.TargetTable)
End Function
 
' Matches on full path so a same-named file from another folder is never mistaken for this one.
Private Function AlreadyOpenWorkbook(path As String) As Workbook
    Dim wb As Workbook
 
    For Each wb In Application.Workbooks
        If StrComp(wb.fullName, path, vbTextCompare) = 0 Then
            Set AlreadyOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function
 
' Opens without prompting. Reports clearly if another program or Excel instance has the file locked.
Private Function OpenReadOnly(path As String) As Workbook
    Dim wb As Workbook
    Dim prevAlerts As Boolean, prevAskLinks As Boolean
 
    prevAlerts = Application.DisplayAlerts
    prevAskLinks = Application.AskToUpdateLinks
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
 
    On Error Resume Next
    Set wb = Application.Workbooks.Open(fileName:=path, _
                                        UpdateLinks:=0, _
                                        ReadOnly:=True, _
                                        IgnoreReadOnlyRecommended:=True, _
                                        Notify:=False)
    If Err.Number <> 0 Then
        MsgBox "Could not open the source workbook:" & vbNewLine & path & vbNewLine & vbNewLine & _
               "Error " & Err.Number & ": " & Err.description & vbNewLine & vbNewLine & _
               "If the file is open in another Excel window, close it and try again.", vbCritical
        Err.Clear
    End If
    On Error GoTo 0
 
    Application.DisplayAlerts = prevAlerts
    Application.AskToUpdateLinks = prevAskLinks
 
    Set OpenReadOnly = wb
End Function
 
Private Sub CloseIfOpened(wb As Workbook, wasOpen As Boolean)
    If Not wasOpen Then wb.Close SaveChanges:=False
End Sub
 
Private Function SheetOrNothing(sheetName As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function
 
' An empty sheet name means "the only sheet", which is how CSV files are handled.
Private Function WorkbookSheetOrNothing(wb As Workbook, sheetName As String) As Worksheet
    If Len(sheetName) = 0 Then
        If wb.Worksheets.count = 1 Then Set WorkbookSheetOrNothing = wb.Worksheets(1)
        Exit Function
    End If
 
    On Error Resume Next
    Set WorkbookSheetOrNothing = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
 
Private Function MissingSheetMessage(wb As Workbook, sheetName As String) As String
    If Len(sheetName) = 0 Then
        MissingSheetMessage = "'" & wb.Name & "' holds " & wb.Worksheets.count & " sheets." & vbNewLine & _
                              "Set sourceSheet in the spec to name the one to import."
    Else
        MissingSheetMessage = "Sheet '" & sheetName & "' was not found in '" & wb.Name & "'."
    End If
End Function
 
Private Function HeadersMatch(headerRng As Range, lo As ListObject) As Boolean
    Dim i As Long
    Dim srcHeader As String, tgtHeader As String
 
    For i = 1 To headerRng.Columns.count
        srcHeader = Trim$(CStr(headerRng.Cells(1, i).value))
        tgtHeader = Trim$(CStr(lo.HeaderRowRange.Cells(1, i).value))
        If StrComp(srcHeader, tgtHeader, vbTextCompare) <> 0 Then
            MsgBox "Column " & i & " does not match." & vbNewLine & _
                   "Source: '" & srcHeader & "'" & vbNewLine & _
                   "Table:  '" & tgtHeader & "'", vbCritical
            Exit Function
        End If
    Next i
 
    HeadersMatch = True
End Function
 
' Contiguous non-blank header cells starting at headerCell.
Private Function SourceColumnCount(srcWs As Worksheet, headerCell As String) As Long
    Dim firstCell As Range
    Dim c As Long
 
    Set firstCell = srcWs.Range(headerCell)
 
    Do While Len(Trim$(CStr(firstCell.Offset(0, c).value))) > 0
        c = c + 1
        If firstCell.column + c > srcWs.Columns.count Then Exit Do
    Loop
 
    SourceColumnCount = c
End Function
 
' Grabs the R1C1 formula from the first body row of each column beyond the imported ones.
Private Function CapturedFormulas(lo As ListObject, srcCols As Long) As Variant
    Dim result() As String
    Dim i As Long
 
    If lo.ListColumns.count <= srcCols Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
 
    ReDim result(srcCols + 1 To lo.ListColumns.count)
    For i = srcCols + 1 To lo.ListColumns.count
        result(i) = lo.ListColumns(i).DataBodyRange.Cells(1, 1).FormulaR1C1
    Next i
 
    CapturedFormulas = result
End Function
 
Private Sub RestoreFormulas(lo As ListObject, srcCols As Long, formulas As Variant)
    Dim i As Long

    If IsEmpty(formulas) Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For i = srcCols + 1 To lo.ListColumns.count
        If Left$(formulas(i), 1) = "=" Then
            lo.ListColumns(i).DataBodyRange.FormulaR1C1 = formulas(i)
        End If
    Next i
End Sub

' ClearTableBody uses Range.Delete, which destroys the cells (including their
' NumberFormat). The write that follows lands in freshly created cells with the
' sheet default (usually General), so a column the user formatted as yyyy-mm-dd
' would come back as a date serial or a locale-default format. Capture before
' the delete and reapply after the resize to keep the user's column formats.
Private Function CapturedNumberFormats(lo As ListObject, srcCols As Long) As Variant
    Dim result() As String
    Dim i As Long
    Dim limit As Long

    If lo.DataBodyRange Is Nothing Then Exit Function

    limit = srcCols
    If limit > lo.ListColumns.count Then limit = lo.ListColumns.count

    ReDim result(1 To limit)
    For i = 1 To limit
        result(i) = lo.ListColumns(i).DataBodyRange.Cells(1).NumberFormat
    Next i

    CapturedNumberFormats = result
End Function

Private Sub RestoreNumberFormats(lo As ListObject, srcCols As Long, formats As Variant)
    Dim i As Long
    Dim limit As Long

    If IsEmpty(formats) Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    limit = srcCols
    If limit > lo.ListColumns.count Then limit = lo.ListColumns.count

    For i = 1 To limit
        lo.ListColumns(i).DataBodyRange.NumberFormat = formats(i)
    Next i
End Sub
 
' Returns the rows below the header up to (but excluding) the first fully blank row.
Private Function SourceDataRange(headerRng As Range) As Range
    Dim ws As Worksheet
    Dim firstRow As Long, r As Long
    Dim rowCount As Long
 
    Set ws = headerRng.Worksheet
    firstRow = headerRng.Row + 1
 
    rowCount = 0
    For r = firstRow To ws.Rows.count
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, headerRng.column), _
                ws.Cells(r, headerRng.column + headerRng.Columns.count - 1))) = 0 Then
            Exit For
        End If
        rowCount = rowCount + 1
    Next r
 
    If rowCount = 0 Then Exit Function
 
    Set SourceDataRange = ws.Cells(firstRow, headerRng.column).Resize(rowCount, headerRng.Columns.count)
End Function
 
Private Sub ClearTableBody(lo As ListObject)
    If lo.ListRows.count = 0 Then Exit Sub
    lo.DataBodyRange.Delete
End Sub


Public Function NewTableImportSpec(key As String, configSheet As String, _
                                   startDateName As String, endDateName As String, _
                                   targetSheet As String, TargetTable As String, _
                                   Optional configTopRow As Long = 16, _
                                   Optional statusCell As String = "") As TableImportSpec
    With NewTableImportSpec
        .key = key
        .configSheet = configSheet
        .configTopRow = configTopRow
        .startDateName = startDateName
        .endDateName = endDateName
        .targetSheet = targetSheet
        .TargetTable = TargetTable
        .statusCell = statusCell
    End With
End Function

Public Sub RunTableImport(spec As TableImportSpec)
    Dim coreSpec        As ImportSpec
    Dim srcWb           As Workbook
    Dim wasOpen         As Boolean
    Dim lo              As ListObject
    Dim entryCount      As Long
    Dim dateRowsArr()   As Long
    Dim descsArr()      As String
    Dim sheetNamesArr() As String
    Dim dataRowsArr()   As Long
    Dim headersArr()    As String
    Dim startDate       As Date
    Dim endDate         As Date
    Dim months()        As Date
    Dim monthCount      As Long
    Dim outputArr()     As Variant
    Dim m As Long, e As Long
    Dim srcWs           As Worksheet
    Dim srcCol          As Long

    On Error GoTo Failed

    coreSpec.key = spec.key
    coreSpec.targetSheet = spec.targetSheet
    coreSpec.TargetTable = spec.TargetTable
    coreSpec.statusCell = spec.statusCell

    Set lo = GetTargetTable(coreSpec)
    If lo Is Nothing Then Exit Sub

    Set srcWb = GetSourceWorkbook(coreSpec, wasOpen)
    RefreshStatus coreSpec
    If srcWb Is Nothing Then Exit Sub

    If Not ReadConfigEntries(spec, entryCount, dateRowsArr, descsArr, sheetNamesArr, dataRowsArr, headersArr) Then
        CloseIfOpened srcWb, wasOpen
        Exit Sub
    End If

    If Not ReadDateRange(spec, startDate, endDate) Then
        CloseIfOpened srcWb, wasOpen
        Exit Sub
    End If

    If lo.ListColumns.count < entryCount + 1 Then
        MsgBox spec.TargetTable & " has " & lo.ListColumns.count & " column(s) but the config has " & _
               entryCount & " entr(ies) plus the Month column (" & entryCount + 1 & " required).", vbCritical
        CloseIfOpened srcWb, wasOpen
        Exit Sub
    End If

    monthCount = BuildMonthList(startDate, endDate, months)
    If monthCount = 0 Then
        MsgBox "No months fall within the specified date range.", vbExclamation
        CloseIfOpened srcWb, wasOpen
        Exit Sub
    End If

    ReDim outputArr(1 To monthCount, 1 To lo.ListColumns.count)

    For m = 1 To monthCount
        outputArr(m, 1) = months(m)

        For e = 1 To entryCount
            Set srcWs = WorkbookSheetOrNothing(srcWb, sheetNamesArr(e))
            If Not srcWs Is Nothing Then
                srcCol = FindColumnForDate(srcWs, dateRowsArr(e), months(m))
                If srcCol > 0 Then
                    outputArr(m, e + 1) = srcWs.Cells(dataRowsArr(e), srcCol).value
                End If
            End If
        Next e
    Next m

    CloseIfOpened srcWb, wasOpen

    Application.ScreenUpdating = False
    ClearTableBody lo
    lo.HeaderRowRange.Offset(1, 0).Resize(monthCount, lo.ListColumns.count).value = outputArr
    lo.Resize lo.HeaderRowRange.Resize(monthCount + 1, lo.ListColumns.count)
    For e = 1 To entryCount
        lo.ListColumns(e + 1).Name = headersArr(e)
    Next e
    Application.ScreenUpdating = True

    SourceLastImportedOn(spec.key) = Now
    UpdateModelCheckForKey spec.key

    MsgBox "Imported " & monthCount & " month(s) into " & spec.TargetTable & ".", vbInformation
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Table import failed at error " & Err.Number & ":" & vbNewLine & Err.description, vbCritical
End Sub

Private Function ReadConfigEntries(spec As TableImportSpec, _
                                   ByRef count As Long, _
                                   ByRef dateRowsArr() As Long, _
                                   ByRef descsArr() As String, _
                                   ByRef sheetNamesArr() As String, _
                                   ByRef dataRowsArr() As Long, _
                                   ByRef headersArr() As String) As Boolean
    Dim cfgWs  As Worksheet
    Dim topRow As Long
    Dim c As Long, n As Long, i As Long

    Set cfgWs = SheetOrNothing(spec.configSheet)
    If cfgWs Is Nothing Then
        MsgBox "Config sheet '" & spec.configSheet & "' was not found in this workbook.", vbCritical
        Exit Function
    End If

    topRow = spec.configTopRow
    n = 0
    c = 2
    Do While c <= cfgWs.Columns.count
        If Len(Trim$(CStr(cfgWs.Cells(topRow, c).value))) = 0 And _
           Len(Trim$(CStr(cfgWs.Cells(topRow + 1, c).value))) = 0 And _
           Len(Trim$(CStr(cfgWs.Cells(topRow + 2, c).value))) = 0 And _
           Len(Trim$(CStr(cfgWs.Cells(topRow + 3, c).value))) = 0 And _
           Len(Trim$(CStr(cfgWs.Cells(topRow + 4, c).value))) = 0 Then Exit Do
        n = n + 1
        c = c + 1
    Loop

    If n = 0 Then
        MsgBox "No config entries found in '" & spec.configSheet & "' at row " & topRow & ".", vbExclamation
        Exit Function
    End If

    ReDim dateRowsArr(1 To n)
    ReDim descsArr(1 To n)
    ReDim sheetNamesArr(1 To n)
    ReDim dataRowsArr(1 To n)
    ReDim headersArr(1 To n)

    For i = 1 To n
        c = i + 1
        dateRowsArr(i) = CLng(cfgWs.Cells(topRow, c).value)
        descsArr(i) = CStr(cfgWs.Cells(topRow + 1, c).value)
        sheetNamesArr(i) = CStr(cfgWs.Cells(topRow + 2, c).value)
        dataRowsArr(i) = CLng(cfgWs.Cells(topRow + 3, c).value)
        headersArr(i) = CStr(cfgWs.Cells(topRow + 4, c).value)
    Next i

    count = n
    ReadConfigEntries = True
End Function

Private Function ReadDateRange(spec As TableImportSpec, _
                               ByRef startDate As Date, _
                               ByRef endDate As Date) As Boolean
    On Error Resume Next

    startDate = CDate(ThisWorkbook.names(spec.startDateName).RefersToRange.value)
    If Err.Number <> 0 Then
        MsgBox "Named range '" & spec.startDateName & "' was not found or is not a valid date.", vbCritical
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    endDate = CDate(ThisWorkbook.names(spec.endDateName).RefersToRange.value)
    If Err.Number <> 0 Then
        MsgBox "Named range '" & spec.endDateName & "' was not found or is not a valid date.", vbCritical
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    If startDate > endDate Then
        MsgBox "Start date (" & Format$(startDate, "dd/mm/yyyy") & ") must not be after end date (" & _
               Format$(endDate, "dd/mm/yyyy") & ").", vbCritical
        Exit Function
    End If

    ReadDateRange = True
End Function

' Steps through month-ending dates from startDate to endDate inclusive.
' Each next date is the last day of the following calendar month.
Private Function BuildMonthList(startDate As Date, endDate As Date, ByRef months() As Date) As Long
    Dim n As Long
    Dim d As Date

    n = 0
    d = startDate
    Do While d <= endDate
        n = n + 1
        d = DateSerial(Year(d), Month(d) + 2, 0)
    Loop

    If n = 0 Then Exit Function

    ReDim months(1 To n)
    n = 0
    d = startDate
    Do While d <= endDate
        n = n + 1
        months(n) = d
        d = DateSerial(Year(d), Month(d) + 2, 0)
    Loop

    BuildMonthList = n
End Function

Private Function FindColumnForDate(ws As Worksheet, dateRow As Long, targetDate As Date) As Long
    Dim lastCol  As Long
    Dim c        As Long
    Dim cellVal  As Variant
    Dim cellDate As Date

    lastCol = ws.Cells(dateRow, ws.Columns.count).End(xlToLeft).column

    For c = 1 To lastCol
        cellVal = ws.Cells(dateRow, c).value
        If Not IsEmpty(cellVal) Then
            On Error Resume Next
            cellDate = CDate(cellVal)
            If Err.Number = 0 Then
                If cellDate = targetDate Then
                    FindColumnForDate = c
                    On Error GoTo 0
                    Exit Function
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next c
End Function


