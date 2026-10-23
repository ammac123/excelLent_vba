Attribute VB_Name = "AppState"
Option Explicit

' Workbook-wide source file paths, readable and writable from any module or sheet code.
' Each caller supplies its own key so separate imports can track separate files.
' Values are stored in hidden defined names so they survive closing the workbook.

Private Const NAME_PREFIX As String = "SourcePath_"
Private Const SELECTED_PREFIX As String = "SourceDate_"
Private Const IMPORTED_PREFIX As String = "ImportDate_"
Private Const VALIDATED_PREFIX As String = "ValidatedDate_"
Private Const CALCULATED_PREFIX As String = "CalculatedDate_"
Private Const SETTING_PREFIX As String = "Setting_"

' Any text value that needs to outlive the session, keyed the same way as the rest.
Public Property Get SettingText(key As String) As String
    SettingText = ReadTextName(SETTING_PREFIX & key)
End Property

Public Property Let SettingText(key As String, value As String)
    WriteTextName SETTING_PREFIX & key, value
End Property

Public Property Get SourceFilePath(key As String) As String
    SourceFilePath = ReadTextName(NAME_PREFIX & key)
End Property

Public Property Let SourceFilePath(key As String, value As String)
    If Len(value) = 0 Then
        ClearSourceFilePath key
        Exit Property
    End If

    WriteTextName NAME_PREFIX & key, value
End Property

Public Sub ClearSourceFilePath(key As String)
    On Error Resume Next
    ThisWorkbook.names(SafeName(NAME_PREFIX & key)).Delete
    ThisWorkbook.names(SafeName(SELECTED_PREFIX & key)).Delete
    ThisWorkbook.names(SafeName(IMPORTED_PREFIX & key)).Delete
    ThisWorkbook.names(SafeName(VALIDATED_PREFIX & key)).Delete
    ThisWorkbook.names(SafeName(CALCULATED_PREFIX & key)).Delete
    On Error GoTo 0
End Sub

' When the file for this key was last selected. Empty if never recorded.
Public Property Get SourceFileSelectedOn(key As String) As Variant
    SourceFileSelectedOn = ReadDateName(SELECTED_PREFIX & key)
End Property

Public Property Let SourceFileSelectedOn(key As String, value As Variant)
    WriteDateName SELECTED_PREFIX & key, value
End Property

' When data for this key was last imported successfully. Empty if never recorded.
Public Property Get SourceLastImportedOn(key As String) As Variant
    SourceLastImportedOn = ReadDateName(IMPORTED_PREFIX & key)
End Property

Public Property Let SourceLastImportedOn(key As String, value As Variant)
    WriteDateName IMPORTED_PREFIX & key, value
End Property

' When the data for this key last passed validation. Empty if never recorded.
Public Property Get LastValidatedOn(key As String) As Variant
    LastValidatedOn = ReadDateName(VALIDATED_PREFIX & key)
End Property

Public Property Let LastValidatedOn(key As String, value As Variant)
    WriteDateName VALIDATED_PREFIX & key, value
End Property

' When the data for this key was last recalculated. Empty if never recorded.
Public Property Get LastCalculatedOn(key As String) As Variant
    LastCalculatedOn = ReadDateName(CALCULATED_PREFIX & key)
End Property

Public Property Let LastCalculatedOn(key As String, value As Variant)
    WriteDateName CALCULATED_PREFIX & key, value
End Property

' Formatted timestamp for display, or a fallback when nothing has been recorded.
Public Function SourceFileSelectedOnText(key As String, _
                                         Optional dateFormat As String = "dd mmm yyyy hh:nn", _
                                         Optional ifNever As String = "Never") As String
    SourceFileSelectedOnText = StampText(SourceFileSelectedOn(key), dateFormat, ifNever)
End Function

Public Function SourceLastImportedOnText(key As String, _
                                         Optional dateFormat As String = "dd mmm yyyy hh:nn", _
                                         Optional ifNever As String = "Never") As String
    SourceLastImportedOnText = StampText(SourceLastImportedOn(key), dateFormat, ifNever)
End Function

Public Function LastValidatedOnText(key As String, _
                                    Optional dateFormat As String = "dd mmm yyyy hh:nn", _
                                    Optional ifNever As String = "Never") As String
    LastValidatedOnText = StampText(LastValidatedOn(key), dateFormat, ifNever)
End Function

Public Function LastCalculatedOnText(key As String, _
                                     Optional dateFormat As String = "dd mmm yyyy hh:nn", _
                                     Optional ifNever As String = "Never") As String
    LastCalculatedOnText = StampText(LastCalculatedOn(key), dateFormat, ifNever)
End Function

Private Function StampText(stamp As Variant, dateFormat As String, ifNever As String) As String
    If IsEmpty(stamp) Then
        StampText = ifNever
    Else
        StampText = Format$(stamp, dateFormat)
    End If
End Function

Private Function ReadTextName(fullName As String) As String
    Dim nm As Name
    Dim raw As String

    On Error Resume Next
    Set nm = ThisWorkbook.names(SafeName(fullName))
    On Error GoTo 0
    If nm Is Nothing Then Exit Function

    raw = nm.RefersTo
    ' Stored as a literal string name, e.g. ="C:\folder\file.xlsx"
    If Left$(raw, 2) = "=" & Chr$(34) And Right$(raw, 1) = Chr$(34) Then
        ReadTextName = Mid$(raw, 3, Len(raw) - 3)
    End If
End Function

Private Sub WriteTextName(fullName As String, value As String)
    ThisWorkbook.names.Add Name:=SafeName(fullName), _
                           RefersTo:="=" & Chr$(34) & value & Chr$(34), _
                           Visible:=False
End Sub

Private Function ReadDateName(fullName As String) As Variant
    Dim nm As Name
    Dim raw As String

    On Error Resume Next
    Set nm = ThisWorkbook.names(SafeName(fullName))
    On Error GoTo 0
    If nm Is Nothing Then Exit Function

    ' Stored as a numeric literal name holding a date serial, e.g. =45932.5
    raw = Mid$(nm.RefersTo, 2)
    If IsNumeric(raw) Then ReadDateName = CDate(CDbl(raw))
End Function

Private Sub WriteDateName(fullName As String, value As Variant)
    If IsEmpty(value) Then
        On Error Resume Next
        ThisWorkbook.names(SafeName(fullName)).Delete
        On Error GoTo 0
        Exit Sub
    End If

    ' Str$ forces a period decimal separator regardless of regional settings.
    ThisWorkbook.names.Add Name:=SafeName(fullName), _
                           RefersTo:="=" & Trim$(Str$(CDbl(CDate(value)))), _
                           Visible:=False
End Sub

' Keys now come from data as well as from code, so anything a defined name cannot
' contain is folded to an underscore before it reaches Names.Add.
Private Function SafeName(rawName As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    For i = 1 To Len(rawName)
        ch = Mid$(rawName, i, 1)
        If ch Like "[A-Za-z0-9_.]" Then
            out = out & ch
        Else
            out = out & "_"
        End If
    Next i

    SafeName = out
End Function

' True when a path has been stored for the key and the file still exists on disk.
Public Function SourceFileExists(key As String) As Boolean
    Dim path As String

    path = SourceFilePath(key)
    If Len(path) = 0 Then Exit Function
    SourceFileExists = (Len(Dir$(path)) > 0)
End Function

' Prompts for a file and stores it against the key. Returns "" if the user cancels.
Public Function SelectSourceFileFor(key As String, Optional title As String = "") As String
    Dim path As String

    If Len(title) = 0 Then title = "Select the source workbook"

    path = PickSourceFile(title)
    If Len(path) = 0 Then Exit Function

    SourceFilePath(key) = path
    SourceFileSelectedOn(key) = Now
    SelectSourceFileFor = path
End Function

Public Function PickSourceFile(Optional title As String = "Select the source workbook") As String
    Dim picked As Variant

    picked = Application.GetOpenFilename( _
                FileFilter:="Data files (*.xlsx;*.xlsm;*.xlsb;*.xls;*.csv),*.xlsx;*.xlsm;*.xlsb;*.xls;*.csv", _
                title:=title)

    ' GetOpenFilename returns False when the user cancels.
    If VarType(picked) = vbBoolean Then Exit Function

    PickSourceFile = CStr(picked)
End Function


