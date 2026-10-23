Attribute VB_Name = "Import_RFC"
Option Explicit

Private Function spec() As ImportSpec
    spec = NewImportSpec(key:="RFC", _
                         sourceSheet:="", _
                         sourceHeaderCell:="A1", _
                         targetSheet:="13_SRC_RFC", _
                         TargetTable:="tblRFC", _
                         statusCell:="B4")
End Function

Private Function tableSpec() As TableImportSpec
    tableSpec = NewTableImportSpec(key:="RFC", _
                         configSheet:="13_SRC_RFC", _
                         configTopRow:=14, _
                         startDateName:="dateRFCStartDate", _
                         endDateName:="dateRFCEndDate", _
                         targetSheet:="13_SRC_RFC", _
                         TargetTable:="tblRFC", _
                         statusCell:="B4")
End Function

Public Sub SelectRFCSource()
    SelectSourceFile spec
End Sub

Public Sub ImportTableRFC()
    RunTableImport tableSpec
End Sub

Public Sub RefreshRFCStatus()
    RefreshStatus spec
End Sub




