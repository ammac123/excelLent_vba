Option Explicit

Private Function spec() As ImportSpec
    spec = NewImportSpec(key:="Data1", _
                        sourceSheet:="", _
                        sourceHeaderCell:="", _
                        targetSheet:="10_SRC_Data", _
                        TargetTable:="tblData", _
                        statusCell:="B4")
End Function

Private Function tableSpec() As TableImportSpec
    tableSpec = NewTableImportSpec(key:="Data1", _
                                   configSheet:="10_SRC_Data", _
                                   startDateName:="StartMonthEnding", _
                                   endDateName:="EndMonthEnding", _
                                   targetSheet:="10_SRC_Data", _
                                   TargetTable:="tblData", _
                                   statusCell:="B4")
End Function

Public Sub SelectDataSource()
    SelectSourceFile spec
End Sub

Public Sub ImportTableData()
    RunTableImport tableSpec
End Sub
