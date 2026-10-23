Attribute VB_Name = "Import_Applications"
Option Explicit

' Applications import. Configuration only - all logic lives in the TableImport module.
' Assign SelectApplicationsSource and ImportApplications to buttons on 14_SRC_Applications.

Private Function spec() As ImportSpec
    spec = NewImportSpec(key:="Applications", _
                         sourceSheet:="", _
                         sourceHeaderCell:="A1", _
                         targetSheet:="14_SRC_Applications", _
                         TargetTable:="tblApplications", _
                         statusCell:="B4")
End Function

Public Sub SelectApplicationsSource()
    SelectSourceFile spec
End Sub

Public Sub ImportApplications()
    RunImport spec
End Sub

Public Sub RefreshApplicationsStatus()
    RefreshStatus spec
End Sub


