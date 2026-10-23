Attribute VB_Name = "Import_Canstar"
Option Explicit

' Applications import. Configuration only - all logic lives in the TableImport module.
' Assign SelectApplicationsSource and ImportApplications to buttons on 14_SRC_Applications.

Private Function spec() As ImportSpec
    spec = NewImportSpec(key:="Canstar", _
                         sourceSheet:="", _
                         sourceHeaderCell:="A1", _
                         targetSheet:="17_SRC_CompetitorRates", _
                         TargetTable:="tblCanstar", _
                         statusCell:="B4")
End Function

Public Sub SelectCanstarSource()
    SelectSourceFile spec
End Sub

Public Sub ImportCanstar()
    RunImport spec
End Sub

Public Sub RefreshCanstarStatus()
    RefreshStatus spec
End Sub




