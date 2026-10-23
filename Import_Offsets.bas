Attribute VB_Name = "Import_Offsets"
Option Explicit

' Offsets import. Configuration only - all logic lives in the TableImport module.
' Assign SelectOffsetsSource and ImportOffsets to buttons on 15_SRC_Offsets.

Private Function spec() As ImportSpec
    spec = NewImportSpec(key:="Offsets", _
                         sourceSheet:="", _
                         sourceHeaderCell:="A1", _
                         targetSheet:="15_SRC_Offsets", _
                         TargetTable:="tblOffsets", _
                         statusCell:="B4")
End Function

Public Sub SelectOffsetsSource()
    SelectSourceFile spec
End Sub

Public Sub ImportOffsets()
    RunImport spec
End Sub

Public Sub RefreshOffsetsStatus()
    RefreshStatus spec
End Sub




