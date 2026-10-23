Attribute VB_Name = "DiscretionsSpec"
Option Explicit

' Configuration for validating and summarising a specific discretions table.
' Each spec bundles the inputs the generic entry points need: which table to read,
' how to label it in messages, where to store the "last validated" timestamp, and
' where to write the summary tables.
'
' To wire up another table, add a factory next to the existing ones and expose a
' button entry point that passes the spec into ValidateDiscretionsFor or
' ValidateAndSummarizeFor.

Public Type DiscretionsSpec
    tableName As String           ' ListObject name, must exist in this workbook
    displayName As String         ' Used in MsgBox and status messages
    stateKey As String            ' Key for AppState.LastValidatedOn
    summarySheet As String        ' Output sheet for the summary; empty => summary disabled
    summaryAnchor As String       ' Top-left cell for the summary tables, e.g. "W10"
    summaryTablePrefix As String  ' Name stem for generated output tables (prefix + ordinal)
End Type

'=======================================================================================
' Factories
'=======================================================================================

Public Function ActiveDiscretionsSpec() As DiscretionsSpec
    Dim s As DiscretionsSpec
    s.tableName = "tblActiveDiscretions"
    s.displayName = "Active discretions"
    ' Kept as "Discretions" so the existing stored timestamp carries over.
    s.stateKey = "Discretions"
    s.summarySheet = "05_Price_Setting_Inputs"
    s.summaryAnchor = "W10"
    ' Kept as "tblMetricsOut" so any formula pointing at the existing tables keeps working.
    s.summaryTablePrefix = "tblMetricsOut"
    ActiveDiscretionsSpec = s
End Function

Public Function CurrentDiscretionsSpec() As DiscretionsSpec
    Dim s As DiscretionsSpec
    s.tableName = "tblCurrentDiscretions"
    s.displayName = "Current discretions"
    s.stateKey = "DiscretionsCurrent"
    ' Fill these in with the sheet and anchor you want the Current summary written to.
    ' Leave empty to skip summary generation for Current.
    s.summarySheet = "08_Discretion_Framework"
    s.summaryAnchor = "W10"
    s.summaryTablePrefix = "tblMetricsOutCurrent"
    CurrentDiscretionsSpec = s
End Function

'=======================================================================================
' Button entry points
'=======================================================================================

Public Sub ValidateDiscretions_Active()
    ValidateDiscretionsFor ActiveDiscretionsSpec()
End Sub

Public Sub ValidateDiscretions_Current()
    ValidateDiscretionsFor CurrentDiscretionsSpec()
End Sub

Public Sub ValidateAndSummarize_Active()
    ValidateAndSummarizeFor ActiveDiscretionsSpec()
End Sub

Public Sub ValidateAndSummarize_Current()
    ValidateAndSummarizeFor CurrentDiscretionsSpec()
End Sub

Public Sub GenerateDiscretionSummary_Active()
    GenerateDiscretionSummaryFor ActiveDiscretionsSpec()
End Sub

Public Sub GenerateDiscretionSummary_Current()
    GenerateDiscretionSummaryFor CurrentDiscretionsSpec()
End Sub

'=======================================================================================
' Chained validate + summary
'=======================================================================================

' Validates the spec's table. If it comes back clean and a summarySheet is configured,
' also runs the summary generation.
Public Sub ValidateAndSummarizeFor(ByRef spec As DiscretionsSpec)
    Dim issues As Collection

    Set issues = RunDiscretionChecks(spec)
    If issues Is Nothing Then Exit Sub
    If issues.count > 0 Then Exit Sub
    If Len(spec.summarySheet) = 0 Then Exit Sub

    GenerateDiscretionSummaryFor spec
End Sub
