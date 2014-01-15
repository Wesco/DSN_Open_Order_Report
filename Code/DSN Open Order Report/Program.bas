Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber = "1.0.0"
Public Const RepositoryName = "DSN_Open_Order_Report"

Sub Main()

End Sub

Sub Clean()
    Dim s As Worksheet

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            Cells.Delete
            Range("A1").Select
        End If
    Next
    
    Sheets("Macro").Select
    Range("C7").Select
End Sub
