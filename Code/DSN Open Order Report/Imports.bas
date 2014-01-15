Attribute VB_Name = "Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ImportMaster
' Date : 1/15/2014
' Desc : Imports the Doosan master list from the network
'---------------------------------------------------------------------------------------
Sub ImportMaster()
    Dim FilePath As String
    Dim FileName As String
    Dim PrevDispAlert As Boolean
    
    FilePath = "\\br3615gaps\gaps\Doosan\Master\"
    FileName = "Doosan Master " & Format(Date, "yyyy") & ".xlsx"
    PrevDispAlert = Application.DisplayAlerts
    
    If FileExists(FilePath & FileName) Then
        Workbooks.Open FilePath & FileName
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = PrevDispAlert
    Else
        Err.Raise Errors.FILE_NOT_FOUND, "ImportMaster", "The Doosan master could not be found"
    End If
End Sub
