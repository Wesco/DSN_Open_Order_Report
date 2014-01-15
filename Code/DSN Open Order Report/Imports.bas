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
        Err.Raise Errors.FILE_NOT_FOUND, "ImportMaster", "The Doosan master could not be found."
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : Import117
' Date : 1/15/2014
' Desc : Imports the most recent 117 report that can be found
'---------------------------------------------------------------------------------------
Sub Import117()
    Dim PrevDispAlert As Boolean
    Dim FileName As String
    Dim FilePath As String
    Dim dt As Date
    Dim i As Long

    FilePath = "\\br3615gaps\gaps\3615 117 Report\DETAIL\ByOutsideSalesperson\1\"

    'Look back up to 30 days for the 117 open order report
    For i = 0 To 30
        dt = Date - i
        FileName = "3615 " & Format(dt, "yyyy-mm-dd") & " ALLORDERS.xlsx"
        If FileExists(FilePath & FileName) Then
            Exit For
        End If
    Next

    'If the 117 open order report was found, import it
    If FileExists(FilePath & FileName) Then
        PrevDispAlert = Application.DisplayAlerts
        Application.DisplayAlerts = False
        Workbooks.Open FilePath & FileName
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("117").Range("A1")
        ActiveWorkbook.Close
        Application.DisplayAlerts = PrevDispAlert
    Else
        Err.Raise Errors.FILE_NOT_FOUND, "Import117", "The 117 report could not be found."
    End If
End Sub
