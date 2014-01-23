Attribute VB_Name = "Exports"
Option Explicit

Sub ExportOOR()
    Dim PrevDispAlert As Boolean
    Dim FilePath As String
    Dim FileName As String
    Dim FileExt As String
    Dim EmailSubj As String

    PrevDispAlert = Application.DisplayAlerts
    FilePath = "\\7938-HP02\Shared\Doosan\Open Order Report\" & Format(Date, "yyyy") & "\" & Format(Date, "mmm") & "\"
    FileExt = ".xlsx"

    If OORType = "aftermarket" Then
        FileName = "Aftermarket OOR " & Format(Date, "yyyy-mm-dd")
        EmailSubj = "DSN Open Order Report (Aftermarket)"
    ElseIf OORType = "production" Then
        FileName = "Production OOR " & Format(Date, "yyyy-mm-dd")
        EmailSubj = "DSN Open Order Report (Production)"
    Else
        Err.Raise CustErr.UNRECOGNIZED_REPORT, "ExportOOR", "The report type could not be verified."
    End If

    If Not FolderExists(FilePath) Then RecMkDir FilePath

    Application.DisplayAlerts = False
    Sheets("Open Order Report").Copy
    ActiveSheet.UsedRange.Columns.AutoFit
    ActiveWorkbook.SaveAs FilePath & FileName & FileExt, xlOpenXMLWorkbook
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert

    Email SendTo:="bhunter@wesco.com; snelson@wesco.com; tyklein@wesco.com", _
          Subject:=EmailSubj, _
          Body:="An updated copy of the Doosan " & OORType & " open order report can be found on the network <a href=""" & FilePath & FileName & FileExt & """" & ">here</a>."
End Sub
