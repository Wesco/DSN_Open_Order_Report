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

'---------------------------------------------------------------------------------------
' Proc : ExportDSNReport
' Date : 1/24/2014
' Desc : Exports DSN Report to the network and emails a copy to Doosan
'---------------------------------------------------------------------------------------
Sub ExportDSNReport()
    Dim PrevDispAlert As Boolean
    Dim FileName As String
    Dim FilePath As String
    Dim EmailTo As String

    PrevDispAlert = Application.DisplayAlerts
    FilePath = "\\7938-HP02\Shared\Doosan\Open Order Report\" & Year(Date) & "\" & Format(Date, "mmm") & "\"

    If OORType = "aftermarket" Then
        FileName = "DSN Report " & Format(Date, "yyyy-mm-dd") & ".xlsx"
        EmailTo = "claude.tutterow@doosan.com"
    ElseIf OORType = "production" Then
        FileName = "DSN Report " & Format(Date, "yyyy-mm-dd") & ".xlsx"
        EmailTo = "dione.guy@doosan.com"
    Else
        Err.Raise CustErr.UNRECOGNIZED_REPORT, "ExportDSNReport", "The report type was not recognized."
    End If

    Application.DisplayAlerts = False
    Sheets("DSN Report").Copy
    ActiveWorkbook.SaveAs FilePath & FileName, xlOpenXMLWorkbook
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert

    Email SendTo:=EmailTo, _
          Subject:="Open Order Report", _
          Body:="Attached is an updated copy of the open order report.", _
          Attachment:=FilePath & FileName
End Sub
