Attribute VB_Name = "Exports"
Option Explicit

Sub ExportOOR()
    Dim PrevDispAlert As Boolean
    Dim FilePath As String
    Dim FileName As String
    Dim FileExt As String

    PrevDispAlert = Application.DisplayAlerts
    FilePath = "\\7938-HP02\Shared\Doosan\Open Order Report\" & Format(Date, "yyyy") & "\" & Format(Date, "mmm") & "\"
    FileName = "OOR " & Format(Date, "yyyy-mm-dd")
    FileExt = ".xlsx"

    If Not FolderExists(FilePath) Then RecMkDir FilePath

    Application.DisplayAlerts = False
    Sheets("Open Order Report").Copy
    ActiveWorkbook.SaveAs FilePath & FileName & FileExt, xlOpenXMLWorkbook
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert

    'Email "abridges@wesco.com", _
          Subject:="IR Open Order Report", _
          CC:="SNelson@wesco.com", _
          Body:="An updated copy of the IR open order report can be found on the network <a href=""" & FilePath & FileName & FileExt & """" & ">here</a>."
End Sub
