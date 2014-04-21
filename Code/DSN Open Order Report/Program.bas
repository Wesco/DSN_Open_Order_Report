Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber = "1.1.3"
Public Const RepositoryName = "DSN_Open_Order_Report"
Public OORType As String    'This will be either aftermarket or production

Sub Main()
    On Error GoTo MAIN_ERR
    Application.ScreenUpdating = False
    UserImportFile DestRange:=Sheets("DSN OOR").Range("A1"), _
                   FileFilter:="XLSX Files (*.xlsx),*.xlsx,XLS Files (*.xls),*.xls,All Files (*.*),*.*", _
                   Title:="Select a Doosan open order report"

    'Determine if the report is aftermarket or production
    If Sheets("DSN OOR").Range("B1").Value = "PPZ Service Parts Inventory Org" Then
        OORType = "aftermarket"
    ElseIf Sheets("DSN OOR").Range("B1").Value = "STA Inventory Org" Then
        OORType = "production"
    Else
        Err.Raise CustErr.UNRECOGNIZED_REPORT, "Main", "The imported report is unrecognized."
    End If

    ImportMaster
    Import117
    ImportPrevOOR
    FormatDSNOOR
    Format117
    CreateOOR
    FormatReport Sheets("Open Order Report")
    ExportOOR
    Clean
    Application.ScreenUpdating = True
    MsgBox "Complete!"
    On Error GoTo 0
    Exit Sub

MAIN_ERR:
    Select Case Err.Number
        Case Errors.FILE_NOT_FOUND, CustErr.INVALID_COLUMN_ORDER, CustErr.UNRECOGNIZED_REPORT:
            MsgBox Err.Description, vbOKOnly, "Oops! Error " & Err.Number & " has occured"

        Case Errors.USER_INTERRUPT:

        Case Else
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") occured in " & Err.Source, vbOKOnly, "Oops! An error has occured"
    End Select
    Clean
End Sub

'---------------------------------------------------------------------------------------
' Proc : SendReport
' Date : 1/24/2014
' Desc : Makes a filtered copy of the report and emails it to doosan
'---------------------------------------------------------------------------------------
Sub SendReport()
    On Error GoTo SEND_ERR
    
    frmSendRep.Show
    If OORType = "" Then
        Clean
        MsgBox "Macro canceled!", vbOKOnly, "Canceled"
        Exit Sub
    End If
    
    ImportPrevOOR
    CreateDSNReport
    FormatReport Sheets("DSN Report")
    FormatDSNReport
    ExportDSNReport
    Clean
    MsgBox "Complete!"
    On Error GoTo 0
    Exit Sub

SEND_ERR:
    Select Case Err.Number
        Case CustErr.UNRECOGNIZED_REPORT
            MsgBox Err.Description, vbOKOnly, "Oops! Error " & Err.Number & " has occured"
        Case Else
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") occured in " & Err.Source, vbOKOnly, "Oops! An error has occured"
    End Select
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
