VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Environ("username") <> "TReische" Then
        Cancel = True
    End If
End Sub

Private Sub Workbook_Open()
    On Error GoTo UPDATE_ERROR
    CheckForUpdates RepositoryName, VersionNumber
    On Error GoTo 0
    Exit Sub

UPDATE_ERROR:
    If MsgBox("An error occured while checking for updates." & vbCrLf & vbCrLf & _
              "Would you like to open the website to download the latest version?", vbYesNo) = vbYes Then

        OpenInBrowser "http://github.com/Wesco/" & RepositoryName & "/releases/"

        If Workbooks.Count = 1 Then
            Application.Quit
        Else
            ThisWorkbook.Close
        End If
    End If
End Sub
