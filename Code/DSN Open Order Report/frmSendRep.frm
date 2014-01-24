VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSendRep 
   Caption         =   "Send Report"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3135
   OleObjectBlob   =   "frmSendRep.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSendRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSend_Click()
    If radAftermarket Then
        OORType = "aftermarket"
        Unload Me
    ElseIf radProduction Then
        OORType = "production"
        Unload Me
    Else
        MsgBox "A report was not selected.", vbOKOnly, "Please select a report"
    End If
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    radAftermarket = True
End Sub
