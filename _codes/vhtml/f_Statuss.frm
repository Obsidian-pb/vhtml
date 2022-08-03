VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_Statuss 
   Caption         =   "Процесс выполнения"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7530
   OleObjectBlob   =   "f_Statuss.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "f_Statuss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub btn_Ok_Click()
    Me.lbl_StatussText.Caption = ""
    Me.Hide
End Sub

Public Sub ShowChanges(ByVal msgtxt As String)
    On Error GoTo ex
    Me.Show

    Me.lbl_StatussText.Caption = Me.lbl_StatussText.Caption & msgtxt & Chr(10)
    Me.Repaint
Exit Sub
ex:
    MsgBox msgtxt
End Sub



Private Sub UserForm_Initialize()
    Me.lbl_StatussText.Caption = ""
End Sub
