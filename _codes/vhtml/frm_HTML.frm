VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_HTML 
   Caption         =   "HTML preview"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9225
   OleObjectBlob   =   "frm_HTML.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_HTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private tmp_path As String

Private Sub UserForm_Activate()
Dim width As Integer
Dim height As Integer
Dim top As Integer
Dim left As Integer

    'Set start consts
    top = 6
    left = 6
    width = Me.InsideWidth - left * 2
    height = 400
    
    wb_Bowser.width = width
    wb_Bowser.height = height
    wb_Bowser.top = top
    wb_Bowser.left = left
    
    cb_Cancel.top = top + height + 6
    
    Me.height = cb_Cancel.top + cb_Cancel.height + 30
    
End Sub


Public Sub ShowHTMLPreview(ByVal html As String)
    
    'Save temporary html page
    tmp_path = Application.ActiveDocument.path & "_.html"
    SaveTextToFile html, tmp_path
    
    'Navigate temporary html page
    wb_Bowser.Navigate tmp_path
    
    'Show form
    Me.Show
End Sub


Private Sub cb_Cancel_Click()
    Me.Hide
    On Error Resume Next
    Kill tmp_path
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    Kill tmp_path
End Sub
