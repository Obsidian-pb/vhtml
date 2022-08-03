Attribute VB_Name = "m_toolbar"
Option Explicit

Private btns As c_Buttons


Sub AddTB()
Dim i As Integer
Dim Bar As CommandBar, Button As CommandBarButton
    
'---Check vhtml toolbar existance------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "vhtml" Then Exit Sub
    Next i

'---Create vhtml toolbar--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "vhtml"
        .Visible = True
    End With
    
'---Add buttons
    AddButtons
End Sub

Sub RemoveTB()
'Remove vhtml toolbar-------------------------------
    On Error Resume Next
    Application.CommandBars("vhtml").Delete
End Sub

Sub AddButtons()
Dim Bar As CommandBar
Dim DocPath As String
    
    On Error GoTo ex
    
    Set Bar = Application.CommandBars("vhtml")
    
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Redact content"
        .Tag = "Redact content"
        .TooltipText = "Redact vhtml shape content"
        .FaceID = 593
        .BeginGroup = True
    End With
    
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Redact attributes"
        .Tag = "Redact attributes"
        .TooltipText = "Redact vhtml shape attributes"
        .FaceID = 1607
        .BeginGroup = False
    End With
    
'---Preview group
'    With Bar.Controls.Add(Type:=msoControlButton)
'        .Caption = "HTML preview selected"
'        .Tag = "HTML preview selected"
'        .TooltipText = "Show HTML preview for selected shape"
'        .FaceID = 558
'        .BeginGroup = True
'    End With

    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "HTML preview all"
        .Tag = "HTML preview all"
        .TooltipText = "Show HTML preview for all shapes on page"
        .FaceID = 109
        .BeginGroup = True
    End With
       
'---Tools group
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Save frames"
        .Tag = "Save frames"
        .TooltipText = "Save frames in doc"
        .FaceID = 1362
        .BeginGroup = True
    End With

    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Render"
        .Tag = "Render"
        .TooltipText = "Render html app"
        .FaceID = 1556
    End With
    
    Set btns = New c_Buttons
    Set Bar = Nothing

Exit Sub
ex:
    Set Bar = Nothing
End Sub


Sub DeleteButtons()
Dim Bar As CommandBar, Button As CommandBarButton
    On Error GoTo ex

    Set Bar = Application.CommandBars("Формулы")
    For Each Button In Bar.Controls
        Button.Delete
    Next Button

Set btns = Nothing
Set Button = Nothing
Set Bar = Nothing

Exit Sub
ex:
    Set Button = Nothing
    Set Bar = Nothing
End Sub

