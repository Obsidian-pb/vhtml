VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public WithEvents btn_RedactContent As Office.CommandBarButton          'Redact content
Attribute btn_RedactContent.VB_VarHelpID = -1
Public WithEvents btn_RedactAttrs As Office.CommandBarButton            'Redact attributes
Attribute btn_RedactAttrs.VB_VarHelpID = -1

'Public WithEvents btn_PreviewSelected As Office.CommandBarButton        'HTML preview selected
Public WithEvents btn_PreviewAllOnPage As Office.CommandBarButton       'HTML preview all
Attribute btn_PreviewAllOnPage.VB_VarHelpID = -1

Public WithEvents btn_SaveFrames As Office.CommandBarButton             'Save frames
Attribute btn_SaveFrames.VB_VarHelpID = -1
Public WithEvents btn_Render As Office.CommandBarButton                 'Render
Attribute btn_Render.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set btn_RedactContent = Application.CommandBars("vhtml").Controls("Redact content")
    Set btn_RedactAttrs = Application.CommandBars("vhtml").Controls("Redact attributes")
    
'    Set btn_PreviewSelected = Application.CommandBars("vhtml").Controls("HTML preview selected")
    Set btn_PreviewAllOnPage = Application.CommandBars("vhtml").Controls("HTML preview all")
    
    Set btn_SaveFrames = Application.CommandBars("vhtml").Controls("Save frames")
    Set btn_Render = Application.CommandBars("vhtml").Controls("Render")
End Sub

Private Sub Class_Terminate()
    Set btn_RedactContent = Nothing
    Set btn_RedactAttrs = Nothing
'    Set btn_PreviewSelected = Nothing
    Set btn_PreviewAllOnPage = Nothing
    Set btn_SaveFrames = Nothing
    Set btn_Render = Nothing
End Sub



Private Sub btn_RedactContent_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
Dim shp As Visio.Shape
    Set shp = FirstSelectedShape
    
    If Not shp Is Nothing Then
        If ShapeHaveCell(shp, CELLNAME_VHTML_TYPE, "0;1;3") Then ShowContentEditForm shp
    Else
        MsgBox "Please select single shape or use popup commands for nested shapes"
    End If
End Sub

Private Sub btn_RedactAttrs_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
Dim shp As Visio.Shape
    Set shp = FirstSelectedShape
    
    If Not shp Is Nothing Then
        If ShapeHaveCell(shp, CELLNAME_VHTML_TYPE, "1;2") Then ShowAttrsEditForm shp
    Else
        MsgBox "Please select single shape or use popup commands for nested shapes"
    End If
End Sub



'Private Sub btn_PreviewSelected_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Dim shp As Visio.Shape
'    Set shp = FirstSelectedShape
'
'    If Not shp Is Nothing Then
'        ShowPreviewForm shp, False
'    Else
'        MsgBox "Please select single shape"
'    End If
'End Sub

Private Sub btn_PreviewAllOnPage_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    ShowPreviewFormForAllInPage False
End Sub





Private Sub btn_SaveFrames_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    SaveAllFrames True
End Sub

Private Sub btn_Render_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    SaveAllPagesCode True
End Sub
