Attribute VB_Name = "m_Edit"
Option Explicit




Public Sub ShowContentEditForm(ByRef shp As Visio.Shape)
    frm_Editor.ShowEditor shp, EDIT_CONTENT
End Sub


Public Sub ShowAttrsEditForm(ByRef shp As Visio.Shape)
    frm_Editor.ShowEditor shp, EDIT_ATTRS
End Sub
