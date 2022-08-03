Attribute VB_Name = "m_Connections"
Option Explicit



Public Sub TurnIntoFormulaConnection(ByRef Connects As IVConnects)
'Procedure for accessing the formula connector (when connecting)
Dim cnct As Visio.Connect
Dim shp As Visio.Shape
Dim rowI As Integer
    
    On Error GoTo EndSub
    
    'Defining the connector shape
    Set shp = Connects(1).FromSheet
    
    '---check if the shape is not a connector shape already
    If ShapeHaveCell(shp, CELLNAME_VHTML_LINK) Then Exit Sub
    
    '---we check whether the figure of the two has connection points
    If shp.Connects.Count <> 2 Then Exit Sub

    '---Check if the shape is a line
    If shp.AreaIU > 0 Then Exit Sub
    
    '---We check whether both connected figures are figures of formulas
    If ShapeHaveCell(shp.Connects(1).ToSheet, CELLNAME_VHTML_TYPE) And ShapeHaveCell(shp.Connects(2).ToSheet, CELLNAME_VHTML_TYPE) Then
        '---Main procedure
        f_LinkShapes.showForm shp.Connects(1).ToSheet, shp.Connects(2).ToSheet, shp
        
        'Add connector props
        shp.AddNamedRow visSectionUser, Split(CELLNAME_VHTML_LINK, ".")(1), visTagDefault
        SetCellVal shp, CELLNAME_VHTML_LINK, True
        SetCellVal shp, "Char.Size", "6pt"
        SetCellVal shp, "Char.Style", "34"
    End If
    
Exit Sub
EndSub:
    Debug.Print Err
'    SaveLog Err, "TurnIntoFormulaConnection"
End Sub
