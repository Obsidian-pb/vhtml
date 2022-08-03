VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_LinkShapes 
   Caption         =   "Connected cells selection"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10830
   OleObjectBlob   =   "f_LinkShapes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_LinkShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private shpTo As Visio.Shape
Private shpFrom As Visio.Shape
Private shpConn As Visio.Shape


Public Sub showForm(ByRef shp1 As Visio.Shape, ByRef shp2 As Visio.Shape, ByRef a_shpConn As Visio.Shape)
Dim i As Integer
    
    'Save shapes vars
    Set shpFrom = shp1
    Set shpTo = shp2
    Set shpConn = a_shpConn
    
    'Fill start cells list
    lb_ShapeFrom.Clear
    For i = 0 To shp1.RowCount(visSectionProp) - 1
        lb_ShapeFrom.AddItem shp1.CellsSRC(visSectionProp, i, visTagDefault).RowNameU
        lb_ShapeFrom.List(i, 1) = shp1.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(visUnitsString)
    Next i

    'Fill end cells list
    lb_ShapeTo.Clear
    For i = 0 To shp2.RowCount(visSectionProp) - 1
        lb_ShapeTo.AddItem shp2.CellsSRC(visSectionProp, i, visTagDefault).RowNameU
        lb_ShapeTo.List(i, 1) = shp2.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(visUnitsString)
    Next i
    
    Me.Show
End Sub

Private Sub cb_Cancel_Click()
    Me.Hide
End Sub

Private Sub cb_Save_Click()
    'Set connections
    ConnectShapes
    
    'Close form
    Me.Hide
End Sub

Private Sub ConnectShapes()
Dim cellFromName As String
Dim cellToName As String
Dim i As Integer
Dim frml As String
    
    'We define the name of the cell from which the value will be obtained
    For i = 0 To Me.lb_ShapeFrom.ListCount - 1
        If Me.lb_ShapeFrom.Selected(i) Then
            cellFromName = Me.lb_ShapeFrom.List(i, 0)
            Exit For
        End If
    Next i
    
    'We define the name of the cell to which the value will be passed
    For i = 0 To Me.lb_ShapeTo.ListCount - 1
        If Me.lb_ShapeTo.Selected(i) Then
            cellToName = Me.lb_ShapeTo.List(i, 0)
            Exit For
        End If
    Next i
      
    'Link cells
    frml = "Sheet." & shpFrom.ID & "!Prop." & cellFromName
    '---in case the user wants to create a new cell
    If cellToName = "Новая" Then
        cellToName = cellToName & "_" & shpFrom.ID
        shpTo.AddNamedRow visSectionProp, cellToName, visTagDefault
    End If
    
    'Adding a cell with a variable to the connector
    shpConn.AddNamedRow visSectionProp, cellFromName, visTagDefault
    shpConn.Cells("Prop." & cellFromName).FormulaU = frml
    'We show the connection parameter in the connector
    frml = Chr(34) & cellFromName & "=>" & cellToName & ": " & Chr(34) & "&" & frml
    shpConn.Characters.AddCustomFieldU frml, visFmtNumGenNoUnits
    
    'Adding a link to the cell with the variable in the connector to the final shape
    frml = "Sheet." & shpConn.ID & "!Prop." & cellFromName
    shpTo.Cells("Prop." & cellToName).FormulaU = frml
    
    
End Sub
