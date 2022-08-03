Attribute VB_Name = "t_Shapes"
Option Explicit

Public Function cellVal(ByRef shps As Variant, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber, Optional defaultValue As Variant = 0) As Variant
'Return shape cell value. If such cell does not exists return defaultValue
Dim shp As Visio.Shape
Dim tmpVal As Variant
    
    On Error GoTo ex
    
    If TypeName(shps) = "Shape" Then
        Set shp = shps
        If shp.CellExists(cellName, 0) Then
            Select Case dataType
                Case Is = visNumber
                    cellVal = shp.Cells(cellName).Result(dataType)
                Case Is = visUnitsString
                    cellVal = shp.Cells(cellName).ResultStr(dataType)
                Case Is = visDate
                    cellVal = shp.Cells(cellName).Result(dataType)
                    If cellVal = 0 Then
                        cellVal = CDate(shp.Cells(cellName).ResultStr(visUnitsString))
                    End If
                Case Else
                    cellVal = shp.Cells(cellName).Result(dataType)
            End Select
        Else
            cellVal = defaultValue
        End If
        Exit Function
    ElseIf TypeName(shps) = "Shapes" Or TypeName(shps) = "Collection" Then     'Если коллекция
        For Each shp In shps
            tmpVal = cellVal(shp, cellName, dataType, defaultValue)
            If tmpVal <> defaultValue Then
                cellVal = tmpVal
                Exit Function
            End If
        Next shp
    End If
    
cellVal = defaultValue
Exit Function
ex:
    cellVal = defaultValue
End Function

Public Function cellFrml(ByRef shps As Variant, ByVal cellName As String, Optional defaultValue As Variant = "") As Variant
'Return shape cell formula. If such cell does not exists return defaultValue
Dim shp As Visio.Shape
Dim tmpVal As Variant
    
    On Error GoTo ex
    
    If TypeName(shps) = "Shape" Then
        Set shp = shps
        If shp.CellExists(cellName, 0) Then
            cellFrml = shp.Cells(cellName).FormulaU
        Else
            cellFrml = defaultValue
        End If
        Exit Function
    End If
    
cellFrml = defaultValue
Exit Function
ex:
    cellFrml = defaultValue
End Function





Public Sub SetCellVal(ByRef shp As Visio.Shape, ByVal cellName As String, ByVal NewVal As Variant)
'Set cell with cellName value. If such cell does not exists, does nothing
Dim cll As Visio.Cell
    
    On Error GoTo ex
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        shp.Cells(cellName).FormulaForce = """" & NewVal & """"
    End If
    
Exit Sub
ex:
'    Debug.Print "Error in t_Shapes module in 'Otcheti'! " & shp.Name & ", " & cellName & ", " & NewVal
End Sub

Public Sub SetCellFrml(ByRef shp As Visio.Shape, ByVal cellName As String, ByVal NewFrml As Variant)
'Set cell with cellName formula. If such cell does not exists, does nothing
Dim cll As Visio.Cell
    
    On Error GoTo ex
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        shp.Cells(cellName).FormulaForceU = NewFrml
    End If
    
Exit Sub
ex:

End Sub

Public Function ShapeHaveCell(ByRef shp As Visio.Shape, ByVal cellName As String, _
                              Optional ByVal val As Variant = "", Optional ByVal delimiter As Variant = ";") As Boolean
'The function returns True if there is such a cell
'If a cell value is specified, it is also checked. If several values are specified, then all of them are checked
Dim vals() As String
Dim curval As String
Dim i As Integer

On Error GoTo ex

    If shp.CellExists(cellName, 0) Then
        If val <> "" Then
            ' Check if one value is passed in the val attribute
            If InStr(1, val, delimiter) > 0 Then
                '---Several values
                vals = Split(val, delimiter)
                For i = 0 To UBound(vals)
                    curval = vals(i)
                    If shp.Cells(cellName).ResultStr(visUnitsString) = curval Then
                        ShapeHaveCell = True
                        Exit Function
                    ElseIf shp.Cells(cellName).Result(visNumber) = curval Then
                        ShapeHaveCell = True
                        Exit Function
                    Else
                        ShapeHaveCell = False
                    End If
                Next i
            Else
                '---Single value
                If shp.Cells(cellName).ResultStr(visUnitsString) = val Then
                    ShapeHaveCell = True
                ElseIf shp.Cells(cellName).Result(visNumber) = val Then
                    ShapeHaveCell = True
                Else
                    ShapeHaveCell = False
                End If
            End If
        Else
            ShapeHaveCell = True
        End If
    Else
        ShapeHaveCell = False
    End If

Exit Function
ex:
    ShapeHaveCell = False
End Function

Public Function FirstSelectedShape() As Visio.Shape
'Get first shape in selection
Dim sel As Visio.Selection
    
    Set sel = Application.ActiveWindow.Selection
    
    If sel.Count = 0 Then
        Set FirstSelectedShape = Nothing
    ElseIf sel.Count = 1 Then
        Set FirstSelectedShape = sel(1)
    Else
        Debug.Print "Selected more then one shape, returned first from selection"
        Set FirstSelectedShape = sel(1)
    End If
End Function
