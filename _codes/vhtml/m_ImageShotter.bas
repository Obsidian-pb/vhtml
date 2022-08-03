Attribute VB_Name = "m_ImageShotter"
Option Explicit

'-----Modeule for images saving---------


Public Sub SaveAllFrames(Optional ByVal showMsg As Boolean = False)
Dim shp As Visio.Shape
Dim pg As Visio.Page
Dim imgPath As String
Dim shpcol As Collection
    
    'Switch layer shapes (for correctly saving)
    LayerPrint "Кадры", True
    
    For Each pg In Application.ActiveDocument.Pages
        Application.ActiveWindow.Page = pg
        For Each shp In pg.Shapes
            If ShapeHaveCell(shp, CELLNAME_VHTML_TYPE, vhtml_types.vt_Frame) Then
                imgPath = Application.ActiveDocument.path & cellVal(shp, "Prop.img_name", visUnitsString)
                SaveImageFromFrame shp, imgPath
                If showMsg Then ShowStatus "Image saved: " & imgPath
            End If
        Next shp
    Next pg
    
    If showMsg Then
        ShowStatus "..."
        ShowStatus "Все изображения успешно сохранены в папку документа"
    End If
    
    'Switch layer shapes (for correctly saving)
    LayerPrint "Кадры", False
End Sub


Public Sub SaveImageFromFrame(ByRef shp As Visio.Shape, Optional imgPath As String = "C:\tst\testimg.jpg")
'Save all images includeed by shape (in frame)
Dim bnd_mainShp() As Double
Dim bnd_frame() As Double

    SelectAllNstedShapes shp, shp.Parent.Index

    
Dim mainW As Double
Dim mainH As Double
Dim frameW As Double
Dim frameH As Double
Dim leftBound As Double
Dim bottomBound As Double
    
    'Get selection frame
    bnd_mainShp = GetShapeBounds(shp)
    bnd_frame = GetSelectionBounds(Application.ActiveWindow.Selection)
    
    frameW = bnd_frame(2) - bnd_frame(0)
    frameH = bnd_frame(3) - bnd_frame(1)
    mainW = bnd_mainShp(2) - bnd_mainShp(0)
    mainH = bnd_mainShp(3) - bnd_mainShp(1)
    leftBound = bnd_mainShp(0) - bnd_frame(0)
    bottomBound = bnd_mainShp(1) - bnd_frame(1)
    
    'Save image (twice because of hide some secondary elements in shapes)
    Application.ActiveWindow.Selection.Export imgPath
    Application.ActiveWindow.Selection.Export imgPath
    WIA_CropImage_Perc imgPath, Round(leftBound / frameW, 5), Round(bottomBound / frameH, 5), _
                                             Round((frameW - leftBound - mainW) / frameW, 5), _
                                             Round((frameH - bottomBound - mainH) / frameH, 5), _
                                             imgPath
    Application.ActiveWindow.DeselectAll
End Sub


Private Function GetSelectionBounds(ByRef slt As Visio.Selection) As Double()
Dim shp As Visio.Shape
Dim bnd(3) As Double    '0 - left, 1 - bottom, 2- right, 3 - top
Dim bnd_shp() As Double
Dim l, b, r, t As Double
Dim intFlags As Integer

    intFlags = visBBoxDrawingCoords
    Application.ActiveWindow.Selection.BoundingBox intFlags + visBBoxExtents, bnd(0), bnd(1), bnd(2), bnd(3)
    
GetSelectionBounds = bnd
End Function

Private Function GetShapeBounds(ByRef shp As Visio.Shape) As Double()
Dim bnd(3) As Double    '0 - left, 1 - bottom, 2- right, 3 - top
Dim intFlags As Integer

    intFlags = visBBoxDrawingCoords
    'Need review!
'    shp.BoundingBox intFlags + visBBoxExtents, bnd(0), bnd(1), bnd(2), bnd(3)
'    shp.BoundingBox intFlags + visBBoxExtents + visBBoxUprightWH, bnd(0), bnd(1), bnd(2), bnd(3)
    bnd(0) = cellVal(shp, "PinX") - cellVal(shp, "Width") / 2
    bnd(1) = cellVal(shp, "Piny") - cellVal(shp, "Height") / 2
    bnd(2) = cellVal(shp, "PinX") + cellVal(shp, "Width") / 2
    bnd(3) = cellVal(shp, "Piny") + cellVal(shp, "Height") / 2
    
    
GetShapeBounds = bnd
End Function

Public Sub SelectAllNstedShapes(ByRef shp As Visio.Shape, Optional ByVal pageIndex As Integer = 0)
Dim shp2 As Visio.Shape
Dim sel As Visio.Selection
Dim pg As Visio.Page
    
    Application.ActiveWindow.DeselectAll
    
    If pageIndex = 0 Then pageIndex = Application.ActivePage.Index
    Set pg = Application.ActiveDocument.Pages(pageIndex)
    
    For Each shp2 In pg.Shapes
        If shp.SpatialRelation(shp2, 0, VisSpatialRelationFlags.visSpatialFrontToBack) > 0 Then
            If Not cellVal(shp2, "User.IndexPers") = 502 Then
                Application.ActiveWindow.Select shp2, visSelect
            End If
        End If
    Next shp2
    Application.ActiveWindow.Select shp, visSelect
    
End Sub


