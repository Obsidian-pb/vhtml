Attribute VB_Name = "m_HTML"
Option Explicit




Public Function GetHTML(ByRef shp As Visio.Shape, Optional intent As Integer = 0) As String
'Get html of single shape and all its children
Dim shp_pattern As String
Dim shp_child As Visio.Shape
Dim content As String
Dim shpCollection As Collection
    
    'check for tag shape
    If Not ShapeHaveCell(shp, CELLNAME_VHTML_TYPE, "1;2;3") Then
'        Debug.Print shp.Name
        GetHTML = ""
        Exit Function
    End If
    
    'get this shape html pattern
    shp_pattern = cellVal(shp, CELLNAME_HTML_PATTERN, visUnitsString, "")
    
    'if this shape html pattern = "" exit function with return "". It means this shape is not vhtml shape
    If shp_pattern = "" Then
        GetHTML = ""
        Exit Function
    End If
    
    'get this shape content
    content = cellVal(shp, CELLNAME_CONTENT, visUnitsString, "")
    
    'if there are none content in this shape check all child shapes. In other case use this shape content
    If content = "" Then
        If shp.Shapes.Count > 0 Then
            'sort child shapes up to down and left to right
            Set shpCollection = FormUniqueCollection(shp.Shapes)
            Set shpCollection = Sort(shpCollection, "PinY")
            
            'fill child shapes content
            For Each shp_child In shpCollection
                'if child shape has not marker name and
                If cellVal(shp_child, CELLNAME_MARKER, visUnitsString, "*") = "" Then
                    content = content & vbNewLine & SP(intent) & GetHTML(shp_child, intent + 4) & vbNewLine & SP(intent - 4)
                End If
                
'                content = content & GetHTML(shp_child)
            Next shp_child
        End If
    End If
    
    GetHTML = Replace(shp_pattern, MARKER_CONTENT, content)
End Function


Public Function GetPageTemplate(Optional ByVal pageIndex As Integer = -1) As String
    
Dim t_shp As Visio.Shape
Dim pg As Visio.Page

    If pageIndex = -1 Then
        pageIndex = Application.ActivePage.Index
    End If
    
    'Search local template (for current page)
    Set t_shp = GetTemplateOnPage(pageIndex)
    If Not t_shp Is Nothing Then
        GetPageTemplate = cellVal(t_shp, CELLNAME_CONTENT, visUnitsString, EMPTY_TEMPLATE_PATTERN)
        Exit Function
    End If
    
    'If page does not have any local template...
    For Each pg In Application.ActiveDocument.Pages
        Set t_shp = GetTemplateOnPage(pg.Index)
        If Not t_shp Is Nothing Then
            If cellVal(t_shp, CELLNAME_GLOBAL_TEMPLATE) = 1 Then
                GetPageTemplate = cellVal(t_shp, CELLNAME_CONTENT, visUnitsString, EMPTY_TEMPLATE_PATTERN)
                Exit Function
            End If
        End If
    Next pg
    
    'In other case return empty template
    GetPageTemplate = EMPTY_TEMPLATE_PATTERN
    
End Function

Public Function GetTemplateOnPage(ByVal pageIndex As Integer) As Visio.Shape
'Get any template shape on page
Dim shp As Visio.Shape
   
    'Search local template (for current page)
    For Each shp In Application.ActiveDocument.Pages.item(pageIndex).Shapes
        If ShapeHaveCell(shp, CELLNAME_VHTML_TYPE, vhtml_types.vt_Template) Then
            Set GetTemplateOnPage = shp
            Exit Function
        End If
    Next shp
    
Set GetTemplateOnPage = Nothing
End Function
