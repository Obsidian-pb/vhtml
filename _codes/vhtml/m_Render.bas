Attribute VB_Name = "m_Render"
Option Explicit

#If VBA7 Then
    #If Win64 Then
        ' Code is running in 64-bit VBA7.
        Private Declare PtrSafe Function ShellExecute _
            Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hwnd As LongPtr , _
            ByVal Operation As String, _
            ByVal filename As String, _
            Optional ByVal Parameters As String, _
            Optional ByVal Directory As String, _
            Optional ByVal WindowStyle As LongPtr  = vbMinimizedFocus _
            ) As LongPtr
    #Else
        ' Code is not running in 64-bit VBA7.
        Private Declare Function ShellExecute _
            Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hwnd As Long, _
            ByVal Operation As String, _
            ByVal filename As String, _
            Optional ByVal Parameters As String, _
            Optional ByVal Directory As String, _
            Optional ByVal WindowStyle As Long = vbMinimizedFocus _
            ) As Long
    #End If
#Else
    #If Win64 Then
        ' Code is running in 64-bit VBA7.
        Private Declare Function ShellExecute _
            Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hwnd As LongLong, _
            ByVal Operation As String, _
            ByVal filename As String, _
            Optional ByVal Parameters As String, _
            Optional ByVal Directory As String, _
            Optional ByVal WindowStyle As LongLong = vbMinimizedFocus _
            ) As LongLong
    #Else
        ' Code is not running in 64-bit VBA7.
        Private Declare Function ShellExecute _
            Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hwnd As Long, _
            ByVal Operation As String, _
            ByVal filename As String, _
            Optional ByVal Parameters As String, _
            Optional ByVal Directory As String, _
            Optional ByVal WindowStyle As Long = vbMinimizedFocus _
            ) As Long
    #End If
#End If

Public tmp_path As String



Public Function RenderPage(ByVal pg As Visio.Page) As String
'Render page code
Dim docPattern As String
Dim content As String
Dim shpCollection As Collection
Dim shp As Visio.Shape

Dim marker As Variant
Dim markersColl As Collection
Dim markeredCodeColl As Collection
    
    'get doc pattern from template
    docPattern = GetPageTemplate
    
    'initialize markers collections
    Set markersColl = New Collection
    Set markeredCodeColl = New Collection
    
    'sort child shapes up to down and left to right
    Set shpCollection = FormUniqueCollection(pg.Shapes)
    Set shpCollection = Sort(shpCollection, "PinY")
    
    'fill child shapes content
    content = ""
    For Each shp In shpCollection
        'if shp is vhtml shape with content
        If ShapeHaveCell(shp, CELLNAME_VHTML_TYPE, "1;2;3") Then
            'if child shape has not marker name and
            marker = cellVal(shp, CELLNAME_MARKER, visUnitsString, "*")
            If marker = "" Then
                content = content & GetHTML(shp) & vbNewLine
            Else
                'Save to markered fragments collection
                markersColl.Add marker
                AddUniqueCollectionItem markeredCodeColl, CStr(GetHTML(shp)), marker
            End If
        End If
    Next shp
    
    'If there is no code on the page, do not save it
    If content = "" Then
        RenderPage = ""
        Exit Function
    End If
    
    'Insert general content
    content = Replace(GetPageTemplate, EMPTY_TEMPLATE_PATTERN, content)
    content = Replace(content, EMPTY_TEMPLATE_TITLE, Split(pg.Name, ".")(0))
    
    'Insert markered fragments
    For Each marker In markersColl
        content = Replace(content, "$" & marker & "$", markeredCodeColl.item(marker))
    Next marker
    
RenderPage = content
End Function


Public Sub SaveAllPagesCode(Optional ByVal showMsg As Boolean = False)
Dim pg As Visio.Page
    
    For Each pg In Application.ActiveDocument.Pages
        SavePageCode pg.Name, showMsg
    Next pg
End Sub


Public Sub SavePageCode(ByVal pageName As String, Optional ByVal showMsg As Boolean = False)
'Save page code to file with name = pageName
Dim pg As Visio.Page
Dim content As String
Dim path As String

    Set pg = Application.ActiveDocument.Pages(pageName)

    'Render active page code
    content = RenderPage(pg)
    
    If content = "" Then Exit Sub
    
    'Save page code
    path = Application.ActiveDocument.path & pageName
    SaveTextToFile content, path
    
    If showMsg Then ShowStatus "Page saved: " & path
End Sub





Public Sub ShowPreviewFormForAllInPage(Optional ByVal inLocalBrowser As Boolean = True)
'Show preview for all vhtml shapes on page
Dim content As String

    'Render active page code
    content = RenderPage(Application.ActivePage)
    
    'Show content in VBA browser or in main webBrowser
    If inLocalBrowser Then
        frm_HTML.ShowHTMLPreview content
    Else
        ShowHTMLInBrowser content
    End If
End Sub


Public Sub ShowPreviewForm(ByRef shp As Visio.Shape, Optional ByVal inLocalBrowser As Boolean = True)
'Show previw for single shape
    'Show content in VBA browser or in main webBrowser
    If inLocalBrowser Then
        frm_HTML.ShowHTMLPreview GetHTML(shp)
    Else
        ShowHTMLInBrowser GetHTML(shp)
    End If
End Sub

Public Sub ShowHTMLInBrowser(ByVal html As String, Optional pageName As String = "_.html")
    
    'Save temporary html page
    tmp_path = Application.ActiveDocument.path & pageName
    SaveTextToFile html, tmp_path
    
    ShellExecute 0, "Open", tmp_path
End Sub
