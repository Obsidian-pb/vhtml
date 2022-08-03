Attribute VB_Name = "m_sourceCOde"
Option Explicit


Public Sub SaveSourceCode()
'Source code export tool
Dim targetPath As String
Dim doc As Visio.Document
    
    targetPath = GetCodePath(ThisDocument)
    ExportVBA ThisDocument, targetPath

    Debug.Print "Source code exported"

End Sub


Public Sub ExportVBA(ByRef doc As Visio.Document, ByVal sDestinationFolder As String)
'Source code export
    Dim oVBComponent As Object
    Dim fullName As String

    For Each oVBComponent In doc.VBProject.VBComponents
        
        If Not oVBComponent.Name = "m_SourceCode" Then
            If oVBComponent.Type = 1 Then
                ' Standard Module
                fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            ElseIf oVBComponent.Type = 2 Then
                ' Class
                fullName = sDestinationFolder & oVBComponent.Name & ".cls"
            ElseIf oVBComponent.Type = 3 Then
                ' Form
                fullName = sDestinationFolder & oVBComponent.Name & ".frm"
            ElseIf oVBComponent.Type = 100 Then
                ' Document
                fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            Else
                ' UNHANDLED/UNKNOWN COMPONENT TYPE
            End If
            
            oVBComponent.Export fullName
            'SaveTextToFile oVBComponent.CodeModule.Lines(1, oVBComponent.CodeModule.CountOfLines), fullName
            Debug.Print "Saved " & fullName
        End If
        
    Next oVBComponent

End Sub

Private Function GetCodePath(ByRef doc As Visio.Document) As String
'Get source code folder path
Dim path As String
Dim docNameWODot As String
    
    '---Current doc path
    path = doc.path
    '---Add code folder name
    path = GetDirPath(path & "_codes")
        
    '---Add current doc codes folder path
    docNameWODot = Split(doc.Name, ".")(0)
    path = GetDirPath(path & "\" & docNameWODot)
    
    GetCodePath = path & "\"
End Function

Private Function GetDirPath(ByVal path As String) As String
'Get folder with path and create it if it does not exists
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    GetDirPath = path
End Function
