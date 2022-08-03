Attribute VB_Name = "t_Collections"
Option Explicit

'--------------------------------Collections tools-------------------------------------
Public Sub AddCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Add newCollection items to oldCollection
Dim item As Object

    For Each item In newCollection
        oldCollection.Add item
    Next item
End Sub

Public Sub AddUniqueCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Add newCollection items (with unique .ID prop) to oldCollection
Dim item As Object

    For Each item In newCollection
       AddUniqueCollectionItem oldCollection, item
    Next item
End Sub

Public Sub AddUniqueCollectionItem(ByRef oldCollection As Collection, ByRef item As Variant, Optional ByVal key As String = "")
'Add item (with unique .key prop) to oldCollection
    On Error GoTo ex
    
    If key = "" Then
        oldCollection.Add item, CStr(item.ID)
    Else
        oldCollection.Add item, key
    End If

Exit Sub
ex:
'    Debug.Print "Item with key='" & item.ID & "' is already exists!)"
End Sub

Public Sub SetCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Refresh old collection items with items from newCollection
Dim item As Object
    On Error GoTo ex
    
    Set oldCollection = New Collection

    For Each item In newCollection
        oldCollection.Add item, item.ID
    Next item
    
Exit Sub
ex:
'    Debug.Print "Item with key='" & item.ID & "' is already exists!)"
End Sub

Public Sub SetUniqueCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Refresh old collection items with items (with unique .key prop) from newCollection
Dim item As Object

    Set oldCollection = New Collection

    For Each item In newCollection
        AddUniqueCollectionItem oldCollection, item
    Next item

End Sub

Public Function FormUniqueCollection(ByRef shps As Visio.Shapes) As Collection
'Form newCollection based on shapes set
Dim newCollection As Collection
Dim shp As Visio.Shape

    Set newCollection = New Collection

    For Each shp In shps
        AddUniqueCollectionItem newCollection, shp
    Next shp

Set FormUniqueCollection = newCollection
End Function

Public Sub RemoveFromCollection(ByRef oldCollection As Collection, ByRef item As Object)
'Remove specific item (with unique .ID prop) from collection
    On Error Resume Next
    oldCollection.Remove CStr(item.ID)
End Sub

Public Function GetFromCollection(ByRef coll As Collection, ByVal ID As String) As Object
'Get specific item (with unique .ID prop) from collection
Dim item As Object

    On Error GoTo ex
    Set item = coll.item(ID)
    If Not item Is Nothing Then
        Set GetFromCollection = item
    Else
        Set GetFromCollection = Nothing
    End If
    
Exit Function
ex:
    Set GetFromCollection = Nothing
End Function

Public Function IsInCollection(ByRef coll As Collection, obj As Object) As Boolean
'Check item (with unique .ID prop) existance in collection
Dim item As Object

    On Error GoTo ex
    
    Set item = coll.item(CStr(obj.ID))
    If Not item Is Nothing Then
        IsInCollection = True
    Else
        IsInCollection = False
    End If
    
Exit Function
ex:
    IsInCollection = False
End Function


Public Function Sort(ByVal shps As Collection, ByVal sortCellName As String) As Collection
'The function returns a sorted collection of shapes. The shapes are sorted by the value in the sortCellName cell - whose is greater than the one above
Dim i As Integer
Dim tmpshp As Visio.Shape
Dim tmpColl As Collection

    
    Set tmpColl = New Collection
    
    Do While shps.Count > 1
        
        Set tmpshp = GetMaxShp(shps, sortCellName)
        
        AddUniqueCollectionItem tmpColl, tmpshp
        RemoveFromCollection shps, tmpshp
        
        i = i + 1
        If i > 100 Then Exit Do
    Loop
    
    Set tmpshp = shps(1)
    AddUniqueCollectionItem tmpColl, tmpshp
    
    Set Sort = tmpColl
End Function

Public Function GetMaxShp(ByRef col As Collection, ByVal sortCellName As String) As Visio.Shape
Dim i As Integer
Dim shp1 As Visio.Shape
Dim shp2 As Visio.Shape
Dim shp1Val As Single
Dim shp2Val As Single
    
    Set shp1 = col(1)
    shp1Val = cellVal(shp1, sortCellName)
    For i = 1 To col.Count
        Set shp2 = col(i)
        shp2Val = cellVal(shp2, sortCellName)
        
        If shp2Val > shp1Val Then
            Set shp1 = shp2
            shp1Val = shp2Val
        End If
    Next i
    
Set GetMaxShp = shp1
End Function
