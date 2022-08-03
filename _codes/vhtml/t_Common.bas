Attribute VB_Name = "t_Common"
Option Explicit



Public Function SaveTextToFile(ByVal txt, ByVal filename) As Boolean
'Save txt in "utf-8" to filename
'It is not my code, but i did not remember where i got it
Dim binaryStream As Object

    On Error Resume Next: Err.Clear

    With CreateObject("ADODB.Stream")
        .Type = 2: .CharSet = "utf-8": .Open
        .WriteText txt

        Set binaryStream = CreateObject("ADODB.Stream")
        binaryStream.Type = 1: binaryStream.Mode = 3: binaryStream.Open
        .Position = 3: .CopyTo binaryStream        'Skip BOM bytes
        .flush: .Close
        binaryStream.SaveToFile filename, 2
        binaryStream.Close
    End With

    SaveTextToFile = Err = 0: DoEvents
End Function

Public Function FixDoubleQuotes(ByVal txt As String) As String
Dim tmp As String
    
    tmp = Mid(txt, 2, Len(txt) - 2)
    
    FixDoubleQuotes = Chr(34) & Replace(tmp, Chr(34), "'") & Chr(34)
End Function

Public Sub LayerPrint(ByVal LayerName As String, toPrint As Boolean)
    On Error Resume Next
    If toPrint Then
        Application.ActivePage.Layers.ItemU("Кадры").CellsC(visLayerPrint).FormulaU = "1"
    Else
        Application.ActivePage.Layers.ItemU("Кадры").CellsC(visLayerPrint).FormulaU = "0"
    End If
End Sub

Public Sub ShowStatus(ByVal msgtxt As String)
    f_Statuss.ShowChanges msgtxt
End Sub

Public Function SP(ByVal num As Integer, Optional str As String = " ") As String
Dim i As Integer
    If num < 0 Then num = 0
    For i = 0 To num - 1
        SP = SP & str
    Next i
End Function
