Attribute VB_Name = "m_imageWork"
'---------------------------------------------------------------------------------------
' Procedure : WIA_CropImage_Perc
' Author    : Daniel Pineault, CARDA Consultants Inc. Modifyed by Malyutin Oleg aka Obsidian
' Website   : http://www.cardaconsultants.com
' Purpose   : Crop an image like WIA_CropImage but sizes in proportions
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Late Binding version  -> None required
'             Early Binding version -> Microsoft Windows Image Acquisition Library vX.X
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sInitialImage     : Fully qualified path and filename of the image file to gcrop
' pLeft             : Distance for the Left most side of the image to start cropping in proportion
' pTop              : Distance for the Top most side of the image to start cropping in proportion
' pRight            : Distance for the Right most side of the image to start cropping in proportion
' pBottom           : Distance for the Bottom most side of the image to start cropping in proportion
' sDestiationImage  : Optional fully qualified path and filename of the cropped image file
'                       If omitted, it overwrites the original sInitialImage image file
'
' Usage:
' ~~~~~~
' Cropped and create a new image
' WIA_CropImage "C:\Temp\Img01.jpg", 0, 0, 1000, 1000, "C:\Temp\Img01_Cropped.jpg"
' Cropped and overwrite the original source image
' WIA_CropImage "C:\Temp\Img01.jpg", 0, 0, 1000, 1000
'
' Revision History:
' Rev       Date(yyyy-mm-dd)        Description
' **************************************************************************************
' 1         2022-06-09              Initial Public Release
'---------------------------------------------------------------------------------------

Public Sub WIA_CropImage_Perc(sInitialImage As String, _
                      pLeft As Single, pBottom As Single, pRight As Single, pTop As Single, _
                      Optional sDestiationImage As String)
    On Error GoTo Error_Handler
    #Const WIA_EarlyBind = False    'True => Early Binding / False => Late Binding
    #If WIA_EarlyBind = True Then
        Dim oIF               As WIA.ImageFile
        Dim oIP               As WIA.ImageProcess
 
        Set oIF = New WIA.ImageFile
        Set oIP = New WIA.ImageProcess
    #Else
        Dim oIF               As Object
        Dim oIP               As Object
 
        Set oIF = CreateObject("WIA.ImageFile")
        Set oIP = CreateObject("WIA.ImageProcess")
    #End If
    Dim sCropImage            As String
 
    'Load and crop
    oIF.LoadFile sInitialImage
    oIP.Filters.Add oIP.FilterInfos("Crop").FilterID
    'Get sizes in pixel
    Dim lLeft As Long
    Dim lTop As Long
    Dim lRight As Long
    Dim lBottom As Long
    lLeft = Int(oIF.width * pLeft)
    lBottom = Int(oIF.height * pBottom)
    lRight = Int(oIF.width * pRight)
    lTop = Int(oIF.height * pTop)
    
    With oIP.Filters(1)
        .Properties("Left") = lLeft
        .Properties("Top") = lTop
        .Properties("Right") = lRight
        .Properties("Bottom") = lBottom
    End With
    Set oIF = oIP.Apply(oIF)
 
    'Save the cropped image
    If sDestiationImage <> "" Then
        'New file specified
        sCropImage = sDestiationImage
    Else
        'Overwrite original if no Dest. file specified
        sCropImage = sInitialImage
    End If
    If Len(Dir(sCropImage)) > 0 Then Kill sCropImage
    oIF.SaveFile sCropImage
 
Error_Handler_Exit:
    On Error Resume Next
    If Not oIP Is Nothing Then Set oIP = Nothing
    If Not oIF Is Nothing Then Set oIF = Nothing
    Exit Sub
 
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: WIA_CropImage" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Sub
