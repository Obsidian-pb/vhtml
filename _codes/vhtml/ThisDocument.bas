VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Public WithEvents app  As Visio.Application
Attribute app.VB_VarHelpID = -1








Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    RemoveTB
    appDeActivate
    
    On Error Resume Next
    Kill Application.ActiveDocument.path & "_.html"
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    appActivate
    AddTB
End Sub




Public Sub appActivate()
    Set app = Visio.Application
    Debug.Print "connection tracing activated"
End Sub
Public Sub appDeActivate()
    Set app = Nothing
    Debug.Print "connection tracing deactivated"
End Sub

Private Sub app_ConnectionsAdded(ByVal Connects As IVConnects)
    TurnIntoFormulaConnection Connects
End Sub



