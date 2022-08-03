VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Editor 
   Caption         =   "Edit"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7170
   OleObjectBlob   =   "frm_Editor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private openFor As Byte
Private forShp As Visio.Shape


Public Sub ShowEditor(ByRef shp As Visio.Shape, ByVal editType As Byte)

    Set forShp = shp
    openFor = editType

    If openFor = EDIT_CONTENT Then
        Me.txt_Formula.Text = cellFrml(shp, CELLNAME_CONTENT, "")
    ElseIf openFor = EDIT_ATTRS Then
        Me.txt_Formula.Text = cellFrml(shp, CELLNAME_ATTRS, "")
    End If
    
    Me.Show
End Sub

Private Sub cb_Cancel_Click()
    Me.Hide
End Sub

Private Sub cb_Save_Click()
    
    On Error GoTo ex
    
    If openFor = EDIT_CONTENT Then
        forShp.Cells(CELLNAME_CONTENT).FormulaForceU = Me.txt_Formula.Text
        Me.lbl_Tip.Caption = "Content updated"
        Me.lbl_Tip.ForeColor = vbGreen
    ElseIf openFor = EDIT_ATTRS Then
        forShp.Cells(CELLNAME_ATTRS).FormulaForceU = Me.txt_Formula.Text
        Me.lbl_Tip.Caption = "Attributes updated"
        Me.lbl_Tip.ForeColor = vbGreen
    End If
    
    
Exit Sub
ex:
    Debug.Print Err.Number & ": " & Err.Description
    Me.lbl_Tip.Caption = "Error: Please check text!"
    Me.lbl_Tip.ForeColor = vbRed
End Sub

Private Sub UserForm_Activate()
Dim width As Integer
Dim height As Integer
Dim top As Integer
Dim left As Integer
    
    'Set form and child elements sizes and positions
    top = 30
    left = 6
    width = Me.InsideWidth - left * 2
    height = 400

    txt_Formula.width = width
    txt_Formula.height = height
    txt_Formula.top = top
    txt_Formula.left = left
    
    cb_Save.top = top + height + 6
    cb_Cancel.top = top + height + 6
    
    Me.height = cb_Save.top + cb_Save.height + 30
    
    lbl_Tip.top = top + height + 6
    
    'Fill cell refs list
    FillMarkersList

    Me.lbl_Tip.Caption = ""
End Sub


Private Sub FillMarkersList()
'Fill cell refs list
Dim i As Integer
Dim rowName As String
Dim cellLabel As String
Dim cellVal As String
    
    Me.cbox_Cells.Clear
    
    For i = 0 To forShp.RowCount(visSectionProp) - 1
        rowName = "Prop." & forShp.CellsSRC(visSectionProp, i, visCustPropsLabel).rowName
        cellLabel = forShp.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(visUnitsString)
        cellVal = forShp.CellsSRC(visSectionProp, i, visCustPropsValue).ResultStr(visUnitsString)

        cbox_Cells.AddItem rowName, i
        cbox_Cells.Column(1, i) = cellLabel
        cbox_Cells.Column(2, i) = cellVal

    Next i
    
End Sub


Private Sub cbox_Cells_Change()
    InsertRef
End Sub

Private Sub cbox_Cells_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        InsertRef
    End If
End Sub

Private Sub InsertRef()
Dim txt1 As String
Dim txt2 As String
Dim i As Integer
Dim l As Long
Dim s As Long

    If Me.cbox_Cells.ListIndex < 0 Then Exit Sub

    l = Me.txt_Formula.SelStart
    If l = 0 Then
        Me.txt_Formula.Text = Me.cbox_Cells.Column(0) & "&" & Me.txt_Formula.Text
    Else
        i = 1
        Do While i <= l
            If Asc(Mid(Me.txt_Formula.Text, i, 1)) = 13 Then
                l = l + 1
            End If
            i = i + 1
            If i > 10000 Then Exit Sub
        Loop
        
        If l = Len(Me.txt_Formula.Text) Then
            Me.txt_Formula.Text = Me.txt_Formula.Text & "&" & Me.cbox_Cells.Column(0)
        Else
            s = Me.txt_Formula.SelLength
            Do While i <= l + s
                If Asc(Mid(Me.txt_Formula.Text, i, 1)) = 13 Then
                    s = s + 1
                End If
                i = i + 1
                If i > 10000 Then Exit Sub
            Loop
            txt1 = left(Me.txt_Formula.Text, l)
            txt2 = Right(Me.txt_Formula.Text, Me.txt_Formula.TextLength - l - s)
            
            Me.txt_Formula.Text = txt1 & Chr(34) & "&" & Me.cbox_Cells.Column(0) & "&" & Chr(34) & txt2
        End If
    End If
End Sub
