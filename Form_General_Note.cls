VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_General_Note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub AddNew_Click()
Clear_Field
Set GNote = New GeneralNote
End Sub

Private Sub Command19_Click()
If (Em_Field) Then Exit Sub


GNote.Description = Me.Description
GNote.Date_ = Me.Date_
GNote.User = Me.User_
GNote.Comment = Me.Comment

Call Save_Note
Call Clear_Field
Me.GeneralNote_subform1.Requery


End Sub

Private Sub Command20_Click()
If Not (GNote Is Nothing) Then Set GNote = Nothing
DoCmd.Close acForm, "General_Note"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Command21_Click()
Delete_Note
Clear_Field
Me.GeneralNote_subform1.Requery
End Sub

Private Sub Form_Load()
Set GNote = New GeneralNote
GNote.Edit = False
End Sub
