VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RequestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Command10_Click()
If Not (Request Is Nothing) Then

    Set Request = Nothing
End If
DoCmd.Close acForm, "RequestForm"
DoCmd.OpenForm "MainMenu"

End Sub

Private Sub Form_Load()
Set Request = New Request
Request.Edit = False
End Sub

Private Sub NewCommand_Click()
Clear_Field
Set Request = New Request
Request.Edit = False
End Sub

Private Sub SaveCommand_Click()
If (Em_Field) Then
    Exit Sub
End If

Set Request = New Request


Request.Name = Me.NameText
Request.Email = Me.Email
Request.Note = Me.Note
Request.WOInstruction = Me.WOInstruction

Call Save_Request
Call Clear_Field
Me.Requestssubform.Requery


End Sub
