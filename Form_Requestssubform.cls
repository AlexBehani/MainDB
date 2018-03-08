VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Requestssubform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub RequestID_Click()

    Set Request = New Request
    Set Request = Load_Request(Me.RequestID)
    
    Forms!RequestForm.NameText = Request.Name
    Forms!RequestForm.Email = Request.Email
    Forms!RequestForm.Note = Request.Note
    Forms!RequestForm.WOInstruction = Request.WOInstruction
    
    Request.Edit = True
    

End Sub
