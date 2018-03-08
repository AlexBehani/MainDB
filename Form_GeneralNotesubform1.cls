VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GeneralNotesubform1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub GID_Click()
On Error GoTo Err
If (Nz(Me.GID, 0) = 0) Then Exit Sub

Dim TempNote As GeneralNote
GNote.ID = Me.GID
Set TempNote = Load_Note()

GNote.Edit = True

Forms!General_Note.Description = TempNote.Description
Forms!General_Note.Date_ = TempNote.Date_
Forms!General_Note.User_ = TempNote.User
Forms!General_Note.Comment = TempNote.Comment

Set TempNote = Nothing


Exit Sub
Err:
If (Err.Number = 91) Then
MsgBox "Please leave the form and come back", vbCritical, "Error"
Else: MsgBox Err.Description, vbCritical, "Error"
End If


End Sub
