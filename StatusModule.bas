Attribute VB_Name = "StatusModule"
Option Compare Database

Public Sub Update_Status_List(value As String, action As String)

Dim db As Database
Dim Status As Recordset

Set db = CurrentDb


If (action = "Add") Then

    Set Status = db.OpenRecordset("Status")
    Status.AddNew
    Status!Status = value
    Status.Update
ElseIf (action = "Remove") Then

    Set Status = db.OpenRecordset("SELECT * FROM Status WHERE Status = '" & value & "'")
    Status.MoveFirst
    Status.Delete

End If

Set Status = Nothing
Set db = Nothing

End Sub
