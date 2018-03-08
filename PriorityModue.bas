Attribute VB_Name = "PriorityModue"
Option Compare Database

Public Sub Update_Priority_List(value As String, action As String)

Dim db As Database
Dim Priority As Recordset

Set db = CurrentDb


If (action = "Add") Then

    Set Priority = db.OpenRecordset("Priority")
    Priority.AddNew
    Priority!Priority = value
    Priority.Update
ElseIf (action = "Remove") Then

    Set Priority = db.OpenRecordset("SELECT * FROM Priority WHERE Priority = '" & value & "'")
    Priority.MoveFirst
    Priority.Delete

End If

Set Priority = Nothing
Set db = Nothing

End Sub

