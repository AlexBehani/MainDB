Attribute VB_Name = "Location"
     Option Compare Database

Public Sub Update_Location_List(value As String, action As String)

Dim db As Database
Dim Location As Recordset

Set db = CurrentDb


If (action = "Add") Then

    Set Location = db.OpenRecordset("Locations")
    Location.AddNew
    Location!Location = value
    Location.Update
ElseIf (action = "Remove") Then

    Set Location = db.OpenRecordset("SELECT * FROM Locations WHERE Location = '" & value & "'")
    Location.MoveFirst
    Location.Delete

End If

Set Location = Nothing
Set db = Nothing

End Sub


Public Function Unique_value(Table_Name As String, Field_value As String, value As String)

Dim db As Database
Dim Rs As Recordset

Set db = CurrentDb

Set Rs = db.OpenRecordset("SELECT * FROM " & Table_Name & " WHERE " & Field_value & " = '" & value & "'")

If (Rs.RecordCount > 0) Then

    Unique_value = False
Else
    Unique_value = True
    
End If

Set db = Nothing
Set Rs = Nothing

End Function



