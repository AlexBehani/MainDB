Attribute VB_Name = "RequestModule"
Option Compare Database

Public Sub Save_Request()

Dim db As Database
Dim RequestRs As Recordset

Set db = CurrentDb

If Not Request.Edit Then
    
    Set RequestRs = db.OpenRecordset("Requests")
    RequestRs.AddNew
    RequestRs!UserName = Request.Name
    RequestRs!Email = Request.Email
    RequestRs!Note = Request.Note
    RequestRs!WOInstruction = Request.WOInstruction
    RequestRs.Update
Else

    Set RequestRs = db.OpenRecordset("SELECT * FROM Requests WHERE RequestID= " & Request.ID)
    RequestRs.MoveFirst
    RequestRs.Edit
    RequestRs!UserName = Request.Name
    RequestRs!Email = Request.Email
    RequestRs!Note = Request.Note
    RequestRs!WOInstruction = Request.WOInstruction
    RequestRs.Update
    
End If

Set RequestRs = Nothing
Set db = Nothing
Request.Edit = False
    
End Sub


Function Load_Request(ID As Integer) As Request

Dim RequestTemp As Request
Dim db As Database
Dim RequestRs As Recordset

Set RequestTemp = New Request
Set db = CurrentDb
Set RequestRs = db.OpenRecordset("SELECT * FROM Requests WHERE RequestID= " & ID)
RequestRs.MoveFirst
RequestTemp.Name = RequestRs!UserName
RequestTemp.Email = RequestRs!Email
RequestTemp.Note = RequestRs!Note
RequestTemp.WOInstruction = RequestRs!WOInstruction

Set Load_Request = RequestTemp
Set db = Nothing
Set RequestRs = Nothing
Set RequestTemp = Nothing



End Function
