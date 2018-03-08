Attribute VB_Name = "WOModule"
Option Compare Database

Public Sub Save_WO()

Dim db As Database
Dim WORs As Recordset

Set db = CurrentDb

If Not WO.Edit Then
    
    Set WORs = db.OpenRecordset("WO")
    WORs.AddNew
    WORs!WODescription = WO.WODescription
    WORs!ModelNumber = WO.ModelNumber
    WORs!Scheduled = WO.Scheduled
    WORs!WOType = WO.WOType
    WORs!WORequest = WO.WORequest
    WORs!AssignedTo = WO.AssignedTo
    WORs!Status = WO.Status
    WORs!Priority = WO.Priority
'    WORs!wonumber = WONumGen
    WORs.Update
Else

    Set WORs = db.OpenRecordset("SELECT * FROM WO WHERE WOID= " & WO.WOID)
    WORs.MoveFirst
    WORs.Edit
    WORs!WODescription = WO.WODescription
    WORs!ModelNumber = WO.ModelNumber
    WORs!Scheduled = WO.Scheduled
    WORs!WOType = WO.WOType
    WORs!WORequest = WO.WORequest
    WORs!AssignedTo = WO.AssignedTo
    WORs!Status = WO.Status
    WORs!Priority = WO.Priority
    WORs.Update
    
End If

Set WORs = Nothing
Set db = Nothing
WO.Edit = False
    
End Sub


Function Load_WO(ID As Integer) As WO

Dim WOTemp As WO
Dim db As Database
Dim WORs As Recordset

Set WOTemp = New WO
Set db = CurrentDb
Set WORs = db.OpenRecordset("SELECT * FROM WO WHERE WOID= " & ID)
WORs.MoveFirst
WOTemp.WODescription = WORs!WODescription
WOTemp.ModelNumber = WORs!ModelNumber
WOTemp.Scheduled = WORs!Scheduled
WOTemp.WOType = WORs!WOType
WOTemp.WORequest = WORs!WORequest
WOTemp.AssignedTo = WORs!AssignedTo
WOTemp.Status = Nz(WORs!Status, "")
WOTemp.Priority = Nz(WORs!Priority, "")
WOTemp.WOID = WORs!WOID
Set Load_WO = WOTemp
Set db = Nothing
Set WORs = Nothing
Set WOTemp = Nothing



End Function

Function Load_WOClosing(ID As Integer) As WOClosing

Dim WOClosingTemp As WOClosing
Dim db As Database
Dim WORs As Recordset

Set WOClosingTemp = New WOClosing
Set db = CurrentDb
Set WORs = db.OpenRecordset("SELECT * FROM WO WHERE WOID= " & ID)
WORs.MoveFirst
'WOClosingTemp.Task = WORs!Task
WOClosingTemp.DateDone = WORs!DateDone
WOClosingTemp.TimeDone = WORs!TimeDone
WOClosingTemp.TaskComment = WORs!TaskComment
WOClosingTemp.Employee = WORs!Employee
WOClosingTemp.Completed = WORs!Completed


Set Load_WOClosing = WOClosingTemp
Set db = Nothing
Set WORs = Nothing
Set WOClosingTemp = Nothing

End Function

Public Sub Save_WOClosing()

Dim db As Database
Dim WORs As Recordset

Set db = CurrentDb

    Set WORs = db.OpenRecordset("SELECT * FROM WO WHERE WOID= " & WOClosing.WOID)
    WORs.MoveFirst
    WORs.Edit
    WORs!DateDone = WOClosing.DateDone
    WORs!TimeDone = WOClosing.TimeDone
    WORs!TaskComment = WOClosing.TaskComment
    WORs!Employee = WOClosing.Employee
    WORs!Completed = WOClosing.Completed
    WORs.Update
    


Set WORs = Nothing
Set db = Nothing
Set WO = Nothing
    
End Sub


Public Function WONumGen() As String

Dim init As String

Dim db As Database
Set db = CurrentDb
Dim WORec As Recordset

init = "WO00001"

Set WORec = db.OpenRecordset("WO")

If Not (WORec.RecordCount > 0) Then
wonumbgen = init
Else

WONumGen = "WO" & Left(DMax("WONumber", "WO"), 1) + 1
End If




End Function
