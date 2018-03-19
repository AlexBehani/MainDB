Attribute VB_Name = "WOModule"
Option Compare Database

Public Function Save_WO()


Dim db As Database
Dim WORs As Recordset

Set db = CurrentDb
Set WORs = db.OpenRecordset("WO")

WORs.AddNew

WORs!WODescription = WO.WODescription
WORs!ModelNumber = WO.ModelNumber
WORs!WOType = WO.WOType
WORs!WORequest = WO.WORequest
WORs!AssignedTo = WO.AssignedTo
WORs!Status = WO.Status
WORs!Completed = False
WORs!RequestedDate = WO.RequestedDate
WORs!DueDate = WO.DueDate
WORs!WONumber = WO.WONumber
WORs!AssetNumber = WO.AssetNumber
WORs!Manufacturer = WO.Manufacturer
WORs!EngineeringComment = WO.EngineeringComment
WORs!RequestBy = WO.RequestBy
WORs!formatwonumber = WO.FormatedWONUmber
WO.WOID = WORs!WOID
WORs.Update


Set db = Nothing
Set WORs = Nothing
   
End Function

'Load PMWO, and also WO
Function Load_WO(pre As String, ID As Integer)

Dim WOTemp As WO
Dim db As Database
Dim WORs As Recordset


Set db = CurrentDb
   Set WOTemp = New WO
If pre = "WO" Then

    Set WORs = db.OpenRecordset("SELECT * FROM WO WHERE WOID= " & ID)

ElseIf pre = "PMWO" Then

    Set WORs = db.OpenRecordset("SELECT * FROM PMWO WHERE WOID= " & ID)

End If

WORs.MoveFirst
WOTemp.WODescription = WORs!WODescription
WOTemp.AssetNumber = WORs!AssetNumber
'WOTemp.ModelNumber = WORs!ModelNumber
WOTemp.WOType = WORs!WOType
WOTemp.WORequest = WORs!WORequest
'WOTemp.AssignedTo = WORs!AssignedTo
WOTemp.Status = Nz(WORs!Status, "")
WOTemp.RequestedDate = Nz(WORs!RequestedDate, 0)
WOTemp.DueDate = WORs!DueDate
WOTemp.WONumber = WORs!WONumber
WOTemp.Manufacturer = WORs!Manufacturer
WOTemp.EngineeringComment = WORs!EngineeringComment
WOTemp.RequestBy = WORs!RequestBy
WOTemp.FormatedWONUmber = WORs!formatwonumber
WOTemp.WOID = WORs!WOID

Set Load_WO = WOTemp

Set db = Nothing
Set WORs = Nothing
Set WOTemp = Nothing



End Function

Function Load_WOClosing(pre As String, ID As Integer) As WOClosing

Dim WOClosingTemp As WOClosing
Dim db As Database
Dim WORs As Recordset

Set db = CurrentDb

If pre = "WO" Then

    Set WORs = db.OpenRecordset("SELECT * FROM WO WHERE WOID= " & ID)

ElseIf pre = "PMWO" Then

    Set WORs = db.OpenRecordset("SELECT * FROM PMWO WHERE WOID= " & ID)

End If

Set WOClosingTemp = New WOClosing


WORs.MoveFirst
'WOClosingTemp.Task = WORs!Task
WOClosingTemp.DateDone = Nz(WORs!DateDone, 0)
WOClosingTemp.taskComment = Nz(WORs!taskComment, 0)
WOClosingTemp.Completed = Nz(WORs!Completed, 0)


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
'    WORs!TimeDone = WOClosing.TimeDone
    WORs!taskComment = WOClosing.taskComment
'    WORs!Employee = WOClosing.Employee
    WORs!Completed = WOClosing.Completed
    WORs.Update
    


Set WORs = Nothing
Set db = Nothing
'Set WO = Nothing
    
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
