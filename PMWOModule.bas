Attribute VB_Name = "PMWOModule"
Option Compare Database

Function Load_PMWO(ID As Integer) As PMWO

Dim PMWOTemp As PMWO
Dim db As Database
Dim WORs As Recordset

Set PMWOTemp = New PMWO
Set db = CurrentDb
Set WORs = db.OpenRecordset("SELECT * FROM PMWO WHERE WOID= " & ID)
WORs.MoveFirst
PMWOTemp.WODescription = WORs!WODescription
PMWOTemp.ModelNumber = WORs!ModelNumber
'pmwotemp.Scheduled = WORs!Scheduled
PMWOTemp.WOType = WORs!WOType
PMWOTemp.WORequest = WORs!WORequest
PMWOTemp.AssignedTo = WORs!AssignedTo
PMWOTemp.Status = Nz(WORs!Status, "")
'pmwotemp.Priority = Nz(WORs!Priority, "")
PMWOTemp.WOID = WORs!WOID
PMWOTemp.AssetNumber = WORs!AssetNumber
PMWOTemp.WONumber = WORs!WONumber
PMWOTemp.FormatedWONUmber = PMWOFormatNo("PMWO", WORs!WONumber)
PMWOTemp.DueDate = WORs!DueDate
PMWOTemp.Manufacturer = Nz(WORs!Manufacturer, "")
PMWOTemp.EngineeringComment = Nz(WORs!EngineeringComment, "")
PMWOTemp.RequestedDate = Nz(WORs!RequestedDate, 0)


Set Load_PMWO = PMWOTemp
Set db = Nothing
Set WORs = Nothing
Set WOTemp = Nothing



End Function


Public Function Save_PMWO()


Dim db As Database
Dim PMRs As Recordset

Set db = CurrentDb
Set PMRs = db.OpenRecordset("SELECT * FROM PMWO WHERE WOID=" & PMWO.WOID)

PMRs.MoveFirst
PMRs.Edit
PMRs!EngineeringComment = PMWO.EngineeringComment
PMRs!Status = PMWO.Status
PMRs!RequestedDate = PMWO.RequestedDate
PMRs.Update

'Set PMRs = Nothing
Set db = Nothing


End Function
