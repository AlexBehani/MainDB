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
'WORs!WOType = WO.WOType
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
WORs!FormatWONumber = WO.FormatedWONUmber
WORs!QRrequired = WO.QRR
WORs!EqDescription = WO.EqDescription
WORs!QAComment = ""
WORs!EngQAComment = WO.EngineeringComment
'WO.WOID = WORs!WOID
WORs.Update


Set db = Nothing
Set WORs = Nothing
   
End Function

Public Function FindWOID(WOFormatnumber As String) As Integer

Dim db As Database
Dim WORecordset As Recordset

Set db = CurrentDb
Set WORecordset = db.OpenRecordset("SELECT ID FROM WO WHERE FormatWONUmber = '" & WOFormatnumber & "'")

If WORecordset.RecordCount > 0 Then
    WORecordset.MoveFirst
    FindWOID = WORecordset!id
Else

    Set WORecordset = db.OpenRecordset("SELECT TOP(ID) FROM WO")
    FindWOID = WORecordset!id

End If

Set db = Nothing
Set WORecordset = Nothing



End Function



Public Function Update_WO(id As Integer)


Dim db As Database
Dim WOR As Recordset

Set db = CurrentDb

                        

Set WOR = db.OpenRecordset("SELECT WODescription, ModelNumber, WORequest, " & _
                            "AssignedTo, Status, Completed, RequestedDate, " & _
                            "DueDate, WONumber, AssetNumber, Manufacturer, " & _
                            "EngineeringComment, RequestBy, formatwonumber, " & _
                            "QRrequired, LockedDown, QAComment, EngQAComment FROM WO WHERE ID =" & id)



WOR.Edit

WOR!WODescription = WO.WODescription
WOR!ModelNumber = WO.ModelNumber
'WOR!WOType = WO.WOType
WOR!WORequest = WO.WORequest
WOR!AssignedTo = WO.AssignedTo
WOR!Status = WO.Status
WOR!Completed = False
WOR!RequestedDate = WO.RequestedDate
WOR!DueDate = WO.DueDate
WOR!WONumber = WO.WONumber
WOR!AssetNumber = WO.AssetNumber
WOR!Manufacturer = WO.Manufacturer
WOR!EngineeringComment = WO.EngineeringComment
WOR!RequestBy = WO.RequestBy
WOR!FormatWONumber = WO.FormatedWONUmber
WOR!QRrequired = WO.QRR
WOR!QAComment = WO.QAComment

If Nz(WO.QAComment, "") <> "" Then
    WOR!EngQAComment = WO.EngineeringComment & vbNewLine & "QA Comment: " & WO.QAComment
Else
    WOR!EngQAComment = WO.EngineeringComment
End If
'WOR!LockedDown = WO.LockedDown
'WO.WOID = WOR!WOID
WOR.Update


Set db = Nothing
Set WOR = Nothing
   
End Function

'Load PMWO, and also WO
Function Load_WO(pre As String, id As Integer)

Dim WOTemp As WO
Dim db As Database
Dim WORs As Recordset


Set db = CurrentDb
   Set WOTemp = New WO
If pre = "WO" Then

    Set WORs = db.OpenRecordset("SELECT * FROM WO WHERE ID= " & id)

ElseIf pre = "PMWO" Then

    Set WORs = db.OpenRecordset("SELECT * FROM PMWO WHERE WOID= " & id)

End If

WORs.MoveFirst
WOTemp.WODescription = Nz(WORs!WODescription, "")
WOTemp.AssetNumber = WORs!AssetNumber
'WOTemp.ModelNumber = WORs!ModelNumber
'WOTemp.WOType = WORs!WOType
WOTemp.WORequest = WORs!WORequest
'WOTemp.AssignedTo = WORs!AssignedTo
WOTemp.Status = Nz(WORs!Status, "")
WOTemp.RequestedDate = Nz(WORs!RequestedDate, 0)
WOTemp.DueDate = WORs!DueDate
WOTemp.WONumber = WORs!WONumber
WOTemp.Manufacturer = WORs!Manufacturer
WOTemp.EngineeringComment = WORs!EngineeringComment
WOTemp.RequestBy = WORs!RequestBy
WOTemp.FormatedWONUmber = WORs!FormatWONumber
WOTemp.EqDescription = Nz(WORs!EqDescription, "")
If pre = "WO" Then WOTemp.QRR = WORs!QRrequired
'If pre = "WO" Then WOTemp.Closed = WORs!Closed
If pre = "WO" Then WOTemp.LockedDown = WORs!LockedDown
If pre = "WO" Then WOTemp.WOID = WORs!id Else WOTemp.WOID = WORs!WOID
If pre = "WO" Then WOTemp.Completed = WORs!Completed
WOTemp.QAComment = Nz(WORs!QAComment, "")

Set Load_WO = WOTemp

Set db = Nothing
Set WORs = Nothing
Set WOTemp = Nothing



End Function

Function Load_WOClosing(pre As String, id As Integer) As WOClosing

Dim WOClosingTemp As WOClosing
Dim db As Database
Dim WORs As Recordset

Set db = CurrentDb

If pre = "WO" Then

    Set WORs = db.OpenRecordset("SELECT * FROM WO WHERE ID= " & id)

ElseIf pre = "PMWO" Then

    Set WORs = db.OpenRecordset("SELECT * FROM PMWO WHERE WOID= " & id)

End If

Set WOClosingTemp = New WOClosing


WORs.MoveFirst
'WOClosingTemp.Task = WORs!Task
WOClosingTemp.DateDone = Nz(WORs!DateDone, 0)
WOClosingTemp.TaskComment = Nz(WORs!TaskComment, 0)
WOClosingTemp.Completed = Nz(WORs!Completed, 0)
'WOClosingTemp.Closed = WORs!Closed


Set Load_WOClosing = WOClosingTemp
Set db = Nothing
Set WORs = Nothing
Set WOClosingTemp = Nothing

End Function

Public Sub Save_WOClosing()

Dim db As Database
Dim WORs As Recordset

Set db = CurrentDb

    Set WORs = db.OpenRecordset("SELECT * FROM WO WHERE ID= " & WOClosing.WOID)
    WORs.MoveFirst
    WORs.Edit
    WORs!DateDone = WOClosing.DateDone
'    WORs!TimeDone = WOClosing.TimeDone
    WORs!TaskComment = WOClosing.TaskComment
'    WORs!Employee = WOClosing.Employee
    WORs!Completed = WOClosing.Completed
    If WOClosing.Completed = True Then WORs!ClosedinDb = Now()
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
Set WORec = Nothing
Set db = Nothing

End Function

Public Function AssetNumberList() As String

Dim db As Database
Dim Eq As Recordset
Dim Str As String

Str = "N/A;"
Set db = CurrentDb
Set Eq = db.OpenRecordset("SELECT AssetN FROM JoinQuery")

If Eq.RecordCount > 0 Then

    Eq.MoveFirst
    
    Do While Not Eq.EOF
        
        Str = Str & Eq!AssetN & ";"
        Eq.MoveNext
        
    Loop
End If
Set Eq = Nothing
Set db = Nothing
AssetNumberList = Str

End Function

Public Function AssetAssociatedData(Asset As String) As String
Dim db As Database
Dim Eq As Recordset
Dim Str As String

Set db = CurrentDb
Set Eq = db.OpenRecordset("SELECT Manufacturer, Status, Description FROM Equipments WHERE AssetN = '" & Asset & "'")

If Eq.RecordCount > 0 Then
    Eq.MoveFirst
    Str = Eq!Manufacturer & ";" & Eq!Status
    WO.EqDescription = Nz(Eq!Description, "")
End If

Set Eq = Nothing
Set db = Nothing
AssetAssociatedData = Str

End Function


'Public Sub Save_Quality_WOClosing()
'
'Dim db As Database
'Dim WORs As Recordset
'
'Set db = CurrentDb
'
'    Set WORs = db.OpenRecordset("SELECT Closed FROM WO WHERE ID= " & WOClosing.WOID)
'    WORs.MoveFirst
'    WORs.Edit
'    WORs!Closed = WOClosing.Closed
'
'    WORs.Update
'
'
'
'Set WORs = Nothing
'Set db = Nothing
''Set WO = Nothing
'
'End Sub

