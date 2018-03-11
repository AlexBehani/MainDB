Attribute VB_Name = "EquipModule"
Option Compare Database

Public Sub Save_Eq()

Dim db As Database
Dim Eq As Recordset


Set db = CurrentDb
Set Eq = db.OpenRecordset("Equipments")


    Eq.AddNew
    Eq!Description = Equip.Description
    Eq!AssetN = Equip.AssetNumber
    Eq!ModelN = Equip.ModelNumber
    Eq!SerialN = Equip.SerialNumber
'    Eq!Location = Equip.Location
'    Eq!Note = Equip.Note
'    Eq!Service = Equip.Service
'    Eq!SafetyInstruction = Equip.SafetyIntstruction
'    Eq!PM = Equip.PM
'    Eq!Safety = Equip.Saftey
    Eq!DateRegistered = Equip.DateRegistered
    Eq!EquipmentN = Equip.EquipmentNumber
'    Eq!GNoteID = Equip.GNoteID
    Eq!Manufacturer = Equip.Manufacturer
    Eq!EquipmentType = Equip.EquipmentType
    Eq!Status = Equip.Status
    Eq!System = Equip.System
    

    Eq.Update

    

    

   
'Else
'
'    Set Eq = db.OpenRecordset("SELECT * FROM Equipments WHERE ID= " & Equip.ID)
'    Eq.MoveFirst
'    Eq.Edit
'    Eq!Description = Equip.Description
'    Eq!AssetN = Equip.AssetNumber
'    Eq!ModelN = Equip.ModelNumber
'    Eq!SerialN = Equip.SerialNumber
'    Eq!Location = Equip.Location
'    Eq!Note = Equip.Note
'    Eq!Service = Equip.Service
'    Eq!SafetyInstruction = Equip.SafetyIntstruction
'    Eq!PM = Equip.PM
'    Eq!Safety = Equip.Saftey
'    Eq!DateRegistered = Equip.DateRegistered
'    Eq!EquipmentN = Equip.EquipmentNumber
'    Eq.Update
'
'End If

Set Eq = Nothing
Set db = Nothing
Equip.Edit = False
    
End Sub


Function Load_Eq(ID As Integer) As Equipments
 
'Dim EqTemp As Equipments
'Dim db As Database
'Dim Eq As Recordset
'
'Set EqTemp = New Equipments
'Set db = CurrentDb
'Set Eq = db.OpenRecordset("SELECT * FROM Equipments WHERE ID= " & ID)
'Eq.MoveFirst
'EqTemp.EquipmentNumber = Eq!EquipmentN
'EqTemp.Description = Eq!Description
'EqTemp.AssetNumber = Eq!AssetN
'EqTemp.ModelNumber = Eq!ModelN
'EqTemp.SerialNumber = Eq!SerialN
'EqTemp.Location = Eq!Location
'EqTemp.Note = Eq!Note
'EqTemp.Service = Eq!Service
'EqTemp.SafetyIntstruction = Eq!SafetyInstruction
'EqTemp.PM = Eq!PM
'EqTemp.Saftey = Eq!Safety
'EqTemp.DateRegistered = Eq!DateRegistered
'EqTemp.GNoteID = Nz(Eq!GNoteID, 0)
'
'
'Set Load_Eq = EqTemp
'Set db = Nothing
'Set Eq = Nothing
'Set EqTemp = Nothing



End Function

Public Sub Delete_Eq()
'On Error GoTo Err
'Dim db As Database
'Dim Eq As Recordset
'
'Set db = CurrentDb
'
'
'    Set Eq = db.OpenRecordset("SELECT * FROM Equipments WHERE ID= " & Equip.ID)
'    Eq.MoveFirst
'    Eq.Delete
'
'
'
'
'Set Eq = Nothing
'Set db = Nothing
'Equip.Edit = False
'
'Exit Sub
'Err:
'If (Err.Number = 3021) Then
'MsgBox "Haven't select any record", vbCritical, "Error"
'Else: MsgBox Err.Description & vbNewLine & Err.Number
'End If


End Sub


Public Sub Save_GenPM(Arr As Variant)

Dim db As Database
Dim Eq As Recordset
Dim GenPM As Recordset

Set db = CurrentDb
Set GenPM = db.OpenRecordset("GeneralPM")

'For n = LBound(Arr) To UBound(Arr) - 1
'MsgBox test(n)
'Next n

'If Not (Equip.Edit) Then

For n = LBound(Arr) To UBound(Arr) - 1
    
    Set PMList = db.OpenRecordset("SELECT * FROM PMTask WHERE PMID=" & Arr(n))
    GenPM.AddNew
    GenPM!DateRegistered = Equip.DateRegistered
    GenPM!AssetNumber = Equip.AssetNumber
    GenPM!ModelNumber = Equip.ModelNumber
    GenPM!SerialNumber = Equip.SerialNumber
'    GenPM!Location = Equip.Location
    GenPM!PMID = Arr(n)
    GenPM!TaskName = PMList!Task_Name
    GenPM!Description = PMList!Description

    GenPM!TaskType = PMList!Type
    GenPM!AssignedTo = PMList!AssignedTo
'    GenPM!DownTime = PMList!DownTime
    GenPM!frequency = PMList!frequency
    GenPM!Manufacturer = Equip.Manufacturer
    GenPM!Status = Equip.Status
    GenPM.Update
    

Next n
    

   
'Else
'
'    Set Eq = db.OpenRecordset("SELECT * FROM Equipments WHERE ID= " & Equip.ID)
'    Eq.MoveFirst
'    Eq.Edit
'    Eq!Description = Equip.Description
'    Eq!AssetN = Equip.AssetNumber
'    Eq!ModelN = Equip.ModelNumber
'    Eq!SerialN = Equip.SerialNumber
'    Eq!Location = Equip.Location
'    Eq!Note = Equip.Note
'    Eq!Service = Equip.Service
'    Eq!SafetyInstruction = Equip.SafetyIntstruction
'    Eq!PM = Equip.PM
'    Eq!Safety = Equip.Saftey
'    Eq!DateRegistered = Equip.DateRegistered
'    Eq!EquipmentN = Equip.EquipmentNumber
'    Eq.Update
'
'End If

Set Eq = Nothing
Set PMList = Nothing
Set db = Nothing
Equip.Edit = False
    
End Sub
