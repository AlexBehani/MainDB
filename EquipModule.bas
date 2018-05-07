Attribute VB_Name = "EquipModule"
Option Compare Database

Public Function Save_Eq() As Long

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
    Eq.Move 0, Eq.LastModified
    Save_Eq = Eq![_ID]
    

    

    

   
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
    
End Function


Public Function Update_Eq(ID As Long)

Dim db As Database
Dim Eq As Recordset


Set db = CurrentDb
Set Eq = db.OpenRecordset("SELECT Description, AssetN,ModelN, SerialN," & _
                            "DateRegistered, EquipmentN, Manufacturer, " & _
                            "EquipmentType, Status, System FROM Equipments" & _
                            " WHERE [_ID]= " & Equip.ID)

    Eq.MoveFirst
    Eq.Edit
    Eq!Description = Equip.Description
    Eq!AssetN = Equip.AssetNumber
    Eq!ModelN = Equip.ModelNumber
    Eq!SerialN = Equip.SerialNumber
    Eq!DateRegistered = Equip.DateRegistered
    Eq!EquipmentN = Equip.EquipmentNumber
    Eq!Manufacturer = Equip.Manufacturer
    Eq!EquipmentType = Equip.EquipmentType
    Eq!Status = Equip.Status
    Eq!System = Equip.System
    

    Eq.Update
    


Set Eq = Nothing
Set db = Nothing
Equip.Edit = False
    
End Function


Function Load_Eq(ID As Long) As Equipments
 
Dim EqTemp As Equipments
Dim db As Database
Dim Eq As Recordset

Set EqTemp = New Equipments
Set db = CurrentDb
Set Eq = db.OpenRecordset("SELECT * FROM Equipments WHERE [_ID] = " & ID)
Eq.MoveFirst
EqTemp.EquipmentNumber = Eq!EquipmentN
EqTemp.Description = Eq!Description
EqTemp.AssetNumber = Eq!AssetN
EqTemp.ModelNumber = Eq!ModelN
EqTemp.SerialNumber = Eq!SerialN
EqTemp.DateRegistered = Eq!DateRegistered
EqTemp.EquipmentNumber = Eq!EquipmentN
EqTemp.EquipmentType = Eq!EquipmentType
EqTemp.Status = Eq!Status
EqTemp.Manufacturer = Eq!Manufacturer
EqTemp.System = Eq!System

Set Load_Eq = EqTemp
Set db = Nothing
Set Eq = Nothing
Set EqTemp = Nothing



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


Public Sub Save_GenPM(Arr As Variant, RowNumber As Long)

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
'    GenPM!TaskName = PMList!Task_Name
    GenPM!Description = PMList!Description

'    GenPM!TaskType = PMList!Type
    GenPM!AssignedTo = PMList!AssignedTo
'    GenPM!DownTime = PMList!DownTime
    GenPM!Frequency = PMList!Frequency
    GenPM!Manufacturer = Equip.Manufacturer
    GenPM!Status = Equip.Status
    GenPM!eqid = RowNumber
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

Public Function PMListRow(ID As Long) As Variant

Dim i As Integer
Dim j As Integer
Dim db As Database
Dim GPM As Recordset

Dim GPMarray() As Variant
Dim PTaskArray() As Variant

Set db = CurrentDb
Set GPM = db.OpenRecordset("SELECT PMID FROM GeneralPM WHERE EQid= " & ID)
If GPM.RecordCount > 0 Then
    GPM.MoveLast
    i = GPM.RecordCount
    ReDim PTaskArray(i)
    GPM.MoveFirst
    For i = 0 To i - 1
        
        PTaskArray(i) = GPM!PMID
        GPM.MoveNext
        
    Next i
   
End If

PMListRow = PMTaskList(PTaskArray)


Set db = Nothing
Set GPM = Nothing

End Function

Public Function PMTaskList(Arr As Variant) As Variant

Dim db As Database
Dim PMTsk As Recordset
Dim PM As Variant
Dim i As Integer
Dim j As Integer

i = UBound(Arr)

ReDim PM(i, 4)
Set db = CurrentDb

For j = 0 To i - 1
    
    Set PMTsk = db.OpenRecordset("SELECT Task_Name, Description, Frequency, PMID FROM PMTask WHERE PMID =" & Arr(j))
        PMTsk.MoveFirst
        PM(j, 0) = PMTsk!Task_Name
        PM(j, 1) = PMTsk!Description
        PM(j, 2) = PMTsk!Frequency
        PM(j, 3) = PMTsk!PMID
        Set PMTsk = Nothing
Next j

PMTaskList = PM
Set db = Nothing
'If Not (PM Is Nothing) Then Set PM = Nothing

End Function
