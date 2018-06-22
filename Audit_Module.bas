Attribute VB_Name = "Audit_Module"
Option Compare Database
Option Explicit

Public Sub PMGnrAudit()

Dim db As Database
Dim PMG As Recordset

Set db = CurrentDb
Set PMG = db.OpenRecordset("PMGenAudit")

PMG.AddNew
With PMG
    
    !User = CUser.FullName
    !gendate = Date
    !gentime = Time()
    .Update
    
End With

Set db = Nothing
Set PMG = Nothing

End Sub

Public Sub EquipmentsAudit(Arr As Variant, RowNumber As Long, St As String)

Dim db As Database
Dim EqAdt As Recordset
Dim PMList As Recordset
Dim n As Integer

Set db = CurrentDb
Set EqAdt = db.OpenRecordset("EquipmentAudit")


For n = LBound(Arr) To UBound(Arr) - 1
    
    Set PMList = db.OpenRecordset("SELECT * FROM PMTask WHERE PMID=" & Arr(n))
    With EqAdt
        .AddNew
        !DateRegistered = Equip.DateRegistered
        !AssetNumber = Equip.AssetNumber
        !ModelNumber = Equip.ModelNumber
        !SerialNumber = Equip.SerialNumber
        !PMID = Arr(n)
        !Description = PMList!Description
        !AssignedTo = PMList!AssignedTo
        !Frequency = PMList!Frequency
        !Manufacturer = Equip.Manufacturer
        !Status = Equip.Status
        !EqDescription = Equip.Description
        !EquipmentN = Equip.EquipmentNumber
        !EquipmentType = Equip.EquipmentType
        !eqid = RowNumber
        !User = CUser.FullName
        !EnterDate = Date
        !enterTime = Time()
        !EnterStatus = St
        .Update
    End With
    

Next n



Set EqAdt = Nothing
Set PMList = Nothing
Set db = Nothing
    
End Sub

Public Sub EquipmentsAudit_noPM(St As String)

Dim db As Database
Dim EqAdt As Recordset

Set db = CurrentDb
Set EqAdt = db.OpenRecordset("EquipmentAudit")

    With EqAdt
        .AddNew
        !DateRegistered = Equip.DateRegistered
        !AssetNumber = Equip.AssetNumber
        !ModelNumber = Equip.ModelNumber
        !SerialNumber = Equip.SerialNumber
        !Manufacturer = Equip.Manufacturer
        !Status = Equip.Status
        !EqDescription = Equip.Description
        !EquipmentN = Equip.EquipmentNumber
        !EquipmentType = Equip.EquipmentType
        !User = CUser.FullName
        !EnterDate = Date
        !enterTime = Time()
        !EnterStatus = St
        .Update
    End With


Set EqAdt = Nothing
Set db = Nothing
    
End Sub
