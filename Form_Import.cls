VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub AssetList_Click()
Dim db As Database
Dim Eq As Recordset
Dim EqList As Recordset

Set db = CurrentDb
Set Eq = db.OpenRecordset("Equipments")
Set EqList = db.OpenRecordset("AssetList")

'Eq.MoveFirst
EqList.MoveFirst

Do While Not EqList.EOF
    
    Eq.AddNew
    With Eq
    !AssetN = Nz(EqList!AssetNumber, "")
    !Description = Nz(EqList!Description, "")
    !ModelN = Nz(EqList!ModelN, "")
    !SerialN = Nz(EqList!SerialN, "")
    !Manufacturer = Nz(EqList!Manufacturer, "")
    !DateRegistered = EqList!CalibrationDate
    !EquipmentN = Nz(EqList!EquipmentN, "")
    !Status = Nz(EqList!Status, "")
        If Nz(EqList!Type, "") = "" Then
            !EquipmentType = ""
        ElseIf Nz(EqList!Type, "") = "equip" Then
            !EquipmentType = "Equipment"
        ElseIf Nz(EqList!Type, "") = "inst" Then
            !EquipmentType = "Instrument"
        ElseIf Nz(EqList!Type, "") = "QC" Then
            !EquipmentType = "QC"
        End If
    
        
    End With
    Eq.Update


EqList.MoveNext
Loop

Set Eq = Nothing
Set EqList = Nothing
Set db = Nothing

End Sub




Private Sub GPMCommand_Click()
Dim db As Database
Dim GPM As Recordset
Dim Qry As Recordset
Dim str As String

str = "SELECT Equipments.DateRegistered, Equipments.AssetN, Equipments.ModelN, " & _
"Equipments.SerialN, PMTask.PMID, PMTask.Description, PMTask.AssignedTo, " & _
"PMTask.Frequency, Equipments.Manufacturer, Equipments.Status, Equipments.[_ID] " & _
"FROM PMTask INNER JOIN Equipments ON (PMTask.ModelN = Equipments.ModelN) AND (PMTask.Manufacturer = Equipments.Manufacturer)"

Set db = CurrentDb
Set GPM = db.OpenRecordset("GeneralPM")
Set Qry = db.OpenRecordset(str)

Qry.MoveFirst

Do While Not Qry.EOF

    GPM.AddNew
    With GPM
    !DateRegistered = Qry!DateRegistered
    !ModelNumber = Qry!ModelN
    !AssetNumber = Qry!AssetN
    !SerialNumber = Qry!SerialN
    !PMID = Qry!PMID
    !Description = Qry!Description
    !AssignedTo = Qry!AssignedTo
    !Frequency = Qry!Frequency
    !Manufacturer = Qry!Manufacturer
    !Status = Qry!Status
    !EQid = Qry![_ID]
      
    
    End With
    GPM.Update
Qry.MoveNext
Loop
MsgBox "finish"
Set db = Nothing
Set GPM = Nothing
Set Qry = Nothing



End Sub

Private Sub TaskListCommand_Click()
Dim db As Database
Dim PM As Recordset
Dim PMList As Recordset

Set db = CurrentDb
Set PM = db.OpenRecordset("PMTask")
Set PMList = db.OpenRecordset("PMTaskList")


PMList.MoveFirst

Do While Not PMList.EOF
    
    PM.AddNew
    With PM
    !Manufacturer = Nz(PMList!Manufacturer, "")
    !ModelN = Nz(PMList!ModelN, "")
    !AssignedTo = PMList!AssignedTo
    !Description = Nz(PMList!PreventiveMaintenance, "")
    
    If Nz(PMList!Frequency, "") = "" Then
    !Frequency = ""
    ElseIf Nz(PMList!Frequency, "") = "3" Then
    !Frequency = "Quarterly"
    ElseIf Nz(PMList!Frequency, "") = "6" Then
    !Frequency = "Semi Annually"
    ElseIf Nz(PMList!Frequency, "") = "12" Then
    !Frequency = "Annually"
    ElseIf Nz(PMList!Frequency, "") = "24" Then
    !Frequency = "Bi Annually"
    End If
    
    
     
    End With
    PM.Update


PMList.MoveNext
Loop

Set PM = Nothing
Set PMList = Nothing
Set db = Nothing

MsgBox "Finish"

End Sub
