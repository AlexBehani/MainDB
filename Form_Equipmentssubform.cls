VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Equipmentssubform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ID_Click()


If (Nz(Me.ID, 0) = 0) Then Exit Sub

Set Equip = Load_Eq(Me.ID)

Equip.ID = Me.ID
Equip.Edit = True
    
Forms!EquipmentForm.EquipmentN = Equip.EquipmentNumber
Forms!EquipmentForm.Description = Equip.Description
Forms!EquipmentForm.AssetNumber = Equip.AssetNumber
Forms!EquipmentForm.SerialNumber = Equip.SerialNumber
Forms!EquipmentForm.ModelNumber = Equip.ModelNumber
Forms!EquipmentForm.LocationCombo = Equip.Location
Forms!EquipmentForm.NoteCombo = Equip.Note
Forms!EquipmentForm.ServiceCheck = Equip.Service
Forms!EquipmentForm.SafetyInstruction = Equip.SafetyIntstruction
Forms!EquipmentForm.PMCombo = Equip.PM
Forms!EquipmentForm.Safety = Equip.Saftey
Forms!EquipmentForm.DateReqistered = Equip.DateRegistered

End Sub
