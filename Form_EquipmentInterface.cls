VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EquipmentInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub EquipmentCombo_AfterUpdate()

Set Equip = Load_Eq(Me.EquipmentCombo.Column(3))
Equip.ID = Me.EquipmentCombo.Column(3)
DoCmd.Close acForm, "EquipmentInterface"
DoCmd.OpenForm "EquipmentEdit"


End Sub

Private Sub GenerateEquipment_Click()
DoCmd.Close acForm, "EquipmentInterface"
DoCmd.OpenForm "EquipmentForm"
End Sub
