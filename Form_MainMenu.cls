VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Administration_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "User_Administration"
End Sub

Private Sub ClosingWO_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "WOClosing"
End Sub

Private Sub Equipment_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "EquipmentForm"
End Sub

Private Sub PM_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "PMGenerationForm"
End Sub

Private Sub Setting_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "SettingForm"
End Sub

Private Sub WorkOrder_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "WorkOrder"
End Sub
