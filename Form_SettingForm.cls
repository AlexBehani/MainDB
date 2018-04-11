VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ExitCommand_Click()
DoCmd.Close acForm, "SettingForm"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub GeneralNote_Click()
DoCmd.Close acForm, "SettingForm"
DoCmd.OpenForm "General_Note"
End Sub

Private Sub LocationCommand_Click()
DoCmd.Close acForm, "SettingForm"
DoCmd.OpenForm "LocationFrm"
End Sub

Private Sub PMList_Click()
DoCmd.Close acForm, "SettingForm"
DoCmd.OpenForm "PM_Task"
End Sub

Private Sub Priority_Click()
DoCmd.Close acForm, "SettingForm"
DoCmd.OpenForm "Priority"
End Sub

Private Sub Request_Click()
DoCmd.Close acForm, "SettingForm"
DoCmd.OpenForm "RequestForm"
End Sub

Private Sub Status_Click()
DoCmd.Close acForm, "SettingForm"
DoCmd.OpenForm "Status"
End Sub
