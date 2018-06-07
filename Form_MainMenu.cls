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

If CUser.AccessLevel > 1 Then
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "WOClosing"
Else
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "CloseWOInterface"
End If
End Sub

Private Sub Equipment_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "EquipmentInterface"
End Sub

Private Sub ExitCommand_Click()
If (MsgBox("Exit?", vbQuestion + vbYesNo, "") = vbYes) Then DoCmd.Quit
End Sub

Private Sub Form_Load()
If Not (CUser Is Nothing) Then
    Me.IDLabel.Caption = CUser.FullName
    Me.IDLabel.ForeColor = RGB(82, 188, 236)
    
    Me.UserIcon.Visible = True
    Me.anonymous.Visible = False
Else
    Me.IDLabel.Caption = "anonymous"
    Me.UserIcon.Visible = False
    Me.anonymous.Visible = True
    Me.IDLabel.ForeColor = RGB(82, 188, 236)
    
    Me.Administration.Enabled = False
    Me.ClosingWO.Enabled = False
    Me.Equipment.Enabled = False
    Exit Sub

End If

If (CUser.AccessLevel > 1) Then
    Me.Administration.Enabled = False
'    Me.ClosingWO.Enabled = False
    Me.Equipment.Enabled = False
    Exit Sub
End If
End Sub

Private Sub PM_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "PMGenerationForm"
End Sub

Private Sub ReportCommand_Click()
DoCmd.OpenForm "ReportInterface"
DoCmd.Close acForm, "MainMenu"
End Sub

Private Sub Setting_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "SettingForm"
End Sub

Private Sub WorkOrder_Click()
DoCmd.Close acForm, "MainMenu"
DoCmd.OpenForm "WOInterfaceForm"
End Sub
