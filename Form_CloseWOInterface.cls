VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CloseWOInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ExitCommand_Click()
DoCmd.OpenForm "MainMenu"
DoCmd.Close acForm, "CloseWOInterface"
End Sub

Private Sub GOCommand_Click()

If Nz(Me.Frame6.value, "") = "" Then
    MsgBox "Please select an option", vbCritical, "What type of WO?"
    Exit Sub
End If

If Me.Frame6.value = 1 Then

DoCmd.OpenForm "PMWOClosing"
DoCmd.Close acForm, "CloseWOInterface"
Else
DoCmd.OpenForm "WOClosing"
DoCmd.Close acForm, "CloseWOInterface"
End If
End Sub
