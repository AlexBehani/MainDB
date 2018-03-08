VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PMGenerationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CancelCommand_Click()
DoCmd.Close acForm, "PMGenerationForm"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Form_Load()
Label7.Caption = "PM Generation for the month of " & MonthName(Month(Date))
End Sub

Private Sub PMGenCommand_Click()

Call DeleteOldPM
Call PMGenerator
End Sub
