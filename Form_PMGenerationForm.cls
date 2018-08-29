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
Label7.Caption = "PM Generation for the month of " & MonthName(Month(Date) + 1)
End Sub

Private Sub PMGenCommand_Click()
Dim i As Integer
Dim j As Integer
MsgBox "It might take a few minutes!" & vbNewLine & vbNewLine _
& "please wait", vbInformation + vbOKOnly, ""

Call PMGnrAudit
Call DeleteOldPM
Call PMGenerator

'i = RightOUtterJOin
i = RightOUtterJOin
j = PMWO_Create(i)
If j = 0 Then
MsgBox "You are all set!", vbInformation, ""
Else
MsgBox "Done" & vbNewLine & vbNewLine & _
"Generated " & j & " PM(s)", vbInformation, ""
End If
'Call DeleteOldPM

End Sub
