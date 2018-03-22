VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReportInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CancelCommand_Click()
DoCmd.Close acForm, "ReportInterface"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Command8_Click()
MsgBox Application.Printer.DeviceName
MsgBox Application.Printer.Port
End Sub

Private Sub Command9_Click()
Dim n As Integer
n = InputBox("No")
MsgBox PMWOFormatNo("PMWO", n)

End Sub

Private Sub ExportAudit_Click()
On Error GoTo Err

DoCmd.OutputTo acOutputTable, "PMAudit", acFormatXLSX

Exit Sub
Err:
If Err.Number = 2501 Then Resume Next
End Sub

Private Sub ExportCommand_Click()
On Error GoTo Err

DoCmd.OutputTo acOutputTable, "Equipments", acFormatXLSX

Exit Sub
Err:
If Err.Number = 2501 Then Resume Next
End Sub

Private Sub PMWOCombo_AfterUpdate()
On Error GoTo Err
Dim m As Integer
Dim n As Integer
n = InputBox("Asset No= " & Me.PMWOCombo.Column(0) & vbNewLine & _
"Word Order Type= " & Me.PMWOCombo.Column(1) & vbNewLine & _
"Status = " & Me.PMWOCombo.Column(2) & vbNewLine _
& "WO Request= " & Me.PMWOCombo.Column(3) & vbNewLine & vbNewLine _
& "How many copies?", "How many copy is needed?")

If Not IsNumeric(n) Then
    MsgBox "Please select a number", vbCritical, "Error"
    Exit Sub
End If

If (n > 10) Then
    MsgBox "Not more than 10 copies!", vbCritical, "Select a number less than or equal to 10!"
    Exit Sub
End If


For m = 1 To n

    DoCmd.OpenReport "PMWOReport", acViewNormal, , "[woid]=" & Me.WOCombo.Column(4)

Next


Exit Sub
Err:
If (Err.Number = 13) Then
MsgBox "Please select a number", vbCritical, "Error"
Exit Sub
End If

End Sub

Private Sub PMWOCommand_Click()

On Error GoTo Err

DoCmd.OutputTo acOutputTable, "PMWO", acFormatXLSX

Exit Sub
Err:
If Err.Number = 2501 Then Resume Next


End Sub

Private Sub TaskListCommand_Click()

On Error GoTo Err

DoCmd.OutputTo acOutputTable, "PMTask", acFormatXLSX

Exit Sub
Err:
If Err.Number = 2501 Then Resume Next

End Sub

Private Sub WOCombo_AfterUpdate()

On Error GoTo Err
Dim m As Integer
Dim n As Integer
n = InputBox("Asset No= " & Me.PMWOCombo.Column(0) & vbNewLine & _
"Word Order Type= " & Me.PMWOCombo.Column(1) & vbNewLine & _
"Status = " & Me.PMWOCombo.Column(2) & vbNewLine _
& "WO Request= " & Me.PMWOCombo.Column(3) & vbNewLine & vbNewLine _
& "How many copies?", "How many copy is needed?")

If Not IsNumeric(n) Then
    MsgBox "Please select a number", vbCritical, "Error"
    Exit Sub
End If

If (n > 10) Then
    MsgBox "Not more than 10 copies!", vbCritical, "Select a number less than or equal to 10!"
    Exit Sub
End If

For m = 1 To n

    DoCmd.OpenReport "WOReport", acViewNormal, , "[woid]=" & Me.WOCombo.Column(4)

Next


Exit Sub
Err:
If (Err.Number = 13) Then
MsgBox "Please select a number", vbCritical, "Error"
Exit Sub
End If

End Sub

Private Sub WOCommand_Click()
On Error GoTo Err

DoCmd.OutputTo acOutputTable, "WO", acFormatXLSX

Exit Sub
Err:
If Err.Number = 2501 Then Resume Next
End Sub
