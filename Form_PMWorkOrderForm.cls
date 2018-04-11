VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PMWorkOrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub CancelCommand_Click()
If Not (PMWO Is Nothing) Then Set PMWO = Nothing
DoCmd.Close acForm, "PMWorkOrderForm"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Form_Load()


Me.WONumber = PMWO.FormatedWONUmber
Me.AssetNumber = PMWO.AssetNumber
Me.DueDate = PMWO.DueDate
Me.RequestedBy = "PM"
'Me.RequestedDate = Me.DueDate
'Me.WOType = PMWO.WOType
Me.RequestedWork = PMWO.WORequest
'Me.Comment = pmwo.EngineeringComment
Me.Description = PMWO.WODescription
Me.Manufacturer = PMWO.Manufacturer
Me.Comment = Nz(PMWO.EngineeringComment, "")
Me.Status = PMWO.Status
If (PMWO.RequestedDate = "12:00:00 AM") Then
    Me.RequestedDate = ""
Else
    Me.RequestedDate = PMWO.RequestedDate
End If
'Me.RequestedDate = Nz(PMWO.RequestedDate, "")

End Sub

Private Sub PrintCommand_Click()
'If (Nz(Me.Comment, "") = "") Then
'
'    MsgBox "Please add 'Engineering Comment", vbInformation, ""
'    Exit Sub
'End If


If (Nz(Me.RequestedDate, "") = "" Or Me.RequestedDate = "12:00:00 AM") Then

    MsgBox "Please add 'Request date'", vbInformation, ""
    Exit Sub
End If


PMWO.WODescription = Nz(Me.Description, "")
PMWO.EngineeringComment = Nz(Me.Comment, "")
PMWO.RequestedDate = Me.RequestedDate
Call Save_PMWO

DoCmd.OpenReport "PMWOReport", acViewNormal, , "[woid]=" & PMWO.WOID

MsgBox "Please pick up your report" & vbNewLine & "It was sent to " & _
Application.Printer.DeviceName, vbInformation, ""

Clear_Field





End Sub
