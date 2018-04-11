VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_WorkOrderEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



'Private Sub CancelCommand_Click()
'If Not (PMWO Is Nothing) Then Set PMWO = Nothing
'DoCmd.Close acForm, "PMWorkOrderForm"
'DoCmd.OpenForm "MainMenu"
'End Sub
'
'Private Sub Form_Load()
'Me.WONumber = PMWO.FormatedWONUmber
'Me.AssetNumber = PMWO.AssetNumber
'Me.DueDate = PMWO.DueDate
'Me.RequestedBy = "PM"
''Me.RequestedDate = Me.DueDate
'Me.WOType = PMWO.WOType
'Me.RequestedWork = PMWO.WORequest
''Me.Comment = pmwo.EngineeringComment
'Me.Description = PMWO.WODescription
'Me.Manufacturer = PMWO.Manufacturer
'Me.Comment = Nz(PMWO.EngineeringComment, "")
'Me.Status = PMWO.Status
'Me.RequestedDate = Nz(PMWO.RequestedDate, 0)
'
'
'
'
'End Sub
'
'Private Sub PrintCommand_Click()
'If (Nz(Me.Comment, "") = "") Then
'
'    MsgBox "Please add 'Engineering Comment", vbInformation, ""
'    Exit Sub
'End If
'
'
'If (Nz(Me.RequestedDate, "") = "" Or Me.RequestedDate = "12:00:00 AM") Then
'
'    MsgBox "Please add 'Request date'", vbInformation, ""
'    Exit Sub
'End If
'
'
'
'PMWO.EngineeringComment = Me.Comment
'PMWO.RequestedDate = Me.RequestedDate
'Call Save_PMWO
'
'DoCmd.OpenReport "WOReport", acViewNormal, , "[woid]=" & PMWO.WOID
'
'
'
'
'
'
'
'End Sub
Private Sub CancelCommand_Click()

If Not (WO Is Nothing) Then Set WO = Nothing
DoCmd.Close acForm, "WorkOrderEditForm"
DoCmd.OpenForm "MainMenu"

End Sub

Private Sub Form_Load()
If CredentialsCheck Then
'Me.RequestedBy = CUser.FullName
End If
Me.WONumber = WO.FormatedWONUmber
Me.AssetNumber = WO.AssetNumber
Me.DueDate = WO.DueDate
Me.RequestedBy = WO.RequestBy
Me.RequestedDate = WO.RequestedDate
Me.Manufacturer = WO.Manufacturer
Me.Status = WO.Status
Me.RequestedWork = Nz(WO.WORequest, "")
Me.Comment = Nz(WO.EngineeringComment, "")
'Me.Description = Nz(WO.WODescription, "")
If (Nz(WO.WODescription, "") = "") Then
    Me.Description = Null
Else
    Me.Description = WO.WODescription
End If
Me.AssetNumber.RowSource = AssetNumberList


End Sub

Private Sub PrintCommand_Click()
If Em_Field Then Exit Sub

'WO.WONumber = Me.WONumber
WO.AssetNumber = Me.AssetNumber
WO.DueDate = Me.DueDate
WO.RequestBy = Me.RequestedBy
WO.RequestedDate = Me.RequestedDate
'WO.WOType = Me.WOType
WO.Manufacturer = Me.Manufacturer
WO.Status = Me.Status
WO.WORequest = Me.RequestedWork
WO.EngineeringComment = Me.Comment
WO.WODescription = Me.Description



Update_WO (WO.WOID)
DoCmd.OpenReport "WOReport", acViewNormal, , "[id]=" & WO.WOID
'Clear_Field


Set WO = Nothing
DoCmd.Close acForm, "WorkOrderEditForm"
DoCmd.OpenForm "MainMenu"

End Sub
