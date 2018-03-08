VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_WorkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command10_Click()
If Not (WO Is Nothing) Then

    Set WO = Nothing
End If
DoCmd.Close acForm, "WorkOrder"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Form_Load()
Set WO = New WO
WO.Edit = False
End Sub

Private Sub NewCommand_Click()
Clear_Field
Me.ScheduledCheck = 0
Set WO = New WO
WO.Edit = False
End Sub

Private Sub SaveCommand_Click()
If (Em_Field) Then
    Exit Sub
End If

'Set WO = New WO

WO.WODescription = Me.WODescription
WO.ModelNumber = Me.EquipN
WO.Scheduled = Nz(Me.ScheduledCheck, 0)
WO.WORequest = Me.WORequest
WO.AssignedTo = Me.AssignedTo
WO.Status = Me.StatusCombo
WO.Priority = Me.PriorityCombo

Call Save_WO
Call Clear_Field
Me.WOsubform.Requery
End Sub
