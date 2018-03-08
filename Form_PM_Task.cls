VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PM_Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub CancelCommand_Click()
If Not (PMTK Is Nothing) Then Set PMTK = Nothing
DoCmd.Close acForm, "PM_Task"
DoCmd.OpenForm "MainMenu"

End Sub

Private Sub Delete_Click()
If (PMTK.Edit = False) Then Exit Sub
Call Delete_PMTask
PMTK.Edit = False
Call Clear_Field
Me.PMTask_subform.Requery
End Sub

Private Sub Form_Load()
Set PMTK = New PMtask
PMTK.Edit = False
End Sub

Private Sub SaveCommand_Click()

If (Em_Field) Then Exit Sub
PMTK.PMTaskName = Me.TaskName
PMTK.PMType = Me.TaskType
PMTK.AssignedTo = Me.AssignedTo
PMTK.Description = Me.Description
PMTK.FrequencyDays = Me.FrequencyDays
PMTK.DownTime = Me.DownTime

Call PMTask_Save
PMTK.Edit = False
Call Clear_Field
Me.PMTask_subform.Requery




End Sub
