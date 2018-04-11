VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PMTasksubform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub PMID_Click()
On Error GoTo Err
If (Nz(Me.PMID, "") = "") Then Exit Sub

Dim PMTemp As PMTask
PMTK.ID = Me.PMID

Set PMTemp = Load_PMTask(Me.PMID)
PMTK.Edit = True

'Forms!PM_Task.TaskName = PMTemp.PMTaskName
'Forms!PM_Task.TaskType = PMTemp.PMType
Forms!PM_Task.AssignedTo = PMTemp.AssignedTo
Forms!PM_Task.Description = PMTemp.Description
Forms!PM_Task.FrequencyDays = PMTemp.FrequencyDays
'Forms!PM_Task.DownTime = PMTemp.DownTime

Set PMTemp = Nothing


Exit Sub

Err:
If (Err.Number = 91) Then
MsgBox "Please leave the form and come back", vbCritical, "Error"
Else: MsgBox Err.Description, vbCritical, "Error"
End If

End Sub
