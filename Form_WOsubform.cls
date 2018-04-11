VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_WOsubform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub WOID_Click()
    Set WO = New WO
    Set WO = Load_WO("WO", Me.WOID)
    
    Forms!WorkOrder.WODescription = WO.WODescription
    Forms!WorkOrder.EquipN = WO.ModelNumber
    Forms!WorkOrder.ScheduledCheck = WO.Scheduled
    Forms!WorkOrder.WORequest = WO.WORequest
    Forms!WorkOrder.AssignedTo = WO.AssignedTo
    Forms!WorkOrder.StatusCombo = WO.Status
'    Forms!WorkOrder.PriorityCombo = WO.Priority
    
    
    WO.Edit = True
End Sub
