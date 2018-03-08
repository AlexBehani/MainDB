VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_WOClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Command15_Click()
MsgBox Nz(WOClosing.Completed, 0)
End Sub



Private Sub Cancel_Click()
DoCmd.Close acForm, "WOClosing"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub CompleteToggle_Click()

If (Me.CompleteToggle) Then
Me.CompleteToggle.Picture = "Done"


Else
Me.CompleteToggle.PictureData = ""

End If
End Sub

Private Sub Form_Load()
Me.CompleteToggle.Picture = ""
End Sub

Private Sub SaveCommand_Click()
On Error GoTo Err
If Nz(Me.WODropDown.Column(0), "") = "" Then

    MsgBox "Please select a Work Order", vbOKOnly, "Work Order?"
    Exit Sub
End If

If (Em_Field) Then
    
    Exit Sub
End If

WOClosing.DateDone = CDate(Me.DateCompleted)
WOClosing.TimeDone = Me.TimeCompleted
WOClosing.TaskComment = Me.TaskComment
WOClosing.Employee = Me.Employee
WOClosing.WOID = Me.WODropDown.Column(0)
If (Me.CompleteToggle) Then WOClosing.Completed = True Else WOClosing.Completed = False

Call Save_WOClosing
MsgBox "Saved", vbOKOnly, ""

Exit Sub
Err:
MsgBox "Error happened!", vbCritical, "Error"
Call Errs(Err.Description, Err.LastDllError, Err.Number, Err.Source)
End Sub

Private Sub WODropDown_AfterUpdate()

Set WO = Load_WO(Me.WODropDown.Column(0))

If Not (Nz(Me.WODropDown.Column(4), "") = "") Then
    Set WOClosing = Load_WOClosing(Me.WODropDown.Column(0))
    
    Me.DateCompleted = WOClosing.DateDone
    Me.TimeCompleted = WOClosing.TimeDone
    Me.TaskComment = WOClosing.TaskComment
    Me.Employee = WOClosing.Employee
    WOClosing.Edit = True
    WOClosing.WOID = Me.WODropDown.Column(0)
    If (WOClosing.Completed) Then
            Me.CompleteToggle.Picture = "Done"
        Else
            Me.CompleteToggle.Picture = ""
    End If
        
    
Else

    Set WOClosing = New WOClosing
    Me.CompleteToggle.Picture = ""
    Clear_Field
End If
    
    Me!SubFormWorkOrder.Form!WODescription = WO.WODescription
    Me!SubFormWorkOrder.Form!ModelN = WO.ModelNumber
    Me!SubFormWorkOrder.Form!ScheduledCheck = WO.Scheduled
    Me!SubFormWorkOrder.Form!WORequest = WO.WORequest
    Me!SubFormWorkOrder.Form!AssignedTo = WO.AssignedTo
    Me!SubFormWorkOrder.Form!StatusCombo = WO.Status
    Me!SubFormWorkOrder.Form!PriorityCombo = WO.Priority
    
    Set WO = Nothing

End Sub
