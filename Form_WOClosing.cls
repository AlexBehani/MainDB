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
On Error GoTo Err_Handel
If Not (WO Is Nothing) Then Set WO = Nothing
If Not (WOClosing Is Nothing) Then Set WOClosing = Nothing
DoCmd.Close acForm, "WOClosing"
DoCmd.OpenForm "MainMenu"
Exit Sub
Err_Handel:
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

Private Sub FilterCombo_AfterUpdate()

If Me.FilterCombo = "Completed" Then
    Me.WODropDown.RowSource = "SELECT WO.FormatWONumber, WO.AssetNumber, WO.Manufacturer, " & _
    "WO.DueDate, WO.ID, WO.taskComment FROM WO WHERE ((Not (Nz(wo.WODescription,'')='')) AND ((WO.Completed)=True));"


ElseIf Me.FilterCombo = "Not Completed" Then
    Me.WODropDown.RowSource = "SELECT WO.FormatWONumber, WO.AssetNumber, WO.Manufacturer, " & _
    "WO.DueDate, WO.ID, WO.taskComment FROM WO WHERE ((Not (Nz(wo.WODescription,'')='')) AND ((WO.Completed)=False));"

End If

End Sub

Private Sub Form_Load()
Me.CompleteToggle.Picture = ""
If CUser.AccessLevel > 1 Then
    Me.WODropDown.RowSource = "SELECT WO.FormatWONumber, WO.AssetNumber, WO.Manufacturer, WO.DueDate, WO.ID, WO.taskComment " & _
    "FROM WO WHERE (((WO.WODescription) Is Not Null) AND ((WO.Invisible)=False) AND ((WO.QRrequired)=True));"
    Me.DateCompleted.Locked = True
    Me.TaskComment.Locked = True
    Me.CompleteToggle.Locked = True
    Me.FilterCombo.Locked = True
    Me.Invisible.Visible = True
    Me.InvisibleLabel.Visible = True

End If
End Sub

Private Sub SaveCommand_Click()

On Error GoTo Err

If CUser.AccessLevel > 1 Then

    If Not WOClosing.Completed Then Exit Sub
    WOClosing.Invisible = Me.Invisible
    Save_Quality_WOClosing
    Exit Sub

End If


If Nz(Me.WODropDown.Column(0), "") = "" Then

    MsgBox "Please select a Work Order", vbOKOnly, "Work Order?"
    Exit Sub
End If

If Not (Me.CompleteToggle) Then

    If MsgBox("You haven't selected completed!" & vbNewLine & "Would you still like to continue?" _
    , vbQuestion + vbYesNo) = vbNo Then
    
        Exit Sub
    End If
End If

If (Em_Field) Then
    
    Exit Sub
End If

'If (WOClosing Is Nothing) Then MsgBox "Closed"

WOClosing.DateDone = Me.DateCompleted
WOClosing.TaskComment = Me.TaskComment
WOClosing.WOID = Me.WODropDown.Column(4)
If (Me.CompleteToggle) Then WOClosing.Completed = True Else WOClosing.Completed = False

Call Save_WOClosing
MsgBox "Saved", vbOKOnly + vbInformation, ""

If (WOClosing.Completed) Then

    If (MsgBox("Would you like to update the Date Registered " & _
    "for an associated PM task?", vbQuestion + vbYesNo, "Update Date Reqister?") = vbYes) Then
    
        DoCmd.Close acForm, "WOClosing"
        DoCmd.OpenForm "GPMTaskList"
        Exit Sub
    End If
End If

'    Set WOClosing = New WOClosing
'    Set WO = New WO
    Set WOClosing = Nothing
    Set WO = Nothing
'    Me.CompleteToggle.Picture = ""
'    Me.SubFormWorkOrder.Requery
    DoCmd.Close acForm, "WOClosing"
    DoCmd.OpenForm "MainMenu"
'    Clear_Field

Exit Sub
Err:
MsgBox "Error happened!", vbCritical, "Error"
Call Errs(Err.Description, Err.LastDllError, Err.Number, Err.Source)
End Sub

Private Sub WODropDown_AfterUpdate()

Set WO = Load_WO("WO", Me.WODropDown.Column(4))

If Not (Nz(Me.WODropDown.Column(5), "") = "") Then
    Set WOClosing = Load_WOClosing("WO", Me.WODropDown.Column(4))
    
    If (WOClosing.DateDone = "12/30/1899") Then
        Me.DateCompleted = Null
    Else
        Me.DateCompleted = WOClosing.DateDone
    End If
    
    If (WOClosing.TaskComment = "") Then
        Me.TaskComment = Null
    Else
        Me.TaskComment = WOClosing.TaskComment
    End If
    
    WOClosing.Edit = True
    WOClosing.WOID = Me.WODropDown.Column(4)
    If (WOClosing.Completed) Then
            Me.CompleteToggle.Picture = "Done"
            Me.CompleteToggle = True
        Else
            Me.CompleteToggle.Picture = ""
    End If
    
    Me.Invisible = WOClosing.Invisible
        
    
Else

    Set WOClosing = New WOClosing
    Me.CompleteToggle.Picture = ""
    Clear_Field
End If
    
    Me!SubFormWorkOrder.Form!WONumber = WO.WONumber
    Me!SubFormWorkOrder.Form!AssetNumber = WO.AssetNumber
    Me!SubFormWorkOrder.Form!DueDate = WO.DueDate
    Me!SubFormWorkOrder.Form!RequestedBy = WO.RequestBy
    Me!SubFormWorkOrder.Form!RequestedDate = WO.RequestedDate
'    Me!SubFormWorkOrder.Form!WOType = WO.WOType
    Me!SubFormWorkOrder.Form!Manufacturer = WO.Manufacturer
    Me!SubFormWorkOrder.Form!Status = WO.Status
    Me!SubFormWorkOrder.Form!WORequest = WO.WORequest
    Me!SubFormWorkOrder.Form!Comment = WO.EngineeringComment
    Me!SubFormWorkOrder.Form!Description = WO.WODescription
    

'    Me!SubFormWorkOrder.Form!PriorityCombo = WO.Priority
    
'    Set WO = Nothing

End Sub
