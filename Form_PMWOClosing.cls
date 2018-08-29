VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PMWOClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
On Error GoTo Err_Handel
If Not (WOClosing Is Nothing) Then Set WOClosing = Nothing
DoCmd.Close acForm, "PMWOClosing"
DoCmd.OpenForm "MainMenu"
Exit Sub
Err_Handel:
DoCmd.Close acForm, "PMWOClosing"
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

        Me.WODropDown.RowSource = "SELECT PMWO.FormatWONumber, PMWO.AssetNumber, PMWO.Manufacturer, PMWO.DueDate, " & _
                                    "PMWO.WOID, PMWO.taskComment, PMWO.EngineeringComment, PMWO.Completed " & _
                                    "FROM PMWO WHERE ((Nz(PMWO.EngineeringComment,'')<>'') AND " & _
                                    "((PMWO.Completed)=True));"
ElseIf Me.FilterCombo = "Not Completed" Then
    
    Me.WODropDown.RowSource = "SELECT PMWO.FormatWONumber, PMWO.AssetNumber, PMWO.Manufacturer, PMWO.DueDate, " & _
                                    "PMWO.WOID, PMWO.taskComment, PMWO.EngineeringComment, PMWO.Completed " & _
                                    "FROM PMWO WHERE ((Nz(PMWO.EngineeringComment,'')<>'') AND " & _
                                    "((PMWO.Completed)=False));"
End If
    


End Sub

Private Sub Form_Load()
Set WOClosing = New WOClosing
Me.CompleteToggle.Picture = ""
End Sub

Private Sub SaveCommand_Click()

'On Error GoTo Err
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

Call Save_PMWOII
'MsgBox "Finish the clear function", vbOKOnly, ""
'Clear_Field
MsgBox "Done", vbInformation, ""

If Not (WOClosing Is Nothing) Then Set WOClosing = Nothing

DoCmd.Close acForm, "PMWOClosing"
DoCmd.OpenForm "MainMenu"
Exit Sub
Err:
If Not (WOClosing Is Nothing) Then Set WOClosing = Nothing
MsgBox "Error happened!", vbCritical, "Error"
Call Errs(Err.Description, Err.LastDllError, Err.Number, Err.Source)
End Sub

Private Sub WODropDown_AfterUpdate()

Set WO = Load_WO("PMWO", Me.WODropDown.Column(4))

If Not (Nz(Me.WODropDown.Column(5), "") = "") Then
    Set WOClosing = Load_WOClosing("PMWO", Me.WODropDown.Column(4))
    
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
            Me.CompleteToggle = False
    End If
        
    
Else

    Set WOClosing = New WOClosing
    Me.CompleteToggle.Picture = ""
'    Clear_Field
End If
    
    Me!WorkOrder.Form!WONumber = WO.WONumber
    Me!WorkOrder.Form!AssetNumber = WO.AssetNumber
    Me!WorkOrder.Form!DueDate = WO.DueDate
    Me!WorkOrder.Form!RequestedBy = WO.RequestBy
    Me!WorkOrder.Form!RequestedDate = WO.RequestedDate
'    Me!WorkOrder.Form!WOType = WO.WOType
    Me!WorkOrder.Form!Manufacturer = WO.Manufacturer
    Me!WorkOrder.Form!Status = WO.Status
    Me!WorkOrder.Form!WORequest = WO.WORequest
    Me!WorkOrder.Form!Comment = WO.EngineeringComment
    Me!WorkOrder.Form!Description = WO.WODescription
    

'    Me!SubFormWorkOrder.Form!PriorityCombo = WO.Priority
    
    Set WO = Nothing

End Sub


