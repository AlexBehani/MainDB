VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Add_command_Click()

If (Nz(Me.Status, "") = "") Then

    MsgBox "Status box is empty", vbCritical, "Null value"
    Exit Sub
    
End If

If (Unique_value("Statuss", "Status", Me.Status)) Then

    Call Update_Status_List(Me.Status, "Add")
    Me.Status = ""
    Me.Status_list.Requery
    
Else
    MsgBox "Duplicated Item", vbCritical, "Duplication not allowed"
    
End If
End Sub

Private Sub AddImage_Click()

Me.AddImage.SpecialEffect = 2
Me.Status_list.SetFocus
If (Nz(Me.Status, "") = "") Then

    MsgBox "Status box is empty", vbCritical, "Null value"
    Me.AddImage.SpecialEffect = 0
    Exit Sub
    
End If

If (Unique_value("Status", "Status", Me.Status)) Then

    Call Update_Status_List(Me.Status, "Add")
    Me.Status = ""
    Me.Status_list.Requery
    
Else
    MsgBox "Duplicated Item", vbCritical, "Duplication not allowed"
    
End If

Me.AddImage.SpecialEffect = 0

End Sub

Private Sub Cancel_Click()
DoCmd.Close acForm, "status"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub MinusImage_Click()

Me.MinusImage.SpecialEffect = 2
Me.Status_list.SetFocus
If (Me.Status_list.ItemsSelected.Count = 0) Then

    MsgBox "Select the item you would like to remove", vbInformation, "Select the Item"
    Me.MinusImage.SpecialEffect = 0
    Exit Sub
End If

Call Update_Status_List(Me.Status_list.Column(0), "Remove")
Me.Status_list.Requery
Me.MinusImage.SpecialEffect = 0

End Sub

Private Sub Remove_command_Click()


If (Me.Status_list.ItemsSelected.Count = 0) Then

    MsgBox "Select the item you would like to remove", vbInformation, "Select the Item"
    Exit Sub
End If

Call Update_Status_List(Me.Status_list.Column(0), "Remove")
Me.Status_list.Requery
End Sub
