VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Priority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddImage_Click()

Me.AddImage.SpecialEffect = 2
Me.Priority_list.SetFocus
If (Nz(Me.Priority, "") = "") Then

    MsgBox "Priority box is empty", vbCritical, "Empty value"
    Me.AddImage.SpecialEffect = 0
    Exit Sub
    
End If

If (Unique_value("Priority", "Priority", Me.Priority)) Then

    Call Update_Priority_List(Me.Priority, "Add")
    Me.Priority = ""
    Me.Priority_list.Requery
    
Else
    MsgBox "Duplicated Item", vbCritical, "Duplication not allowed"
    
End If

Me.AddImage.SpecialEffect = 0
End Sub

Private Sub Cancel_Click()
DoCmd.Close acForm, "Priority"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub MinusImage_Click()

Me.MinusImage.SpecialEffect = 2
Me.Priority_list.SetFocus
If (Me.Priority_list.ItemsSelected.Count = 0) Then

    MsgBox "Select the item you would like to remove", vbInformation, "Select the Item"
    Me.MinusImage.SpecialEffect = 0
    Exit Sub
End If

Call Update_Priority_List(Me.Priority_list.Column(0), "Remove")
Me.Priority_list.Requery
Me.MinusImage.SpecialEffect = 0

End Sub
