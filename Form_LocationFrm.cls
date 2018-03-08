VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LocationFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub AddToList_Click()
If (Nz(Me.LocationText, "") = "") Then

    MsgBox "Location box is empty", vbCritical, "Null value"
    Exit Sub
    
End If

If (Unique_value("Locations", "Location", Me.LocationText)) Then

    Call Update_Location_List(Me.LocationText, "Add")
    Me.LocationText = ""
    Me.Location_list.Requery
    
Else
    MsgBox "Duplicated Item", vbCritical, "Duplication not allowed"
    
End If

End Sub

Private Sub Command7_Click()
DoCmd.Close acForm, "LocationFrm"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub RemoveFromList_Click()

If (Me.Location_list.ItemsSelected.Count = 0) Then

    MsgBox "Select the item you would like to remove", vbInformation, "Select the Item"
    Exit Sub
End If

Call Update_Location_List(Me.Location_list.Column(0), "Remove")
Me.Location_list.Requery


End Sub
