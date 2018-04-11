VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EquipmentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Cancel_Click()

If Not (Equip Is Nothing) Then Set Equip = Nothing
DoCmd.Close acForm, "EquipmentForm"
DoCmd.OpenForm "MainMenu"

End Sub

Private Sub Delete_Click()

If (Nz(Equip.ID, 0) = 0) Then Exit Sub

If (MsgBox("You're about to delete a record" & vbNewLine & "Are you sure?", vbQuestion + vbYesNo, "Delete?") = vbNo) Then Exit Sub
Call Delete_Eq
Clear_Field
'Me.ServiceCheck = 0
Me.Equipments_subform.Requery

End Sub

Private Sub Form_Current()
Set Equip = New Equipments
Equip.Edit = False



End Sub

Private Sub PMCombo_AfterUpdate()

Dim Str As String
Str = ""

If Not (Nz(Me.PMList.ListCount, 0) = 0) Then
    
    For i = 0 To Me.PMList.ListCount
    
        If Me.PMCombo.Column(3) = Me.PMList.Column(3, i) Then
        
            MsgBox "No duplication allwoed", vbCritical, "Error"
            Exit Sub
        End If
    Next
    
End If

For i = 0 To 3
Str = Str & Me.PMCombo.Column(i) & ";"

Next i

Me.PMList.AddItem Str


End Sub

Private Sub PMList_DblClick(Cancel As Integer)
If (Nz(Me.PMList.value, "") = "") Then Exit Sub
Me.PMList.RemoveItem (Me.PMList.value)
Me.PMList.value = ""

Me.PMList.Requery
End Sub

Private Sub Save_Click()
If (Me.StatusCombo = "Spare" And Me.EquipmentTypeCombo = "Instrument") Then
    Me.AssetNumber.Tag = ""
    Me.EquipmentN.Tag = ""
    Me.SystemCombo.Tag = ""
End If

If (Em_Field) Then Exit Sub


Dim arr() As Variant
Dim i As Integer
Dim j As Integer

j = Me.PMList.ListCount

If Not (j > 0) Then
MsgBox "Please select PM", vbCritical, ""
Exit Sub
End If
'i = Me.PMList.ListCount
'ReDim test(1 To i)
ReDim arr(j)

For i = 0 To j - 1

    arr(i) = Me.PMList.Column(3, i)

Next i


Equip.EquipmentNumber = Nz(Me.EquipmentN, 0)
Equip.Description = Me.Description
Equip.AssetNumber = Nz(Me.AssetNumber, "")
Equip.SerialNumber = Me.SerialNumber
Equip.ModelNumber = Me.ModelNumber
Equip.DateRegistered = Me.DateReqistered
Equip.EquipmentType = Me.EquipmentTypeCombo
Equip.Manufacturer = Me.Manufacturer
Equip.System = Nz(Me.SystemCombo, "")
Equip.Status = Me.StatusCombo

If Equip.Status = "In-Service" Or (Equip.Status = "Spare" And Equip.EquipmentType = "Instrument") Then
    Call Save_GenPM(arr())
End If
Call Save_Eq
Call Clear_Field
Me.PMList.RowSource = ""
Me.Equipments_subform.Requery



End Sub


Private Sub StatusCombo_AfterUpdate()
If Me.StatusCombo = "Out-Of-Service" Then
    MsgBox "Database will not generate PM for " & Me.StatusCombo & " status", vbCritical, "Please note"
End If
End Sub
