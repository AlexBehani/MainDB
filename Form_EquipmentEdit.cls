VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EquipmentEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Cancel_Click()

If Not (Equip Is Nothing) Then Set Equip = Nothing
DoCmd.Close acForm, "EquipmentEdit"
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



Private Sub Form_Load()
Dim Row As Variant
Dim i As Integer
Dim j As Integer

Me.AssetNumber = Equip.AssetNumber
Me.Description = Equip.Description
Me.EquipmentN = Equip.EquipmentNumber
Me.EquipmentTypeCombo = Equip.EquipmentType
Me.SerialNumber = Equip.SerialNumber
Me.ModelNumber = Equip.ModelNumber
Me.Manufacturer = Equip.Manufacturer
Me.StatusCombo = Equip.Status
Me.SystemCombo = Equip.System
Me.DateReqistered = Equip.DateRegistered
Row = PMListRow(Equip.ID)

i = UBound(Row)

For j = 0 To i - 1

    Me.PMList.AddItem (Row(j, 0) & ";" & Row(j, 1) & ";" & Row(j, 2) & ";" & Row(j, 3))

Next j


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

Dim RowNumber As Long

If (Me.StatusCombo = "Spare" And Me.EquipmentTypeCombo = "Instrument") Then
    Me.AssetNumber.Tag = ""
    Me.EquipmentN.Tag = ""
    Me.SystemCombo.Tag = ""
End If

If (Em_Field) Then Exit Sub


Dim Arr() As Variant
Dim i As Integer
Dim j As Integer

j = Me.PMList.ListCount

If Not (j > 0) Then
MsgBox "Please select PM", vbCritical, ""
Exit Sub
End If

ReDim Arr(j)

For i = 0 To j - 1

    Arr(i) = Me.PMList.Column(3, i)

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


Update_Eq (Equip.ID)
DeleteGenPM (Equip.ID)

If Equip.Status = "In-Service" Or (Equip.Status = "Spare" And Equip.EquipmentType = "Instrument") Then
    Call Save_GenPM(Arr(), Equip.ID)
End If

'Call Clear_Field
'Me.PMList.RowSource = ""
'Me.Equipments_subform.Requery

DoCmd.OpenForm "MainMenu"
DoCmd.Close acForm, "EquipmentEdit"



End Sub


Private Sub StatusCombo_AfterUpdate()
If Me.StatusCombo = "Out-Of-Service" Then
    MsgBox "Database will not generate PM for " & Me.StatusCombo & " status", vbCritical, "Please note"
End If
End Sub
