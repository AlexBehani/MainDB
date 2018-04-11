Attribute VB_Name = "ClearFields"
Option Compare Database

Public Sub Clear_Field()
On Error GoTo err_handel

Dim ctrl As Control



'  For Each ctrl In Screen.ActiveForm.Controls
'MsgBox ctrl.Tag
'    If (Len(ctrl.Tag) <> 0 And ctrl.ControlType <> acCheckBox And ctrl.Visible = True) Then
'Debug.Print Screen.ActiveControl.Form.Tag
'Debug.Print ctrl.value
'        Set ctrl.value = Null
'    End If
'Next



Exit Sub
err_handel:
If (Err.Number = 2455 Or Err.Number = 438) Then
Resume Next
Else
MsgBox Err.Description & Err.Number
Stop
End If
End Sub
