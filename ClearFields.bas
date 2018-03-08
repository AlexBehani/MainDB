Attribute VB_Name = "ClearFields"
Option Compare Database

Public Sub Clear_Field()
On Error GoTo err_handel

Dim ctrl As Control


For Each ctrl In Screen.ActiveForm.Controls
    If Len(ctrl.Tag) <> 0 And ctrl.ControlType <> acCheckBox And ctrl.Visible = True Then
        ctrl = Null
    End If
Next



Exit Sub
err_handel:
MsgBox Err.Description

End Sub
