VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GPMTaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CancelCommand_Click()
DoCmd.Close acForm, "GPMTaskList"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Form_Load()

Dim Row As String

Row = ReturnPMRecord(WO.AssetNumber)
Me.GPMList.RowSource = Row
 

End Sub

Private Sub GPMList_DblClick(Cancel As Integer)
On Error GoTo Err
  
  If (MsgBox("If you press Yes, you will update Date Registerd from " & vbNewLine & Me.GPMList.Column(0) & _
  " to " & WOClosing.DateDone & vbNewLine & vbNewLine & "Are you sure?", vbQuestion + vbYesNo, "") = vbYes) Then

        Call UPdate_DateRegister(Me.GPMList.Column(5), WOClosing.DateDone)
        
        
  End If
  
Me.GPMList.RowSource = ReturnPMRecord(WO.AssetNumber)
Me.GPMList.Requery
 
Exit Sub
Err:
MsgBox Err.Number & vbNewLine & Err.Description
Resume Next
End Sub
