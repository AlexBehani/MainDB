VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_WOInterfaceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
Me.PMWorkOrderCombo.RowSource = PMWORowlist("PMWO")

End Sub

Private Sub GenerateWO_Click()

Set WO = New WO
WO.WONumber = MinAvailWONo
WO.FormatedWONUmber = PMWOFormatNo("WO", WO.WONumber)
WO.Edit = False




DoCmd.Close acForm, "WOInterfaceForm"
DoCmd.OpenForm "WorkOrderForm"
End Sub

Private Sub PMWorkOrderCombo_AfterUpdate()
'Dim str As String
'
'str = Me.PMWorkOrderCombo.Column(0)
Set PMWO = Load_PMWO(Me.PMWorkOrderCombo.Column(1))
DoCmd.Close acForm, "WOInterfaceForm"
DoCmd.OpenForm "PMWorkOrderForm"
End Sub

Private Sub PMWorkOrderCombo_Change()
'MsgBox "hello"
End Sub
