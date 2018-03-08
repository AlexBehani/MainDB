VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command6_Click()
Dim UserName As String
Dim Password As String
Dim Condition As Integer

UserName = Nz(Me.UserName, "")
Password = Nz(Me.Password, "")
If (UserName = "" Or Password = "") Then
MsgBox "Please enter User name and Password", vbCritical + vbExclamation, "Invalid Input"
Exit Sub
End If

Condition = Login(UserName, Password)
If (Condition = 0) Then

    DoCmd.Close acForm, "Login"
    DoCmd.OpenForm ("MainMenu")
    
ElseIf (Condition = 2) Then

    MsgBox "User Name and Password does not match", vbCritical, "Invalid User name and Password"
    Me.UserName = ""
    Me.Password = ""

ElseIf (Condition = 1) Then
    
    DoCmd.Close acForm, "Login"
    DoCmd.OpenForm ("Update_Password")
End If
End Sub
