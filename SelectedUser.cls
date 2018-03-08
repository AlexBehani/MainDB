VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectedUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public FName As String
Public LName As String
Public FullName As String
Public UserID As Integer



Property Get GetFullName() As String
GetFullName = FName & " " & LName

End Property
