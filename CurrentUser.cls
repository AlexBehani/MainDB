VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CurrentUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private CurrentUser As String
Public FName As String
Public LName As String


Property Let User(value As String)
CurrentUser = value
End Property

Property Get User() As String
User = CurrentUser
End Property


Function FullName()
FullName = FName & " " & LName
End Function

