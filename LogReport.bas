Attribute VB_Name = "LogReport"
Option Compare Database
Option Explicit

Public Sub LoginReport()

Dim db As Database
Dim Lg As Recordset

Set db = CurrentDb
Set Lg = db.OpenRecordset("LogTable")

Lg.AddNew
With Lg
    !User = CUser.FullName
    !logdate = Date
    !logtime = Time()
    !Status = "Log in"
    .Update
End With

Set db = Nothing
Set Lg = Nothing
    
End Sub

Public Sub LogOutReport()

Dim db As Database
Dim Lg As Recordset

Set db = CurrentDb
Set Lg = db.OpenRecordset("LogTable")

Lg.AddNew
With Lg
    !User = CUser.FullName
    !logdate = Date
    !logtime = Time()
    !Status = "Log out"
    .Update
End With

Set db = Nothing
Set Lg = Nothing
    
End Sub

Public Sub LogError(UserName As String)

Dim db As Database
Dim Lg As Recordset

Set db = CurrentDb
Set Lg = db.OpenRecordset("LogTable")

Lg.AddNew
With Lg
    !User = UserName
    !logdate = Date
    !logtime = Time()
    !Status = "Log Error"
    .Update
End With

Set db = Nothing
Set Lg = Nothing
End Sub
