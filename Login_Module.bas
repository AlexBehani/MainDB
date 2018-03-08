Attribute VB_Name = "Login_Module"
Option Compare Database
Public CUser As CurrentUser
Public SUser As SelectedUser
Public Request As Request
Public WO As WO
Public WOClosing As WOClosing
Public GNote As GeneralNote
Public Equip As Equipments
Public PMTK As PMtask


' login=0 User & Password correct, also no need to update
' login=1 User & Password correct, need to update
' login=2 User & password not correct

Public Function Login(UserName As String, Pass As String) As Integer


Dim db As dao.Database
Dim Rs As dao.Recordset
Dim txt As String


Set db = CurrentDb

Set Rs = db.OpenRecordset("SELECT * FROM Users WHERE UserName='" & UserName & "' AND Password='" & BASE64SHA1(Pass) & "'")

If (Rs.RecordCount > 0) Then

        If (Rs!PWDRst = 0) Then
        
            Login = 0
        ElseIf (Rs!PWDRst = -1) Then
            Login = 1
        End If
    Set CUser = New CurrentUser
    CUser.User = Nz(Rs!FName, "") & " " & Nz(Rs!LName, "")
    CUser.FName = Nz(Rs!FName, "")
    CUser.LName = Nz(Rs!LName, "")
    
Else
Login = 2
    
End If

db.Close
Set Rs = Nothing
Set db = Nothing

End Function


Public Sub Register_User(FName As String, LName As String, Optional var3 As Integer)
Dim db As Database

Dim PWR As String
Dim Rs As Recordset

PWR = "-1"
Set db = CurrentDb
Set Rs = db.OpenRecordset("SELECT * FROM Users WHERE FName = '" & FName & "' AND LName = '" & LName & "'")

If Rs.RecordCount > 0 Then
    
    MsgBox "User info is not unique", vbCritical, "Duplicated information"
    Set Rs = Nothing
    Set db = Nothing
    Exit Sub
End If

Set Rs = db.OpenRecordset("Users")
    
    Rs.AddNew
    Rs!FName = FName
    Rs!LName = LName
    Rs!Password = BASE64SHA1("welcome1")
    Rs!PWDRst = -1
    Rs.Update
    
MsgBox "New User is added", vbInformation, "Done"
Set Rs = Nothing
Set db = Nothing

End Sub


Public Function Change_User_info(FName As String, LName As String, UserID As Integer)
On Error GoTo Err
Dim db As Database
Dim User As Recordset

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT * FROM Users WHERE UserID = " & UserID)
User.MoveFirst
User.Edit
User!FName = FName
User!LName = LName
User.Update


Set db = Nothing
Set Rs = Nothing

Exit Function
Err:
MsgBox Err.Number, vbCritical, "Error"
Set db = Nothing
Set Rs = Nothing

End Function


Public Sub Reset_password(UserID As Integer)
Dim db As Database
Dim User As Recordset

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT * FROM Users WHERE UserID = " & UserID)

User.MoveFirst
User.Edit
User!Password = BASE64SHA1("welcome")
User.Update

Set db = Nothing
Set User = Nothing
End Sub


Public Sub DeleteUser(UserID As Integer)
Dim db As Database
Dim User As Recordset

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT * FROM Users WHERE UserID = " & UserID)

User.MoveFirst
User.Delete



Set db = Nothing
Set User = Nothing
End Sub



Public Sub UpdatePassword(Pass As String)

Dim db As Database
Dim User As Recordset

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT * FROM Users WHERE FName = '" & CUser.FName & "' AND LName = '" & CUser.LName & "'")

User.MoveFirst
User.Edit
User!Password = BASE64SHA1(Pass)
User!PWDRst = 0
User.Update

Set db = Nothing
Set User = Nothing




End Sub
