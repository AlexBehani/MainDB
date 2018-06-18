Attribute VB_Name = "Login_Module"
Option Compare Database
Public CUser As CurrentUser
Public SUser As SelectedUser
Public Request As Request
Public WO As WO
Public WOClosing As WOClosing
Public GNote As GeneralNote
Public Equip As Equipments
Public PMTK As PMTask
Public PMWO As PMWO


' login=0 User & Password correct, also no need to update
' login=1 User & Password correct, need to update
' login=2 User & password not correct

Public Function Login(UserName As String, Pass As String) As Integer


Dim db As dao.Database
Dim Rs As dao.Recordset
Dim txt As String
Dim PassTxt As String


Set db = CurrentDb

PassTxt = BASE64SHA1(Pass)
Set Rs = db.OpenRecordset("SELECT * FROM Users WHERE UserName='" & UserName & "' AND Password='" & PassTxt & "'")

If (Rs.RecordCount > 0) Then

        If (Rs!pwdrst = 0) Then
        
            Login = 0
        ElseIf (Rs!pwdrst = -1) Then
            Login = 1
        End If
    Set CUser = New CurrentUser
    CUser.User = Nz(Rs!FName, "") & " " & Nz(Rs!LName, "")
    CUser.FName = Nz(Rs!FName, "")
    CUser.LName = Nz(Rs!LName, "")
    CUser.AccessLevel = Nz(Rs!AccessLevel, 0)
    
    
Else
Login = 2
    
End If

db.Close
Set Rs = Nothing
Set db = Nothing

End Function


Public Function Register_User(FName As String, LName As String, var3 As Integer, EmailAdd As String)
Dim db As Database

Dim PWR As String
Dim Rs As Recordset
Dim Pss As String

Pass = RandString
PWR = "-1"
Set db = CurrentDb
Set Rs = db.OpenRecordset("SELECT * FROM Users WHERE FName = '" & FName & "' AND LName = '" & LName & "'")

If Rs.RecordCount > 0 Then
    
    MsgBox "User info is not unique", vbCritical, "Duplicated information"
    Set Rs = Nothing
    Set db = Nothing
    Exit Function
End If

Set Rs = db.OpenRecordset("Users")
    
    Rs.AddNew
    Rs!FName = FName
    Rs!LName = LName
    Rs!Password = BASE64SHA1(Pass)
    Rs!pwdrst = -1
    Rs!UserName = FName & LName
    Rs!AccessLevel = var3
    Rs!EmailAddress = EmailAdd
    Rs.Update
    
Register_User = Pass
MsgBox "New User is added", vbInformation, "Done"
Set Rs = Nothing
Set db = Nothing

End Function


Public Function Change_User_info(FName As String, LName As String, UserID As Integer, AccessLevel As Integer, EmailAdd As String)
On Error GoTo Err
Dim db As Database
Dim User As Recordset

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT FName, LName, AccessLevel, UserName, EmailAddress FROM Users WHERE UserID = " & UserID)
User.MoveFirst
User.Edit
User!FName = FName
User!LName = LName
User!AccessLevel = AccessLevel
User!UserName = FName & LName
User!EmailAddress = EmailAdd
User.Update


Set db = Nothing
Set Rs = Nothing

Exit Function
Err:
MsgBox Err.Number, vbCritical, "Error"
Set db = Nothing
Set Rs = Nothing

End Function


Public Function Reset_password(UserID As Integer) As String
Dim db As Database
Dim User As Recordset
Dim Str

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT * FROM Users WHERE UserID = " & UserID)

Str = RandString()
User.MoveFirst
User.Edit
User!Password = BASE64SHA1(Str)
User!pwdrst = -1
User!Locked = False
User!NoAttempt = 0
User.Update
Reset_password = Str

Set db = Nothing
Set User = Nothing
End Function

Function RandString()

    Dim s As String * 8
    Dim n As Integer
    Dim ch As Integer
    For n = 1 To Len(s)
        Do
            ch = Rnd() * 127
            
        Loop While ch < 48 Or ch > 57 And ch < 65 Or ch > 90 And ch < 97 Or ch > 122
        Mid(s, n, 1) = Chr(ch)
    Next

    RandString = s

End Function

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
User!pwdrst = False
User.Update

Set db = Nothing
Set User = Nothing

End Sub

Public Function CheckUserLocked(User As String) As Boolean

Dim db As Database
Dim Users As Recordset

Set db = CurrentDb
Set Users = db.OpenRecordset("SELECT Locked FROM Users WHERE UserName ='" & User & "'")
If Users.RecordCount > 0 Then
    Users.MoveFirst
    If Users!Locked Then
        CheckUserLocked = True
        Set db = Nothing
        Set Users = Nothing
        Exit Function
    End If
End If
    
CheckUserLocked = False


End Function


Public Sub ResetNoAttempt(User As String)

Dim db As Database
Dim Users As Recordset

Set db = CurrentDb
Set Users = db.OpenRecordset("SELECT NoAttempt FROM Users WHERE UserName ='" & User & "'")
    If Users.RecordCount > 0 Then
        Users.Edit
        Users!NoAttempt = 0
        Users.Update
    End If

Set db = Nothing
Set Users = Nothing
End Sub


Public Function AggregateNoAttempt(User As String) As Boolean
On Error GoTo Err_Handel
Dim db As Database
Dim Users As Recordset

Set db = CurrentDb
Set Users = db.OpenRecordset("SELECT Locked, NoAttempt, FName, LName FROM Users WHERE UserName = '" & User & "'")
If Users.RecordCount > 0 Then

    Users.MoveFirst
    Users.Edit
    Users!NoAttempt = Nz(Users!NoAttempt, 0) + 1
    If Users!NoAttempt > 2 Then
        Users!Locked = True
        AggregateNoAttempt = True
    End If
    Users.Update
    If Users!NoAttempt > 2 Then
        LockedUserEmail (Users!FName & " " & Users!LName)
    End If
    Set db = Nothing
    Set Users = Nothing
    Exit Function
End If

Set db = Nothing
Set Users = Nothing

Exit Function
Err_Handel:
Set db = Nothing
Set Users = Nothing
AggregateNoAttempt = False
Resume Next


End Function




