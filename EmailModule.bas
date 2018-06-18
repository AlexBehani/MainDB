Attribute VB_Name = "EmailModule"
Option Compare Database
Option Explicit

Public Sub AlertQuality(WON As String)

Dim Em As Recordset
Dim Eml As Recordset
Dim db As Database

If CheckEmailSent(WON) Then Exit Sub

Set db = CurrentDb
Set Em = GetEmailAddress(3)
Set Eml = db.OpenRecordset("EmailTable")



Em.MoveFirst
Do While Not Em.EOF
Eml.AddNew
With Eml
    !FullName = Em!FName & " " & Em!LName
    !EmailAddress = Em!EmailAddress
    !Status = "Pending"
    !EmailSubject = "Work Order = " & WON
    !EmailContent = "Hi " & Em!FName & "," & vbNewLine & vbNewLine & "Please review " & _
    "the Work Order #" & WON & ". The WO has been selected for Quality Review by " & CUser.FullName & _
    vbNewLine & vbNewLine & "Best regards,"
    !EmailTopic = WON
    .Update


End With
Em.MoveNext
Loop


Set Em = Nothing
Set Eml = Nothing
Set db = Nothing

End Sub

Function GetEmailAddress(n As Integer) As Recordset
Dim db As Database
Dim Em As Recordset

Set db = CurrentDb
Set Em = db.OpenRecordset("SELECT EmailAddress, FName, LName FROM Users WHERE Recipient = True AND AccessLevel =" & n, dbReadOnly)

Set GetEmailAddress = Em

Set db = Nothing
Set Em = Nothing

End Function


Function GetEmailAddressByID(UserID As Integer) As Recordset
Dim db As Database
Dim Em As Recordset

Set db = CurrentDb
Set Em = db.OpenRecordset("SELECT EmailAddress, FName, LName FROM Users WHERE Recipient = True AND UserID =" & UserID, dbReadOnly)

Set GetEmailAddressByID = Em

Set db = Nothing
Set Em = Nothing

End Function


Function GetEmailAddressByID_RstPass(UserID As Integer) As Recordset
Dim db As Database
Dim Em As Recordset

Set db = CurrentDb
Set Em = db.OpenRecordset("SELECT EmailAddress, FName, LName FROM Users WHERE UserID =" & UserID, dbReadOnly)

Set GetEmailAddressByID_RstPass = Em

Set db = Nothing
Set Em = Nothing

End Function

Function CheckEmailSent(WON As String) As Boolean
On Error GoTo Err_Handel
Dim db As Database
Dim EmT As Recordset

Set db = CurrentDb
Set EmT = db.OpenRecordset("SELECT id FROM EmailTable WHERE EmailTopic ='" & WON & "'")

If EmT.RecordCount > 0 Then
    CheckEmailSent = True
Else
    CheckEmailSent = False
End If

Set db = Nothing
Set EmT = Nothing
Exit Function
Err_Handel:
If Err.Number = 3044 Then
MsgBox "Error happend while trying to connect to the Back-end" & vbNewLine & _
"Database was not able to send emial", vbCritical, ""
Else
MsgBox Err.Description
End If
Resume Next
End Function


Public Sub LockedUserEmail(User As String)

Dim Em As Recordset
Dim Eml As Recordset
Dim db As Database

Set db = CurrentDb
Set Em = GetEmailAddress(1)
Set Eml = db.OpenRecordset("EmailTable")



Em.MoveFirst
Do While Not Em.EOF
Eml.AddNew
With Eml
    !FullName = Em!FName & " " & Em!LName
    !EmailAddress = Em!EmailAddress
    !Status = "Pending"
    !EmailSubject = User & " - CMMS account has been locked"
    !EmailContent = "Hi " & Em!FName & "," & vbNewLine & vbNewLine & "Please be informed " & _
    "that " & User & " account has been locked" & _
    vbNewLine & vbNewLine & "Best regards,"
    !EmailTopic = User & " account locked"
    .Update


End With
Em.MoveNext
Loop


Set Em = Nothing
Set Eml = Nothing
Set db = Nothing

End Sub

Public Sub EmailPassword(UserID As Integer, Pass As String)
Dim Em As Recordset
Dim Eml As Recordset
Dim db As Database

Set db = CurrentDb
Set Em = GetEmailAddressByID_RstPass(UserID)
Set Eml = db.OpenRecordset("EmailTable")


Em.MoveFirst
Do While Not Em.EOF
Eml.AddNew
With Eml
    !FullName = Em!FName & " " & Em!LName
    !EmailAddress = Em!EmailAddress
    !Status = "Pending"
    !EmailSubject = Em!FName & " " & Em!LName & " - Temporary Password"
    !EmailContent = "Hi " & Em!FName & "," & vbNewLine & vbNewLine & "Your temporary password is " & _
    Pass & _
    vbNewLine & vbNewLine & "Best regards,"
    !EmailTopic = "Temporary Password"
    .Update


End With
Em.MoveNext
Loop


Set Em = Nothing
Set Eml = Nothing
Set db = Nothing
End Sub

Public Sub EmailNewRegisteredUser(FName As String, LName As String, EmailAdd As String, Pass As String)
Dim Eml As Recordset
Dim db As Database

Set db = CurrentDb
Set Eml = db.OpenRecordset("EmailTable")

Eml.AddNew
With Eml
    !FullName = FName & " " & LName
    !EmailAddress = EmailAdd
    !Status = "Pending"
    !EmailSubject = "CMMS Database"
    !EmailContent = "Hi " & FName & "," & vbNewLine & vbNewLine & "You've been added to the CMMS database. Following are your credentials: " & _
    vbNewLine & vbNewLine & "User name: " & FName & LName & vbNewLine & "Password: " & Pass & _
    vbNewLine & vbNewLine & "Best regards,"
    !EmailTopic = "New Register - Temporary Password"
    .Update


End With

Set Eml = Nothing
Set db = Nothing
End Sub
