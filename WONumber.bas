Attribute VB_Name = "WONumber"
Option Compare Database

Public Function PMWOFormatNo(pre As String, n As Integer)
 
Dim C As Integer, L As Integer
Dim str As String, j As Integer

C = 6
 str = ""

 L = C - Len(CStr(n))

For j = 0 To L - 1
    str = str & 0
Next j

PMWOFormatNo = pre & str & n
End Function


Function MinAvailWONo() As Integer

Dim db As Database
Dim WO As Recordset

Set db = CurrentDb
Set WO = db.OpenRecordset("WO")

If WO.RecordCount > 0 Then

    WO.MoveFirst
    MinAvailWONo = DMax("WONumber", "WO") + 1
Else
    MinAvailWONo = 1
End If

Set db = Nothing
Set WO = Nothing

End Function


Function MinAvailPMWONo() As Integer

Dim db As Database
Dim PMWO As Recordset
Dim str As String


Set db = CurrentDb
Set PMWO = db.OpenRecordset("PMWO")

If PMWO.RecordCount > 0 Then

    MinAvailPMWONo = DMax("WONumber", "PMWO") + 1
Else
    MinAvailPMWONo = 1
End If

Set db = Nothing
Set PMWO = Nothing

End Function

Public Function PMWORowlist(pre As String, Optional Filter As String) As String


Dim db As Database
Dim PMWO As Recordset
Dim str As String
Set db = CurrentDb

If Filter = "New" Then

    Set PMWO = db.OpenRecordset("SELECT WONumber, WOID,FormatWONumber FROM PMWO WHERE EngineeringComment IS NULL")
    
    If Not (PMWO.RecordCount > 0) Then
        PMWORowlist = ""
        Exit Function
    End If

    PMWO.MoveFirst

    Do While Not PMWO.EOF

        str = str & PMWO!FormatWONumber & ";" & PMWO!WOID & ";"
        PMWO.MoveNext
    Loop

ElseIf Filter = "Existed ones" Then

    Set PMWO = db.OpenRecordset("SELECT WONumber, WOID,FormatWONumber FROM PMWO WHERE EngineeringComment IS NOT NULL")
    
    If Not (PMWO.RecordCount > 0) Then
        PMWORowlist = ""
        Exit Function
    End If

    PMWO.MoveFirst

    Do While Not PMWO.EOF

        str = str & PMWO!FormatWONumber & ";" & PMWO!WOID & ";"
        PMWO.MoveNext
    Loop
    
Else
    Set PMWO = db.OpenRecordset("SELECT WONumber, WOID,FormatWONumber FROM PMWO")
    
    If Not (PMWO.RecordCount > 0) Then
        PMWORowlist = ""
        Exit Function
    End If

    PMWO.MoveFirst

    Do While Not PMWO.EOF

        str = str & PMWO!FormatWONumber & ";" & PMWO!WOID & ";"
        PMWO.MoveNext
    Loop

End If


PMWORowlist = str

End Function



Public Function WORowlist() As String
On Error GoTo Err_Handel

Dim db As Database
Dim WO As Recordset
Dim str As String
Set db = CurrentDb


If (CUser.AccessLevel > 2) Then
    str = ""
    Set WO = db.OpenRecordset("SELECT FormatWONumber, ID FROM WO WHERE QRrequired =True AND Invisible =False")
    
    If Not (WO.RecordCount > 0) Then
        WORowlist = ""
        Exit Function
    End If

    WO.MoveFirst

    Do While Not WO.EOF

        str = str & WO!FormatWONumber & ";" & WO!ID & ";"
        WO.MoveNext
    Loop

    WORowlist = str

Else
    str = ""
    Set WO = db.OpenRecordset("SELECT FormatWONumber, ID FROM WO WHERE Invisible =False")
    
    If Not (WO.RecordCount > 0) Then
        WORowlist = ""
        Exit Function
    End If

    WO.MoveFirst

    Do While Not WO.EOF

        str = str & WO!FormatWONumber & ";" & WO!ID & ";"
        WO.MoveNext
    Loop

    WORowlist = str
End If

Exit Function
Err_Handel:
If Err.Number = 91 Then
MsgBox "Lost your credentials, please logout and log back in", vbCritical, ""
Set db = Nothing
Set WO = Nothing
WORowlist = ""
Exit Function
Else
Set db = Nothing
Set WO = Nothing
WORowlist = ""
Exit Function
End If

End Function





