Attribute VB_Name = "WONumber"
Option Compare Database

Public Function PMWOFormatNo(pre As String, n As Integer)
 
Dim C As Integer, L As Integer
Dim Str As String, j As Integer

C = 6
 Str = ""
 L = C - Len(n)

For j = 0 To L - 1
    Str = Str & 0
Next j

PMWOFormatNo = pre & Str & n
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
Dim Str As String


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

Public Function PMWORowlist(pre As String) As String


Dim db As Database
Dim PMWO As Recordset
Dim Str As String

Set db = CurrentDb
Set PMWO = db.OpenRecordset("SELECT WONumber, WOID FROM PMWO")
If Not (PMWO.RecordCount > 0) Then
    pmworlist = ""
    Exit Function
End If

PMWO.MoveFirst

Do While Not PMWO.EOF

    Str = Str & PMWOFormatNo(pre, PMWO!WONumber) & ";" & PMWO!WOID & ";"
    PMWO.MoveNext
Loop

PMWORowlist = Str

End Function




