Attribute VB_Name = "WONumber"
Option Compare Database

Public Function PMWOFormatNo(pre As String, n As Integer)
 
Dim C As Integer, L As Integer
Dim Str As String, j As Integer

C = 6
 Str = ""

 L = C - Len(CStr(n))

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

Public Function PMWORowlist(pre As String, Optional Filter As String) As String


Dim db As Database
Dim PMWO As Recordset
Dim Str As String
Set db = CurrentDb

If Filter = "New" Then

    Set PMWO = db.OpenRecordset("SELECT WONumber, WOID,FormatWONumber FROM PMWO WHERE EngineeringComment IS NULL")
    
    If Not (PMWO.RecordCount > 0) Then
        pmworlist = ""
        Exit Function
    End If

    PMWO.MoveFirst

    Do While Not PMWO.EOF

        Str = Str & PMWO!formatwonumber & ";" & PMWO!WOID & ";"
        PMWO.MoveNext
    Loop

ElseIf Filter = "Existed ones" Then

    Set PMWO = db.OpenRecordset("SELECT WONumber, WOID,FormatWONumber FROM PMWO WHERE EngineeringComment IS NOT NULL")
    
    If Not (PMWO.RecordCount > 0) Then
        pmworlist = ""
        Exit Function
    End If

    PMWO.MoveFirst

    Do While Not PMWO.EOF

        Str = Str & PMWO!formatwonumber & ";" & PMWO!WOID & ";"
        PMWO.MoveNext
    Loop
    
Else
    Set PMWO = db.OpenRecordset("SELECT WONumber, WOID,FormatWONumber FROM PMWO")
    
    If Not (PMWO.RecordCount > 0) Then
        pmworlist = ""
        Exit Function
    End If

    PMWO.MoveFirst

    Do While Not PMWO.EOF

        Str = Str & PMWO!formatwonumber & ";" & PMWO!WOID & ";"
        PMWO.MoveNext
    Loop

End If


PMWORowlist = Str

End Function





