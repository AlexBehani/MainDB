Attribute VB_Name = "OutterJoinPMWO"
Option Compare Database
Option Explicit

Function RightOUtterJOin() As Integer

Dim db As Database
Dim HyRs As Recordset
Dim PMWO As Recordset
Dim Str As String
Dim i As Integer
Dim j As Integer
Dim AvailNo As Integer

j = 0
Set db = CurrentDb

Set PMWO = db.OpenRecordset("PMWO")

Set HyRs = db.OpenRecordset("SELECT * FROM TempStorePMWO " & _
                            "LEFT JOIN PMWO ON TempStorePMWO.GPMid = PMWO.GPMid " & _
                            "WHERE PMWO.GPMid IS Null")
                            
If HyRs.RecordCount > 0 Then
AvailNo = MinAvailPMWONo
    
    HyRs.MoveFirst
    Do While Not HyRs.EOF
        PMWO.AddNew
        For i = 0 To 22
        
        Debug.Print PMWO(i) & " i= " & i & " name= " & PMWO(i).Name
        If i = 16 Then
            PMWO(i) = AvailNo
        ElseIf i = 21 Then
            PMWO(i) = PMWOFormatNo("PMWO", AvailNo)
        Else
        
            PMWO(i) = HyRs(i)
        End If
        Next i
        PMWO.Update
        HyRs.MoveNext
        j = j + 1
        AvailNo = AvailNo + 1
    
    Loop
End If

Set PMWO = Nothing
Set HyRs = Nothing
Set db = Nothing
RightOUtterJOin = j


                            
End Function
