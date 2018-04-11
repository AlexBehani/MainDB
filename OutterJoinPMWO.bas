Attribute VB_Name = "OutterJoinPMWO"
Option Compare Database
Option Explicit

Function RightOUtterJOin() As Integer

Dim db As Database
Dim NewDateR As Recordset
Dim NewEq As Recordset
Dim PMWO As Recordset
Dim Str As String
Dim i As Integer
Dim j As Integer
Dim AvailNo As Integer

Set db = CurrentDb

Set PMWO = db.OpenRecordset("PMWO")

Set NewEq = db.OpenRecordset("SELECT * FROM TempStorePMWO " & _
                            "LEFT JOIN PMWO ON TempStorePMWO.GPMid = PMWO.GPMid " & _
                            "WHERE PMWO.GPMid IS Null")

                            
j = 0
AvailNo = MinAvailPMWONo

If NewEq.RecordCount > 0 Then
    
    NewEq.MoveFirst
    Do While Not NewEq.EOF
        PMWO.AddNew
        For i = 0 To 24
            If i = 16 Then
                PMWO(i) = AvailNo
            ElseIf i = 21 Then
                PMWO(i) = PMWOFormatNo("PMWO", AvailNo)
            Else
    
                PMWO(i) = NewEq(i)
            End If

        Next i
        PMWO.Update

        NewEq.MoveNext
        j = j + 1
        AvailNo = AvailNo + 1
    
    Loop
    

End If



Set NewDateR = db.OpenRecordset("SELECT TempStorePMWO.WOID, TempStorePMWO.WODescription, " & _
                            "TempStorePMWO.ModelNumber, TempStorePMWO.WOType, TempStorePMWO.WORequest, " & _
                            "TempStorePMWO.AssignedTo, TempStorePMWO.Status, TempStorePMWO.Completed, " & _
                            "TempStorePMWO.RequestedDate, TempStorePMWO.DateDone, TempStorePMWO.taskComment, " & _
                            "TempStorePMWO.DueDate, TempStorePMWO.WONumber, TempStorePMWO.AssetNumber, " & _
                            "TempStorePMWO.Manufacturer, TempStorePMWO.EngineeringComment, TempStorePMWO.RequestBy, " & _
                            "TempStorePMWO.FormatWONumber, TempStorePMWO.GPMid, TempStorePMWO.DateRegistered " & _
                            "FROM TempStorePMWO LEFT JOIN PMWO ON TempStorePMWO.GPMid=PMWO.GPMid AND " & _
                            "TempStorePMWO.DateRegistered=PMWO.DateRegistered " & _
                            "WHERE IsNull(PMWO.DateRegistered);")

If NewDateR.RecordCount > 0 Then


NewDateR.MoveFirst

    Do While Not NewDateR.EOF
        PMWO.AddNew
        PMWO!WODescription = NewDateR!WODescription
        PMWO!ModelNumber = NewDateR!ModelNumber
        PMWO!WOType = Nz(NewDateR!WOType, "")
        PMWO!WORequest = NewDateR!WORequest
        PMWO!AssignedTo = NewDateR!AssignedTo
        PMWO!Status = NewDateR!Status
        PMWO!Completed = False
        PMWO!RequestedDate = NewDateR!RequestedDate
        PMWO!DateDone = NewDateR!DateDone
        PMWO!TaskComment = NewDateR!TaskComment
        PMWO!DueDate = NewDateR!DueDate
        PMWO!WONumber = AvailNo
        PMWO!AssetNumber = Nz(NewDateR!AssetNumber, "")
        PMWO!Manufacturer = NewDateR!Manufacturer
        PMWO!EngineeringComment = NewDateR!EngineeringComment
        PMWO!RequestBy = NewDateR!RequestBy
        PMWO!formatwonumber = PMWOFormatNo("PMWO", AvailNo)
        PMWO!gpmid = NewDateR!gpmid
        PMWO!DateRegistered = NewDateR!DateRegistered
      
        PMWO.Update
        AvailNo = AvailNo + 1
        j = j + 1
        NewDateR.MoveNext
        
    Loop
End If



Set PMWO = Nothing
Set NewEq = Nothing
Set NewDateR = Nothing
Set db = Nothing
RightOUtterJOin = j

                          
End Function
