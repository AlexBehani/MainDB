Attribute VB_Name = "OutterJoinPMWO"
Option Compare Database
Option Explicit

'Function RightOUtterJOin() As Integer
'
'Dim db As Database
'Dim NewDateR As Recordset
'Dim NewEq As Recordset
'Dim PMWO As Recordset
'Dim str As String
'Dim i As Integer
'Dim j As Integer
'Dim AvailNo As Integer
'
'Set db = CurrentDb
'
'Set PMWO = db.OpenRecordset("PMWO")
'
''Set NewEq = db.OpenRecordset("SELECT * FROM TempStorePMWO " & _
''                            "LEFT JOIN PMWO ON TempStorePMWO.GPMid = PMWO.GPMid " & _
''                            "WHERE PMWO.GPMid IS Null")
'
'
'Set NewEq = db.OpenRecordset("SELECT TempStorePMWO.WOID, TempStorePMWO.ModelNumber, TempStorePMWO.WORequest, " & _
'"TempStorePMWO.AssignedTo, TempStorePMWO.Status, TempStorePMWO.DueDate, TempStorePMWO.AssetNumber, " & _
'"TempStorePMWO.Manufacturer, TempStorePMWO.RequestBy, TempStorePMWO.GPMid, TempStorePMWO.DateRegistered, " & _
'"TempStorePMWO.EqDescription, TempStorePMWO.WODescription FROM TempStorePMWO " & _
'"LEFT JOIN PMWO ON TempStorePMWO.GPMid = PMWO.GPMid " & _
'"WHERE (((PMWO.GPMid) Is Null));")
'
'
'j = 0
'AvailNo = MinAvailPMWONo
'
'If NewEq.RecordCount > 0 Then
'    NewEq.MoveFirst
'
'    Do While Not NewEq.EOF
'        PMWO.AddNew
'        PMWO!WODescription = NewEq!WODescription
'        PMWO!ModelNumber = NewEq!ModelNumber
''        PMWO!WOType = Nz(NewDateR!WOType, "")
'        PMWO!WORequest = NewEq!WORequest
'        PMWO!AssignedTo = NewEq!AssignedTo
'        PMWO!Status = NewEq!Status
'        PMWO!Completed = False
''        PMWO!RequestedDate = NewEq!RequestedDate
''        PMWO!DateDone = NewEq!DateDone
''        PMWO!TaskComment = NewEq!TaskComment
'        PMWO!DueDate = NewEq!DueDate
'        PMWO!WONumber = AvailNo
'        PMWO!AssetNumber = Nz(NewEq!AssetNumber, "")
'        PMWO!Manufacturer = NewEq!Manufacturer
''        PMWO!EngineeringComment = NewEq!EngineeringComment
'        PMWO!RequestBy = NewEq!RequestBy
'        PMWO!FormatWONumber = PMWOFormatNo("PMWO", AvailNo)
'        PMWO!GPMid = NewEq!GPMid
'        PMWO!DateRegistered = NewEq!DateRegistered
'        PMWO!EqDescription = NewEq!EqDescription
'
'        PMWO.Update
'        AvailNo = AvailNo + 1
'        j = j + 1
'        NewEq.MoveNext
'
'    Loop
'End If
'
'
''If NewEq.RecordCount > 0 Then
''
''    NewEq.MoveFirst
''    Do While Not NewEq.EOF
''        PMWO.AddNew
''        For i = 0 To 25
'''            If i = 16 Then
'''                PMWO(i) = AvailNo
'''            ElseIf i = 21 Then
'''                PMWO(i) = PMWOFormatNo("PMWO", AvailNo)
'''            Else
'''
'''                PMWO(i) = NewEq(i)
'''            End If
''    MsgBox NewEq(i) & "   " & i
''
''        Next i
''        PMWO.Update
''
''        NewEq.MoveNext
''        j = j + 1
''        AvailNo = AvailNo + 1
''
''    Loop
''
''
''End If
'
'
'
'Set NewDateR = db.OpenRecordset("SELECT TempStorePMWO.WOID, TempStorePMWO.WODescription, " & _
'                            "TempStorePMWO.ModelNumber, TempStorePMWO.WORequest, " & _
'                            "TempStorePMWO.AssignedTo, TempStorePMWO.Status, TempStorePMWO.Completed, " & _
'                            "TempStorePMWO.RequestedDate, TempStorePMWO.DateDone, TempStorePMWO.taskComment, " & _
'                            "TempStorePMWO.DueDate, TempStorePMWO.WONumber, TempStorePMWO.AssetNumber, " & _
'                            "TempStorePMWO.Manufacturer, TempStorePMWO.EngineeringComment, TempStorePMWO.RequestBy, TempStorePMWO.EqDescription, " & _
'                            "TempStorePMWO.FormatWONumber, TempStorePMWO.GPMid, TempStorePMWO.DateRegistered " & _
'                            "FROM TempStorePMWO LEFT JOIN PMWO ON TempStorePMWO.GPMid=PMWO.GPMid AND " & _
'                            "TempStorePMWO.DateRegistered=PMWO.DateRegistered " & _
'                            "WHERE IsNull(PMWO.DateRegistered);")
'
'If NewDateR.RecordCount > 0 Then
'
'
'NewDateR.MoveFirst
'
'    Do While Not NewDateR.EOF
'        PMWO.AddNew
'        PMWO!WODescription = NewDateR!WODescription
'        PMWO!ModelNumber = NewDateR!ModelNumber
''        PMWO!WOType = Nz(NewDateR!WOType, "")
'        PMWO!WORequest = NewDateR!WORequest
'        PMWO!AssignedTo = NewDateR!AssignedTo
'        PMWO!Status = NewDateR!Status
'        PMWO!Completed = False
'        PMWO!RequestedDate = NewDateR!RequestedDate
'        PMWO!DateDone = NewDateR!DateDone
'        PMWO!TaskComment = NewDateR!TaskComment
'        PMWO!DueDate = NewDateR!DueDate
'        PMWO!WONumber = AvailNo
'        PMWO!AssetNumber = Nz(NewDateR!AssetNumber, "")
'        PMWO!Manufacturer = NewDateR!Manufacturer
'        PMWO!EngineeringComment = NewDateR!EngineeringComment
'        PMWO!RequestBy = NewDateR!RequestBy
'        PMWO!FormatWONumber = PMWOFormatNo("PMWO", AvailNo)
'        PMWO!GPMid = NewDateR!GPMid
'        PMWO!DateRegistered = NewDateR!DateRegistered
'        PMWO!EqDescription = NewDateR!EqDescription
'
'        PMWO.Update
'        AvailNo = AvailNo + 1
'        j = j + 1
'        NewDateR.MoveNext
'
'    Loop
'End If
'
'
'
'Set PMWO = Nothing
'Set NewEq = Nothing
'Set NewDateR = Nothing
'Set db = Nothing
'RightOUtterJOin = j
'
'
'End Function

Dim i As Integer
Dim AvailNo As Integer
Dim j As Integer

Function RightOUtterJOin() As Integer

Dim db As Database
'Dim NewDateR As Recordset
Dim NewEq As Recordset
Dim PMWO As Recordset
Dim str As String
'Dim i As Integer
'Dim j As Integer
Dim AvailNo As Integer

Set db = CurrentDb

Set PMWO = db.OpenRecordset("PMWO")

'Set NewEq = db.OpenRecordset("SELECT * FROM TempStorePMWO " & _
'                            "LEFT JOIN PMWO ON TempStorePMWO.GPMid = PMWO.GPMid " & _
'                            "WHERE PMWO.GPMid IS Null")


Set NewEq = db.OpenRecordset("SELECT TempStorePMWO.WOID, TempStorePMWO.ModelNumber, TempStorePMWO.WORequest, " & _
"TempStorePMWO.AssignedTo, TempStorePMWO.Status, TempStorePMWO.DueDate, TempStorePMWO.AssetNumber, " & _
"TempStorePMWO.Manufacturer, TempStorePMWO.RequestBy, TempStorePMWO.GPMid, TempStorePMWO.DateRegistered, " & _
"TempStorePMWO.EqDescription, TempStorePMWO.WODescription FROM TempStorePMWO " & _
"LEFT JOIN PMWO ON TempStorePMWO.GPMid = PMWO.GPMid " & _
"WHERE (((PMWO.GPMid) Is Null));")


j = 0
AvailNo = MinAvailPMWONo

If NewEq.RecordCount > 0 Then
    NewEq.MoveFirst

    Do While Not NewEq.EOF
        PMWO.AddNew
        PMWO!WODescription = NewEq!WODescription
        PMWO!ModelNumber = NewEq!ModelNumber
'        PMWO!WOType = Nz(NewDateR!WOType, "")
        PMWO!WORequest = NewEq!WORequest
        PMWO!AssignedTo = NewEq!AssignedTo
        PMWO!Status = NewEq!Status
        PMWO!Completed = False
'        PMWO!RequestedDate = NewEq!RequestedDate
'        PMWO!DateDone = NewEq!DateDone
'        PMWO!TaskComment = NewEq!TaskComment
        PMWO!DueDate = NewEq!DueDate
        PMWO!WONumber = AvailNo
        PMWO!AssetNumber = Nz(NewEq!AssetNumber, "")
        PMWO!Manufacturer = NewEq!Manufacturer
'        PMWO!EngineeringComment = NewEq!EngineeringComment
        PMWO!RequestBy = NewEq!RequestBy
        PMWO!FormatWONumber = PMWOFormatNo("PMWO", AvailNo)
        PMWO!GPMid = NewEq!GPMid
        PMWO!DateRegistered = NewEq!DateRegistered
        PMWO!EqDescription = NewEq!EqDescription
      
        PMWO.Update
        AvailNo = AvailNo + 1
        j = j + 1
        NewEq.MoveNext
        
    Loop
End If





'
'Set NewDateR = db.OpenRecordset("SELECT TempStorePMWO.WOID, TempStorePMWO.WODescription, " & _
'                            "TempStorePMWO.ModelNumber, TempStorePMWO.WORequest, " & _
'                            "TempStorePMWO.AssignedTo, TempStorePMWO.Status, TempStorePMWO.Completed, " & _
'                            "TempStorePMWO.RequestedDate, TempStorePMWO.DateDone, TempStorePMWO.taskComment, " & _
'                            "TempStorePMWO.DueDate, TempStorePMWO.WONumber, TempStorePMWO.AssetNumber, " & _
'                            "TempStorePMWO.Manufacturer, TempStorePMWO.EngineeringComment, TempStorePMWO.RequestBy, TempStorePMWO.EqDescription, " & _
'                            "TempStorePMWO.FormatWONumber, TempStorePMWO.GPMid, TempStorePMWO.DateRegistered " & _
'                            "FROM TempStorePMWO LEFT JOIN PMWO ON TempStorePMWO.GPMid=PMWO.GPMid AND " & _
'                            "TempStorePMWO.DateRegistered=PMWO.DateRegistered " & _
'                            "WHERE IsNull(PMWO.DateRegistered) OR (TempStorePMWO.DueDate<> PMWO.DueDate);")
'
'If NewDateR.RecordCount > 0 Then
'
'
'NewDateR.MoveFirst
'
'    Do While Not NewDateR.EOF
'        PMWO.AddNew
'        PMWO!WODescription = NewDateR!WODescription
'        PMWO!ModelNumber = NewDateR!ModelNumber
''        PMWO!WOType = Nz(NewDateR!WOType, "")
'        PMWO!WORequest = NewDateR!WORequest
'        PMWO!AssignedTo = NewDateR!AssignedTo
'        PMWO!Status = NewDateR!Status
'        PMWO!Completed = False
'        PMWO!RequestedDate = NewDateR!RequestedDate
'        PMWO!DateDone = NewDateR!DateDone
'        PMWO!taskComment = NewDateR!taskComment
'        PMWO!DueDate = NewDateR!DueDate
'        PMWO!WONumber = AvailNo
'        PMWO!AssetNumber = Nz(NewDateR!AssetNumber, "")
'        PMWO!Manufacturer = NewDateR!Manufacturer
'        PMWO!EngineeringComment = NewDateR!EngineeringComment
'        PMWO!RequestBy = NewDateR!RequestBy
'        PMWO!FormatWONumber = PMWOFormatNo("PMWO", AvailNo)
'        PMWO!GPMid = NewDateR!GPMid
'        PMWO!DateRegistered = NewDateR!DateRegistered
'        PMWO!EqDescription = NewDateR!EqDescription
'
'        PMWO.Update
'        AvailNo = AvailNo + 1
'        j = j + 1
'        NewDateR.MoveNext
'
'    Loop
'End If



RightOUtterJOin = j
Set PMWO = Nothing
Set NewEq = Nothing
'Set NewDateR = Nothing
Set db = Nothing


                          
End Function

Function RightOUtterJOin_2(AvailNo As Integer) As Integer

Dim db As Database
Dim NewDateR As Recordset
Dim PMWO As Recordset
Dim str As String



Set db = CurrentDb

Set PMWO = db.OpenRecordset("PMWO")


'Set NewDateR = db.OpenRecordset("SELECT TempStorePMWO.WOID, TempStorePMWO.WODescription, " & _
'                            "TempStorePMWO.ModelNumber, TempStorePMWO.WORequest, " & _
'                            "TempStorePMWO.AssignedTo, TempStorePMWO.Status, TempStorePMWO.Completed, " & _
'                            "TempStorePMWO.RequestedDate, TempStorePMWO.DateDone, TempStorePMWO.taskComment, " & _
'                            "TempStorePMWO.DueDate, TempStorePMWO.WONumber, TempStorePMWO.AssetNumber, " & _
'                            "TempStorePMWO.Manufacturer, TempStorePMWO.EngineeringComment, TempStorePMWO.RequestBy, TempStorePMWO.EqDescription, " & _
'                            "TempStorePMWO.FormatWONumber, TempStorePMWO.GPMid, TempStorePMWO.DateRegistered " & _
'                            "FROM TempStorePMWO LEFT JOIN PMWO ON TempStorePMWO.GPMid=PMWO.GPMid AND " & _
'                            "TempStorePMWO.DateRegistered=PMWO.DateRegistered " & _
'                            "WHERE IsNull(PMWO.DateRegistered) OR (TempStorePMWO.DueDate<> PMWO.DueDate);")


Set NewDateR = db.OpenRecordset("SELECT TempStorePMWO.WOID, TempStorePMWO.WODescription, " & _
                           "TempStorePMWO.ModelNumber, TempStorePMWO.WORequest, " & _
                            "TempStorePMWO.AssignedTo, TempStorePMWO.Status, TempStorePMWO.Completed, " & _
                            "TempStorePMWO.RequestedDate, TempStorePMWO.DateDone, TempStorePMWO.taskComment, " & _
                            "TempStorePMWO.DueDate, TempStorePMWO.WONumber, TempStorePMWO.AssetNumber, " & _
                            "TempStorePMWO.Manufacturer, TempStorePMWO.EngineeringComment, TempStorePMWO.RequestBy, TempStorePMWO.EqDescription, " & _
                            "TempStorePMWO.FormatWONumber, TempStorePMWO.GPMid, TempStorePMWO.DateRegistered " & _
                            "FROM TempStorePMWO LEFT JOIN InnerJoinPMWO_TempPMWO ON TempStorePMWO.DueDate = InnerJoinPMWO_TempPMWO.DueDate" & _
" WHERE (((InnerJoinPMWO_TempPMWO.DueDate) Is Null));")


If NewDateR.RecordCount > 0 Then

NewDateR.MoveLast



NewDateR.MoveFirst

    Do While Not NewDateR.EOF
        PMWO.AddNew
        PMWO!WODescription = NewDateR!WODescription
        PMWO!ModelNumber = NewDateR!ModelNumber
'        PMWO!WOType = Nz(NewDateR!WOType, "")
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
        PMWO!FormatWONumber = PMWOFormatNo("PMWO", AvailNo)
        PMWO!GPMid = NewDateR!GPMid
        PMWO!DateRegistered = NewDateR!DateRegistered
        PMWO!EqDescription = NewDateR!EqDescription
      
        PMWO.Update
        AvailNo = AvailNo + 1
        j = j + 1
        NewDateR.MoveNext
        
    Loop
End If




Set PMWO = Nothing

Set NewDateR = Nothing
Set db = Nothing
RightOUtterJOin_2 = j

                          
End Function


'Public Function PMWO_Create(n As Integer) As Integer
'
'Dim db As Database
'Dim TempPMWO As Recordset
'Dim PMWO As Recordset
'Dim PMWO_Wrt As Recordset
'Dim i As Integer
'Dim j As Integer
'Dim AvailNo As Integer
'Dim LeftOver As Boolean
'
'LeftOver = True
'j = n
'AvailNo = MinAvailPMWONo
'
'Set db = CurrentDb
'
'Set PMWO = db.OpenRecordset("SELECT GPMID, DueDate FROM PMWO", dbOpenSnapshot, dbReadOnly)
'Set TempPMWO = db.OpenRecordset("TempStorePMWO", , dbReadOnly)
'Set PMWO_Wrt = db.OpenRecordset("PMWO", dbAppendOnly)
'TempPMWO.MoveFirst
'PMWO.MoveFirst
'Do While Not TempPMWO.EOF
'
'    Do While Not PMWO.EOF
'
'        If TempPMWO!GPMid = PMWO!GMPid Then
'            If TempPMWO!DueDate = PMWO!DueDate Then
'                LeftOver = False
'            End If
'        End If
'
'        PMWO.MoveNext
'
'
'    Loop
'
'    PMWO.MoveFirst
'
'    If LeftOver Then
'
'        PMWO_Wrt.AddNew
'        With PMWO_Wrt
'
'            !WODescription = TempPMWO!WODescription
'            !ModelNumber = TempPMWO!ModelNumber
'            !WORequest = TempPMWO!WORequest
'            !AssignedTo = TempPMWO!AssignedTo
'            !Status = TempPMWO!Status
'            !Completed = False
'            !RequestedDate = TempPMWO!RequestedDate
'            !DateDone = TempPMWO!DateDone
'            !TaskComment = TempPMWO!TaskComment
'            !DueDate = TempPMWO!DueDate
'            !WONumber = AvailNo
'            !AssetNumber = Nz(TempPMWO!AssetNumber, "")
'            !Manufacturer = TempPMWO!Manufacturer
'            !EngineeringComment = TempPMWO!EngineeringComment
'            !RequestBy = TempPMWO!RequestBy
'            !FormatWONumber = PMWOFormatNo("", AvailNo)
'            !GPMid = TempPMWO!GPMid
'            !DateRegistered = TempPMWO!DateRegistered
'            !EqDescription = TempPMWO!EqDescription
'
'            .Update
'            AvailNo = AvailNo + 1
'            j = j + 1
'
'
'        End With
'
'    End If
'    LeftOver = True
'
'TempPMWO.MoveNext
'
'Loop
'PMWO_Create = j
'Set PMWO = Nothing
'Set TempPMWO = Nothing
'Set PMWO_Wrt = Nothing
'Set db = Nothing
'
'
'
'End Function

Public Function PMWO_Create(n As Integer) As Integer

On Error GoTo Err_Handel
Dim db As Database
Dim TempPMWO As Recordset
Dim PMWO As Recordset
Dim PMWO_Wrt As Recordset
Dim i As Double
Dim j As Double
Dim AvailNo As Integer
Dim LeftOver As Boolean

LeftOver = True
j = n
AvailNo = MinAvailPMWONo

Set db = CurrentDb

Set PMWO = db.OpenRecordset("SELECT GPMID, DueDate FROM PMWO", dbOpenSnapshot, dbReadOnly)
Set TempPMWO = db.OpenRecordset("TempStorePMWO", , dbReadOnly)
Set PMWO_Wrt = db.OpenRecordset("PMWO")
TempPMWO.MoveFirst
PMWO.MoveFirst
Do While Not TempPMWO.EOF
    
    Do While Not PMWO.EOF
    
        If TempPMWO!GPMid = PMWO!GPMid Then
            If TempPMWO!DueDate = PMWO!DueDate Then
                LeftOver = False
            End If
        End If
        
        PMWO.MoveNext
    
    
    Loop
    
    PMWO.MoveFirst
    
    If LeftOver Then
    
        PMWO_Wrt.AddNew
        With PMWO_Wrt

            !WODescription = TempPMWO!WODescription
            !ModelNumber = TempPMWO!ModelNumber
            !WORequest = TempPMWO!WORequest
            !AssignedTo = TempPMWO!AssignedTo
            !Status = TempPMWO!Status
            !Completed = False
            !RequestedDate = TempPMWO!RequestedDate
            !DateDone = TempPMWO!DateDone
            !TaskComment = TempPMWO!TaskComment
            !DueDate = TempPMWO!DueDate
            !WONumber = AvailNo
            !AssetNumber = Nz(TempPMWO!AssetNumber, "")
            !Manufacturer = TempPMWO!Manufacturer
            !EngineeringComment = TempPMWO!EngineeringComment
            !RequestBy = TempPMWO!RequestBy
            !FormatWONumber = PMWOFormatNo("PMWO", AvailNo)
            !GPMid = TempPMWO!GPMid
            !DateRegistered = TempPMWO!DateRegistered
            !EqDescription = TempPMWO!EqDescription
                  .Update
            AvailNo = AvailNo + 1
            j = j + 1

        
        End With
        
    End If
    LeftOver = True

TempPMWO.MoveNext

Loop
PMWO_Create = j
Set PMWO = Nothing
Set TempPMWO = Nothing
Set PMWO_Wrt = Nothing
Set db = Nothing

Exit Function
Err_Handel:
If Err.Number <> 3021 Then
MsgBox Err.Description
End If
Resume Next

End Function

