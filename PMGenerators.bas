Attribute VB_Name = "PMGenerators"
Option Compare Database

Public Function PMGenerator()
'Dim AvailNo As Integer
Dim db As Database
Dim GenPM As Recordset
Dim PMWO As Recordset
Dim Mdate As Integer, Ydate As Integer, MCr As Integer, YCr As Integer, i As Integer

i = 0
Set db = CurrentDb
Set GenPM = db.OpenRecordset("GeneralPM")
Set PMWO = db.OpenRecordset("TempStorePMWO")

If (Month(Date) = 12) Then

    MCr = 1
    YCr = Year(Date) + 1
Else

    MCr = Month(Date) + 1
    YCr = Year(Date)
End If


'AvailNo = MinAvailPMWONo
GenPM.MoveFirst
Do While Not GenPM.EOF


Mdate = Month(GenPM!DateRegistered)
Ydate = Year(GenPM!DateRegistered)


    If Not (YCr = Ydate And MCr = Mdate) Then
        
        If GenPM!frequency = "Bi Annaully" Then
            If (Ydate <> YCr And Mdate = MCr And YCr > Ydate) Then
                If ((YCr - Ydate) Mod 2) = 0 Then
                    
                    PMWO.AddNew
'                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
'                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!Description
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
'                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
'                    PMWO!WONumber = AvailNo
'                    PMWO!formatwonumber = PMWOFormatNo("PMWO", AvailNo)
                    PMWO!gpmid = GenPM!ID
                    PMWO!DateRegistered = GenPM!DateRegistered
                    PMWO.Update
'                    AvailNo = AvailNo + 1
                    i = i + 1
                    
                End If
            End If
            
        ElseIf GenPM!frequency = "Annually" Then
            If (Ydate <> YCr And Mdate = MCr) Then
                    
                    PMWO.AddNew
'                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
'                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!Description
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
'                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
'                    PMWO!WONumber = AvailNo
'                    PMWO!formatwonumber = PMWOFormatNo("PMWO", AvailNo)
                    PMWO!gpmid = GenPM!ID
                    PMWO!DateRegistered = GenPM!DateRegistered
                    PMWO.Update
'                    AvailNo = AvailNo + 1
                    i = i + 1
            End If
            
        ElseIf GenPM!frequency = "Semi Annually" Then
            If (Ydate = YCr And Mdate <> MCr And MCr > Mdate) Then
                If ((MCr - Mdate) Mod 6) Then
                    
                    PMWO.AddNew
'                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
'                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!Description
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
'                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
'                    PMWO!WONumber = AvailNo
'                    PMWO!formatwonumber = PMWOFormatNo("PMWO", AvailNo)
                    PMWO!gpmid = GenPM!ID
                    PMWO!DateRegistered = GenPM!DateRegistered
                    PMWO.Update
'                    AvailNo = AvailNo + 1
                    i = i + 1
                End If
            ElseIf (Ydate <> YCr And Mdate <> MCr) Then
                If (Abs(Mdate - MCr) Mod 6) = 0 Then
                
                    PMWO.AddNew
'                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
'                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!Description
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
'                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
'                    PMWO!WONumber = AvailNo
'                    PMWO!formatwonumber = PMWOFormatNo("PMWO", AvailNo)
                    PMWO!gpmid = GenPM!ID
                    PMWO!DateRegistered = GenPM!DateRegistered
                    PMWO.Update
'                    AvailNo = AvailNo + 1
                    i = i + 1
                End If
            ElseIf (Ydate <> YCr And Mdate = MCr) Then
            
                    PMWO.AddNew
'                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
'                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!Description
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
'                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
'                    PMWO!WONumber = AvailNo
'                    PMWO!formatwonumber = PMWOFormatNo("PMWO", AvailNo)
                    PMWO!gpmid = GenPM!ID
                    PMWO!DateRegistered = GenPM!DateRegistered
                    PMWO.Update
'                    AvailNo = AvailNo + 1
                    i = i + 1
            End If
            
        ElseIf GenPM!frequency = "Quarterly" Then
            If (Ydate = YCr And Mdate <> MCr And MCr > Mdate) Then
                If ((MCr - Mdate) Mod 3) = 0 Then
                    
                    PMWO.AddNew
'                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
'                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!Description
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
'                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
'                    PMWO!WONumber = AvailNo
'                    PMWO!formatwonumber = PMWOFormatNo("PMWO", AvailNo)
                    PMWO!gpmid = GenPM!ID
                    PMWO!DateRegistered = GenPM!DateRegistered
                    PMWO.Update
'                    AvailNo = AvailNo + 1
                    i = i + 1
                End If
            
            ElseIf (Ydate <> YCr And Mdate <> MCr And YCr > Ydate) Then
                If (Abs(MCr - Mdate) Mod 3 = 0) Then
                    
                    PMWO.AddNew
'                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
'                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!Description
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
'                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
'                    PMWO!WONumber = AvailNo
'                    PMWO!formatwonumber = PMWOFormatNo("PMWO", AvailNo)
                    PMWO!gpmid = GenPM!ID
                    PMWO!DateRegistered = GenPM!DateRegistered
                    PMWO.Update
'                    AvailNo = AvailNo + 1
                    i = i + 1
                End If
                
            ElseIf (Ydate <> YCr And MCr = Mdate) Then
            
                    PMWO.AddNew
'                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
'                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!Description
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
'                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
'                    PMWO!WONumber = AvailNo
'                    PMWO!formatwonumber = PMWOFormatNo("PMWO", AvailNo)
                    PMWO!gpmid = GenPM!ID
                    PMWO!DateRegistered = GenPM!DateRegistered
                    PMWO.Update
'                    AvailNo = AvailNo + 1
                    i = i + 1
            End If
        End If
    End If
 
'Debug.Print GenPM!ID

GenPM.MoveNext


Loop

Set GenPM = Nothing
Set PMWO = Nothing
Set db = Nothing
PMGenerator = i
End Function

Function DueDate(RegDate As Date) As Date

If (Day(RegDate) > Mdays) Then
DueDate = DateAdd("m", 1, DateSerial(Year(Date), Month(Date), Mdays))

Else
DueDate = DateAdd("m", 1, DateSerial(Year(Date), Month(Date), Day(RegDate)))


End If

End Function


Function Mdays() As Integer
Mdays = Day(DateSerial(Year(Date), myMonth + 1, 1) - 1)
'Mdays = Day(DateSerial(YCr, MCr + 1, 1) - 1)
End Function



Public Sub DeleteOldPM()
On Error GoTo Err
Dim db As Database
'Dim GPM As Recordset
Dim Str As String
Set db = CurrentDb
Str = "DELETE * FROM TempStorePMWO"

db.Execute Str

'Set GPM = db.OpenRecordset("DELETE FROM TempStorePMWO WHERE month(DueDate) =" & Month(Date) & " AND Year(DueDate)=" & Year(Date))



'If Not (GPM.RecordCount > 0) Then Exit Sub
'GPM.MoveLast
'GPM.MoveFirst
'
'Do While Not GPM.EOF
'
'GPM.Delete
'GPM.MoveNext
'Loop

Set db = Nothing
'Set GPM = Nothing


Exit Sub

Err:
MsgBox Err.Description
MsgBox Err.Number
Set db = Nothing
'Set GPM = Nothing
End Sub

Public Function UPdate_DateRegister(ID As Integer, DateReg As Date)

Dim db As Database
Dim GPM As Recordset

' have to make the record more specific
Set db = CurrentDb
Set GPM = db.OpenRecordset("SELECT * FROM GeneralPM WHERE ID= " & ID)

GPM.MoveFirst
If (GPM!DateRegistered = DateReg) Then

    Set db = Nothing
    Set GPM = Nothing
    Exit Function
End If

Call GPMAudit(GPM, DateReg)
    GPM.Edit
    GPM!DateRegistered = DateReg
    GPM.Update

Set GPM = Nothing
Set db = Nothing





End Function


Public Function GPMAudit(GPM As Recordset, DateReg As Date)
Dim db As Database
Dim PMAudit As Recordset

Set db = CurrentDb
Set PMAudit = db.OpenRecordset("PMAudit")
    PMAudit.AddNew
    PMAudit!DateRegistered = GPM!DateRegistered
    PMAudit!UpdatedDateRegistered = DateReg
    PMAudit!ModelNumber = GPM!ModelNumber
    PMAudit!AssetNumber = GPM!SerialNumber
    PMAudit!PMID = GPM!PMID
'    PMAudit!TaskName = GPM!TaskName
'    PMAudit!TaskType = GPM!TaskType
    PMAudit!Description = GPM!Description
    PMAudit!AssignedTo = GPM!AssignedTo
    PMAudit!frequency = GPM!frequency
    PMAudit!Status = GPM!Status
    PMAudit!Manufacturer = GPM!Manufacturer
    PMAudit!modified = GPM!DateRegistered & " updated to " & DateReg
    PMAudit!DateTime = Now
    If Not (CUser Is Nothing) Then PMAudit!User = CUser.FullName
    PMAudit.Update
    
'    GPM!dateregesterd = DateReg
'    GPM.Update

Set PMAudit = Nothing
'Set GPM = Nothing
Set db = Nothing

End Function
