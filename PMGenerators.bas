Attribute VB_Name = "PMGenerators"
Option Compare Database

Public Function PMGenerator() As Integer

Dim AvailNo As Integer
Dim db As Database
Dim GenPM As Recordset
Dim PMWO As Recordset
Dim Mdate As Integer, Ydate As Integer, MCr As Integer, YCr As Integer, i As Integer

i = 0
Set db = CurrentDb
Set GenPM = db.OpenRecordset("GeneralPM")
Set PMWO = db.OpenRecordset("PMWO")
AvailNo = MinAvailPMWONo
GenPM.MoveFirst
Do While Not GenPM.EOF

MCr = Month(Date)
YCr = Year(Date)
Mdate = Month(GenPM!DateRegistered)
Ydate = Year(GenPM!DateRegistered)

    If Not (YCr = Ydate And MCr = Mdate) Then
        
        If GenPM!frequency = "Bi Annaully" Then
            If (Ydate <> YCr And Mdate = MCr And YCr > Ydate) Then
                If ((YCr - Ydate) Mod 2) = 0 Then
                    
                    PMWO.AddNew
                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!TaskName
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
                    PMWO!WONumber = AvailNo
                    PMWO.Update
                    AvailNo = AvailNo + 1
                    i = i + 1
                    
                End If
            End If
            
        ElseIf GenPM!frequency = "Annually" Then
            If (Ydate <> YCr And Mdate = MCr) Then
                    
                    PMWO.AddNew
                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!TaskName
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
                    PMWO!WONumber = AvailNo
                    PMWO.Update
                    AvailNo = AvailNo + 1
                    i = i + 1
            End If
            
        ElseIf GenPM!frequency = "Semi Annually" Then
            If (Ydate = YCr And Mdate <> MCr And MCr > Mdate) Then
                If ((MCr - Mdate) Mod 6) Then
                    
                    PMWO.AddNew
                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!TaskName
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
                    PMWO!WONumber = AvailNo
                    PMWO.Update
                    AvailNo = AvailNo + 1
                    i = i + 1
                End If
            ElseIf (Ydate <> YCr And Mdate <> MCr) Then
                If (Abs(Mdate - MCr) Mod 6) = 0 Then
                
                    PMWO.AddNew
                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!TaskName
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
                    PMWO!WONumber = AvailNo
                    PMWO.Update
                    AvailNo = AvailNo + 1
                    i = i + 1
                End If
            ElseIf (Ydate <> YCr And Mdate = MCr) Then
            
                    PMWO.AddNew
                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!TaskName
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
                    PMWO!WONumber = AvailNo
                    PMWO.Update
                    AvailNo = AvailNo + 1
                    i = i + 1
            End If
            
        ElseIf GenPM!frequency = "Quarterly" Then
            If (Ydate = YCr And Mdate <> MCr And MCr > Mdate) Then
                If ((MCr - Mdate) Mod 3) = 0 Then
                    
                    PMWO.AddNew
                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!TaskName
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
                    PMWO!WONumber = AvailNo
                    PMWO.Update
                    AvailNo = AvailNo + 1
                    i = i + 1
                End If
            
            ElseIf (Ydate <> YCr And Mdate <> MCr And YCr > Ydate) Then
                If (Abs(MCr - Mdate) Mod 3 = 0) Then
                    
                    PMWO.AddNew
                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!TaskName
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
                    PMWO!WONumber = AvailNo
                    PMWO.Update
                    AvailNo = AvailNo + 1
                    i = i + 1
                End If
                
            ElseIf (Ydate <> YCr And MCr = Mdate) Then
            
                    PMWO.AddNew
                    PMWO!WODescription = GenPM!Description
                    PMWO!ModelNumber = GenPM!ModelNumber
                    PMWO!AssetNumber = GenPM!AssetNumber
                    PMWO!WOType = GenPM!TaskType
                    PMWO!WORequest = GenPM!TaskName
                    PMWO!RequestBy = "PM"
                    PMWO!AssignedTo = GenPM!AssignedTo
                    PMWO!Completed = False
                    PMWO!Task = GenPM!TaskName
                    PMWO!Status = GenPM!Status
                    PMWO!DueDate = DueDate(GenPM!DateRegistered)
                    PMWO!Manufacturer = GenPM!Manufacturer
                    PMWO!WONumber = AvailNo
                    PMWO.Update
                    AvailNo = AvailNo + 1
                    i = i + 1
            End If
        End If
    End If
 

GenPM.MoveNext

Loop

Set GenPM = Nothing
Set PMWO = Nothing
Set db = Nothing
PMGenerator = i
End Function

Function DueDate(RegDate As Date) As Date

If (Day(RegDate) > Mdays) Then
DueDate = DateSerial(Year(Date), Month(Date), Mdays)
Else
DueDate = DateSerial(Year(Date), Month(Date), Day(RegDate))

End If

End Function


Function Mdays() As Integer
Mdays = Day(DateSerial(Year(Date), myMonth + 1, 1) - 1)
End Function



Public Sub DeleteOldPM()
On Error GoTo Err
Dim db As Database
Dim GPM As Recordset

Set db = CurrentDb


Set GPM = db.OpenRecordset("SELECT * FROM PMWO WHERE month(DueDate) =" & Month(Date) & " AND Year(DueDate)=" & Year(Date))


If Not (GPM.RecordCount > 0) Then Exit Sub
GPM.MoveLast
GPM.MoveFirst

Do While Not GPM.EOF

GPM.Delete
GPM.MoveNext
Loop

Set db = Nothing
Set GPM = Nothing


Exit Sub

Err:
MsgBox Err.Description
MsgBox Err.Number
Set db = Nothing
Set GPM = Nothing
End Sub
