Attribute VB_Name = "PMGenerators"
Option Compare Database

Public Sub PMGenerator()

Dim db As Database
Dim GenPM As Recordset
Dim WO As Recordset
Dim Mdate As Integer, Ydate As Integer, MCr As Integer, YCr As Integer

Set db = CurrentDb
Set GenPM = db.OpenRecordset("GeneralPM")
Set WO = db.OpenRecordset("WO")

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
                    
                    WO.AddNew
                    WO!WODescription = GenPM!Description
                    WO!ModelNumber = GenPM!ModelNumber
                    WO!Scheduled = False
                    WO!Type = GenPM!TaskType
                    WO!WORequest = "PM"
                    WO!AssignedTo = GenPM!AssignedTo
                    WO!Completed = False
                    WO!Task = GenPM!TaskName
                    WO!DueDate = DueDate(GenPM!DateRegistered)
                    WO.Update
                    
                End If
            End If
            
        ElseIf GenPM!frequency = "Annually" Then
            If (Ydate <> YCr And Mdate = MCr) Then
                    
                    WO.AddNew
                    WO!WODescription = GenPM!Description
                    WO!ModelNumber = GenPM!ModelNumber
                    WO!Scheduled = False
                    WO!Type = GenPM!TaskType
                    WO!WORequest = "PM"
                    WO!AssignedTo = GenPM!AssignedTo
                    WO!Completed = False
                    WO!Task = GenPM!TaskName

                    WO!DueDate = DueDate(GenPM!DateRegistered)

                    WO.Update
                    
            End If
            
        ElseIf GenPM!frequency = "Semi Annually" Then
            If (Ydate = YCr And Mdate <> MCr And MCr > Mdate) Then
                If ((MCr - Mdate) Mod 6) Then
                    
                    WO.AddNew
                    WO!WODescription = GenPM!Description
                    WO!ModelNumber = GenPM!ModelNumber
                    WO!Scheduled = False
                    WO!Type = GenPM!TaskType
                    WO!WORequest = "PM"
                    WO!AssignedTo = GenPM!AssignedTo
                    WO!Completed = False
                    WO!Task = GenPM!TaskName

                    WO!DueDate = DueDate(GenPM!DateRegistered)

                    WO.Update
                    
                End If
            ElseIf (Ydate <> YCr And Mdate <> MCr) Then
                If (Abs(Mdate - MCr) Mod 6) = 0 Then
                
                    WO.AddNew
                    WO!WODescription = GenPM!Description
                    WO!ModelNumber = GenPM!ModelNumber
                    WO!Scheduled = False
                    WO!WOType = GenPM!TaskType
                    WO!WORequest = "PM"
                    WO!AssignedTo = GenPM!AssignedTo
                    WO!Completed = False
                    WO!Task = GenPM!TaskName
                    WO!DueDate = DueDate(GenPM!DateRegistered)
                    WO.Update
                    
                End If
            ElseIf (Ydate <> YCr And Mdate = MCr) Then
            
                    WO.AddNew
                    WO!WODescription = GenPM!Description
                    WO!ModelNumber = GenPM!ModelNumber
                    WO!Scheduled = False
                    WO!WOType = GenPM!TaskType
                    WO!WORequest = "PM"
                    WO!AssignedTo = GenPM!AssignedTo
                    WO!Completed = False
                    WO!Task = GenPM!TaskName
                    WO!DueDate = DueDate(GenPM!DateRegistered)
                    WO.Update
                    
            End If
            
        ElseIf GenPM!frequency = "Quarterly" Then
            If (Ydate = YCr And Mdate <> MCr And MCr > Mdate) Then
                If ((MCr - Mdate) Mod 3) = 0 Then
                    
                    WO.AddNew
                    WO!WODescription = GenPM!Description
                    WO!ModelNumber = GenPM!ModelNumber
                    WO!Scheduled = False
                    WO!WOType = GenPM!TaskType
                    WO!WORequest = "PM"
                    WO!AssignedTo = GenPM!AssignedTo
                    WO!Completed = False
                    WO!Task = GenPM!TaskName
                    WO!DueDate = DueDate(GenPM!DateRegistered)
                    WO.Update
                    
                End If
            
            ElseIf (Ydate <> YCr And Mdate <> MCr And YCr > Ydate) Then
                If (Abs(MCr - Mdate) Mod 3 = 0) Then
                    
                    WO.AddNew
                    WO!WODescription = GenPM!Description
                    WO!ModelNumber = GenPM!ModelNumber
                    WO!Scheduled = False
                    WO!WOType = GenPM!TaskType
                    WO!WORequest = "PM"
                    WO!AssignedTo = GenPM!AssignedTo
                    WO!Completed = False
                    WO!Task = GenPM!TaskName
                    WO!DueDate = DueDate(GenPM!DateRegistered)
                    WO.Update
                    
                End If
                
            ElseIf (Ydate <> YCr And MCr = Mdate) Then
            
                    WO.AddNew
                    WO!WODescription = GenPM!Description
                    WO!ModelNumber = GenPM!ModelNumber
                    WO!Scheduled = False
                    WO!WOType = GenPM!TaskType
                    WO!WORequest = "PM"
                    WO!AssignedTo = GenPM!AssignedTo
                    WO!Completed = False
                    WO!Task = GenPM!TaskName
                    WO!DueDate = DueDate(GenPM!DateRegistered)
                    WO.Update
                    
            End If
        End If
    End If
 

GenPM.MoveNext

Loop

Set GenPM = Nothing
Set WO = Nothing
Set db = Nothing

End Sub

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


Set GPM = db.OpenRecordset("SELECT * FROM WO WHERE month(DueDate) =" & Month(Date) & " AND Year(DueDate)=" & Year(Date))


If Not (GPM.RecordCount > 0) Then Exit Sub
GPM.MoveLast
GPM.MoveFirst
MsgBox GPM.RecordCount
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
