Attribute VB_Name = "PMTaskModule"
Option Compare Database

Public Sub PMTask_Save()

Dim db As Database
Dim PM As Recordset

Set db = CurrentDb


If (PMTK.Edit = False) Then
    Set PM = db.OpenRecordset("PMTask")
    PM.AddNew
    PM!Task_Name = PMTK.PMTaskName
    PM!Type = PMTK.PMType
    PM!Description = PMTK.Description
    PM!AssignedTo = PMTK.AssignedTo
    PM!DownTime = PMTK.DownTime
    PM!frequency = PMTK.FrequencyDays
    PM.Update
    
ElseIf (PMTK.Edit = True) Then

    Set PM = db.OpenRecordset("SELECT * FROM PMTask WHERE PMID =" & PMTK.ID)
    
    PM.MoveFirst
    PM.Edit
    PM!Task_Name = PMTK.PMTaskName
    PM!Type = PMTK.PMType
    PM!Description = PMTK.Description
    PM!AssignedTo = PMTK.AssignedTo
    PM!DownTime = PMTK.DownTime
    PM.Update
    
'ElseIf (action = "Delete") Then
'
'    Set PM = db.OpenRecordset("SELECT * FROM PMTask WHERE PMID =" & PMTk.Id)
'    PM.MoveFirst
'    PM.Delete
    
End If

Set db = Nothing
Set PM = Nothing


End Sub



Public Function Load_PMTask(ID As Integer) As PMtask


Dim db As Database
Dim PMTK As Recordset
Dim PMTaskTemp As PMtask

Set PMTaskTemp = New PMtask
Set db = CurrentDb
Set PMTK = db.OpenRecordset("SELECT * FROM PMTask WHERE PMID= " & ID)
PMTK.MoveFirst

PMTaskTemp.PMTaskName = PMTK!Task_Name
PMTaskTemp.PMType = PMTK!Type
PMTaskTemp.AssignedTo = PMTK!AssignedTo
PMTaskTemp.Description = PMTK!Description
PMTaskTemp.FrequencyDays = PMTK!frequency
PMTaskTemp.DownTime = PMTK!DownTime

Set Load_PMTask = PMTaskTemp

Set PMTK = Nothing
Set PMTaskTemp = Nothing
Set db = Nothing


End Function


Public Sub Delete_PMTask()


Dim db As Database
Dim PMtask As Recordset

Set db = CurrentDb
Set PMtask = db.OpenRecordset("SELECT * FROM PMTask WHERE PMID=" & PMTK.ID)
PMtask.MoveFirst
PMtask.Delete


Set PMtask = Nothing
Set db = Nothing


End Sub



