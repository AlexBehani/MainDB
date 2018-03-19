Attribute VB_Name = "GenPM"
Option Compare Database
Option Explicit

Public Function ReturnPMRecord(Asset As String) As String

Dim db As Database
Dim GPM As Recordset
Dim Str As String

Set db = CurrentDb
Set GPM = db.OpenRecordset("SELECT DateRegistered, TaskName, Description, AssetNumber, Frequency, ID" & _
                            " FROM GeneralPM WHERE AssetNumber = '" & Asset & "'")
                            
                            
If GPM.RecordCount > 0 Then

GPM.MoveFirst

Do While Not GPM.EOF

    Str = Str & GPM!DateRegistered & ";" & GPM!TaskName & ";" & GPM!Description & ";" & _
        GPM!AssetNumber & ";" & GPM!frequency & ";" & GPM!ID & ";"
    GPM.MoveNext
    
Loop

ReturnPMRecord = Str

Set GPM = Nothing
Set db = Nothing


End If

End Function
