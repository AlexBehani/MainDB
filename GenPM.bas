Attribute VB_Name = "GenPM"
Option Compare Database
Option Explicit

Public Function ReturnPMRecord(Asset As String) As String

Dim db As Database
Dim GPM As Recordset
Dim str As String

Set db = CurrentDb

If Asset = "N/A" Then

    Set GPM = db.OpenRecordset("SELECT DateRegistered, TaskName, Description, AssetNumber, Manufacturer, ModelNumber, Frequency, ID" & _
                            " FROM GeneralPM WHERE Status = 'Spare'")
Else

    Set GPM = db.OpenRecordset("SELECT DateRegistered, TaskName, Description, AssetNumber, Manufacturer, ModelNumber, Frequency, ID" & _
                            " FROM GeneralPM WHERE AssetNumber = '" & Asset & "'")
End If
                            
If GPM.RecordCount > 0 Then

GPM.MoveFirst

Do While Not GPM.EOF

    str = str & GPM!DateRegistered & ";" & GPM!TaskName & ";" & GPM!Description & ";" & _
        GPM!AssetNumber & ";" & GPM!Manufacturer & ";" & GPM!ModelNumber & ";" & _
        GPM!Frequency & ";" & GPM!ID & ";"
    GPM.MoveNext
    
Loop

ReturnPMRecord = str

Set GPM = Nothing
Set db = Nothing


End If

End Function


Public Sub DeleteGenPM(ID As Long)

Dim db As Database
Dim GeneralPM As Recordset
Dim str As String

Set db = CurrentDb
str = "DELETE * FROM GeneralPM WHERE EQid =" & ID

db.Execute str


Set db = Nothing
Set GeneralPM = Nothing




End Sub

