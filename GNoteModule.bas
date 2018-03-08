Attribute VB_Name = "GNoteModule"
Option Compare Database

Public Sub Save_Note()

Dim db As Database
Dim Note As Recordset

Set db = CurrentDb

If Not GNote.Edit Then
    
    Set Note = db.OpenRecordset("GeneralNote")
    Note.AddNew
    Note!Title = GNote.Title
    Note!Description = GNote.Description
    Note!DateVar = GNote.Date_
    Note!User = GNote.User
    Note!Comment = GNote.Comment

    Note.Update
Else

    Set Note = db.OpenRecordset("SELECT * FROM GeneralNote WHERE ID= " & GNote.ID)
    Note.MoveFirst
    Note.Edit
    Note!Title = GNote.Title
    Note!Description = GNote.Description
    Note!DateVar = GNote.Date_
    Note!User = GNote.User
    Note!Comment = GNote.Comment
    Note.Update
    
End If

Set Note = Nothing
Set db = Nothing
'Set GNote = Nothing
    
End Sub


Function Load_Note() As GeneralNote

Dim NoteTemp As GeneralNote
Dim db As Database
Dim Note As Recordset

Set NoteTemp = New GeneralNote
Set db = CurrentDb
Set Note = db.OpenRecordset("SELECT * FROM GeneralNote WHERE ID= " & GNote.ID)
Note.MoveFirst

NoteTemp.Description = Note!Description

NoteTemp.User = Note!User
NoteTemp.Comment = Note!Comment
NoteTemp.Date_ = Note!DateVar

Set Load_Note = NoteTemp
Set db = Nothing
Set Note = Nothing
Set NoteTemp = Nothing


End Function


Public Sub Delete_Note()
On Error GoTo Err
Dim db As Database
Dim Note As Recordset

Set db = CurrentDb


    Set Note = db.OpenRecordset("SELECT * FROM GeneralNote WHERE ID= " & GNote.ID)
    Note.MoveFirst
    Note.Delete

    


Set Note = Nothing
Set db = Nothing
GNote.Edit = False
    
Exit Sub
Err:
If (Err.Number = 3021) Then
MsgBox "Haven't select any record", vbCritical, "Error"
Else: MsgBox Err.Description & vbNewLine & Err.Number
End If


End Sub
