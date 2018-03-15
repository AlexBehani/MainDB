Attribute VB_Name = "Commas"
Option Compare Database
Option Explicit

Public Function CheckForComma(Str As String) As Boolean

If InStr(1, Str, ",") > 0 Then

    CheckForComma = True
Else
    CheckForComma = False
End If

End Function

