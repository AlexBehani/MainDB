VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WOClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public WOID As Integer
Public Edit As Boolean
Public Task As String
Public DateDone As Date
Public TimeDone As Date
Public TaskComment As String
Public Employee As String
Public Completed As Boolean
Public LockedDown As Boolean
Public Invisible As Boolean
