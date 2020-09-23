Attribute VB_Name = "Global"
'Option Explicit
Public TimesAround As Long
Public db As Database
Public rstInfo As Recordset
Public dUsers As Dictionary
Public stopSearch As Boolean
Public imLoading As Boolean
Public socketMode As String
Public currentUser As String
Public doingDL As Boolean
Public currIndex As Long
Public strFilePath As String
