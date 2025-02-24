VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Simple logging to a text file
Private Const LOG_FILE As String = "error_log.txt"
Private fileNum As Integer

' Initialize logger
Private Sub Class_Initialize()
    fileNum = FreeFile
End Sub

' Log an error
Public Sub LogError(ByVal message As String)
    On Error Resume Next
    
    ' Open log file in append mode
    Open ThisDocument.Path & "\" & LOG_FILE For Append As fileNum
    
    ' Write log entry
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " ERROR: " & message
    
    ' Close file
    Close fileNum
End Sub

' Log an informational message
Public Sub LogInfo(ByVal message As String)
    On Error Resume Next
    
    ' Open log file in append mode
    Open ThisDocument.Path & "\" & LOG_FILE For Append As fileNum
    
    ' Write log entry
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " INFO: " & message
    
    ' Close file
    Close fileNum
End Sub 