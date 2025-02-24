VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APIClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Simple configuration
Private baseURL As String
Private apiKey As String
Private httpClient As Object

' Initialize the client
Private Sub Class_Initialize()
    Set httpClient = CreateObject("WinHTTP.WinHTTPRequest.5.1")
End Sub

' Clean up
Private Sub Class_Terminate()
    Set httpClient = Nothing
End Sub

' Configure the client
Public Sub Configure(ByVal url As String, ByVal key As String)
    baseURL = url
    apiKey = key
End Sub

' Make an API request
Public Function MakeRequest(ByVal prompt As String) As String
    On Error GoTo ErrorHandler
    
    ' Prepare and send request
    With httpClient
        .Open "POST", baseURL & "/v1/chat/completions", False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & apiKey
        
        ' Prepare chat completion payload
        Dim payload As String
        payload = "{" & _
                 """model"": """ & API_MODEL & """," & _
                 """messages"": [" & _
                 "{""role"": ""system"", ""content"": ""You are a helpful assistant that analyzes and enhances Visio diagram elements.""}," & _
                 "{""role"": ""user"", ""content"": """ & EscapeJSON(prompt) & """}" & _
                 "]," & _
                 """max_tokens"": " & MAX_TOKENS & "," & _
                 """temperature"": 0.7" & _
                 "}"
        
        .Send payload
        
        ' Check response
        If .Status >= 200 And .Status < 300 Then
            ' Parse response to get actual content
            Dim responseText As String
            responseText = .ResponseText
            
            ' Extract content from response (basic parsing)
            Dim startPos As Long, endPos As Long
            startPos = InStr(1, responseText, """content"":""") + 11
            endPos = InStr(startPos, responseText, """")
            
            If startPos > 11 And endPos > startPos Then
                MakeRequest = Mid(responseText, startPos, endPos - startPos)
            Else
                MakeRequest = "Error: Could not parse response"
            End If
        Else
            Err.Raise vbObjectError + 1, "APIClient", "API request failed with status " & .Status
        End If
    End With
    
    Exit Function

ErrorHandler:
    MakeRequest = "Error: " & Err.Description
End Function

' Helper function to escape JSON strings
Private Function EscapeJSON(ByVal text As String) As String
    EscapeJSON = Replace(Replace(Replace(Replace(text, "\", "\\"), """", "\"""), vbCrLf, "\n"), vbTab, "\t")
End Function 