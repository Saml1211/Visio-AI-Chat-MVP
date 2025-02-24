Attribute VB_Name = "Main"
Option Explicit

' Global objects
Private processor As BatchProcessor
Private logger As ErrorLogger

' Initialize the system
Public Sub Initialize(Optional ByVal apiKey As String = "")
    On Error GoTo ErrorHandler
    
    ' Create objects
    Set processor = New BatchProcessor
    Set logger = New ErrorLogger
    
    ' Get API key if not provided
    If apiKey = "" Then
        apiKey = GetSetting("VisioAIChat", "Config", "APIKey", "")
        If apiKey = "" Then
            apiKey = InputBox("Please enter your OpenAI API key:", "API Configuration")
            If apiKey = "" Then
                MsgBox ERR_API_KEY_MISSING, vbCritical
                Exit Sub
            End If
            SaveSetting "VisioAIChat", "Config", "APIKey", apiKey
        End If
    End If
    
    ' Configure API
    processor.Configure API_ENDPOINT, apiKey
    
    ' Log initialization
    logger.LogInfo "System initialized successfully"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error initializing system: " & Err.Description, vbCritical
End Sub

' Process selected shapes
Public Sub ProcessSelectedShapes()
    On Error GoTo ErrorHandler
    
    ' Check if processor is initialized
    If processor Is Nothing Then
        MsgBox ERR_NOT_INITIALIZED, vbExclamation
        Exit Sub
    End If
    
    ' Check if shapes are selected
    If Application.ActiveWindow.Selection.Count = 0 Then
        MsgBox ERR_NO_SELECTION, vbExclamation
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "Processing shapes..."
    
    ' Get selected shapes
    Dim shape As Visio.Shape
    For Each shape In Application.ActiveWindow.Selection
        ' Generate prompt from shape text
        Dim prompt As String
        prompt = Replace(DEFAULT_PROMPT_TEMPLATE, "{{text}}", shape.Text)
        
        ' Add to processing queue
        processor.AddToQueue shape, prompt
    Next shape
    
    ' Process all shapes in queue
    processor.ProcessQueue
    
    ' Log completion
    logger.LogInfo "Processed " & Application.ActiveWindow.Selection.Count & " shapes"
    
    ' Clear status bar
    Application.StatusBar = ""
    Exit Sub
    
ErrorHandler:
    logger.LogError "Error processing shapes: " & Err.Description
    MsgBox "Error processing shapes. Check error log for details.", vbExclamation
    Application.StatusBar = ""
End Sub

' Clean up resources
Public Sub Cleanup()
    Set processor = Nothing
    Set logger = Nothing
    Application.StatusBar = ""
End Sub

' Reset API key
Public Sub ResetAPIKey()
    DeleteSetting "VisioAIChat", "Config", "APIKey"
    MsgBox "API key has been reset. Please reinitialize the system.", vbInformation
    Cleanup
End Sub 