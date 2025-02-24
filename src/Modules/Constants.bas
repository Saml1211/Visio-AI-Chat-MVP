Attribute VB_Name = "Constants"
Option Explicit

' API Configuration
Public Const API_ENDPOINT As String = "https://api.openai.com"
Public Const API_MODEL As String = "gpt-3.5-turbo"  ' or your preferred model
Public Const MAX_TOKENS As Integer = 150

' Prompt Templates
Public Const DEFAULT_PROMPT_TEMPLATE As String = "Analyze and enhance the following Visio diagram element: {{text}}"

' Error Messages
Public Const ERR_NOT_INITIALIZED As String = "System not initialized. Please run Initialize first."
Public Const ERR_NO_SELECTION As String = "No shapes selected. Please select at least one shape."
Public Const ERR_API_KEY_MISSING As String = "API key not configured. Please set your API key in the Initialize sub." 