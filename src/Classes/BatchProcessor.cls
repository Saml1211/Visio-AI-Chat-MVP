VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BatchProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Simple queue for shapes and prompts
Private Type QueueItem
    Shape As Visio.Shape
    Prompt As String
End Type

Private queue() As QueueItem
Private queueSize As Long
Private apiClient As APIClient

' Initialize the processor
Private Sub Class_Initialize()
    ReDim queue(0 To 99) ' Start with space for 100 items
    queueSize = 0
    Set apiClient = New APIClient
End Sub

' Clean up
Private Sub Class_Terminate()
    Set apiClient = Nothing
End Sub

' Configure the processor
Public Sub Configure(ByVal url As String, ByVal key As String)
    apiClient.Configure url, key
End Sub

' Add a shape to the queue
Public Sub AddToQueue(ByVal shape As Visio.Shape, ByVal prompt As String)
    ' Resize array if needed
    If queueSize = UBound(queue) Then
        ReDim Preserve queue(0 To UBound(queue) * 2)
    End If
    
    ' Add to queue
    With queue(queueSize)
        Set .Shape = shape
        .Prompt = prompt
    End With
    queueSize = queueSize + 1
End Sub

' Process all items in queue
Public Sub ProcessQueue()
    Dim i As Long
    
    For i = 0 To queueSize - 1
        With queue(i)
            ' Get AI response
            Dim response As String
            response = apiClient.MakeRequest(.Prompt)
            
            ' Update shape if successful
            If Not response Like "Error:*" Then
                UpdateShape .Shape, response
            End If
        End With
    Next i
    
    ' Clear queue
    queueSize = 0
End Sub

' Update a shape with AI response
Private Sub UpdateShape(ByVal shape As Visio.Shape, ByVal response As String)
    On Error Resume Next
    
    ' Simple text update
    shape.Text = response
    
    ' Add timestamp to shape data
    shape.AddSection visSectionUser
    shape.AddRow visSectionUser, visRowUser
    shape.CellsSRC(visSectionUser, visRowUser, visUserValue).FormulaU = """" & Now & """"
End Sub 