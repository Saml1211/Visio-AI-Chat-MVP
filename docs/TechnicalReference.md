# Visio AI Chat MVP - Technical Reference

## Architecture Overview

The solution consists of five main components:

1. **APIClient** (`APIClient.cls`)
   - Handles HTTP communication with OpenAI API
   - Manages API authentication
   - Formats requests and parses responses

2. **BatchProcessor** (`BatchProcessor.cls`)
   - Manages queue of shapes to process
   - Handles batch processing logic
   - Updates Visio shapes with AI responses

3. **ErrorLogger** (`ErrorLogger.cls`)
   - Provides logging functionality
   - Writes to text-based log file
   - Timestamps all entries

4. **Constants** (`Constants.bas`)
   - Centralizes configuration
   - Defines error messages
   - Stores API settings

5. **Main** (`Main.bas`)
   - Provides main entry points
   - Manages system lifecycle
   - Handles user interaction

## Component Details

### APIClient Class

```vba
' Key Methods
Configure(url As String, key As String)
MakeRequest(prompt As String) As String
```

- Uses WinHTTP for API communication
- Implements OpenAI Chat API format
- Handles JSON escaping and parsing
- Returns plain text responses

### BatchProcessor Class

```vba
' Key Methods
Configure(url As String, key As String)
AddToQueue(shape As Visio.Shape, prompt As String)
ProcessQueue()
```

- Uses dynamic array for queue management
- Processes shapes in sequence
- Updates shape text and metadata
- Handles errors per shape

### ErrorLogger Class

```vba
' Key Methods
LogError(message As String)
LogInfo(message As String)
```

- File-based logging
- Append-only operation
- Thread-safe file handling
- Automatic timestamps

### Constants Module

```vba
' Key Constants
API_ENDPOINT As String
API_MODEL As String
MAX_TOKENS As Integer
DEFAULT_PROMPT_TEMPLATE As String
```

- Centralizes configuration
- Defines error messages
- Stores API parameters

### Main Module

```vba
' Key Methods
Initialize(Optional apiKey As String = "")
ProcessSelectedShapes()
Cleanup()
ResetAPIKey()
```

- Manages system lifecycle
- Handles API key storage
- Provides user feedback
- Implements error handling

## Data Flow

1. User selects shapes and initiates processing
2. Main module validates selection and state
3. BatchProcessor queues shapes
4. For each shape:
   - Text is extracted
   - Prompt is generated
   - APIClient makes request
   - Response is parsed
   - Shape is updated
5. Results are logged
6. Status is updated

## Error Handling

### Hierarchy
1. Shape-level errors (contained)
2. Batch-level errors (logged)
3. System-level errors (displayed)

### Implementation
```vba
On Error GoTo ErrorHandler
' ... code ...
ErrorHandler:
    logger.LogError "Error: " & Err.Description
    MsgBox "Error occurred", vbExclamation
```

## Security Considerations

1. API Key Storage
   - Uses VBA's `SaveSetting`/`GetSetting`
   - Key never stored in code
   - Can be reset if compromised

2. HTTP Security
   - Uses HTTPS only
   - TLS 1.2 minimum
   - Bearer token auth

3. Error Logging
   - No sensitive data in logs
   - Local file system only
   - Basic access controls

## Performance

### Optimization
- Batch processing
- Array-based queue
- Minimal API calls
- Efficient string handling

### Limitations
- Synchronous processing
- Single-threaded operation
- Memory scales with batch size

## Extension Points

### Adding Features
1. Create new module/class
2. Update Main for initialization
3. Add to Constants if needed
4. Document in UserGuide

### Modifying Behavior
1. Update Constants for configuration
2. Modify prompt templates
3. Adjust batch parameters
4. Enhance error handling

## Best Practices

### Development
1. Use Option Explicit
2. Handle all errors
3. Clean up resources
4. Log operations

### Deployment
1. Test in isolated environment
2. Verify references
3. Check security settings
4. Document changes

## Testing

### Manual Tests
1. Initialize system
2. Process single shape
3. Process multiple shapes
4. Handle errors
5. Reset configuration

### Error Cases
1. Invalid API key
2. Network failure
3. Invalid shapes
4. Queue overflow
5. File system issues 