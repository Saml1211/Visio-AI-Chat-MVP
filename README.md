# Visio AI Chat MVP

A Microsoft Visio VBA integration for AI-powered shape interactions and automated diagram enhancements.

## Features

- API client with configurable endpoints and rate limiting
- Batch request queue system with retry logic
- JSON payload builder with extensible parameters
- Shape-text to AI prompt auto-generation
- Response mapping to shape metadata
- Real-time preview in modeless UserForm
- Version control for API configurations

## Prerequisites

- Microsoft Visio 2019 or later
- VBA Development environment enabled
- Windows OS (for DPAPI integration)

## Project Structure

```
src/
├── Classes/
│   ├── APIClient.cls         - Base API client implementation
│   ├── BatchProcessor.cls    - Request queue management
│   ├── JSONBuilder.cls      - Payload construction
│   ├── ErrorLogger.cls      - Logging implementation
│   └── ConfigManager.cls    - API configuration management
├── Modules/
│   ├── Main.bas            - Entry point and core functions
│   ├── Utils.bas           - Helper utilities
│   └── Constants.bas       - Project constants
└── Forms/
    └── PreviewForm.frm     - Real-time preview UI
```

## Setup Instructions

1. Enable VBA Development in Visio:
   - File → Options → Trust Center → Trust Center Settings
   - Enable "Trust access to the VBA project object model"

2. Import Required References:
   - Microsoft Scripting Runtime
   - Microsoft WinHTTP Services
   - Microsoft Visual Basic for Applications Extensibility

3. Import the VBA Modules:
   - Import all .cls, .bas, and .frm files from the src directory
   - Ensure all dependencies are properly referenced

4. Configure API Settings:
   - Open the ConfigManager and set your API credentials
   - Adjust rate limiting settings if needed (default: 60 RPM)

## Usage

1. Basic Shape Integration:
```vba
' Select a shape and generate AI response
Public Sub UpdateShapeWithAI()
    Dim selectedShape As Visio.Shape
    Set selectedShape = Application.ActiveWindow.Selection(1)
    Call UpdateShapeWithAIResponse(selectedShape)
End Sub
```

2. Batch Processing:
```vba
' Process multiple shapes
Public Sub ProcessShapeBatch()
    Dim shapes As New Collection
    ' Add shapes to collection
    Call ProcessPromptBatch(shapes)
End Sub
```

## Error Handling

The system includes comprehensive error logging:
- All errors are logged to a dedicated worksheet
- Critical errors trigger user notifications
- Automatic retry logic for transient failures

## Security Notes

- API keys are stored securely using Windows DPAPI
- All HTTP communications use TLS 1.2+
- No sensitive data is stored in plain text

## Contributing

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## License

MIT License - See LICENSE file for details

## Support

For support and feature requests, please open an issue in the repository.
