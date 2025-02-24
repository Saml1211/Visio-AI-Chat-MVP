# Visio AI Chat MVP - User Guide

## Overview
Visio AI Chat MVP is a VBA-based tool that enhances your Visio diagrams using OpenAI's GPT models. It allows you to automatically analyze and improve the text content of your Visio shapes.

## Prerequisites
- Microsoft Visio 2019 or later
- An OpenAI API key ([Get one here](https://platform.openai.com/api-keys))
- VBA enabled in Visio

## Installation

1. Open Visio and press `Alt + F11` to open the VBA editor
2. In the VBA editor, go to `File > Import File` and import these files:
   - `src/Classes/APIClient.cls`
   - `src/Classes/BatchProcessor.cls`
   - `src/Classes/ErrorLogger.cls`
   - `src/Modules/Constants.bas`
   - `src/Modules/Main.bas`

3. Enable necessary references:
   - In the VBA editor, go to `Tools > References`
   - Check these references:
     - Microsoft Scripting Runtime
     - Microsoft WinHTTP Services

## First-Time Setup

1. In the VBA editor, run the `Initialize` function:
   ```vba
   Initialize
   ```
2. When prompted, enter your OpenAI API key
3. The key will be securely saved for future use

## Using the Tool

### Basic Usage

1. Select one or more shapes in your Visio diagram
2. In the VBA editor, run:
   ```vba
   ProcessSelectedShapes
   ```
3. The tool will:
   - Analyze the text in each selected shape
   - Generate improved content using AI
   - Update the shapes with enhanced text
   - Add a timestamp to the shape's metadata

### Tips
- Select multiple shapes to process them in batch
- The status bar will show progress during processing
- Check the error log file if something goes wrong

## Troubleshooting

### Reset API Key
If you need to change your API key:
```vba
ResetAPIKey
Initialize  ' Will prompt for new key
```

### Common Issues

1. "System not initialized"
   - Solution: Run `Initialize` first

2. "No shapes selected"
   - Solution: Select at least one shape before running `ProcessSelectedShapes`

3. "API request failed"
   - Check your internet connection
   - Verify your API key is valid
   - Check the error log for details

### Error Logs
- Errors are logged to `error_log.txt` in your document's directory
- Each log entry includes timestamp and error details

## Best Practices

1. Save your document before processing large numbers of shapes
2. Process related shapes together for more consistent results
3. Review AI-generated content before saving your document

## Support
For support and feature requests, please:
1. Check the error log for detailed error messages
2. Contact your system administrator
3. Submit issues to the project repository

## Cleanup
When you're done using the tool:
```vba
Cleanup
```
This will properly close all resources and clear the status bar. 