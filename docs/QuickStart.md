# Visio AI Chat MVP - Quick Start Guide

Get started with Visio AI Chat in minutes! This guide covers the essential steps to enhance your Visio diagrams using AI.

## 1. Installation (2 minutes)

1. Open Visio
2. Press `Alt + F11` to open VBA editor
3. Import these files:
   ```
   src/Classes/APIClient.cls
   src/Classes/BatchProcessor.cls
   src/Classes/ErrorLogger.cls
   src/Modules/Constants.bas
   src/Modules/Main.bas
   ```

## 2. Setup (1 minute)

1. Get your OpenAI API key from [platform.openai.com/api-keys](https://platform.openai.com/api-keys)
2. In the VBA editor, type and run:
   ```vba
   Initialize
   ```
3. Enter your API key when prompted

## 3. Enhance Your Diagrams (30 seconds)

1. Select one or more shapes in your diagram
2. In the VBA editor, run:
   ```vba
   ProcessSelectedShapes
   ```
3. Watch as your shapes are enhanced with AI-generated content!

## Common Commands

```vba
' Start the system
Initialize

' Process selected shapes
ProcessSelectedShapes

' Change API key
ResetAPIKey

' Clean up when done
Cleanup
```

## Tips & Tricks

- Select multiple shapes to process them together
- Watch the status bar for progress
- Check `error_log.txt` if something goes wrong
- Save your document before processing large diagrams

## Need Help?

- See `UserGuide.md` for detailed instructions
- Check `TechnicalReference.md` for advanced features
- Contact support for additional assistance 