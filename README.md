# Visio AI Chat MVP

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/version-0.1.0-blue.svg)](https://github.com/yourusername/Visio-AI-Chat-MVP/releases/tag/v0.1.0)

A Microsoft Visio VBA integration that enhances your diagrams using OpenAI's GPT models. This tool automatically analyzes and improves the text content of your Visio shapes using AI.

![Visio AI Chat Demo](docs/images/demo.gif)

## 🌟 Features

- 🤖 OpenAI GPT integration for intelligent shape text enhancement
- 📦 Batch processing of multiple shapes
- 🔑 Secure API key management
- 📊 Progress tracking and status updates
- 📝 Comprehensive error logging
- 🔄 Automatic shape metadata updates

## 🚀 Quick Start

1. **Prerequisites**
   - Microsoft Visio 2019 or later
   - OpenAI API key ([Get one here](https://platform.openai.com/api-keys))
   - VBA enabled in Visio

2. **Installation**
   ```
   1. Open Visio and press Alt + F11
   2. Import all .cls and .bas files from src/
   3. Enable required references in Tools > References
   ```

3. **First Run**
   ```vba
   ' Initialize the system (will prompt for API key)
   Initialize
   
   ' Process selected shapes
   ProcessSelectedShapes
   ```

👉 [View the Quick Start Guide](docs/QuickStart.md) for detailed instructions.

## 📚 Documentation

- [User Guide](docs/UserGuide.md) - Complete usage instructions
- [Technical Reference](docs/TechnicalReference.md) - Architecture and development details
- [Quick Start Guide](docs/QuickStart.md) - Get up and running quickly
- [Changelog](CHANGELOG.md) - Version history and updates

## 🏗️ Architecture

The solution consists of five main components:

1. **APIClient** - Handles OpenAI API communication
2. **BatchProcessor** - Manages shape processing queue
3. **ErrorLogger** - Provides logging functionality
4. **Constants** - Centralizes configuration
5. **Main** - Manages system lifecycle

## 🔒 Security

- API keys stored securely using VBA's built-in settings
- HTTPS-only communication
- No sensitive data in logs
- Basic access controls

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- OpenAI for their powerful GPT models
- Microsoft Visio team for their extensible platform
- All contributors and users of this project

## 📞 Support

If you encounter any issues or have questions:

1. Check the [User Guide](docs/UserGuide.md)
2. Review the error log (`error_log.txt`)
3. [Open an issue](../../issues) on GitHub

---
Made with ❤️ by the Visio AI Chat team
