# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ChatCAD is an AutoCAD plugin that integrates AI chatbot capabilities directly into the AutoCAD environment. The plugin enables users to interact with AI models (primarily Qwen/DashScope API) through a sidebar interface within AutoCAD, allowing for AI-assisted CAD operations, drawing analysis, and automated design tasks.

## Build and Development Commands

### Build the Plugin
```bash
# Build Debug version
msbuild chatcad.sln /p:Configuration=Debug

# Build Release version  
msbuild chatcad.sln /p:Configuration=Release

# Alternative using dotnet (if applicable)
dotnet build chatcad.sln --configuration Debug
dotnet build chatcad.sln --configuration Release
```

### Package Management
```bash
# Restore NuGet packages
nuget restore chatcad.sln
```

### Deploy Plugin to AutoCAD
The plugin builds to `bin\Debug\` or `bin\Release\` and should be loaded into AutoCAD using the `NETLOAD` command or by placing the DLL in AutoCAD's plugin directory.

## Architecture Overview

### Core Components

**Command.cs** - Main entry point and command handler
- Implements `IExtensionApplication` for AutoCAD plugin lifecycle
- Contains AutoCAD command definitions (like `TT`, `QWENSDK`)
- Manages plugin initialization and palette setup
- Serves as the bridge between AutoCAD and the chat interface

**PaletteSetDlg.cs** - Main UI Component
- WinForms-based user control that provides the chat interface
- Uses WebView2 for rendering markdown-formatted AI responses
- Handles user input and message display
- Manages conversation flow and UI interactions

**QwenClient.cs** - AI Model Integration
- Custom SDK implementation for Qwen/DashScope API
- Supports both streaming and non-streaming completions
- Handles function calling and tool integration
- Manages HTTP client for API communications

**CadDrawingHelper.cs** - AutoCAD Drawing Operations
- Contains utilities for CAD drawing manipulation
- Handles geometric calculations and entity processing
- Manages coordinate transformations and drawing analysis
- Integrates model responses with CAD operations

**CadPlotUploader.cs** - Image Processing Integration
- Handles uploading CAD plot images to external services
- Manages coordinate system transformations
- Facilitates AI analysis of CAD drawings through image processing

### Key Dependencies

- **AutoCAD 2018 .NET API** - Core CAD functionality
- **Newtonsoft.Json** - JSON serialization for API communications
- **Microsoft.Web.WebView2** - Web rendering for chat interface
- **Markdig** - Markdown processing for AI responses
- **OpenCvSharp4** - Image processing capabilities
- **ClosedXML/NPOI** - Excel/Office document processing

### Plugin Architecture Flow

1. **Initialization**: Command class registers with AutoCAD and sets up palette
2. **UI Setup**: PaletteSetDlg creates the chat interface sidebar
3. **User Interaction**: Users interact through the chat interface
4. **AI Processing**: QwenClient handles communication with AI models
5. **CAD Integration**: CadDrawingHelper applies AI responses to drawings
6. **Image Analysis**: CadPlotUploader enables image-based AI analysis

### Configuration

The plugin uses `app.config` for configuration:
- `DASHSCOPE_API_KEY` - API key for Qwen model access
- `DASHSCOPE_BASE_URL` - Base URL for DashScope API

### AutoCAD Commands

- `TT` - Initializes the main chat interface palette
- `QWENSDK` - Direct access to Qwen SDK functionality
- Additional plot and analysis commands defined in Command.cs

### Development Notes

- Target Framework: .NET Framework 4.6.2
- AutoCAD Version: 2018
- The plugin creates a persistent sidebar (PaletteSet) for user interaction
- Server communication defaults to `http://127.0.0.1:8000` for local development
- Image processing and analysis capabilities are integrated for CAD drawing interpretation
- The architecture supports both local AI processing and cloud-based model integration