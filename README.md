# 🏗️ BIM Automation Tool v2.0

**Advanced Export Manager for Autodesk Revit**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-blue)](https://dotnet.microsoft.com/)
[![Revit](https://img.shields.io/badge/Revit-2022+-red)](https://www.autodesk.com/products/revit/)

A professional-grade add-in for batch export automation with modern UI and comprehensive format support.

**Quick Stats:** ~40% time reduction • 500+ views per cycle • ~95% error reduction • 5 export formats

## ⚡ Quick Start

### Installation (5 minutes)
1. Open `BimAutomationTool.sln` in Visual Studio
2. Update Revit API references to match your installation
3. Build Solution (Ctrl+Shift+B)
4. Copy `BimAutomationTool.dll` and `BimAutomationTool.addin` to:
   ```
   %AppData%\Autodesk\Revit\Addins\2022\
   ```
5. Launch Revit → Add-Ins → External Tools → **BIM Automation Tool**

### Basic Usage (3 steps)
1. **Configure**: Select output folder and export formats (DWG, PDF, IFC, NWC, Schedules)
2. **Select Views**: Use filters or select manually from the view list
3. **Export**: Click "Start Export" and monitor real-time progress

---

## 🚀 Key Features

### Multi-Format Export
- **DWG** - AutoCAD format with layer mapping
- **PDF** - High-quality documents (72-600 DPI)
- **IFC** - Industry Foundation Classes (2x2, 2x3, 4)
- **NWC** - Navisworks coordination
- **CSV/Excel** - Schedule data extraction

### Modern Interface
- Tabbed UI with real-time progress tracking
- Advanced view filtering and bulk selection
- Configuration save/load for reusable workflows
- Statistics dashboard with time estimation

### Smart Features
- Organized subfolders by view type
- Flexible naming conventions
- Comprehensive error handling and logging
- Cancellation support

## 📊 Performance

| Metric | Value |
|--------|-------|
| Time Reduction | ~40% |
| Views per Cycle | 500+ |
| Error Reduction | ~95% |
| Export Formats | 5 |
| User Actions | 3 clicks |

## 🛠️ Tech Stack

**C# / .NET Framework 4.8** • **Revit API 2022+** • **WPF** • **System.Text.Json** • **Async/Await**

## 📁 Project Structure

```
BIM-Automation-Tool/
├── src/
│   ├── CmdMain.cs              # Main command entry point
│   ├── MainWindow.xaml         # WPF UI definition
│   ├── MainWindow.xaml.cs      # UI logic and event handlers
│   ├── ExportManager.cs        # Core export engine
│   ├── ExportConfig.cs         # Configuration model
│   ├── ConfigManager.cs        # Config save/load
│   ├── ExportLogger.cs         # Logging system
│   ├── BimAutomationTool.csproj # Project file
│   └── Properties/
│       └── AssemblyInfo.cs     # Assembly metadata
├── BimAutomationTool.addin     # Revit add-in manifest
├── BimAutomationTool.sln       # Visual Studio solution
├── README.md                   # This file
└── LICENSE                     # MIT License
```

## 🔧 Installation

### Prerequisites
- Visual Studio 2019 or later
- .NET Framework 4.8
- Autodesk Revit 2022 or later

### Build Instructions

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/BIM-Automation-Tool.git
   cd BIM-Automation-Tool
   ```

2. **Open in Visual Studio**
   - Open `BimAutomationTool.sln`

3. **Update Revit API References**
   - Right-click project → Properties → Reference Paths
   - Update paths to match your Revit installation:
     - `C:\Program Files\Autodesk\Revit 2022\RevitAPI.dll`
     - `C:\Program Files\Autodesk\Revit 2022\RevitAPIUI.dll`

4. **Build the Solution**
   - Build → Build Solution (Ctrl+Shift+B)
   - Output: `src\bin\Release\BimAutomationTool.dll`

5. **Deploy to Revit**
   - Copy `BimAutomationTool.dll` to:
     ```
     %AppData%\Autodesk\Revit\Addins\2022\
     ```
   - Copy `BimAutomationTool.addin` to the same location
   - Update the `.addin` file if using a different Revit version

6. **Launch Revit**
   - Open Revit
   - Go to Add-Ins → External Tools → **BIM Automation Tool**

## 📖 Usage Guide

### Basic Workflow

1. **Open a Revit Project**
   - Launch the tool from Add-Ins menu

2. **Configure Export Options**
   - **Output Location**: Choose destination folder
   - **Export Formats**: Select DWG, PDF, IFC, NWC, or Schedules
   - **View Selection**: Filter by type or select manually
   - **Naming Convention**: Choose naming pattern

3. **Review View List**
   - Switch to "View List" tab
   - Use filters to find specific views
   - Select/deselect views as needed

4. **Configure Settings**
   - DWG: Layer mapping, line weights, colors
   - PDF: Quality, rasterization, combine options
   - IFC: Version selection, base quantities

5. **Start Export**
   - Click "Start Export"
   - Monitor progress in real-time
   - Review log file after completion

### Configuration Management

- **Save Config**: Save current settings for reuse
- **Load Config**: Load previously saved settings
- Configs stored as JSON files
- Default config auto-loaded on startup

### Advanced Features

#### View Filtering
- Filter by view type (Floor Plans, Sections, etc.)
- Filter by sheet placement
- Text search across view names and types

#### Batch Operations
- Select All / Deselect All
- Export multiple formats simultaneously
- Automatic subfolder organization

#### Error Handling
- Detailed error logging
- Continue on error (doesn't stop entire batch)
- Export summary with success/failure counts

## 🎯 Export Format Details

### DWG Export
- Layer mapping (AIA standard)
- Line weight preservation
- Color management
- Shared coordinates
- AutoCAD 2018 format

### PDF Export
- Quality levels: Low (72 DPI) to Very High (600 DPI)
- Single or combined PDF output
- Rasterization control
- Hidden element filtering

### IFC Export
- IFC 2x2, 2x3, or IFC 4
- Base quantities export
- Space boundary levels
- Wall and column splitting

### NWC Export
- Model scope export
- Shared coordinates
- Element properties
- Room geometry and attributes

### Schedule Export
- CSV format with proper escaping
- Excel-compatible output
- Header and body sections
- UTF-8 encoding

## 🔍 Troubleshooting

### Common Issues

**Tool doesn't appear in Revit**
- Verify `.addin` file is in correct location
- Check Assembly path in `.addin` matches DLL location
- Ensure Revit version matches

**Export fails**
- Check write permissions on output folder
- Verify views are not in use
- Review log file for specific errors

**UI doesn't load**
- Ensure .NET Framework 4.8 is installed
- Check for missing WPF dependencies
- Review Revit journal for error messages

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with Autodesk Revit API
- Inspired by BIM coordination workflows
- Community feedback and contributions

## 📧 Contact

For questions, issues, or suggestions, please open an issue on GitHub.

---

**Version 2.0** - Major upgrade with modern UI and multi-format support


