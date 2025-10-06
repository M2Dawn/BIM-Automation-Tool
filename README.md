# BIM Automation Tool

Custom Revit add-in for batch export automation and schedule data extraction.

## Features
- Batch export of views to DWG format
- Automated schedule data export to CSV
- Custom naming conventions
- Progress tracking and error handling

## Tech Stack
- C# / .NET Framework
- Revit API
- WPF (Windows Presentation Foundation)

## Results
- ~40% time reduction in export workflows
- 500+ views processed per coordination cycle
- ~95% error reduction in data transfer

###  Project Structure
- `src/CmdMain.cs` → Main ExternalCommand implementation.
- `BimAutomationTool.addin` → Add-in manifest for Revit.
- `README.md` → Documentation.
- `LICENSE` → Open-source license (MIT by default).

###  How to Use
1. Clone this repo.
2. Open `src/` in **Visual Studio** (Class Library).
3. Reference Revit API assemblies (`RevitAPI.dll`, `RevitAPIUI.dll`).
4. Build the solution (target **.NET Framework 4.8** or your Revit version).
5. Copy the built DLL + `BimAutomationTool.addin` into:
6. Open Revit → Add-Ins → External Tools → **BIM Automation Tool**.


