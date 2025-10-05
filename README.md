# BIM Automation Tool

A lightweight Revit add-in (C#) that automates repetitive BIM tasks:

- Batch export of selected views/sheets to **DWG**.
- Export of Revit schedule data to **CSV**.

###  Features
- Saves hours during coordination and documentation stages.
- Built with the Revit API (`RevitAPI.dll`, `RevitAPIUI.dll`).
- Simple and extensible — easy to adapt to other tasks.

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


