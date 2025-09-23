# BIM Automation Tool

**BIM Automation Tool** is a lightweight Revit add-in (C#) that automates:
- Batch export of selected views to DWG.
- Exporting Revit schedule tables to CSV for reporting/analysis.

## Features
- Export multiple views (sheets, plans) to DWG with DWGExportOptions.
- Extract schedule rows and create a CSV per schedule.
- Simple UI (Revit ExternalCommand) — select views/schedules in the Revit UI and run.

## Requirements
- Revit 2020 / 2021 / 2022 / 2023 (adjust target framework & references)
- Visual Studio 2019/2022
- .NET Framework 4.8 (or the Revit target version)
- Revit API assemblies (RevitAPI.dll, RevitAPIUI.dll) referenced from Revit install folder.

## Installation
1. Build the solution in Visual Studio (Release).
2. Copy the generated `BimAutomationTool.dll` and any dependencies to:
   `%APPDATA%\Autodesk\Revit\Addins\<RevitYear>\`
3. Also place `BimAutomationTool.addin` into the same folder.

## Usage
- Open Revit. Load a project/host model.
- Open the views/schedules you want to export (or select them in Project Browser).
- From Add-Ins → External Tools, run **BIM Automation Tool**.
- Choose output folder for DWG and CSV exports in the dialog.

## Notes
- The tool uses Revit API `Document.Export()` for DWG export and `TableData`/`SchedulableField` for schedule reads.
- If you require PDF export, integrate a DWG → PDF step or use Revit's PrintManager API.

## License
MIT — see LICENSE file.
