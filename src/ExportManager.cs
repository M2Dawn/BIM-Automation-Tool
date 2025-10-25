using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace BimAutomationTool
{
    public class ExportManager
    {
        private Document _doc;
        private bool _cancelRequested = false;

        public ExportManager(Document doc)
        {
            _doc = doc;
        }

        public void CancelExport()
        {
            _cancelRequested = true;
        }

        public void ExecuteExport(
            List<View> views, 
            List<ViewSchedule> schedules, 
            ExportConfig config, 
            IProgress<ExportProgress> progress,
            ExportLogger logger)
        {
            _cancelRequested = false;
            int totalOperations = CalculateTotalOperations(views, schedules, config);
            int currentOperation = 0;

            logger.LogInfo($"Starting export - {views.Count} views, {schedules.Count} schedules");
            logger.LogInfo($"Output path: {config.OutputPath}");

            try
            {
                // Create output directories
                Directory.CreateDirectory(config.OutputPath);

                // Export views to various formats
                if (config.ExportDWG)
                {
                    currentOperation = ExportViewsToDWG(views, config, progress, logger, currentOperation, totalOperations);
                    if (_cancelRequested) return;
                }

                if (config.ExportPDF)
                {
                    currentOperation = ExportViewsToPDF(views, config, progress, logger, currentOperation, totalOperations);
                    if (_cancelRequested) return;
                }

                if (config.ExportIFC)
                {
                    currentOperation = ExportToIFC(config, progress, logger, currentOperation, totalOperations);
                    if (_cancelRequested) return;
                }

                if (config.ExportNWC)
                {
                    currentOperation = ExportToNWC(config, progress, logger, currentOperation, totalOperations);
                    if (_cancelRequested) return;
                }

                // Export schedules
                if (config.ExportSchedulesCSV)
                {
                    currentOperation = ExportSchedulesToCSV(schedules, config, progress, logger, currentOperation, totalOperations);
                    if (_cancelRequested) return;
                }

                if (config.ExportSchedulesExcel)
                {
                    currentOperation = ExportSchedulesToExcel(schedules, config, progress, logger, currentOperation, totalOperations);
                    if (_cancelRequested) return;
                }

                logger.LogInfo("Export completed successfully");
                ReportProgress(progress, 100, "Export completed!", $"All operations finished successfully");
            }
            catch (Exception ex)
            {
                logger.LogError($"Export failed: {ex.Message}");
                logger.LogError($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        private int CalculateTotalOperations(List<View> views, List<ViewSchedule> schedules, ExportConfig config)
        {
            int total = 0;
            if (config.ExportDWG) total += views.Count;
            if (config.ExportPDF) total += views.Count;
            if (config.ExportIFC) total += 1;
            if (config.ExportNWC) total += 1;
            if (config.ExportSchedulesCSV) total += schedules.Count;
            if (config.ExportSchedulesExcel) total += schedules.Count;
            return total;
        }

        private int ExportViewsToDWG(
            List<View> views, 
            ExportConfig config, 
            IProgress<ExportProgress> progress,
            ExportLogger logger,
            int currentOp, 
            int totalOps)
        {
            string dwgFolder = config.CreateSubfolders 
                ? Path.Combine(config.OutputPath, "DWG") 
                : config.OutputPath;
            Directory.CreateDirectory(dwgFolder);

            logger.LogInfo($"Exporting {views.Count} views to DWG format");

            DWGExportOptions dwgOptions = new DWGExportOptions
            {
                LayerMapping = config.DWGLayerMapping ? "AIA" : null,
                ExportOfSolids = SolidGeometry.Polymesh,
                SharedCoords = true,
                MergedViews = false,
                FileVersion = ACADVersion.R2018,
                Colors = config.DWGColors ? ExportColorMode.TrueColorPerView : ExportColorMode.IndexColors,
                LineScaling = LineScaling.ViewScale
            };

            foreach (var view in views)
            {
                if (_cancelRequested) break;

                try
                {
                    string fileName = GetFileName(view, config);
                    string viewFolder = config.CreateSubfolders 
                        ? Path.Combine(dwgFolder, view.ViewType.ToString()) 
                        : dwgFolder;
                    Directory.CreateDirectory(viewFolder);

                    ViewSet viewSet = new ViewSet();
                    viewSet.Insert(view);

                    _doc.Export(viewFolder, fileName, viewSet, dwgOptions);

                    currentOp++;
                    ReportProgress(progress, (currentOp * 100) / totalOps, 
                        $"Exporting to DWG: {view.Name}", 
                        $"Completed {currentOp} of {totalOps} operations");

                    logger.LogInfo($"Exported DWG: {view.Name}");
                }
                catch (Exception ex)
                {
                    logger.LogError($"Failed to export DWG for view '{view.Name}': {ex.Message}");
                }
            }

            return currentOp;
        }

        private int ExportViewsToPDF(
            List<View> views, 
            ExportConfig config, 
            IProgress<ExportProgress> progress,
            ExportLogger logger,
            int currentOp, 
            int totalOps)
        {
            string pdfFolder = config.CreateSubfolders 
                ? Path.Combine(config.OutputPath, "PDF") 
                : config.OutputPath;
            Directory.CreateDirectory(pdfFolder);

            logger.LogInfo($"Exporting {views.Count} views to PDF format");

            PDFExportOptions pdfOptions = new PDFExportOptions
            {
                FileName = "export",
                Combine = config.PDFCombine,
                RasterQuality = GetPDFRasterQuality(config.PDFQuality),
                AlwaysUseRaster = config.PDFRaster,
                HideCropBoundaries = true,
                HideReferencePlane = true,
                HideScopeBoxes = true,
                HideUnreferencedViewTags = true,
                StopOnError = false
            };

            if (config.PDFCombine)
            {
                // Export all views to a single PDF
                try
                {
                    ViewSet viewSet = new ViewSet();
                    foreach (var view in views)
                        viewSet.Insert(view);

                    string fileName = Path.Combine(pdfFolder, "Combined_Export.pdf");
                    _doc.Export(pdfFolder, "Combined_Export", viewSet, pdfOptions);

                    currentOp += views.Count;
                    ReportProgress(progress, (currentOp * 100) / totalOps, 
                        "Exporting combined PDF", 
                        $"Completed {currentOp} of {totalOps} operations");

                    logger.LogInfo($"Exported combined PDF with {views.Count} views");
                }
                catch (Exception ex)
                {
                    logger.LogError($"Failed to export combined PDF: {ex.Message}");
                }
            }
            else
            {
                // Export each view to separate PDF
                foreach (var view in views)
                {
                    if (_cancelRequested) break;

                    try
                    {
                        string fileName = GetFileName(view, config);
                        string viewFolder = config.CreateSubfolders 
                            ? Path.Combine(pdfFolder, view.ViewType.ToString()) 
                            : pdfFolder;
                        Directory.CreateDirectory(viewFolder);

                        ViewSet viewSet = new ViewSet();
                        viewSet.Insert(view);

                        _doc.Export(viewFolder, fileName, viewSet, pdfOptions);

                        currentOp++;
                        ReportProgress(progress, (currentOp * 100) / totalOps, 
                            $"Exporting to PDF: {view.Name}", 
                            $"Completed {currentOp} of {totalOps} operations");

                        logger.LogInfo($"Exported PDF: {view.Name}");
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Failed to export PDF for view '{view.Name}': {ex.Message}");
                    }
                }
            }

            return currentOp;
        }

        private int ExportToIFC(
            ExportConfig config, 
            IProgress<ExportProgress> progress,
            ExportLogger logger,
            int currentOp, 
            int totalOps)
        {
            string ifcFolder = config.CreateSubfolders 
                ? Path.Combine(config.OutputPath, "IFC") 
                : config.OutputPath;
            Directory.CreateDirectory(ifcFolder);

            logger.LogInfo("Exporting to IFC format");

            try
            {
                IFCExportOptions ifcOptions = new IFCExportOptions
                {
                    FileVersion = GetIFCVersion(config.IFCVersion),
                    SpaceBoundaryLevel = 1,
                    ExportBaseQuantities = config.IFCExportBaseQuantities,
                    WallAndColumnSplitting = true
                };

                string fileName = Path.Combine(ifcFolder, $"{_doc.Title}_IFC.ifc");
                _doc.Export(ifcFolder, $"{_doc.Title}_IFC", ifcOptions);

                currentOp++;
                ReportProgress(progress, (currentOp * 100) / totalOps, 
                    "Exporting to IFC", 
                    $"Completed {currentOp} of {totalOps} operations");

                logger.LogInfo("Exported IFC successfully");
            }
            catch (Exception ex)
            {
                logger.LogError($"Failed to export IFC: {ex.Message}");
            }

            return currentOp;
        }

        private int ExportToNWC(
            ExportConfig config, 
            IProgress<ExportProgress> progress,
            ExportLogger logger,
            int currentOp, 
            int totalOps)
        {
            string nwcFolder = config.CreateSubfolders 
                ? Path.Combine(config.OutputPath, "NWC") 
                : config.OutputPath;
            Directory.CreateDirectory(nwcFolder);

            logger.LogInfo("Exporting to NWC format");

            try
            {
                NavisworksExportOptions nwcOptions = new NavisworksExportOptions
                {
                    ExportScope = NavisworksExportScope.Model,
                    ExportLinks = false,
                    Coordinates = NavisworksCoordinates.Shared,
                    ConvertElementProperties = true,
                    ExportRoomAsAttribute = true,
                    ExportRoomGeometry = true
                };

                string fileName = Path.Combine(nwcFolder, $"{_doc.Title}.nwc");
                _doc.Export(nwcFolder, $"{_doc.Title}", nwcOptions);

                currentOp++;
                ReportProgress(progress, (currentOp * 100) / totalOps, 
                    "Exporting to NWC", 
                    $"Completed {currentOp} of {totalOps} operations");

                logger.LogInfo("Exported NWC successfully");
            }
            catch (Exception ex)
            {
                logger.LogError($"Failed to export NWC: {ex.Message}");
            }

            return currentOp;
        }

        private int ExportSchedulesToCSV(
            List<ViewSchedule> schedules, 
            ExportConfig config, 
            IProgress<ExportProgress> progress,
            ExportLogger logger,
            int currentOp, 
            int totalOps)
        {
            string csvFolder = config.CreateSubfolders 
                ? Path.Combine(config.OutputPath, "Schedules", "CSV") 
                : Path.Combine(config.OutputPath, "Schedules");
            Directory.CreateDirectory(csvFolder);

            logger.LogInfo($"Exporting {schedules.Count} schedules to CSV");

            foreach (var schedule in schedules)
            {
                if (_cancelRequested) break;

                try
                {
                    ExportScheduleToCsv(schedule, csvFolder);

                    currentOp++;
                    ReportProgress(progress, (currentOp * 100) / totalOps, 
                        $"Exporting schedule: {schedule.Name}", 
                        $"Completed {currentOp} of {totalOps} operations");

                    logger.LogInfo($"Exported schedule to CSV: {schedule.Name}");
                }
                catch (Exception ex)
                {
                    logger.LogError($"Failed to export schedule '{schedule.Name}' to CSV: {ex.Message}");
                }
            }

            return currentOp;
        }

        private void ExportScheduleToCsv(ViewSchedule schedule, string outFolder)
        {
            TableData tableData = schedule.GetTableData();
            TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);
            int rows = sectionData.NumberOfRows;
            int cols = sectionData.NumberOfColumns;

            StringBuilder sb = new StringBuilder();

            // Export header
            TableSectionData headerData = tableData.GetSectionData(SectionType.Header);
            if (headerData != null && headerData.NumberOfRows > 0)
            {
                for (int r = 0; r < headerData.NumberOfRows; r++)
                {
                    List<string> rowVals = new List<string>();
                    for (int c = 0; c < headerData.NumberOfColumns; c++)
                    {
                        string cell = headerData.GetCellText(r, c);
                        rowVals.Add(EscapeCSV(cell ?? ""));
                    }
                    sb.AppendLine(string.Join(",", rowVals));
                }
            }

            // Export body
            for (int r = 0; r < rows; r++)
            {
                List<string> rowVals = new List<string>();
                for (int c = 0; c < cols; c++)
                {
                    string cell = sectionData.GetCellText(r, c);
                    rowVals.Add(EscapeCSV(cell ?? ""));
                }
                sb.AppendLine(string.Join(",", rowVals));
            }

            string filename = Path.Combine(outFolder, SanitizeFileName(schedule.Name) + ".csv");
            File.WriteAllText(filename, sb.ToString(), Encoding.UTF8);
        }

        private int ExportSchedulesToExcel(
            List<ViewSchedule> schedules, 
            ExportConfig config, 
            IProgress<ExportProgress> progress,
            ExportLogger logger,
            int currentOp, 
            int totalOps)
        {
            // Note: Excel export would require additional libraries like EPPlus or ClosedXML
            // For now, we'll export enhanced CSV that can be opened in Excel
            string excelFolder = config.CreateSubfolders 
                ? Path.Combine(config.OutputPath, "Schedules", "Excel") 
                : Path.Combine(config.OutputPath, "Schedules");
            Directory.CreateDirectory(excelFolder);

            logger.LogInfo($"Exporting {schedules.Count} schedules to Excel-compatible format");

            foreach (var schedule in schedules)
            {
                if (_cancelRequested) break;

                try
                {
                    // Export as CSV with Excel-friendly formatting
                    ExportScheduleToCsv(schedule, excelFolder);

                    currentOp++;
                    ReportProgress(progress, (currentOp * 100) / totalOps, 
                        $"Exporting schedule to Excel: {schedule.Name}", 
                        $"Completed {currentOp} of {totalOps} operations");

                    logger.LogInfo($"Exported schedule to Excel format: {schedule.Name}");
                }
                catch (Exception ex)
                {
                    logger.LogError($"Failed to export schedule '{schedule.Name}' to Excel: {ex.Message}");
                }
            }

            return currentOp;
        }

        private string GetFileName(View view, ExportConfig config)
        {
            string name = "";

            if (config.UseSheetNumber)
            {
                try
                {
                    var sheetIds = view.GetAllPlacedViews();
                    if (sheetIds != null && sheetIds.Count > 0)
                    {
                        var sheet = _doc.GetElement(sheetIds.First()) as ViewSheet;
                        if (sheet != null)
                            name = $"{sheet.SheetNumber}_{view.Name}";
                    }
                }
                catch { }
            }

            if (string.IsNullOrEmpty(name))
            {
                name = !string.IsNullOrEmpty(config.CustomPrefix) 
                    ? $"{config.CustomPrefix}_{view.Name}" 
                    : view.Name;
            }

            return SanitizeFileName(name);
        }

        private string SanitizeFileName(string fileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string sanitized = new string(fileName.Select(c => invalidChars.Contains(c) ? '_' : c).ToArray());
            return sanitized;
        }

        private string EscapeCSV(string value)
        {
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
            {
                return $"\"{value.Replace("\"", "\"\"")}\"";
            }
            return value;
        }

        private RasterQualityType GetPDFRasterQuality(int qualityIndex)
        {
            switch (qualityIndex)
            {
                case 0: return RasterQualityType.Low;
                case 1: return RasterQualityType.Medium;
                case 2: return RasterQualityType.High;
                case 3: return RasterQualityType.Presentation;
                default: return RasterQualityType.Medium;
            }
        }

        private IFCVersion GetIFCVersion(int versionIndex)
        {
            switch (versionIndex)
            {
                case 0: return IFCVersion.IFC2x2;
                case 1: return IFCVersion.IFC2x3;
                case 2: return IFCVersion.IFC4;
                default: return IFCVersion.IFC4;
            }
        }

        private void ReportProgress(IProgress<ExportProgress> progress, int percentage, string operation, string details)
        {
            progress?.Report(new ExportProgress
            {
                Percentage = percentage,
                CurrentOperation = operation,
                Details = details
            });
        }
    }

    public class ExportProgress
    {
        public int Percentage { get; set; }
        public string CurrentOperation { get; set; }
        public string Details { get; set; }
    }
}
