using System;

namespace BimAutomationTool
{
    [Serializable]
    public class ExportConfig
    {
        // Output settings
        public string OutputPath { get; set; }
        public bool CreateSubfolders { get; set; }
        public bool OverwriteExisting { get; set; }
        public bool OpenFolderAfter { get; set; }

        // Export formats
        public bool ExportDWG { get; set; }
        public bool ExportPDF { get; set; }
        public bool ExportIFC { get; set; }
        public bool ExportNWC { get; set; }
        public bool ExportSchedulesCSV { get; set; }
        public bool ExportSchedulesExcel { get; set; }

        // View selection
        public bool OnlySheetViews { get; set; }

        // Naming convention
        public bool UseViewName { get; set; }
        public bool UseSheetNumber { get; set; }
        public string CustomPrefix { get; set; }

        // DWG settings
        public bool DWGLayerMapping { get; set; }
        public bool DWGLineWeights { get; set; }
        public bool DWGColors { get; set; }

        // PDF settings
        public bool PDFCombine { get; set; }
        public bool PDFRaster { get; set; }
        public int PDFQuality { get; set; } // 0=Low, 1=Medium, 2=High, 3=VeryHigh

        // IFC settings
        public int IFCVersion { get; set; } // 0=IFC2x2, 1=IFC2x3, 2=IFC4
        public bool IFCExportBaseQuantities { get; set; }

        public ExportConfig()
        {
            // Set defaults
            OutputPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop), 
                "BIM_Exports");
            CreateSubfolders = true;
            OverwriteExisting = false;
            OpenFolderAfter = true;

            ExportDWG = true;
            ExportPDF = false;
            ExportIFC = false;
            ExportNWC = false;
            ExportSchedulesCSV = true;
            ExportSchedulesExcel = false;

            OnlySheetViews = false;

            UseViewName = true;
            UseSheetNumber = false;
            CustomPrefix = "";

            DWGLayerMapping = true;
            DWGLineWeights = false;
            DWGColors = true;

            PDFCombine = false;
            PDFRaster = false;
            PDFQuality = 1;

            IFCVersion = 2;
            IFCExportBaseQuantities = true;
        }
    }
}
