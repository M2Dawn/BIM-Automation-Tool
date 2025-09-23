using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace BimAutomationTool
{
    [Transaction(TransactionMode.Manual)]
    public class CmdMain : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc?.Document;

            if (doc == null)
            {
                message = "Please open a project before running.";
                return Result.Failed;
            }

            string outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string dwgFolder = Path.Combine(outputFolder, "BIM_Exports\\DWG");
            string csvFolder = Path.Combine(outputFolder, "BIM_Exports\\Schedules");
            Directory.CreateDirectory(dwgFolder);
            Directory.CreateDirectory(csvFolder);

            try
            {
                // Export views
                var views = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate && !(v is View3D) && v.CanBePrinted)
                    .ToList();

                if (views.Any())
                {
                    DWGExportOptions dwgOptions = new DWGExportOptions();
                    ViewSet vs = new ViewSet();
                    foreach (var v in views) vs.Insert(v);
                    doc.Export(dwgFolder, "views_export", vs, dwgOptions);
                }

                // Export schedules
                var schedules = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule)).Cast<ViewSchedule>();
                foreach (var sched in schedules)
                {
                    ExportScheduleToCsv(doc, sched, csvFolder);
                }

                TaskDialog.Show("BIM Automation Tool", "Export complete.\nSaved in Desktop/BIM_Exports");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void ExportScheduleToCsv(Document doc, ViewSchedule schedule, string outFolder)
        {
            TableData tableData = schedule.GetTableData();
            TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);
            int rows = sectionData.NumberOfRows;
            int cols = sectionData.NumberOfColumns;

            StringBuilder sb = new StringBuilder();
            for (int r = 0; r < rows; r++)
            {
                List<string> rowVals = new List<string>();
                for (int c = 0; c < cols; c++)
                {
                    string cell = sectionData.GetCellText(r, c);
                    rowVals.Add(cell?.Replace(",", ";") ?? "");
                }
                sb.AppendLine(string.Join(",", rowVals));
            }

            string filename = Path.Combine(outFolder, schedule.Name + ".csv");
            File.WriteAllText(filename, sb.ToString(), Encoding.UTF8);
        }
    }
}
