using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

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

            try
            {
                // Launch the WPF UI
                MainWindow window = new MainWindow(uiDoc);
                window.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error launching BIM Automation Tool: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}
