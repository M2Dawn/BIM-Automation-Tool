using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace BimAutomationTool
{
    public partial class MainWindow : Window
    {
        private Document _doc;
        private UIDocument _uiDoc;
        private ObservableCollection<ViewItem> _viewItems;
        private List<ViewSchedule> _schedules;
        private bool _isExporting = false;
        private ExportManager _exportManager;

        public MainWindow(UIDocument uiDoc)
        {
            InitializeComponent();
            _uiDoc = uiDoc;
            _doc = uiDoc.Document;
            _viewItems = new ObservableCollection<ViewItem>();
            _schedules = new List<ViewSchedule>();
            _exportManager = new ExportManager(_doc);

            InitializeUI();
            LoadViews();
            LoadSchedules();
            UpdateStatistics();
        }

        private void InitializeUI()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            txtOutputPath.Text = Path.Combine(desktopPath, "BIM_Exports");
            dgViews.ItemsSource = _viewItems;
        }

        private void LoadViews()
        {
            var views = new FilteredElementCollector(_doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.CanBePrinted)
                .OrderBy(v => v.ViewType.ToString())
                .ThenBy(v => v.Name);

            foreach (var view in views)
            {
                string sheetNumber = "";
                try
                {
                    var sheetIds = view.GetAllPlacedViews();
                    if (sheetIds != null && sheetIds.Count > 0)
                    {
                        var sheet = _doc.GetElement(sheetIds.First()) as ViewSheet;
                        if (sheet != null)
                            sheetNumber = sheet.SheetNumber;
                    }
                }
                catch { }

                _viewItems.Add(new ViewItem
                {
                    View = view,
                    Name = view.Name,
                    ViewType = view.ViewType.ToString(),
                    SheetNumber = sheetNumber,
                    IsSelected = ShouldSelectByDefault(view)
                });
            }
        }

        private bool ShouldSelectByDefault(View view)
        {
            // Select floor plans, ceiling plans, elevations, and sections by default
            return view.ViewType == ViewType.FloorPlan ||
                   view.ViewType == ViewType.CeilingPlan ||
                   view.ViewType == ViewType.Elevation ||
                   view.ViewType == ViewType.Section;
        }

        private void LoadSchedules()
        {
            _schedules = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .ToList();
        }

        private void UpdateStatistics()
        {
            txtTotalViews.Text = _viewItems.Count.ToString();
            int selected = _viewItems.Count(v => v.IsSelected);
            txtSelectedViews.Text = selected.ToString();
            txtTotalSchedules.Text = _schedules.Count.ToString();

            // Estimate time (rough calculation: 2 seconds per view per format)
            int formats = 0;
            if (chkDWG.IsChecked == true) formats++;
            if (chkPDF.IsChecked == true) formats++;
            if (chkIFC.IsChecked == true) formats++;
            if (chkNWC.IsChecked == true) formats++;

            int totalOperations = selected * formats;
            if (chkExportSchedules.IsChecked == true)
                totalOperations += _schedules.Count;

            int estimatedSeconds = totalOperations * 2;
            if (estimatedSeconds > 0)
            {
                TimeSpan ts = TimeSpan.FromSeconds(estimatedSeconds);
                txtEstimatedTime.Text = ts.TotalMinutes < 1 
                    ? $"{ts.Seconds}s" 
                    : $"{(int)ts.TotalMinutes}m {ts.Seconds}s";
            }
            else
            {
                txtEstimatedTime.Text = "--";
            }
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select output folder for exports",
                SelectedPath = txtOutputPath.Text
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtOutputPath.Text = dialog.SelectedPath;
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _viewItems)
                item.IsSelected = true;
            UpdateStatistics();
        }

        private void BtnDeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _viewItems)
                item.IsSelected = false;
            UpdateStatistics();
        }

        private void TxtFilter_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filter = txtFilter.Text.ToLower();
            if (string.IsNullOrWhiteSpace(filter))
            {
                dgViews.ItemsSource = _viewItems;
            }
            else
            {
                var filtered = _viewItems.Where(v =>
                    v.Name.ToLower().Contains(filter) ||
                    v.ViewType.ToLower().Contains(filter) ||
                    v.SheetNumber.ToLower().Contains(filter));
                dgViews.ItemsSource = new ObservableCollection<ViewItem>(filtered);
            }
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_isExporting) return;

            // Validation
            if (!_viewItems.Any(v => v.IsSelected) && chkExportSchedules.IsChecked != true)
            {
                MessageBox.Show("Please select at least one view or enable schedule export.", 
                    "No Items Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!chkDWG.IsChecked.GetValueOrDefault() && 
                !chkPDF.IsChecked.GetValueOrDefault() && 
                !chkIFC.IsChecked.GetValueOrDefault() && 
                !chkNWC.IsChecked.GetValueOrDefault() &&
                !chkExportSchedules.IsChecked.GetValueOrDefault())
            {
                MessageBox.Show("Please select at least one export format.", 
                    "No Format Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                _isExporting = true;
                btnExport.IsEnabled = false;
                btnCancel.IsEnabled = true;
                progressBar.Value = 0;

                var config = BuildExportConfig();
                var selectedViews = _viewItems.Where(v => v.IsSelected).Select(v => v.View).ToList();

                var progress = new Progress<ExportProgress>(p =>
                {
                    progressBar.Value = p.Percentage;
                    txtCurrentOperation.Text = p.CurrentOperation;
                    txtProgressDetails.Text = p.Details;
                });

                var logger = new ExportLogger(Path.Combine(config.OutputPath, "export_log.txt"));
                
                await Task.Run(() => _exportManager.ExecuteExport(selectedViews, _schedules, config, progress, logger));

                MessageBox.Show($"Export completed successfully!\n\nExported to: {config.OutputPath}", 
                    "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);

                if (chkOpenFolderAfter.IsChecked == true)
                {
                    System.Diagnostics.Process.Start("explorer.exe", config.OutputPath);
                }

                logger.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed:\n{ex.Message}", 
                    "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _isExporting = false;
                btnExport.IsEnabled = true;
                btnCancel.IsEnabled = false;
                txtCurrentOperation.Text = "Export completed";
                progressBar.Value = 100;
            }
        }

        private ExportConfig BuildExportConfig()
        {
            return new ExportConfig
            {
                OutputPath = txtOutputPath.Text,
                ExportDWG = chkDWG.IsChecked.GetValueOrDefault(),
                ExportPDF = chkPDF.IsChecked.GetValueOrDefault(),
                ExportIFC = chkIFC.IsChecked.GetValueOrDefault(),
                ExportNWC = chkNWC.IsChecked.GetValueOrDefault(),
                ExportSchedulesCSV = chkExportSchedules.IsChecked.GetValueOrDefault(),
                ExportSchedulesExcel = chkExportSchedulesExcel.IsChecked.GetValueOrDefault(),
                CreateSubfolders = chkCreateSubfolders.IsChecked.GetValueOrDefault(),
                OverwriteExisting = chkOverwriteExisting.IsChecked.GetValueOrDefault(),
                UseViewName = rbViewName.IsChecked.GetValueOrDefault(),
                UseSheetNumber = rbSheetNumber.IsChecked.GetValueOrDefault(),
                CustomPrefix = rbCustom.IsChecked.GetValueOrDefault() ? txtCustomPrefix.Text : "",
                OnlySheetViews = chkOnlySheets.IsChecked.GetValueOrDefault(),
                DWGLayerMapping = chkDWGLayerMapping.IsChecked.GetValueOrDefault(),
                DWGLineWeights = chkDWGLineWeights.IsChecked.GetValueOrDefault(),
                DWGColors = chkDWGColors.IsChecked.GetValueOrDefault(),
                PDFCombine = chkPDFCombine.IsChecked.GetValueOrDefault(),
                PDFRaster = chkPDFRaster.IsChecked.GetValueOrDefault(),
                PDFQuality = cmbPDFQuality.SelectedIndex,
                IFCVersion = cmbIFCVersion.SelectedIndex,
                IFCExportBaseQuantities = chkIFCExportBaseQuantities.IsChecked.GetValueOrDefault()
            };
        }

        private void BtnSaveConfig_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "JSON Config|*.json",
                DefaultExt = "json",
                FileName = "export_config.json"
            };

            if (dialog.ShowDialog() == true)
            {
                var config = BuildExportConfig();
                ConfigManager.SaveConfig(config, dialog.FileName);
                MessageBox.Show("Configuration saved successfully!", "Save Config", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnLoadConfig_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "JSON Config|*.json",
                DefaultExt = "json"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var config = ConfigManager.LoadConfig(dialog.FileName);
                    ApplyConfig(config);
                    MessageBox.Show("Configuration loaded successfully!", "Load Config", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to load configuration:\n{ex.Message}", 
                        "Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ApplyConfig(ExportConfig config)
        {
            txtOutputPath.Text = config.OutputPath;
            chkDWG.IsChecked = config.ExportDWG;
            chkPDF.IsChecked = config.ExportPDF;
            chkIFC.IsChecked = config.ExportIFC;
            chkNWC.IsChecked = config.ExportNWC;
            chkExportSchedules.IsChecked = config.ExportSchedulesCSV;
            chkExportSchedulesExcel.IsChecked = config.ExportSchedulesExcel;
            chkCreateSubfolders.IsChecked = config.CreateSubfolders;
            chkOverwriteExisting.IsChecked = config.OverwriteExisting;
            rbViewName.IsChecked = config.UseViewName;
            rbSheetNumber.IsChecked = config.UseSheetNumber;
            rbCustom.IsChecked = !string.IsNullOrEmpty(config.CustomPrefix);
            txtCustomPrefix.Text = config.CustomPrefix;
            chkOnlySheets.IsChecked = config.OnlySheetViews;
            chkDWGLayerMapping.IsChecked = config.DWGLayerMapping;
            chkDWGLineWeights.IsChecked = config.DWGLineWeights;
            chkDWGColors.IsChecked = config.DWGColors;
            chkPDFCombine.IsChecked = config.PDFCombine;
            chkPDFRaster.IsChecked = config.PDFRaster;
            cmbPDFQuality.SelectedIndex = config.PDFQuality;
            cmbIFCVersion.SelectedIndex = config.IFCVersion;
            chkIFCExportBaseQuantities.IsChecked = config.IFCExportBaseQuantities;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            // Implement cancellation logic
            _exportManager.CancelExport();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            if (_isExporting)
            {
                var result = MessageBox.Show("Export is in progress. Are you sure you want to close?", 
                    "Confirm Close", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                    return;
            }
            Close();
        }
    }

    public class ViewItem : INotifyPropertyChanged
    {
        private bool _isSelected;

        public View View { get; set; }
        public string Name { get; set; }
        public string ViewType { get; set; }
        public string SheetNumber { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
