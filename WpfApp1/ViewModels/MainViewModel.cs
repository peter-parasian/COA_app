using System;
using System.Linq;
using System.Windows.Input;
using WpfApp1.Core.Models;
using WpfApp1.Core.Services;
using WpfApp1.Data.Database;
using WpfApp1.Data.Repositories;

namespace WpfApp1.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private SqliteContext _dbContext;
        private BusbarRepository _repository;
        private ExcelImportService _importService;

        private readonly object _lockObject = new object();

        private string _debugLog = string.Empty;
        public string DebugLog
        {
            get => _debugLog;
            set { _debugLog = value; OnPropertyChanged(); }
        }

        public int TotalFilesFound { get; private set; }
        public int TotalRowsInserted { get; private set; }

        private bool _isBusy = false;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                _isBusy = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsNotBusy));
            }
        }

        public bool IsNotBusy => !IsBusy;

        private bool _showBlankPage = false;
        public bool ShowBlankPage
        {
            get => _showBlankPage;
            set { _showBlankPage = value; OnPropertyChanged(); }
        }

        public event System.Action<string>? OnShowMessage;

        public System.Collections.ObjectModel.ObservableCollection<string> Years { get; set; } = new System.Collections.ObjectModel.ObservableCollection<string>();
        public System.Collections.ObjectModel.ObservableCollection<string> Months { get; set; } = new System.Collections.ObjectModel.ObservableCollection<string>();
        public System.Collections.ObjectModel.ObservableCollection<string> Standards { get; set; } = new System.Collections.ObjectModel.ObservableCollection<string>();

        private string? _selectedYear;
        public string? SelectedYear
        {
            get => _selectedYear;
            set
            {
                if (_selectedYear != value)
                {
                    _selectedYear = value;
                    OnPropertyChanged();
                    SetDefaultProductionDate();
                }
            }
        }

        private string? _selectedMonth;
        public string? SelectedMonth
        {
            get => _selectedMonth;
            set
            {
                if (_selectedMonth != value)
                {
                    _selectedMonth = value;
                    OnPropertyChanged();
                    SetDefaultProductionDate();
                }
            }
        }

        private System.DateTime? _selectedDate;
        public System.DateTime? SelectedDate
        {
            get => _selectedDate;
            set { _selectedDate = value; OnPropertyChanged(); }
        }

        private string? _selectedStandard;
        public string? SelectedStandard
        {
            get => _selectedStandard;
            set { _selectedStandard = value; OnPropertyChanged(); }
        }

        // --- NEW PROPERTIES FOR PRINT COA ---
        private string _customerName = string.Empty;
        public string CustomerName
        {
            get => _customerName;
            set { _customerName = value; OnPropertyChanged(); }
        }

        private string _poNumber = string.Empty;
        public string PoNumber
        {
            get => _poNumber;
            set { _poNumber = value; OnPropertyChanged(); }
        }
        // -----------------------------------

        private System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem> _searchResults = new System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem>();
        public System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem> SearchResults
        {
            get => _searchResults;
            set { _searchResults = value; OnPropertyChanged(); }
        }

        public System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarExportItem> ExportList { get; set; }
            = new System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarExportItem>();

        public System.Windows.Input.ICommand FindCommand { get; }
        public System.Windows.Input.ICommand AddToExportCommand { get; }
        public System.Windows.Input.ICommand RemoveFromExportCommand { get; }
        public System.Windows.Input.ICommand PrintCoaCommand { get; }

        public MainViewModel()
        {
            _dbContext = new SqliteContext();
            _repository = new BusbarRepository(_dbContext);
            _importService = new ExcelImportService(_repository);

            _importService.OnDebugMessage += (msg) => {
                lock (_lockObject)
                {
                    if (DebugLog.Length > 5000) DebugLog = string.Empty;
                    DebugLog += msg + System.Environment.NewLine;
                }
            };

            InitializeSearchData();

            LoadAvailableYears();

            FindCommand = new RelayCommand(ExecuteFind);
            AddToExportCommand = new RelayCommand(ExecuteAddToExport);
            RemoveFromExportCommand = new RelayCommand(ExecuteRemoveFromExport);
            PrintCoaCommand = new RelayCommand(ExecutePrintCoa);
        }

        public void ImportExcelToSQLite()
        {
            try
            {
                _dbContext.EnsureDatabaseFolderExists();
                ResetCounters();

                using var connection = _dbContext.CreateConnection();
                using var transaction = connection.BeginTransaction();

                _importService.Import(connection, transaction);

                transaction.Commit();

                TotalFilesFound = _importService.TotalFilesFound;
                TotalRowsInserted = _importService.TotalRowsInserted;

                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    //OnShowMessage?.Invoke($"IMPORT SELESAI\n\nFile ditemukan : {TotalFilesFound}\nBaris disimpan : {TotalRowsInserted}\n\nDebug Log:\n{DebugLog}");

                    LoadAvailableYears();
                });
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        public void ButtonMode2_Click()
        {
            OnShowMessage?.Invoke("MODE 2 belum diimplementasikan");
        }

        public void ButtonMode3_Click()
        {
            OnShowMessage?.Invoke("MODE 3 belum diimplementasikan");
        }

        public void ButtonMode4_Click()
        {
            OnShowMessage?.Invoke("MODE 4 belum diimplementasikan");
        }

        public void BackToMenu()
        {
            ResetSearchData();
            CustomerName = string.Empty;
            PoNumber = string.Empty;
            ExportList.Clear();
            ShowBlankPage = false;
        }

        private void ResetCounters()
        {
            TotalFilesFound = 0;
            TotalRowsInserted = 0;
            DebugLog = "";
        }

        private void InitializeSearchData()
        {
            Months.Add("January");
            Months.Add("February");
            Months.Add("March");
            Months.Add("April");
            Months.Add("May");
            Months.Add("June");
            Months.Add("July");
            Months.Add("August");
            Months.Add("September");
            Months.Add("October");
            Months.Add("November");
            Months.Add("December");

            Standards.Add("JIS");
            Standards.Add("DIN");
            Standards.Add("ASTM");

            SearchResults.Clear();
        }

        private void LoadAvailableYears()
        {
            try
            {
                var dbYears = _repository.GetAvailableYears();

                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    Years.Clear();
                    foreach (var year in dbYears)
                    {
                        Years.Add(year);
                    }
                });
            }
            catch (System.Exception ex)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    System.Windows.MessageBox.Show($"Error loading years: {ex.Message}");
                });
            }
        }

        private void SetDefaultProductionDate()
        {
            if (!string.IsNullOrWhiteSpace(SelectedYear) && !string.IsNullOrWhiteSpace(SelectedMonth))
            {
                if (int.TryParse(SelectedYear, out int year))
                {
                    string engMonth = ConvertMonthToEnglish(SelectedMonth);

                    int month = WpfApp1.Shared.Helpers.DateHelper.GetMonthNumber(engMonth);

                    if (month > 0 && month <= 12)
                    {
                        SelectedDate = new System.DateTime(year, month, 1);
                    }
                }
            }
        }

        private void ExecuteFind(object? parameter)
        {
            if (string.IsNullOrWhiteSpace(SelectedYear))
            {
                OnShowMessage?.Invoke("Harap memilih YEAR.");
                return;
            }

            if (string.IsNullOrWhiteSpace(SelectedMonth))
            {
                OnShowMessage?.Invoke("Harap memilih MONTH.");
                return;
            }

            if (SelectedDate == null)
            {
                OnShowMessage?.Invoke("Harap memilih PRODUCTION DATE.");
                return;
            }

            try
            {
                string dbMonth = ConvertMonthToEnglish(SelectedMonth);
                string dbDate = SelectedDate.Value.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                var data = _repository.SearchBusbarRecords(SelectedYear, dbMonth, dbDate);

                SearchResults.Clear();
                foreach (var item in data)
                {
                    SearchResults.Add(item);
                }

                if (SearchResults.Count == 0)
                {
                    OnShowMessage?.Invoke("Data tidak ditemukan untuk kriteria tersebut.");
                }
                //else
                //{
                //    OnShowMessage?.Invoke($"Ditemukan {SearchResults.Count} data.");
                //}
            }
            catch (System.Exception ex)
            {
                OnShowMessage?.Invoke($"Terjadi kesalahan saat pencarian: {ex.Message}");
            }
        }

        private void ExecuteAddToExport(object? parameter)
        {
            if (parameter is BusbarSearchItem selectedItem)
            {
                bool exists = ExportList.Any(x => x.RecordData.Id == selectedItem.FullRecord.Id);

                if (!exists)
                {
                    var exportItem = new WpfApp1.Core.Models.BusbarExportItem(selectedItem.FullRecord);
                    ExportList.Add(exportItem);
                }
                else
                {
                    OnShowMessage?.Invoke("Data ini sudah ada dalam daftar Export.");
                }
            }
        }

        private void ExecuteRemoveFromExport(object? parameter)
        {
            if (parameter is WpfApp1.Core.Models.BusbarExportItem itemToRemove)
            {
                ExportList.Remove(itemToRemove);
            }
        }

        // --- UPDATED LOGIC FOR PRINT COA ---
        private void ExecutePrintCoa(object? parameter)
        {
            if (ExportList.Count == 0)
            {
                System.Windows.MessageBox.Show("Export List kosong. Silakan pilih data terlebih dahulu.", "Peringatan", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(SelectedStandard))
            {
                System.Windows.MessageBox.Show("Silakan pilih Standard terlebih dahulu.", "Peringatan", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(CustomerName))
            {
                System.Windows.MessageBox.Show("Silakan input Nama Perusahaan.", "Peringatan", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(PoNumber))
            {
                System.Windows.MessageBox.Show("Silakan input Nomor PO.", "Peringatan", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Tentukan nilai standar (Prototype Logic)
                int standardValue = 0;
                switch (SelectedStandard?.ToUpper())
                {
                    case "JIS":
                        standardValue = 10;
                        break;
                    case "DIN":
                        standardValue = 20;
                        break;
                    case "ASTM":
                        standardValue = 30;
                        break;
                    default:
                        standardValue = 0;
                        break;
                }

                using (var workbook = new ClosedXML.Excel.XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("COA");

                    worksheet.Cell("B2").Value = "Standard Code (Value):";
                    worksheet.Cell("C2").Value = standardValue;

                    worksheet.Cell("B3").Value = "Company Name:";
                    worksheet.Cell("C3").Value = CustomerName;

                    worksheet.Cell("B4").Value = "PO Number:";
                    worksheet.Cell("C4").Value = PoNumber;

                    int startRow = 6;
                    worksheet.Cell(startRow, 1).Value = "Batch_no";
                    worksheet.Cell(startRow, 2).Value = "Size_mm";
                    worksheet.Cell(startRow, 3).Value = "Thickness_mm";
                    worksheet.Cell(startRow, 4).Value = "Width_mm";
                    worksheet.Cell(startRow, 5).Value = "Length";
                    worksheet.Cell(startRow, 6).Value = "Radius";
                    worksheet.Cell(startRow, 7).Value = "Chamber_mm";
                    worksheet.Cell(startRow, 8).Value = "Electric_IACS";
                    worksheet.Cell(startRow, 9).Value = "Weight"; 
                    worksheet.Cell(startRow, 10).Value = "Elongation";
                    worksheet.Cell(startRow, 11).Value = "Tensile";
                    worksheet.Cell(startRow, 12).Value = "Bend_test";
                    worksheet.Cell(startRow, 13).Value = "Spectro_Cu";

                    int currentRow = startRow + 1;
                    foreach (var item in ExportList)
                    {
                        var rec = item.RecordData;

                        worksheet.Cell(currentRow, 1).Value = rec.BatchNo;
                        worksheet.Cell(currentRow, 2).Value = rec.Size;
                        worksheet.Cell(currentRow, 3).Value = rec.Thickness;
                        worksheet.Cell(currentRow, 4).Value = rec.Width;
                        worksheet.Cell(currentRow, 5).Value = rec.Length;
                        worksheet.Cell(currentRow, 6).Value = rec.Radius;
                        worksheet.Cell(currentRow, 7).Value = rec.Chamber;
                        worksheet.Cell(currentRow, 8).Value = rec.Electric;
                        worksheet.Cell(currentRow, 9).Value = rec.Resistivity;
                        worksheet.Cell(currentRow, 10).Value = rec.Elongation;
                        worksheet.Cell(currentRow, 11).Value = rec.Tensile;
                        worksheet.Cell(currentRow, 12).Value = rec.BendTest;
                        worksheet.Cell(currentRow, 13).Value = rec.Spectro;

                        currentRow++;
                    }

                    string desktopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                    string fileName = $"COA_{System.DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                    string fullPath = System.IO.Path.Combine(desktopPath, fileName);

                    workbook.SaveAs(fullPath);

                    System.Windows.MessageBox.Show($"Data berhasil diexport ke:\n{fullPath}", "Sukses", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);

                    CustomerName = string.Empty;
                    PoNumber = string.Empty;
                    SelectedStandard = null;
                    ExportList.Clear();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"Gagal membuat Excel:\n{ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
        // -----------------------------------

        private string ConvertMonthToEnglish(string indoMonth)
        {
            switch (indoMonth)
            {
                case "January": return "January";
                case "February": return "February";
                case "March": return "March";
                case "April": return "April";
                case "May": return "May";
                case "June": return "June";
                case "July": return "July";
                case "August": return "August";
                case "September": return "September";
                case "October": return "October";
                case "November": return "November";
                case "December": return "December";
                default: return indoMonth;
            }
        }

        private void ResetSearchData()
        {
            SelectedYear = null;
            SelectedMonth = null;
            SelectedDate = null;
            SelectedStandard = null;
            SearchResults.Clear();
        }

    }

    public class RelayCommand : System.Windows.Input.ICommand
    {
        private readonly System.Action<object?> _execute;
        private readonly System.Func<object?, bool>? _canExecute;

        public RelayCommand(System.Action<object?> execute, System.Func<object?, bool>? canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public event System.EventHandler? CanExecuteChanged
        {
            add { System.Windows.Input.CommandManager.RequerySuggested += value; }
            remove { System.Windows.Input.CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object? parameter) => _canExecute == null || _canExecute(parameter);

        public void Execute(object? parameter) => _execute(parameter);
    }
}