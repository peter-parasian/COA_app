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

        private string _numberDO = string.Empty;

        public string DoNumber
        {
            get => _numberDO;
            set { _numberDO = value; OnPropertyChanged(); }
        }

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
            DoNumber = string.Empty;
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

        private string GetRomanMonth(int month)
        {
            switch (month)
            {
                case 1: return "I";
                case 2: return "II";
                case 3: return "III";
                case 4: return "IV";
                case 5: return "V";
                case 6: return "VI";
                case 7: return "VII";
                case 8: return "VIII";
                case 9: return "IX";
                case 10: return "X";
                case 11: return "XI";
                case 12: return "XII";
                default: return "";
            }
        }

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

            if (string.IsNullOrWhiteSpace(DoNumber))
            {
                System.Windows.MessageBox.Show("Silakan input Nomor DO (Delivery Order).", "Peringatan", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            try
            {
                string templatePath = @"C:\Users\mrrx\Documents\My Web Sites\H\TEMPLATE_COA_BUSBAR.xlsx";

                if (!System.IO.File.Exists(templatePath))
                {
                    System.Windows.MessageBox.Show($"File template tidak ditemukan:\n{templatePath}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    return;
                }

                // Membuka Template Excel
                using (var workbook = new ClosedXML.Excel.XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // --- 1. PENGISIAN HEADER ---
                    worksheet.Cell("C12").Value = ": " + PoNumber;
                    worksheet.Cell("J12").Value = ": " + CustomerName;

                    System.DateTime now = System.DateTime.Now;
                    worksheet.Cell("J13").Value = ": " + now.ToString("dd/MM/yyyy");

                    string basePath = @"C:\Users\mrrx\Documents\My Web Sites\H\COA";
                    string yearFolder = now.ToString("yyyy");

                    var cultureIndo = new System.Globalization.CultureInfo("id-ID");
                    string monthName = cultureIndo.DateTimeFormat.GetMonthName(now.Month);
                    monthName = cultureIndo.TextInfo.ToTitleCase(monthName);
                    string monthFolder = $"{now.Month}. {monthName}";
                    string finalDirectory = System.IO.Path.Combine(basePath, yearFolder, monthFolder);

                    if (!System.IO.Directory.Exists(finalDirectory))
                    {
                        System.IO.Directory.CreateDirectory(finalDirectory);
                    }

                    string[] existingFiles = System.IO.Directory.GetFiles(finalDirectory, "*.xlsx");
                    int nomorFile = existingFiles.Length + 1;

                    string romanMonth = GetRomanMonth(now.Month);
                    worksheet.Cell("J14").Value = ": " + $"{nomorFile}/{romanMonth}/{now.Year}";

                    // --- 2. PENGISIAN DATA GRID ---

                    int dataCount = ExportList.Count;
                    int startRowTable1 = 20;
                    int originalStartRowTable2 = 30;
                    int startRowTable2 = originalStartRowTable2;

                    // Insert baris jika data > 3
                    if (dataCount > 3)
                    {
                        int rowsToInsert = dataCount - 3;
                        worksheet.Row(22).InsertRowsBelow(rowsToInsert);
                        startRowTable2 = originalStartRowTable2 + rowsToInsert;
                    }

                    // --- LOOP PENGISIAN DATA ---
                    for (int i = 0; i < dataCount; i++)
                    {
                        var rec = ExportList[i].RecordData;

                        // --- ISI TABEL 1 (ATAS) ---
                        int r1 = startRowTable1 + i;

                        // Data
                        worksheet.Cell(r1, 2).Value = rec.BatchNo;  // B
                        worksheet.Cell(r1, 3).Value = rec.Size;     // C

                        // Merge D & E + Text Multiline
                        var cellD = worksheet.Cell(r1, 4);
                        cellD.Value = "No Dirty\nNo Blackspot\nNo Blisters";
                        cellD.Style.Alignment.WrapText = true;
                        worksheet.Range(r1, 4, r1, 5).Merge(); // Merge D dan E

                        worksheet.Cell(r1, 6).Value = rec.Thickness; // F
                        worksheet.Cell(r1, 7).Value = rec.Width;     // G
                        worksheet.Cell(r1, 8).Value = rec.Length;    // H
                        worksheet.Cell(r1, 9).Value = rec.Radius;    // I
                        worksheet.Cell(r1, 10).Value = rec.Chamber; // J

                        // Kolom K: OK
                        var cellK = worksheet.Cell(r1, 11); // K
                        cellK.Value = "OK";

                        // STYLE TABEL 1 (B sampai K): Bold, Middle Align, Middle Center
                        var rangeT1 = worksheet.Range(r1, 2, r1, 11);
                        rangeT1.Style.Font.Bold = true;
                        rangeT1.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                        rangeT1.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;


                        // --- ISI TABEL 2 (BAWAH) ---
                        int r2 = startRowTable2 + i;

                        // Data
                        worksheet.Cell(r2, 2).Value = rec.BatchNo;    // B
                        worksheet.Cell(r2, 3).Value = rec.Size;       // C
                        worksheet.Cell(r2, 4).Value = rec.Electric;    // D
                        worksheet.Cell(r2, 5).Value = rec.Resistivity;// E
                        worksheet.Cell(r2, 6).Value = rec.Elongation;  // F
                        worksheet.Cell(r2, 7).Value = rec.Tensile;    // G

                        // Kolom H: No Crack
                        worksheet.Cell(r2, 8).Value = "No Crack";

                        worksheet.Cell(r2, 9).Value = rec.Spectro;    // I
                        worksheet.Cell(r2, 10).Value = rec.Oxygen;    // J

                        // Kolom K: OK (Ditambahkan agar ada K30 kebawah)
                        var cellK2 = worksheet.Cell(r2, 11); // K
                        cellK2.Value = "OK";

                        // STYLE TABEL 2 (B sampai K): Bold, Middle Align, Middle Center
                        // Rentang diperluas sampai kolom 11 (K) agar OK ikut diformat
                        var rangeT2 = worksheet.Range(r2, 2, r2, 11);
                        rangeT2.Style.Font.Bold = true;
                        rangeT2.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                        rangeT2.Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                    }

                    // --- 3. FORMATTING FULL BORDER ---

                    // Border Table 1 (B sampai K)
                    var rangeBorder1 = worksheet.Range(startRowTable1, 2, startRowTable1 + dataCount - 1, 11);
                    rangeBorder1.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                    rangeBorder1.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                    rangeBorder1.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                    rangeBorder1.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                    // Inside Border untuk garis pemisah antar sel
                    rangeBorder1.Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;

                    // Border Table 2 (B sampai K)
                    var rangeBorder2 = worksheet.Range(startRowTable2, 2, startRowTable2 + dataCount - 1, 11);
                    rangeBorder2.Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                    rangeBorder2.Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                    rangeBorder2.Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                    rangeBorder2.Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                    // Inside Border untuk garis pemisah antar sel
                    rangeBorder2.Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;

                    // Fit to Page
                    worksheet.PageSetup.PagesTall = 1;
                    worksheet.PageSetup.PagesWide = 1;

                    // --- 4. SIMPAN FILE ---
                    string fileName = $"{nomorFile}. COA {CustomerName} {DoNumber}.xlsx";
                    string fullPath = System.IO.Path.Combine(finalDirectory, fileName);

                    workbook.SaveAs(fullPath);

                    System.Windows.MessageBox.Show($"Data berhasil diexport ke:\n{fullPath}", "Sukses", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);

                    CustomerName = string.Empty;
                    PoNumber = string.Empty;
                    DoNumber = string.Empty;
                    SelectedStandard = null;
                    ExportList.Clear();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"Gagal membuat Excel:\n{ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

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