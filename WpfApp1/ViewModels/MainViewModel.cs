using System;
using System.Linq;
using System.Windows.Input;
using System.Windows.Threading;
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
        private CoaPrintService _printService;

        private System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem>> _searchCache;

        private readonly object _logLock = new object();
        private System.Text.StringBuilder _logBuffer = new System.Text.StringBuilder();
        private System.Windows.Threading.DispatcherTimer _logTimer;

        #region Properties

        private string _debugLog = string.Empty;
        public string DebugLog
        {
            get => _debugLog;
            set { _debugLog = value; OnPropertyChanged(); }
        }

        private int _progressValue = 0;
        public int ProgressValue
        {
            get => _progressValue;
            set { _progressValue = value; OnPropertyChanged(); }
        }

        private int _progressMaximum = 100;
        public int ProgressMaximum
        {
            get => _progressMaximum;
            set { _progressMaximum = value; OnPropertyChanged(); }
        }

        private string _progressText = "Ready";
        public string ProgressText
        {
            get => _progressText;
            set { _progressText = value; OnPropertyChanged(); }
        }

        public int TotalFilesFound { get; private set; }
        public int TotalRowsInserted { get; private set; }

        private bool _isBusy = false;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                if (_isBusy != value)
                {
                    _isBusy = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsNotBusy));
                    OnPropertyChanged(nameof(IsBusyVisibility));

                    if (_isBusy) StartLogTimer();
                    else StopLogTimer();
                }
            }
        }

        public bool IsNotBusy => !IsBusy;

        public System.Windows.Visibility IsBusyVisibility =>
            IsBusy ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

        private string _busyMessage = "Processing...";
        public string BusyMessage
        {
            get => _busyMessage;
            set { _busyMessage = value; OnPropertyChanged(); }
        }

        private bool _showBlankPage = false;
        public bool ShowBlankPage
        {
            get => _showBlankPage;
            set { _showBlankPage = value; OnPropertyChanged(); }
        }

        private string _notificationMessage = string.Empty;
        public string NotificationMessage
        {
            get => _notificationMessage;
            set { _notificationMessage = value; OnPropertyChanged(); }
        }

        private bool _isNotificationVisible = false;
        public bool IsNotificationVisible
        {
            get => _isNotificationVisible;
            set
            {
                _isNotificationVisible = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NotificationVisibility));
            }
        }

        public System.Windows.Visibility NotificationVisibility =>
            IsNotificationVisible ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

        public System.Collections.ObjectModel.ObservableCollection<string> Years { get; set; } = new System.Collections.ObjectModel.ObservableCollection<string>();
        public System.Collections.ObjectModel.ObservableCollection<string> Months { get; set; } = new System.Collections.ObjectModel.ObservableCollection<string>();
        public System.Collections.ObjectModel.ObservableCollection<string> Standards { get; set; } = new System.Collections.ObjectModel.ObservableCollection<string>();

        public System.Collections.ObjectModel.ObservableCollection<string> TypeOptions { get; set; } = new System.Collections.ObjectModel.ObservableCollection<string>
        {
            "Select", "RD", "FR", "TP", "NONE"
        };

        private string? _selectedYear;
        public string? SelectedYear
        {
            get => _selectedYear;
            set { if (_selectedYear != value) { _selectedYear = value; OnPropertyChanged(); SetDefaultProductionDate(); } }
        }

        private string? _selectedMonth;
        public string? SelectedMonth
        {
            get => _selectedMonth;
            set { if (_selectedMonth != value) { _selectedMonth = value; OnPropertyChanged(); SetDefaultProductionDate(); } }
        }

        private System.DateTime? _selectedDate;
        public System.DateTime? SelectedDate { get => _selectedDate; set { _selectedDate = value; OnPropertyChanged(); } }

        private string? _selectedStandard;
        public string? SelectedStandard { get => _selectedStandard; set { _selectedStandard = value; OnPropertyChanged(); } }

        private string _customerName = string.Empty;
        public string CustomerName { get => _customerName; set { _customerName = value; OnPropertyChanged(); } }

        private string _poNumber = string.Empty;
        public string PoNumber { get => _poNumber; set { _poNumber = value; OnPropertyChanged(); } }

        private string _numberDO = string.Empty;
        public string DoNumber { get => _numberDO; set { _numberDO = value; OnPropertyChanged(); } }

        private System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem> _searchResults = new System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem>();
        public System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem> SearchResults { get => _searchResults; set { _searchResults = value; OnPropertyChanged(); } }

        public System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarExportItem> ExportList { get; set; }
            = new System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarExportItem>();

        #endregion

        #region Commands
        public System.Windows.Input.ICommand FindCommand { get; }
        public System.Windows.Input.ICommand AddToExportCommand { get; }
        public System.Windows.Input.ICommand RemoveFromExportCommand { get; }
        public System.Windows.Input.ICommand PrintCoaCommand { get; }
        #endregion

        public event System.Action<string>? OnShowMessage;

        public MainViewModel()
        {
            _dbContext = new SqliteContext();
            _repository = new BusbarRepository(_dbContext);
            _importService = new ExcelImportService(_repository);
            _printService = new CoaPrintService();

            _searchCache = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem>>();

            _logTimer = new System.Windows.Threading.DispatcherTimer();
            _logTimer.Interval = System.TimeSpan.FromMilliseconds(2000);
            _logTimer.Tick += LogTimer_Tick;

            _importService.OnDebugMessage += (msg) => {
                lock (_logLock)
                {
                    _logBuffer.AppendLine(msg);
                }
            };

            _importService.OnProgress += (current, total) => {
                System.Windows.Application.Current.Dispatcher.InvokeAsync(() => UpdateProgress(current, total));
            };

            InitializeSearchData();

            FindCommand = new RelayCommand(ExecuteFind);
            AddToExportCommand = new RelayCommand(ExecuteAddToExport);
            RemoveFromExportCommand = new RelayCommand(ExecuteRemoveFromExport);
            PrintCoaCommand = new RelayCommand(ExecutePrintCoa);
        }

        #region Logging & Progress

        private void LogTimer_Tick(object? sender, System.EventArgs e)
        {
            lock (_logLock)
            {
                string newLogs = _logBuffer.ToString();
                if (string.IsNullOrWhiteSpace(newLogs))
                {
                    return;
                }

                _logBuffer.Clear();

                if (DebugLog.Length > 3000)
                {
                    DebugLog = "...[Log truncated]..." + System.Environment.NewLine;
                }
                DebugLog += newLogs;
            }
        }

        private void StartLogTimer()
        {
            if (!_logTimer.IsEnabled) _logTimer.Start();
        }

        private void StopLogTimer()
        {
            _logTimer.Stop();
            LogTimer_Tick(null, System.EventArgs.Empty);
        }

        private void UpdateProgress(int current, int total)
        {
            ProgressValue = current;
            ProgressMaximum = total;
            ProgressText = total > 0 ? $"Processing file {current} of {total}" : "Scanning files...";
        }

        private async void TriggerSuccessNotification(string message)
        {
            NotificationMessage = message;
            IsNotificationVisible = true;
            await System.Threading.Tasks.Task.Delay(3000);
            IsNotificationVisible = false;
        }

        #endregion

        #region Import Logic

        public void ImportExcelToSQLite()
        {
            try
            {
                _dbContext.EnsureDatabaseFolderExists();
                ResetCounters();

                using var connection = _dbContext.CreateConnection();
                if (connection.State != System.Data.ConnectionState.Open) connection.Open();

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "PRAGMA synchronous = OFF; PRAGMA journal_mode = MEMORY; PRAGMA temp_store = MEMORY;";
                    command.ExecuteNonQuery();
                }

                using var transaction = connection.BeginTransaction();
                try
                {
                    _importService.Import(connection, transaction);
                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }

                TotalFilesFound = _importService.TotalFilesFound;
                TotalRowsInserted = _importService.TotalRowsInserted;
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        private void ResetCounters()
        {
            TotalFilesFound = 0;
            TotalRowsInserted = 0;
            DebugLog = "";
            _searchCache.Clear();
        }

        #endregion

        #region Navigation & Menu

        public void ButtonMode2_Click() { OnShowMessage?.Invoke("MODE 2 belum diimplementasikan"); }
        public void ButtonMode3_Click() { OnShowMessage?.Invoke("MODE 3 belum diimplementasikan"); }
        public void ButtonMode4_Click() { OnShowMessage?.Invoke("MODE 4 belum diimplementasikan"); }

        public void BackToMenu()
        {
            ResetSearchData();
            CustomerName = string.Empty;
            PoNumber = string.Empty;
            DoNumber = string.Empty;
            ExportList.Clear();
            ShowBlankPage = false;
            IsNotificationVisible = false;
            _searchCache.Clear();
        }

        #endregion

        #region Search Logic

        private void InitializeSearchData()
        {
            Months.Add("January"); Months.Add("February"); Months.Add("March");
            Months.Add("April"); Months.Add("May"); Months.Add("June");
            Months.Add("July"); Months.Add("August"); Months.Add("September");
            Months.Add("October"); Months.Add("November"); Months.Add("December");

            Standards.Add("JIS");

            SearchResults.Clear();
            _ = LoadAvailableYears();
        }

        private async System.Threading.Tasks.Task LoadAvailableYears()
        {
            try
            {
                var dbYears = await System.Threading.Tasks.Task.Run(() => _repository.GetAvailableYears());

                Years.Clear();
                foreach (var year in dbYears) Years.Add(year);
            }
            catch (System.Exception ex)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() => { System.Windows.MessageBox.Show($"Error loading years: {ex.Message}"); });
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
                    if (month > 0 && month <= 12) SelectedDate = new System.DateTime(year, month, 1);
                }
            }
        }

        private async void ExecuteFind(object? parameter)
        {
            if (string.IsNullOrWhiteSpace(SelectedYear)) { OnShowMessage?.Invoke("Harap memilih YEAR."); return; }
            if (string.IsNullOrWhiteSpace(SelectedMonth)) { OnShowMessage?.Invoke("Harap memilih MONTH."); return; }
            if (SelectedDate == null) { OnShowMessage?.Invoke("Harap memilih PRODUCTION DATE."); return; }

            try
            {
                string dbMonth = ConvertMonthToEnglish(SelectedMonth);

                string dateSql = SelectedDate.Value.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                string dateIndo = SelectedDate.Value.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                string cacheKey = $"{SelectedYear}_{dbMonth}_{dateSql}";

                if (_searchCache.ContainsKey(cacheKey))
                {
                    var cachedList = _searchCache[cacheKey];
                    if (cachedList != null && cachedList.Count > 0)
                    {
                        SearchResults = new System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem>(cachedList);
                        return;
                    }
                    else
                    {
                        _searchCache.Remove(cacheKey);
                    }
                }

                var data = await System.Threading.Tasks.Task.Run(() => _repository.SearchBusbarRecords(SelectedYear, dbMonth, dateSql));

                if (data == null || !data.Any())
                {
                    data = await System.Threading.Tasks.Task.Run(() => _repository.SearchBusbarRecords(SelectedYear, dbMonth, dateIndo));
                }

                var newResults = new System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem>();
                var listForCache = new System.Collections.Generic.List<BusbarSearchItem>();

                if (data != null)
                {
                    foreach (var item in data)
                    {
                        newResults.Add(item);
                        listForCache.Add(item);
                    }
                }

                if (listForCache.Count > 0)
                {
                    _searchCache[cacheKey] = listForCache;
                }

                SearchResults = newResults;

                if (SearchResults.Count == 0)
                {
                    OnShowMessage?.Invoke("Data tidak ditemukan.");
                }
            }
            catch (System.Exception ex)
            {
                OnShowMessage?.Invoke($"Terjadi kesalahan saat pencarian: {ex.Message}");
            }
        }

        private void ResetSearchData()
        {
            SelectedYear = null; SelectedMonth = null; SelectedDate = null; SelectedStandard = null; SearchResults.Clear();
        }

        #endregion

        #region Export & Print Logic

        private void ExecuteAddToExport(object? parameter)
        {
            if (parameter is BusbarSearchItem selectedItem)
            {
                bool exists = ExportList.Any(x => x.RecordData.Id == selectedItem.FullRecord.Id);
                if (!exists) { var exportItem = new WpfApp1.Core.Models.BusbarExportItem(selectedItem.FullRecord); ExportList.Add(exportItem); }
                else { OnShowMessage?.Invoke("Data ini sudah ada dalam daftar Export."); }
            }
        }

        private void ExecuteRemoveFromExport(object? parameter)
        {
            if (parameter is WpfApp1.Core.Models.BusbarExportItem itemToRemove) ExportList.Remove(itemToRemove);
        }

        private string FormatDoNumber(string doNumber)
        {
            if (int.TryParse(doNumber, out int doVal)) { if (doVal < 10) return "0" + doVal.ToString(); }
            return doNumber;
        }

        private async void ExecutePrintCoa(object? parameter)
        {
            bool hasUnselectedType = false;
            foreach (var item in ExportList)
            {
                if (item.SelectedType == "Select")
                {
                    hasUnselectedType = true;
                    break;
                }
            }

            if (hasUnselectedType)
            {
                System.Windows.MessageBox.Show("Mohon lengkapi kolom 'Type' untuk semua data sebelum melakukan Export COA.", "Validasi Data", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            if (ExportList.Count == 0)
            {
                System.Windows.MessageBox.Show("Export List kosong. Silakan pilih data terlebih dahulu.", "Peringatan", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(SelectedStandard) || string.IsNullOrWhiteSpace(CustomerName) ||
                string.IsNullOrWhiteSpace(PoNumber) || string.IsNullOrWhiteSpace(DoNumber))
            {
                System.Windows.MessageBox.Show("Silakan lengkapi data (Standard, Customer, PO, DO).", "Peringatan", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            DoNumber = FormatDoNumber(DoNumber);

            if (IsBusy) return;

            BusyMessage = "Generating COA Document...";
            IsBusy = true;

            try
            {
                var itemsToExport = new System.Collections.Generic.List<WpfApp1.Core.Models.BusbarExportItem>(ExportList);
                string custName = CustomerName;
                string po = PoNumber;
                string doNum = DoNumber;
                string std = SelectedStandard ?? string.Empty;

                string savedExcelPath = await _printService.GenerateCoaExcel(custName, po, doNum, itemsToExport, std);

                CustomerName = string.Empty;
                PoNumber = string.Empty;
                DoNumber = string.Empty;
                SelectedStandard = null;
                ExportList.Clear();

                TriggerSuccessNotification("COA Generated Successfully!");
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"Gagal membuat Dokumen:\n{ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                _printService.ClearCache();
            }
        }

        #endregion

        #region Helpers

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

        #endregion
    }

    public class RelayCommand : System.Windows.Input.ICommand
    {
        private readonly System.Action<object?> _execute;
        private readonly System.Func<object?, bool>? _canExecute;
        public RelayCommand(System.Action<object?> execute, System.Func<object?, bool>? canExecute = null) { _execute = execute; _canExecute = canExecute; }
        public event System.EventHandler? CanExecuteChanged { add { System.Windows.Input.CommandManager.RequerySuggested += value; } remove { System.Windows.Input.CommandManager.RequerySuggested -= value; } }
        public bool CanExecute(object? parameter) => _canExecute == null || _canExecute(parameter);
        public void Execute(object? parameter) => _execute(parameter);
    }
}