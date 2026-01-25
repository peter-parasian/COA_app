namespace WpfApp1.ViewModels
{
    public class MainViewModel : WpfApp1.ViewModels.BaseViewModel
    {
        private WpfApp1.Data.Database.SqliteContext _dbContext;
        private WpfApp1.Data.Repositories.BusbarRepository _repository;
        private WpfApp1.Core.Services.ExcelImportService _importService;
        private WpfApp1.Core.Services.CoaPrintService _printService;

        private System.Collections.Generic.Dictionary<System.String, System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem>> _searchCache;

        private readonly System.Object _logLock = new System.Object();
        private System.Text.StringBuilder _logBuffer = new System.Text.StringBuilder();
        private System.Windows.Threading.DispatcherTimer _logTimer;

        #region Properties

        private System.String _debugLog = System.String.Empty;
        public System.String DebugLog
        {
            get => _debugLog;
            set { _debugLog = value; OnPropertyChanged(); }
        }

        private System.Int32 _progressValue = 0;
        public System.Int32 ProgressValue
        {
            get => _progressValue;
            set { _progressValue = value; OnPropertyChanged(); }
        }

        private System.Int32 _progressMaximum = 100;
        public System.Int32 ProgressMaximum
        {
            get => _progressMaximum;
            set { _progressMaximum = value; OnPropertyChanged(); }
        }

        private System.String _progressText = "Ready";
        public System.String ProgressText
        {
            get => _progressText;
            set { _progressText = value; OnPropertyChanged(); }
        }

        public System.Int32 TotalFilesFound { get; private set; }
        public System.Int32 TotalRowsInserted { get; private set; }

        private System.Boolean _isBusy = false;
        public System.Boolean IsBusy
        {
            get => _isBusy;
            set
            {
                if (_isBusy != value)
                {
                    _isBusy = value;
                    OnPropertyChanged();
                    OnPropertyChanged("IsNotBusy");
                    OnPropertyChanged("IsBusyVisibility");

                    if (_isBusy)
                    {
                        StartLogTimer();
                    }
                    else
                    {
                        StopLogTimer();
                    }
                }
            }
        }

        public System.Boolean IsNotBusy => !IsBusy;

        public System.Windows.Visibility IsBusyVisibility =>
            IsBusy ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

        private System.String _busyMessage = "Processing...";
        public System.String BusyMessage
        {
            get => _busyMessage;
            set { _busyMessage = value; OnPropertyChanged(); }
        }

        private System.Boolean _showBlankPage = false;
        public System.Boolean ShowBlankPage
        {
            get => _showBlankPage;
            set { _showBlankPage = value; OnPropertyChanged(); }
        }

        private System.String _notificationMessage = System.String.Empty;
        public System.String NotificationMessage
        {
            get => _notificationMessage;
            set { _notificationMessage = value; OnPropertyChanged(); }
        }

        private System.Boolean _isNotificationVisible = false;
        public System.Boolean IsNotificationVisible
        {
            get => _isNotificationVisible;
            set
            {
                _isNotificationVisible = value;
                OnPropertyChanged();
                OnPropertyChanged("NotificationVisibility");
            }
        }

        public System.Windows.Visibility NotificationVisibility =>
            IsNotificationVisible ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

        public System.Collections.ObjectModel.ObservableCollection<System.String> Years { get; set; } =
            new System.Collections.ObjectModel.ObservableCollection<System.String>();
        public System.Collections.ObjectModel.ObservableCollection<System.String> Months { get; set; } =
            new System.Collections.ObjectModel.ObservableCollection<System.String>();
        public System.Collections.ObjectModel.ObservableCollection<System.String> Standards { get; set; } =
            new System.Collections.ObjectModel.ObservableCollection<System.String>();

        public System.Collections.ObjectModel.ObservableCollection<System.String> TypeOptions { get; set; } =
            new System.Collections.ObjectModel.ObservableCollection<System.String>
        {
            "Select", "RD", "FR", "TP", "NONE"
        };

        private System.String? _selectedYear;
        public System.String? SelectedYear
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

        private System.String? _selectedMonth;
        public System.String? SelectedMonth
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
        public System.DateTime? SelectedDate { get => _selectedDate; set { _selectedDate = value; OnPropertyChanged(); } }

        private System.String? _selectedStandard;
        public System.String? SelectedStandard { get => _selectedStandard; set { _selectedStandard = value; OnPropertyChanged(); } }

        private System.String _customerName = System.String.Empty;
        public System.String CustomerName { get => _customerName; set { _customerName = value; OnPropertyChanged(); } }

        private System.String _poNumber = System.String.Empty;
        public System.String PoNumber { get => _poNumber; set { _poNumber = value; OnPropertyChanged(); } }

        private System.String _numberDO = System.String.Empty;
        public System.String DoNumber { get => _numberDO; set { _numberDO = value; OnPropertyChanged(); } }

        private System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarSearchItem> _searchResults =
            new System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarSearchItem>();
        public System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarSearchItem> SearchResults { get => _searchResults; set { _searchResults = value; OnPropertyChanged(); } }

        public System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarExportItem> ExportList { get; set; }
            = new System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarExportItem>();

        #endregion

        #region Commands
        public System.Windows.Input.ICommand FindCommand { get; }
        public System.Windows.Input.ICommand AddToExportCommand { get; }
        public System.Windows.Input.ICommand RemoveFromExportCommand { get; }
        public System.Windows.Input.ICommand PrintCoaCommand { get; }
        #endregion

        public event System.Action<System.String>? OnShowMessage;

        public MainViewModel()
        {
            _dbContext = new WpfApp1.Data.Database.SqliteContext();
            _repository = new WpfApp1.Data.Repositories.BusbarRepository(_dbContext);
            _importService = new WpfApp1.Core.Services.ExcelImportService(_repository);
            _printService = new WpfApp1.Core.Services.CoaPrintService();

            _searchCache = new System.Collections.Generic.Dictionary<System.String, System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem>>();

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

        private void LogTimer_Tick(System.Object? sender, System.EventArgs e)
        {
            lock (_logLock)
            {
                System.String newLogs = _logBuffer.ToString();
                if (System.String.IsNullOrWhiteSpace(newLogs))
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
            if (!_logTimer.IsEnabled)
            {
                _logTimer.Start();
            }
        }

        private void StopLogTimer()
        {
            _logTimer.Stop();
            LogTimer_Tick(null, System.EventArgs.Empty);
        }

        private void UpdateProgress(System.Int32 current, System.Int32 total)
        {
            ProgressValue = current;
            ProgressMaximum = total;
            ProgressText = total > 0 ? $"Processing file {current} of {total}" : "Scanning files...";
        }

        private async void TriggerSuccessNotification(System.String message)
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

                Microsoft.Data.Sqlite.SqliteConnection? connection = null;
                Microsoft.Data.Sqlite.SqliteTransaction? transaction = null;

                try
                {
                    connection = _dbContext.CreateConnection();

                    if (connection.State != System.Data.ConnectionState.Open)
                    {
                        connection.Open();
                    }

                    Microsoft.Data.Sqlite.SqliteCommand? pragmaCommand = null;
                    try
                    {
                        pragmaCommand = connection.CreateCommand();
                        pragmaCommand.CommandText = "PRAGMA synchronous = OFF; PRAGMA journal_mode = MEMORY; PRAGMA temp_store = MEMORY;";
                        pragmaCommand.ExecuteNonQuery();
                    }
                    finally
                    {
                        if (pragmaCommand != null)
                        {
                            pragmaCommand.Dispose();
                        }
                    }

                    transaction = connection.BeginTransaction();
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
                }
                finally
                {
                    if (transaction != null)
                    {
                        transaction.Dispose();
                    }
                    if (connection != null)
                    {
                        connection.Dispose();
                    }
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
            CustomerName = System.String.Empty;
            PoNumber = System.String.Empty;
            DoNumber = System.String.Empty;
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
                System.Collections.Generic.List<System.String> dbYears =
                    await System.Threading.Tasks.Task.Run(() => _repository.GetAvailableYears());

                Years.Clear();
                foreach (System.String year in dbYears)
                {
                    Years.Add(year);
                }
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
            if (!System.String.IsNullOrWhiteSpace(SelectedYear) && !System.String.IsNullOrWhiteSpace(SelectedMonth))
            {
                if (System.Int32.TryParse(SelectedYear, out System.Int32 year))
                {
                    System.String engMonth = ConvertMonthToEnglish(SelectedMonth);
                    System.Int32 month = WpfApp1.Shared.Helpers.DateHelper.GetMonthNumber(engMonth);
                    if (month > 0 && month <= 12)
                    {
                        SelectedDate = new System.DateTime(year, month, 1);
                    }
                }
            }
        }

        private async void ExecuteFind(System.Object? parameter)
        {
            if (System.String.IsNullOrWhiteSpace(SelectedYear))
            {
                OnShowMessage?.Invoke("Harap memilih YEAR.");
                return;
            }
            if (System.String.IsNullOrWhiteSpace(SelectedMonth))
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
                System.String dbMonth = ConvertMonthToEnglish(SelectedMonth);

                System.String dateSql = SelectedDate.Value.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                System.String dateIndo = SelectedDate.Value.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                System.String cacheKey = $"{SelectedYear}_{dbMonth}_{dateSql}";

                if (_searchCache.ContainsKey(cacheKey))
                {
                    System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem>? cachedList = _searchCache[cacheKey];
                    if (cachedList != null && cachedList.Count > 0)
                    {
                        SearchResults = new System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarSearchItem>(cachedList);
                        return;
                    }
                    else
                    {
                        _searchCache.Remove(cacheKey);
                    }
                }

                System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem> data =
                    await System.Threading.Tasks.Task.Run(() => _repository.SearchBusbarRecords(SelectedYear, dbMonth, dateSql));

                if (data == null || data.Count == 0)
                {
                    data = await System.Threading.Tasks.Task.Run(() => _repository.SearchBusbarRecords(SelectedYear, dbMonth, dateIndo));
                }

                System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarSearchItem> newResults =
                    new System.Collections.ObjectModel.ObservableCollection<WpfApp1.Core.Models.BusbarSearchItem>();
                System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem> listForCache =
                    new System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem>();

                if (data != null)
                {
                    foreach (WpfApp1.Core.Models.BusbarSearchItem item in data)
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

        private void ExecuteAddToExport(System.Object? parameter)
        {
            if (parameter is WpfApp1.Core.Models.BusbarSearchItem selectedItem)
            {
                System.Boolean exists = false;
                foreach (WpfApp1.Core.Models.BusbarExportItem x in ExportList)
                {
                    if (x.RecordData.Id == selectedItem.FullRecord.Id)
                    {
                        exists = true;
                        break;
                    }
                }

                if (!exists)
                {
                    WpfApp1.Core.Models.BusbarExportItem exportItem = new WpfApp1.Core.Models.BusbarExportItem(selectedItem.FullRecord);
                    ExportList.Add(exportItem);
                }
                else
                {
                    OnShowMessage?.Invoke("Data ini sudah ada dalam daftar Export.");
                }
            }
        }

        private void ExecuteRemoveFromExport(System.Object? parameter)
        {
            if (parameter is WpfApp1.Core.Models.BusbarExportItem itemToRemove)
            {
                ExportList.Remove(itemToRemove);
            }
        }

        private System.String FormatDoNumber(System.String doNumber)
        {
            if (System.Int32.TryParse(doNumber, out System.Int32 doVal))
            {
                if (doVal < 10)
                {
                    return "0" + doVal.ToString();
                }
            }
            return doNumber;
        }

        private async void ExecutePrintCoa(System.Object? parameter)
        {
            System.Boolean hasUnselectedType = false;
            foreach (WpfApp1.Core.Models.BusbarExportItem item in ExportList)
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

            if (System.String.IsNullOrWhiteSpace(SelectedStandard) || System.String.IsNullOrWhiteSpace(CustomerName) ||
                System.String.IsNullOrWhiteSpace(PoNumber) || System.String.IsNullOrWhiteSpace(DoNumber))
            {
                System.Windows.MessageBox.Show("Silakan lengkapi data (Standard, Customer, PO, DO).", "Peringatan", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            DoNumber = FormatDoNumber(DoNumber);

            if (IsBusy)
            {
                return;
            }

            BusyMessage = "Generating COA Document...";
            IsBusy = true;

            try
            {
                System.Collections.Generic.List<WpfApp1.Core.Models.BusbarExportItem> itemsToExport =
                    new System.Collections.Generic.List<WpfApp1.Core.Models.BusbarExportItem>(ExportList);

                System.String custName = CustomerName;
                System.String po = PoNumber;
                System.String doNum = DoNumber;
                System.String std = SelectedStandard ?? System.String.Empty;

                System.String savedExcelPath = await _printService.GenerateCoaExcel(custName, po, doNum, itemsToExport, std);

                CustomerName = System.String.Empty;
                PoNumber = System.String.Empty;
                DoNumber = System.String.Empty;
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

        private System.String ConvertMonthToEnglish(System.String indoMonth)
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
        private readonly System.Action<System.Object?> _execute;
        private readonly System.Func<System.Object?, System.Boolean>? _canExecute;

        public RelayCommand(System.Action<System.Object?> execute, System.Func<System.Object?, System.Boolean>? canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public event System.EventHandler? CanExecuteChanged
        {
            add { System.Windows.Input.CommandManager.RequerySuggested += value; }
            remove { System.Windows.Input.CommandManager.RequerySuggested -= value; }
        }

        public System.Boolean CanExecute(System.Object? parameter) => _canExecute == null || _canExecute(parameter);
        public void Execute(System.Object? parameter) => _execute(parameter);
    }
}