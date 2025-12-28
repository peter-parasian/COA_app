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

        private System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem> _searchResults = new System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem>();
        public System.Collections.ObjectModel.ObservableCollection<BusbarSearchItem> SearchResults
        {
            get => _searchResults;
            set { _searchResults = value; OnPropertyChanged(); }
        }

        public System.Windows.Input.ICommand FindCommand { get; }

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

        private void ExecuteFind()
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
        private readonly System.Action _execute;
        private readonly System.Func<bool>? _canExecute;

        public RelayCommand(System.Action execute, System.Func<bool>? canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public event System.EventHandler? CanExecuteChanged
        {
            add { System.Windows.Input.CommandManager.RequerySuggested += value; }
            remove { System.Windows.Input.CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object? parameter) => _canExecute == null || _canExecute();

        public void Execute(object? parameter) => _execute();
    }
}