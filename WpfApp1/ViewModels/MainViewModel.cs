using System;
using System.Windows;
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

        public event Action<string> OnShowMessage;

        public MainViewModel()
        {
            _dbContext = new SqliteContext();
            _repository = new BusbarRepository();
            _importService = new ExcelImportService(_repository);

            _importService.OnDebugMessage += (msg) => {
                lock (_lockObject)
                {
                    if (DebugLog.Length > 5000) DebugLog = string.Empty;
                    DebugLog += msg + System.Environment.NewLine;
                }
            };
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

                OnShowMessage?.Invoke($"IMPORT SELESAI\n\nFile ditemukan : {TotalFilesFound}\nBaris disimpan : {TotalRowsInserted}\n\nDebug Log:\n{DebugLog}");
            }
            catch (System.Exception ex)
            {
                // Bubble up exception agar View yang menangani error message
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
            ShowBlankPage = false;
        }

        private void ResetCounters()
        {
            TotalFilesFound = 0;
            TotalRowsInserted = 0;
            DebugLog = "";
        }
    }
}