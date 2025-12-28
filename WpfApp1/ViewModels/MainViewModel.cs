using System;
using System.Collections.Generic;
using System.Text;
using WpfApp1.Core.Services;
using WpfApp1.Data.Database;
using WpfApp1.Data.Repositories;

namespace WpfApp1.ViewModels
{
    public class MainViewModel
    {
        private SqliteContext _dbContext;
        private BusbarRepository _repository;
        private ExcelImportService _importService;

        private int _totalFilesFound;
        private int _totalRowsInserted;
        private string _debugLog;

        public MainViewModel()
        {
            _dbContext = new SqliteContext();
            _repository = new BusbarRepository();
            _importService = new ExcelImportService(_repository);

            // Subscribe to debug logs from service
            _importService.OnDebugMessage += (msg) => {
                lock (_debugLog)
                {
                    if (_debugLog.Length < 1000) _debugLog += msg + System.Environment.NewLine;
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

                _totalFilesFound = _importService.TotalFilesFound;
                _totalRowsInserted = _importService.TotalRowsInserted;

                ShowFinalReport();
            }
            catch (System.Exception ex)
            {
                throw; // Bubble up to View to show MessageBox
            }
        }

        public void ButtonMode2_Click()
        {
            System.Windows.MessageBox.Show(
                "MODE 2 belum diimplementasikan",
                "Info",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }

        public void ButtonMode3_Click()
        {
            System.Windows.MessageBox.Show(
                "MODE 3 belum diimplementasikan",
                "Info",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }

        public void ButtonMode4_Click()
        {
            System.Windows.MessageBox.Show(
                "MODE 4 belum diimplementasikan",
                "Info",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }

        private void ResetCounters()
        {
            _totalFilesFound = 0;
            _totalRowsInserted = 0;
            _debugLog = "";
        }

        private void ShowFinalReport()
        {
            System.Windows.MessageBox.Show(
                $"IMPORT SELESAI\n\nFile ditemukan : {_totalFilesFound}\nBaris disimpan : {_totalRowsInserted}\n\nDebug Log:\n{_debugLog}",
                "Laporan Import",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }
    }
}