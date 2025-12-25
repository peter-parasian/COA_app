using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        private const string ExcelRootFolder = @"C:\Users\mrrx\Documents\My Web Sites\H\OPERATOR\COPPER BUSBAR & STRIP";
        private const string DbPath = @"C:\sqLite\data_qc.db";

        private int _totalFilesFound;
        private int _totalRowsInserted;
        private string _debugLog;

        public MainWindow()
        {
            InitializeComponent();
            ImportExcelToSQLite();
        }

        private void ImportExcelToSQLite()
        {
            try
            {
                EnsureDatabaseFolderExists();

                using var connection = new SqliteConnection($"Data Source={DbPath}");
                connection.Open();

                CreateBusbarTable(connection);

                using var transaction = connection.BeginTransaction();

                TraverseFoldersAndImport(connection, transaction);

                transaction.Commit();

                ShowFinalReport();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"ERROR FATAL:\n{ex.Message}\n\n{ex.StackTrace}",
                    "Import Gagal",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void EnsureDatabaseFolderExists()
        {
            string folder = System.IO.Path.GetDirectoryName(DbPath);

            if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }

        private void CreateBusbarTable(SqliteConnection connection)
        {
            using var cmd = connection.CreateCommand();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS Busbar;

            CREATE TABLE IF NOT EXISTS Busbar (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Batch_no TEXT,
                Prod_date TEXT,
                Size_mm TEXT,
                Thickness_mm REAL,
                Width_mm REAL,
                Radius REAL,
                Length REAL,
                Chamber_mm REAL,
                Status TEXT,
                Year_folder TEXT,
                Month_folder TEXT,
                Oxygen_analizer REAL,
                Spectro_Cu TEXT,
                Electric_iacs REAL,
                Weight_resistivity REAL,
                Elongation_pct REAL,
                Bend_test TEXT, 
                Tensile_strength REAL
            );
            ";

            cmd.ExecuteNonQuery();
        }

        private void TraverseFoldersAndImport(SqliteConnection connection, SqliteTransaction transaction)
        {
            ResetCounters();

            if (!Directory.Exists(ExcelRootFolder))
            {
                throw new DirectoryNotFoundException("Folder root Excel tidak ditemukan.");
            }

            foreach (var yearDir in Directory.GetDirectories(ExcelRootFolder))
            {
                string year = new DirectoryInfo(yearDir).Name;

                foreach (var monthDir in Directory.GetDirectories(yearDir))
                {
                    string month = new DirectoryInfo(monthDir).Name;

                    foreach (var file in Directory.GetFiles(monthDir, "*.xlsx"))
                    {
                        if (System.IO.Path.GetFileName(file).StartsWith("~$"))
                            continue;

                        _totalFilesFound++;

                        ProcessSingleExcelFile(
                            connection,
                            transaction,
                            file,
                            year,
                            month);
                    }
                }
            }
        }

        private void ProcessSingleExcelFile(
            SqliteConnection connection,
            SqliteTransaction transaction,
            string filePath,
            string year,
            string month)
        {
            try
            {
                using var workbook = new XLWorkbook(filePath);

                var sheet = workbook.Worksheets
                    .FirstOrDefault(w =>
                        w.Name.Trim()
                        .Equals("YLB 50", StringComparison.OrdinalIgnoreCase));

                if (sheet == null)
                {
                    AppendDebug($"SKIP: Sheet YLB 50 tidak ada -> {System.IO.Path.GetFileName(filePath)}");
                    return;
                }

                int row = 3;

                while (true)
                {
                    string sizeValue = sheet.Cell(row, 3).GetValue<string>();

                    if (string.IsNullOrWhiteSpace(sizeValue))
                        break;

                    InsertBusbarRow(connection, transaction, sizeValue, year, month);

                    _totalRowsInserted++;
                    row += 2;
                }
            }
            catch (Exception ex)
            {
                AppendDebug($"ERROR FILE: {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }
        }

        private void InsertBusbarRow(
            SqliteConnection connection,
            SqliteTransaction transaction,
            string size,
            string year,
            string month)
        {
            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            cmd.CommandText = @"
                INSERT INTO Busbar (Size_mm, Year_folder, Month_folder)
                VALUES (@Size, @Year, @Month);
            ";

            cmd.Parameters.AddWithValue("@Size", size);
            cmd.Parameters.AddWithValue("@Year", year);
            cmd.Parameters.AddWithValue("@Month", month);

            cmd.ExecuteNonQuery();
        }

        private void ResetCounters()
        {
            _totalFilesFound = 0;
            _totalRowsInserted = 0;
            _debugLog = "";
        }

        private void AppendDebug(string message)
        {
            if (_debugLog.Length < 800)
            {
                _debugLog += message + Environment.NewLine;
            }
        }

        private void ShowFinalReport()
        {
            MessageBox.Show(
                $"IMPORT SELESAI\n\n" +
                $"File ditemukan : {_totalFilesFound}\n" +
                $"Baris disimpan : {_totalRowsInserted}\n\n" +
                $"Debug (awal):\n{_debugLog}",
                "Laporan Import",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}