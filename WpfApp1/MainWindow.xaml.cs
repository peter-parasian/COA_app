using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
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
                // 1. Pastikan folder database ada
                EnsureDatabaseFolderExists();

                // 2. Buka koneksi database
                using var connection = new SqliteConnection($"Data Source={DbPath}");
                connection.Open();

                // 3. Buat ulang tabel (struktur fresh)
                CreateBusbarTable(connection);

                // 4. Mulai transaksi untuk performa insert batch
                using var transaction = connection.BeginTransaction();

                // 5. Traverse folder dan proses file
                TraverseFoldersAndImport(connection, transaction);

                // 6. Commit perubahan ke database
                transaction.Commit();

                // 7. Tampilkan laporan akhir
                ShowFinalReport();
            }
            catch (Exception ex)
            {
                // Logging error yang jelas ke UI
                MessageBox.Show(
                    $"ERROR FATAL:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Import Gagal",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void EnsureDatabaseFolderExists()
        {
            string folder = System.IO.Path.GetDirectoryName(DbPath);

            if (!string.IsNullOrEmpty(folder) && !System.IO.Directory.Exists(folder))
            {
                System.IO.Directory.CreateDirectory(folder);
            }
        }

        private void CreateBusbarTable(SqliteConnection connection)
        {
            using var cmd = connection.CreateCommand();

            // Menggunakan multi-line string untuk keterbacaan SQL
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

            if (!System.IO.Directory.Exists(ExcelRootFolder))
            {
                throw new System.IO.DirectoryNotFoundException($"Folder root Excel tidak ditemukan: {ExcelRootFolder}");
            }

            // Loop Folder Tahun
            foreach (var yearDir in System.IO.Directory.GetDirectories(ExcelRootFolder))
            {
                string year = new System.IO.DirectoryInfo(yearDir).Name;

                // Loop Folder Bulan
                foreach (var monthDir in System.IO.Directory.GetDirectories(yearDir))
                {
                    string month = new System.IO.DirectoryInfo(monthDir).Name;

                    // Loop File Excel
                    foreach (var file in System.IO.Directory.GetFiles(monthDir, "*.xlsx"))
                    {
                        string fileName = System.IO.Path.GetFileName(file);

                        // Skip file temporary Excel (yang sedang dibuka)
                        if (fileName.StartsWith("~$"))
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
                // Menggunakan ClosedXML untuk membaca Excel tanpa instalasi Office Interop
                using var workbook = new XLWorkbook(filePath);

                // Cari sheet bernama "YLB 50" (case insensitive)
                var sheet = workbook.Worksheets
                    .FirstOrDefault(w =>
                        w.Name.Trim()
                        .Equals("YLB 50", StringComparison.OrdinalIgnoreCase));

                if (sheet == null)
                {
                    AppendDebug($"SKIP: Sheet 'YLB 50' tidak ditemukan -> {System.IO.Path.GetFileName(filePath)}");
                    return;
                }

                // Mulai baca dari baris 3 (sesuai logika kode asli)
                int row = 3;

                while (true)
                {
                    // Ambil nilai dari kolom ke-3 (Column C) untuk ukuran
                    // Gunakan GetString agar konsisten, cek null/empty setelahnya
                    string sizeValue = sheet.Cell(row, 3).GetString();

                    // Jika sel kosong, diasumsikan data selesai
                    if (string.IsNullOrWhiteSpace(sizeValue))
                        break;

                    // Simpan ke database
                    InsertBusbarRow(connection, transaction, sizeValue, year, month);

                    _totalRowsInserted++;

                    // Lompat 1 baris (sesuai logika kode asli: baris data selang seling)
                    row += 2;
                }
            }
            catch (Exception ex)
            {
                // Tangkap error per-file agar proses lain tidak terhenti
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

            // Parameterized Query untuk mencegah SQL Injection dan handle tipe data dengan aman
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
            // Batasi log agar memori tidak membengkak
            if (_debugLog.Length < 1000)
            {
                _debugLog += message + System.Environment.NewLine;
            }
        }

        private void ShowFinalReport()
        {
            MessageBox.Show(
                $"IMPORT SELESAI\n\n" +
                $"File ditemukan : {_totalFilesFound}\n" +
                $"Baris disimpan : {_totalRowsInserted}\n\n" +
                $"Debug Log:\n{_debugLog}",
                "Laporan Import",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}