using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Data.Sqlite;
using System.Windows;

namespace WpfApp1
{
    public partial class MainWindow : System.Windows.Window
    {
        private const string ExcelRootFolder = @"C:\Users\mrrx\Documents\My Web Sites\H\OPERATOR\COPPER BUSBAR & STRIP";
        private const string DbPath = @"C:\sqLite\data_qc.db";

        private int _totalFilesFound;
        private int _totalRowsInserted;
        private string _debugLog;

        // --- BATCHING BUFFERS ---
        private System.Collections.Generic.List<BusbarRecord> _busbarBatchBuffer = new System.Collections.Generic.List<BusbarRecord>();
        private System.Collections.Generic.List<TLJRecord> _tlj350BatchBuffer = new System.Collections.Generic.List<TLJRecord>();
        private System.Collections.Generic.List<TLJRecord> _tlj500BatchBuffer = new System.Collections.Generic.List<TLJRecord>();
        private const int BATCH_SIZE = 500; // Ukuran batch insert

        // Struct sederhana untuk menampung data sementara (Human-Centric DTO)
        private struct BusbarRecord
        {
            public string Size, Year, Month, ProdDate, BendTest;
            public double Thickness, Width, Length, Radius, Chamber, Electric, Resistivity, Elongation, Tensile, Spectro, Oxygen;
        }

        private struct TLJRecord
        {
            public string Size, Year, Month, ProdDate, BatchNo;
            public System.DateTime ParsedDate; // Helper untuk sorting di memori
        }

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

                // Bersihkan buffer sebelum mulai
                _busbarBatchBuffer.Clear();
                _tlj350BatchBuffer.Clear();
                _tlj500BatchBuffer.Clear();

                using var connection = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={DbPath}");
                connection.Open();

                // Konfigurasi performa SQLite
                using (var pragmaCmd = connection.CreateCommand())
                {
                    pragmaCmd.CommandText = "PRAGMA synchronous = OFF; PRAGMA journal_mode = MEMORY;";
                    pragmaCmd.ExecuteNonQuery();
                }

                CreateBusbarTable(connection);

                using var transaction = connection.BeginTransaction();

                TraverseFoldersAndImport(connection, transaction);

                // Flush sisa data yang belum mencapai limit batch
                FlushBusbarBatch(connection, transaction);
                FlushTLJBatch(connection, transaction, "TLJ350", _tlj350BatchBuffer);
                FlushTLJBatch(connection, transaction, "TLJ500", _tlj500BatchBuffer);

                UpdateBusbarBatchNumbers(connection, transaction);

                transaction.Commit();

                ShowFinalReport();
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"ERROR FATAL:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Import Gagal",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
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

        private void CreateBusbarTable(Microsoft.Data.Sqlite.SqliteConnection connection)
        {
            using var cmd = connection.CreateCommand();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS Busbar;
                CREATE TABLE IF NOT EXISTS Busbar (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT, Month_folder TEXT, Batch_no TEXT, Prod_date TEXT, Size_mm TEXT,
                    Thickness_mm REAL, Width_mm REAL, Length INTEGER, Radius REAL, Chamber_mm REAL,
                    Electric_IACS REAL, Weight REAL, Elongation REAL, Tensile REAL, Bend_test TEXT,
                    Spectro_Cu REAL, Oxygen REAL
                );
                -- OPTIMASI INDEX 
                CREATE INDEX IF NOT EXISTS IDX_Busbar_LookUp ON Busbar(Size_mm, Prod_date);
            ";
            cmd.ExecuteNonQuery();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS TLJ500;
                CREATE TABLE IF NOT EXISTS TLJ500 (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT, Month_folder TEXT, Batch_no TEXT, Prod_date TEXT, Size_mm TEXT
                );
                -- OPTIMASI INDEX
                CREATE INDEX IF NOT EXISTS IDX_TLJ500_LookUp ON TLJ500(Size_mm, Prod_date);
            ";
            cmd.ExecuteNonQuery();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS TLJ350;
                CREATE TABLE IF NOT EXISTS TLJ350 (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT, Month_folder TEXT, Batch_no TEXT, Prod_date TEXT, Size_mm TEXT
                );
                -- OPTIMASI INDEX
                CREATE INDEX IF NOT EXISTS IDX_TLJ350_LookUp ON TLJ350(Size_mm, Prod_date);
            ";
            cmd.ExecuteNonQuery();
        }

        private void TraverseFoldersAndImport(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            ResetCounters();

            if (!System.IO.Directory.Exists(ExcelRootFolder))
            {
                throw new System.IO.DirectoryNotFoundException($"Folder root Excel tidak ditemukan: {ExcelRootFolder}");
            }

            foreach (string yearDir in System.IO.Directory.GetDirectories(ExcelRootFolder))
            {
                string year = new System.IO.DirectoryInfo(yearDir).Name.Trim();

                foreach (string monthDir in System.IO.Directory.GetDirectories(yearDir))
                {
                    string rawMonth = new System.IO.DirectoryInfo(monthDir).Name.Trim();
                    string normalizedMonth = NormalizeMonthFolder(rawMonth);

                    foreach (string file in System.IO.Directory.GetFiles(monthDir, "*.xlsx"))
                    {
                        string fileName = System.IO.Path.GetFileName(file);

                        if (fileName.StartsWith("~$"))
                            continue;

                        _totalFilesFound++;
                        ProcessSingleExcelFile(connection, transaction, file, year, normalizedMonth);
                    }
                }
            }
        }

        private void ProcessSingleExcelFile(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string filePath,
            string year,
            string month)
        {
            // Membuka file tetap per-file karena logic excel closedXML butuh stream
            using var workbook = new ClosedXML.Excel.XLWorkbook(filePath);
            int row = 3;

            // --- Process YLB Sheet ---
            try
            {
                var sheet_YLB = workbook.Worksheets
                    .FirstOrDefault(w => w.Name.Trim().Equals("YLB 50", System.StringComparison.OrdinalIgnoreCase));

                if (sheet_YLB == null)
                {
                    AppendDebug($"SKIP: Sheet 'YLB 50' tidak ditemukan -> {System.IO.Path.GetFileName(filePath)}");
                }
                else
                {
                    string currentProdDate = string.Empty;
                    int folderMonthNum = GetMonthNumber(month);
                    int.TryParse(year, out int folderYearNum);

                    while (true)
                    {
                        string sizeValue_YLB = sheet_YLB.Cell(row, "C").GetString();
                        if (string.IsNullOrWhiteSpace(sizeValue_YLB)) break;

                        string rawDateFromCell = sheet_YLB.Cell(row, "B").GetString().Trim();
                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        // Parse Data
                        BusbarRecord record = new BusbarRecord();
                        record.Size = CleanSizeText(sizeValue_YLB);
                        record.Year = year;
                        record.Month = month;
                        record.ProdDate = currentProdDate;

                        record.Thickness = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "G").GetString()), 2);
                        record.Width = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "I").GetString()), 2);
                        record.Radius = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "J").GetString()), 2);
                        record.Chamber = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "L").GetString()), 2);
                        record.Electric = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "U").GetString()), 2);
                        record.Oxygen = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "X").GetString()), 2);

                        record.Spectro = ParseCustomDecimal(sheet_YLB.Cell(row, "Y").GetString());
                        record.Resistivity = ParseCustomDecimal(sheet_YLB.Cell(row, "T").GetString());

                        record.Length = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "K").GetString()), 0);

                        record.Elongation = System.Math.Round(GetMergedOrAverageValue(sheet_YLB, row, "R"), 2);
                        record.Tensile = System.Math.Round(GetMergedOrAverageValue(sheet_YLB, row, "Q"), 2);

                        record.BendTest = sheet_YLB.Cell(row, "W").GetString();

                        InsertBusbarRow(connection, transaction, record);
                        _totalRowsInserted++;
                        row += 2;
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (YLB): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }

            // --- Process TLJ 350 Sheet ---
            row = 3;
            try
            {
                var sheet_TLJ350 = workbook.Worksheets.FirstOrDefault(w => w.Name.Trim().Equals("TLJ 350", System.StringComparison.OrdinalIgnoreCase));
                if (sheet_TLJ350 != null)
                {
                    string currentProdDate = string.Empty;
                    int folderMonthNum = GetMonthNumber(month);
                    int.TryParse(year, out int folderYearNum);

                    while (true)
                    {
                        string sizeValue = sheet_TLJ350.Cell(row, "D").GetString();
                        if (string.IsNullOrWhiteSpace(sizeValue)) break;

                        string rawDate = sheet_TLJ350.Cell(row, "B").GetString().Trim();
                        if (!string.IsNullOrEmpty(rawDate))
                        {
                            currentProdDate = StandardizeDate(rawDate, folderMonthNum, folderYearNum);
                        }

                        TLJRecord record = new TLJRecord
                        {
                            Size = CleanSizeText(sizeValue),
                            Year = year,
                            Month = month,
                            ProdDate = currentProdDate,
                            BatchNo = sheet_TLJ350.Cell(row, "C").GetString()
                        };

                        InsertTLJ350_Row(connection, transaction, record);
                        _totalRowsInserted++;
                        row += 2;
                    }
                }
            }
            catch (System.Exception ex) { AppendDebug($"ERROR FILE (TLJ350): {ex.Message}"); }

            // --- Process TLJ 500 Sheet ---
            row = 3;
            try
            {
                var sheet_TLJ500 = workbook.Worksheets.FirstOrDefault(w => w.Name.Trim().Equals("TLJ 500", System.StringComparison.OrdinalIgnoreCase));
                if (sheet_TLJ500 != null)
                {
                    string currentProdDate = string.Empty;
                    int folderMonthNum = GetMonthNumber(month);
                    int.TryParse(year, out int folderYearNum);

                    while (true)
                    {
                        string sizeValue = sheet_TLJ500.Cell(row, "D").GetString();
                        if (string.IsNullOrWhiteSpace(sizeValue)) break;

                        string rawDate = sheet_TLJ500.Cell(row, "B").GetString().Trim();
                        if (!string.IsNullOrEmpty(rawDate))
                        {
                            currentProdDate = StandardizeDate(rawDate, folderMonthNum, folderYearNum);
                        }

                        TLJRecord record = new TLJRecord
                        {
                            Size = CleanSizeText(sizeValue),
                            Year = year,
                            Month = month,
                            ProdDate = currentProdDate,
                            BatchNo = sheet_TLJ500.Cell(row, "C").GetString()
                        };

                        InsertTLJ500_Row(connection, transaction, record);
                        _totalRowsInserted++;
                        row += 2;
                    }
                }
            }
            catch (System.Exception ex) { AppendDebug($"ERROR FILE (TLJ500): {ex.Message}"); }
        }

        // --- NEW BATCH INSERT METHODS ---

        private void InsertBusbarRow(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            BusbarRecord record)
        {
            _busbarBatchBuffer.Add(record);
            if (_busbarBatchBuffer.Count >= BATCH_SIZE)
            {
                FlushBusbarBatch(connection, transaction);
            }
        }

        private void FlushBusbarBatch(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            if (_busbarBatchBuffer.Count == 0) return;

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            // Membangun perintah SQL Batch
            System.Text.StringBuilder sqlBuilder = new System.Text.StringBuilder();
            sqlBuilder.Append(@"INSERT INTO Busbar (
                Size_mm, Year_folder, Month_folder, Prod_date, 
                Thickness_mm, Width_mm, Length, Radius, Chamber_mm,
                Electric_IACS, Weight, Elongation, Tensile,
                Bend_test, Spectro_Cu, Oxygen
            ) VALUES ");

            for (int i = 0; i < _busbarBatchBuffer.Count; i++)
            {
                if (i > 0) sqlBuilder.Append(",");
                // Menggunakan parameter index @p0_1, @p0_2 dll untuk keamanan
                sqlBuilder.Append($"(@s{i}, @y{i}, @m{i}, @d{i}, @t{i}, @w{i}, @l{i}, @r{i}, @c{i}, @e{i}, 0, @el{i}, @tn{i}, @bt{i}, @sp{i}, @ox{i})");

                var item = _busbarBatchBuffer[i];
                cmd.Parameters.AddWithValue($"@s{i}", item.Size);
                cmd.Parameters.AddWithValue($"@y{i}", item.Year.Trim());
                cmd.Parameters.AddWithValue($"@m{i}", item.Month.Trim());
                cmd.Parameters.AddWithValue($"@d{i}", item.ProdDate);
                cmd.Parameters.AddWithValue($"@t{i}", item.Thickness);
                cmd.Parameters.AddWithValue($"@w{i}", item.Width);
                cmd.Parameters.AddWithValue($"@l{i}", item.Length);
                cmd.Parameters.AddWithValue($"@r{i}", item.Radius);
                cmd.Parameters.AddWithValue($"@c{i}", item.Chamber);
                cmd.Parameters.AddWithValue($"@e{i}", item.Electric);
                cmd.Parameters.AddWithValue($"@el{i}", item.Elongation);
                cmd.Parameters.AddWithValue($"@tn{i}", item.Tensile);
                cmd.Parameters.AddWithValue($"@bt{i}", string.IsNullOrEmpty(item.BendTest) ? (object)System.DBNull.Value : item.BendTest);
                cmd.Parameters.AddWithValue($"@sp{i}", item.Spectro);
                cmd.Parameters.AddWithValue($"@ox{i}", item.Oxygen);
            }

            cmd.CommandText = sqlBuilder.ToString();
            cmd.ExecuteNonQuery();
            _busbarBatchBuffer.Clear();
        }

        private void InsertTLJ350_Row(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            TLJRecord record)
        {
            _tlj350BatchBuffer.Add(record);
            if (_tlj350BatchBuffer.Count >= BATCH_SIZE)
            {
                FlushTLJBatch(connection, transaction, "TLJ350", _tlj350BatchBuffer);
            }
        }

        private void InsertTLJ500_Row(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            TLJRecord record)
        {
            _tlj500BatchBuffer.Add(record);
            if (_tlj500BatchBuffer.Count >= BATCH_SIZE)
            {
                FlushTLJBatch(connection, transaction, "TLJ500", _tlj500BatchBuffer);
            }
        }

        private void FlushTLJBatch(
           Microsoft.Data.Sqlite.SqliteConnection connection,
           Microsoft.Data.Sqlite.SqliteTransaction transaction,
           string tableName,
           System.Collections.Generic.List<TLJRecord> buffer)
        {
            if (buffer.Count == 0) return;

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            System.Text.StringBuilder sqlBuilder = new System.Text.StringBuilder();
            sqlBuilder.Append($"INSERT INTO {tableName} (Size_mm, Year_folder, Month_folder, Prod_date, Batch_no) VALUES ");

            for (int i = 0; i < buffer.Count; i++)
            {
                if (i > 0) sqlBuilder.Append(",");
                sqlBuilder.Append($"(@s{i}, @y{i}, @m{i}, @d{i}, @b{i})");

                var item = buffer[i];
                cmd.Parameters.AddWithValue($"@s{i}", item.Size);
                cmd.Parameters.AddWithValue($"@y{i}", item.Year.Trim());
                cmd.Parameters.AddWithValue($"@m{i}", item.Month.Trim());
                cmd.Parameters.AddWithValue($"@d{i}", item.ProdDate);
                cmd.Parameters.AddWithValue($"@b{i}", string.IsNullOrEmpty(item.BatchNo) ? (object)System.DBNull.Value : item.BatchNo);
            }

            cmd.CommandText = sqlBuilder.ToString();
            cmd.ExecuteNonQuery();
            buffer.Clear();
        }

        // --- OPTIMIZED UPDATE LOGIC (IN-MEMORY LOOKUP) ---

        private void UpdateBusbarBatchNumbers(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            try
            {
                // 1. LOAD DATA REFERENSI KE MEMORY (Cache)
                var cache350 = LoadTLJCache(connection, transaction, "TLJ350");
                var cache500 = LoadTLJCache(connection, transaction, "TLJ500");

                // 2. LOAD BUSBAR YANG PERLU UPDATE
                using var selectBusbarCmd = connection.CreateCommand();
                selectBusbarCmd.Transaction = transaction;
                selectBusbarCmd.CommandText = @"
                    SELECT Id, Size_mm, Prod_date 
                    FROM Busbar 
                    WHERE (Batch_no IS NULL OR Batch_no = '')
                ";

                // Kita gunakan List untuk update agar Reader tidak konflik dengan Update Command
                var updates = new System.Collections.Generic.List<(int Id, string Batch)>();

                using (var reader = selectBusbarCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int id = reader.GetInt32(0);
                        string size = reader.GetString(1);
                        string dateStr = reader.GetString(2);

                        // Tentukan cache mana yang dipakai
                        string targetTable = DetermineTLJTable(size);
                        var targetCache = (targetTable == "TLJ350") ? cache350 : cache500;

                        // Lookup di Memory
                        string batchNo = FindBatchInMemory(targetCache, size, dateStr);

                        if (!string.IsNullOrEmpty(batchNo))
                        {
                            updates.Add((id, batchNo));
                        }
                    }
                }

                // 3. EKSEKUSI UPDATE DALAM BATCH (Transaction yang sama)
                if (updates.Count > 0)
                {
                    using var updateCmd = connection.CreateCommand();
                    updateCmd.Transaction = transaction;
                    updateCmd.CommandText = "UPDATE Busbar SET Batch_no = @b WHERE Id = @id";

                    var pBatch = updateCmd.CreateParameter(); pBatch.ParameterName = "@b";
                    var pId = updateCmd.CreateParameter(); pId.ParameterName = "@id";
                    updateCmd.Parameters.Add(pBatch);
                    updateCmd.Parameters.Add(pId);

                    foreach (var up in updates)
                    {
                        pBatch.Value = up.Batch;
                        pId.Value = up.Id;
                        updateCmd.ExecuteNonQuery();
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR UpdateBusbarBatchNumbers: {ex.Message}");
                throw;
            }
        }

        private System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<TLJRecord>> LoadTLJCache(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction trans,
            string tableName)
        {
            var cache = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<TLJRecord>>();

            using var cmd = connection.CreateCommand();
            cmd.Transaction = trans;
            cmd.CommandText = $"SELECT Size_mm, Prod_date, Batch_no FROM {tableName} WHERE Batch_no IS NOT NULL AND Batch_no != ''";

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string size = reader.GetString(0);
                string dateStr = reader.GetString(1);
                string batch = reader.GetString(2);

                if (!System.DateTime.TryParseExact(dateStr, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out System.DateTime dt))
                {
                    continue; // Skip tanggal invalid
                }

                if (!cache.ContainsKey(size))
                {
                    cache[size] = new System.Collections.Generic.List<TLJRecord>();
                }

                cache[size].Add(new TLJRecord
                {
                    Size = size,
                    ProdDate = dateStr,
                    ParsedDate = dt,
                    BatchNo = batch
                });
            }

            // Sort setiap list berdasarkan tanggal agar pencarian "tanggal sebelumnya" mudah
            foreach (var key in cache.Keys)
            {
                cache[key].Sort((a, b) => a.ParsedDate.CompareTo(b.ParsedDate));
            }

            return cache;
        }

        private string FindBatchInMemory(
            System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<TLJRecord>> cache,
            string size,
            string targetDateStr)
        {
            if (!cache.ContainsKey(size)) return string.Empty;

            if (!System.DateTime.TryParseExact(targetDateStr, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out System.DateTime targetDate))
            {
                return string.Empty;
            }

            var list = cache[size];

            // 1. Cari Tanggal Sama Persis (Reverse untuk ambil yang terbaru jika ada duplikat tanggal)
            for (int i = list.Count - 1; i >= 0; i--)
            {
                if (list[i].ParsedDate == targetDate)
                {
                    return ProcessRawBatchString(list[i].BatchNo);
                }
            }

            // 2. Cari Tanggal Sebelum Target (Paling dekat)
            // Karena sudah di-sort Ascending, kita cari dari belakang, yang pertama kali lebih kecil dari target
            for (int i = list.Count - 1; i >= 0; i--)
            {
                if (list[i].ParsedDate < targetDate)
                {
                    return ProcessRawBatchString(list[i].BatchNo);
                }
            }

            return string.Empty;
        }

        private string ProcessRawBatchString(string rawBatch)
        {
            if (string.IsNullOrEmpty(rawBatch)) return string.Empty;

            // Logic ExtractBatchNumbers yang lama dipindah kesini (in-memory)
            var batchList = new System.Collections.Generic.List<string>();
            string[] batches = rawBatch.Split(new[] { '\n', '\r' }, System.StringSplitOptions.RemoveEmptyEntries);
            foreach (string b in batches)
            {
                string t = b.Trim();
                if (!string.IsNullOrEmpty(t)) batchList.Add(t);
            }
            return System.String.Join("\n", batchList);
        }

        // --- EXISTING HELPERS (Unchanged Logic, just formatting) ---

        private string StandardizeDate(string rawDate, int expectedMonth, int expectedYear)
        {
            if (string.IsNullOrWhiteSpace(rawDate)) return string.Empty;

            if (System.DateTime.TryParse(rawDate, out System.DateTime parsedDate))
            {
                if (expectedYear > 2000 && parsedDate.Year != expectedYear)
                {
                    parsedDate = new System.DateTime(expectedYear, parsedDate.Month, parsedDate.Day);
                }

                if (expectedMonth > 0 && parsedDate.Month != expectedMonth)
                {
                    if (parsedDate.Day <= 12)
                    {
                        int newMonth = parsedDate.Day;
                        int newDay = parsedDate.Month;

                        if (newMonth == expectedMonth)
                        {
                            parsedDate = new System.DateTime(parsedDate.Year, newMonth, newDay);
                        }
                    }
                }

                return parsedDate.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            }

            return rawDate;
        }

        private int GetMonthNumber(string monthName)
        {
            if (string.IsNullOrWhiteSpace(monthName)) return 0;
            try
            {
                return System.DateTime.ParseExact(monthName, "MMMM", System.Globalization.CultureInfo.InvariantCulture).Month;
            }
            catch
            {
                return 0;
            }
        }

        private double GetMergedOrAverageValue(ClosedXML.Excel.IXLWorksheet sheet_YLB, int startRow, string columnLetter)
        {
            var cellFirst = sheet_YLB.Cell(startRow, columnLetter);
            if (cellFirst.IsMerged()) return ParseCustomDecimal(cellFirst.GetString());

            var val1 = ParseCustomDecimal(cellFirst.GetString());
            var val2 = ParseCustomDecimal(sheet_YLB.Cell(startRow + 1, columnLetter).GetString());

            if (val1 == 0) return val2;
            if (val2 == 0) return val1;

            return (val1 + val2) / 2.0;
        }

        private double ParseCustomDecimal(string rawInput)
        {
            if (string.IsNullOrWhiteSpace(rawInput)) return 0.0;
            string cleanInput = rawInput.Replace(",", ".").Trim();
            if (double.TryParse(cleanInput, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result)) return result;
            return 0.0;
        }

        private string NormalizeMonthFolder(string rawMonth)
        {
            if (string.IsNullOrWhiteSpace(rawMonth)) return string.Empty;
            for (int i = 0; i < rawMonth.Length; i++)
            {
                if (!char.IsDigit(rawMonth[i])) continue;
                int start = i;
                while (i < rawMonth.Length && char.IsDigit(rawMonth[i])) i++;
                string numberText = rawMonth.Substring(start, i - start);
                if (!int.TryParse(numberText, out int monthNumber)) continue;
                if (monthNumber < 1 || monthNumber > 12) continue;
                return new System.Globalization.DateTimeFormatInfo().GetMonthName(monthNumber);
            }
            return string.Empty;
        }

        private System.String CleanSizeText(System.String raw)
        {
            if (System.String.IsNullOrWhiteSpace(raw)) return System.String.Empty;

            System.String text = raw.ToUpper();

            int start = -1;
            for (int i = 0; i < text.Length; i++)
            {
                if (System.Char.IsDigit(text[i]))
                {
                    start = i;
                    break;
                }
            }

            if (start == -1) return System.String.Empty;

            System.String substring = text.Substring(start);

            int xIndex = substring.IndexOf('X');
            if (xIndex == -1) return System.String.Empty;

            if (xIndex + 1 >= substring.Length || !System.Char.IsDigit(substring[xIndex + 1]))
                return System.String.Empty;

            int end = xIndex + 1;
            while (end < substring.Length && System.Char.IsDigit(substring[end]))
                end++;

            System.String result = substring.Substring(0, end).Trim();

            System.String remaining = substring.Substring(end).Trim();

            System.String keyword = System.String.Empty;

            System.String cleanRemaining = System.String.Empty;
            for (int i = 0; i < remaining.Length; i++)
            {
                if (System.Char.IsLetterOrDigit(remaining[i]))
                {
                    cleanRemaining += remaining[i];
                }
            }

            if (cleanRemaining.Contains("FR"))
            {
                int frIndex = cleanRemaining.IndexOf("FR");
                if (frIndex >= 0)
                {
                    keyword = "FR";
                }
            }
            else if (!System.String.IsNullOrEmpty(cleanRemaining))
            {
                for (int i = 0; i < cleanRemaining.Length; i++)
                {
                    if (cleanRemaining[i] == 'B' && i + 1 < cleanRemaining.Length &&
                        System.Char.IsDigit(cleanRemaining[i + 1]))
                    {
                        int bStart = i;
                        int bEnd = i + 1;
                        while (bEnd < cleanRemaining.Length && System.Char.IsDigit(cleanRemaining[bEnd]))
                        {
                            bEnd++;
                        }

                        if (bEnd - bStart >= 2)
                        {
                            keyword = cleanRemaining.Substring(bStart, bEnd - bStart);
                            break;
                        }
                    }
                }
            }

            if (!System.String.IsNullOrEmpty(keyword))
            {
                result = result + " " + keyword;
            }

            return result.Trim();
        }

        private string DetermineTLJTable(string size_mm)
        {
            string cleanSize = size_mm.ToUpper().Replace(" ", "");

            int xIndex = cleanSize.IndexOf('X');
            if (xIndex == -1) return "TLJ500";

            string beforeX = cleanSize.Substring(0, xIndex);
            string afterX = cleanSize.Substring(xIndex + 1);

            string afterXDigits = "";
            for (int i = 0; i < afterX.Length; i++)
            {
                if (System.Char.IsDigit(afterX[i]))
                {
                    afterXDigits += afterX[i];
                }
                else
                {
                    break;
                }
            }

            if (int.TryParse(beforeX, out int firstDimension) &&
                int.TryParse(afterXDigits, out int secondDimension))
            {
                if (firstDimension <= 10 && secondDimension <= 100)
                {
                    return "TLJ350";
                }
            }

            return "TLJ500";
        }

        private void ResetCounters()
        {
            _totalFilesFound = 0;
            _totalRowsInserted = 0;
            _debugLog = "";
        }

        private void AppendDebug(string message)
        {
            if (_debugLog.Length < 1000) _debugLog += message + System.Environment.NewLine;
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