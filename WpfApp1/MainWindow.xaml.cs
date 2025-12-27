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
        private const int BatchSize = 100;

        private int _totalFilesFound;
        private int _totalRowsInserted;
        private string _debugLog = string.Empty; // FIX: Initialize with empty string

        // Optimasi: Cache untuk menghindari parsing berulang
        private System.Collections.Generic.Dictionary<string, int> _monthCache =
            new System.Collections.Generic.Dictionary<string, int>();

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

                using var connection = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={DbPath}");
                connection.Open();

                CreateBusbarTable(connection);

                using var transaction = connection.BeginTransaction();

                TraverseFoldersAndImport(connection, transaction);

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
            // FIX: Use null-conditional operator and check for null
            string? folder = System.IO.Path.GetDirectoryName(DbPath);

            if (!System.String.IsNullOrEmpty(folder) && !System.IO.Directory.Exists(folder))
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
                    Year_folder TEXT,
                    Month_folder TEXT,
                    Batch_no TEXT,
                    Prod_date TEXT,
                    Size_mm TEXT,
                    Thickness_mm REAL,
                    Width_mm REAL,
                    Length INTEGER,
                    Radius REAL,
                    Chamber_mm REAL,
                    Electric_IACS REAL,
                    Weight REAL,
                    Elongation REAL,
                    Tensile REAL,
                    Bend_test TEXT,
                    Spectro_Cu REAL,
                    Oxygen REAL
                );
            ";

            cmd.ExecuteNonQuery();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS TLJ500;

                CREATE TABLE IF NOT EXISTS TLJ500 (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT,
                    Month_folder TEXT,
                    Batch_no TEXT,
                    Prod_date TEXT,
                    Size_mm TEXT
                );
            ";

            cmd.ExecuteNonQuery();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS TLJ350;

                CREATE TABLE IF NOT EXISTS TLJ350 (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT,
                    Month_folder TEXT,
                    Batch_no TEXT,
                    Prod_date TEXT,
                    Size_mm TEXT
                );
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

            var yearDirs = System.IO.Directory.GetDirectories(ExcelRootFolder);

            foreach (string yearDir in yearDirs)
            {
                string year = new System.IO.DirectoryInfo(yearDir).Name.Trim();

                var monthDirs = System.IO.Directory.GetDirectories(yearDir);

                foreach (string monthDir in monthDirs)
                {
                    string rawMonth = new System.IO.DirectoryInfo(monthDir).Name.Trim();
                    string normalizedMonth = NormalizeMonthFolder(rawMonth);

                    var files = System.IO.Directory.GetFiles(monthDir, "*.xlsx");

                    ProcessFileBatch(connection, transaction, files, year, normalizedMonth);
                }
            }
        }

        private void ProcessFileBatch(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            System.String[] files,
            string year,
            string month)
        {
            for (int i = 0; i < files.Length; i++)
            {
                string file = files[i];
                string fileName = System.IO.Path.GetFileName(file);

                if (fileName.StartsWith("~$"))
                    continue;

                _totalFilesFound++;

                ProcessSingleExcelFile(connection, transaction, file, year, month);

                if ((i + 1) % 10 == 0)
                {
                    AppendDebug($"Diproses: {i + 1} dari {files.Length} file di {year}/{month}");
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
            // FIX: Simplified - remove problematic LoadOptions
            using var workbook = new ClosedXML.Excel.XLWorkbook(filePath);

            int folderMonthNum = GetMonthNumberCached(month);
            int folderYearNum = 0;
            int.TryParse(year, out folderYearNum);

            // --- Process YLB Sheet ---
            try
            {
                var sheet_YLB = workbook.Worksheets
                    .FirstOrDefault(w =>
                        w.Name.Trim().Equals(
                            "YLB 50",
                            System.StringComparison.OrdinalIgnoreCase));

                if (sheet_YLB == null)
                {
                    AppendDebug($"SKIP: Sheet 'YLB 50' tidak ditemukan -> {System.IO.Path.GetFileName(filePath)}");
                }
                else
                {
                    var ylbBatch = new System.Collections.Generic.List<object[]>();
                    string currentProdDate = string.Empty;
                    int row = 3;

                    while (true)
                    {
                        string sizeValue_YLB = sheet_YLB.Cell(row, "C").GetString();
                        if (string.IsNullOrWhiteSpace(sizeValue_YLB))
                            break;

                        string rawDateFromCell = sheet_YLB.Cell(row, "B").GetString().Trim();

                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        string cleanSize_YLB = CleanSizeText(sizeValue_YLB);

                        var values = ExtractYLBValues(sheet_YLB, row);

                        ylbBatch.Add(new object[]
                        {
                            cleanSize_YLB, year, month, currentProdDate,
                            values.Thickness, values.Width, values.Length,
                            values.Radius, values.Chamber, values.Electric,
                            values.Resistivity, values.Elongation, values.Tensile,
                            values.BendTest, values.Spectro, values.Oxygen
                        });

                        if (ylbBatch.Count >= BatchSize)
                        {
                            BatchInsertBusbarRows(connection, transaction, ylbBatch);
                            ylbBatch.Clear();
                        }

                        _totalRowsInserted++;
                        row += 2;
                    }

                    if (ylbBatch.Count > 0)
                    {
                        BatchInsertBusbarRows(connection, transaction, ylbBatch);
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (YLB): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }

            // --- Process TLJ 350 Sheet ---
            try
            {
                var sheet_TLJ350 = workbook.Worksheets.FirstOrDefault(w => w.Name.Trim().Equals("TLJ 350", System.StringComparison.OrdinalIgnoreCase));
                if (sheet_TLJ350 == null)
                {
                    AppendDebug($"SKIP: Sheet 'TLJ350' tidak ditemukan -> {System.IO.Path.GetFileName(filePath)}");
                }
                else
                {
                    var tlj350Batch = new System.Collections.Generic.List<object[]>();
                    string currentProdDate = string.Empty;
                    int row = 3;

                    while (true)
                    {
                        string sizeValue_TLJ350 = sheet_TLJ350.Cell(row, "D").GetString();

                        if (string.IsNullOrWhiteSpace(sizeValue_TLJ350))
                            break;

                        string rawDateFromCell = sheet_TLJ350.Cell(row, "B").GetString().Trim();

                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        string cleanSize_TLJ350 = CleanSizeText(sizeValue_TLJ350);
                        string batchValue = sheet_TLJ350.Cell(row, "C").GetString();

                        tlj350Batch.Add(new object[]
                        {
                            cleanSize_TLJ350, year, month, currentProdDate, batchValue
                        });

                        if (tlj350Batch.Count >= BatchSize)
                        {
                            BatchInsertTLJ350Rows(connection, transaction, tlj350Batch);
                            tlj350Batch.Clear();
                        }

                        _totalRowsInserted++;
                        row += 2;
                    }

                    if (tlj350Batch.Count > 0)
                    {
                        BatchInsertTLJ350Rows(connection, transaction, tlj350Batch);
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (TLJ350): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }

            // --- Process TLJ 500 Sheet ---
            try
            {
                var sheet_TLJ500 = workbook.Worksheets.FirstOrDefault(w => w.Name.Trim().Equals("TLJ 500", System.StringComparison.OrdinalIgnoreCase));
                if (sheet_TLJ500 == null)
                {
                    AppendDebug($"SKIP: Sheet 'TLJ500' tidak ditemukan -> {System.IO.Path.GetFileName(filePath)}");
                }
                else
                {
                    var tlj500Batch = new System.Collections.Generic.List<object[]>();
                    string currentProdDate = string.Empty;
                    int row = 3;

                    while (true)
                    {
                        string sizeValue_TLJ500 = sheet_TLJ500.Cell(row, "D").GetString();

                        if (string.IsNullOrWhiteSpace(sizeValue_TLJ500))
                            break;

                        string rawDateFromCell = sheet_TLJ500.Cell(row, "B").GetString().Trim();

                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        string cleanSize_TLJ500 = CleanSizeText(sizeValue_TLJ500);
                        string batchValue = sheet_TLJ500.Cell(row, "C").GetString();

                        tlj500Batch.Add(new object[]
                        {
                            cleanSize_TLJ500, year, month, currentProdDate, batchValue
                        });

                        if (tlj500Batch.Count >= BatchSize)
                        {
                            BatchInsertTLJ500Rows(connection, transaction, tlj500Batch);
                            tlj500Batch.Clear();
                        }

                        _totalRowsInserted++;
                        row += 2;
                    }

                    if (tlj500Batch.Count > 0)
                    {
                        BatchInsertTLJ500Rows(connection, transaction, tlj500Batch);
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (TLJ500): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }
        }

        private (double Thickness, double Width, double Length, double Radius,
                 double Chamber, double Electric, double Resistivity,
                 double Elongation, double Tensile, string BendTest,
                 double Spectro, double Oxygen) ExtractYLBValues(ClosedXML.Excel.IXLWorksheet sheet, int row)
        {
            double rawThickness = ParseCustomDecimal(sheet.Cell(row, "G").GetString());
            double valThickness = System.Math.Round(rawThickness, 2);

            double rawWidth = ParseCustomDecimal(sheet.Cell(row, "I").GetString());
            double valWidth = System.Math.Round(rawWidth, 2);

            double rawRadius = ParseCustomDecimal(sheet.Cell(row, "J").GetString());
            double valRadius = System.Math.Round(rawRadius, 2);

            double rawChamber = ParseCustomDecimal(sheet.Cell(row, "L").GetString());
            double valChamber = System.Math.Round(rawChamber, 2);

            double rawElectric = ParseCustomDecimal(sheet.Cell(row, "U").GetString());
            double valElectric = System.Math.Round(rawElectric, 2);

            double rawOxygen = ParseCustomDecimal(sheet.Cell(row, "X").GetString());
            double valOxygen = System.Math.Round(rawOxygen, 2);

            double valSpectro = ParseCustomDecimal(sheet.Cell(row, "Y").GetString());
            double valResistivity = ParseCustomDecimal(sheet.Cell(row, "T").GetString());

            double rawLength = ParseCustomDecimal(sheet.Cell(row, "K").GetString());
            double valLength = System.Math.Round(rawLength, 0);

            double rawElongation = GetMergedOrAverageValue(sheet, row, "R");
            double valElongation = System.Math.Round(rawElongation, 2);

            double rawTensile = GetMergedOrAverageValue(sheet, row, "Q");
            double valTensile = System.Math.Round(rawTensile, 2);

            string valBendTest = sheet.Cell(row, "W").GetString();

            return (valThickness, valWidth, valLength, valRadius, valChamber,
                    valElectric, valResistivity, valElongation, valTensile,
                    valBendTest, valSpectro, valOxygen);
        }

        private void BatchInsertBusbarRows(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            System.Collections.Generic.List<object[]> batchData)
        {
            if (batchData.Count == 0) return;

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            var parameters = new System.Text.StringBuilder();
            var paramList = new System.Collections.Generic.List<Microsoft.Data.Sqlite.SqliteParameter>();

            for (int i = 0; i < batchData.Count; i++)
            {
                if (i > 0) parameters.Append(",\n");
                parameters.Append($"(@Size{i}, @Year{i}, @Month{i}, @ProdDate{i}, " +
                                 $"@Thickness{i}, @Width{i}, @Length{i}, @Radius{i}, @Chamber{i}, " +
                                 $"@Electric{i}, @Resistivity{i}, @Elongation{i}, @Tensile{i}, " +
                                 $"@Bend{i}, @Spectro{i}, @Oxygen{i})");

                var row = batchData[i];

                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Size{i}", row[0]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Year{i}", row[1]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Month{i}", row[2]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@ProdDate{i}", row[3]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Thickness{i}", row[4]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Width{i}", row[5]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Length{i}", row[6]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Radius{i}", row[7]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Chamber{i}", row[8]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Electric{i}", row[9]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Resistivity{i}", row[10]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Elongation{i}", row[11]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Tensile{i}", row[12]));

                object bendVal = (row[13] == null || System.String.IsNullOrEmpty(row[13] as string)) ?
                    System.DBNull.Value : row[13];
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Bend{i}", bendVal));

                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Spectro{i}", row[14]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Oxygen{i}", row[15]));
            }

            cmd.CommandText = $@"
                INSERT INTO Busbar (
                    Size_mm, Year_folder, Month_folder, Prod_date, 
                    Thickness_mm, Width_mm, Length, Radius, Chamber_mm,
                    Electric_IACS, Weight, Elongation, Tensile,
                    Bend_test, Spectro_Cu, Oxygen
                ) VALUES {parameters}";

            cmd.Parameters.AddRange(paramList.ToArray());
            cmd.ExecuteNonQuery();
        }

        private void BatchInsertTLJ350Rows(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            System.Collections.Generic.List<object[]> batchData)
        {
            if (batchData.Count == 0) return;

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            var parameters = new System.Text.StringBuilder();
            var paramList = new System.Collections.Generic.List<Microsoft.Data.Sqlite.SqliteParameter>();

            for (int i = 0; i < batchData.Count; i++)
            {
                if (i > 0) parameters.Append(",\n");
                parameters.Append($"(@Size{i}, @Year{i}, @Month{i}, @ProdDate{i}, @Batch{i})");

                var row = batchData[i];

                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Size{i}", row[0]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Year{i}", row[1]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Month{i}", row[2]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@ProdDate{i}", row[3]));

                object batchVal = (row[4] == null || System.String.IsNullOrEmpty(row[4] as string)) ?
                    System.DBNull.Value : row[4];
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Batch{i}", batchVal));
            }

            cmd.CommandText = $@"
                INSERT INTO TLJ350 (Size_mm, Year_folder, Month_folder, Prod_date, Batch_no)
                VALUES {parameters}";

            cmd.Parameters.AddRange(paramList.ToArray());
            cmd.ExecuteNonQuery();
        }

        private void BatchInsertTLJ500Rows(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            System.Collections.Generic.List<object[]> batchData)
        {
            if (batchData.Count == 0) return;

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            var parameters = new System.Text.StringBuilder();
            var paramList = new System.Collections.Generic.List<Microsoft.Data.Sqlite.SqliteParameter>();

            for (int i = 0; i < batchData.Count; i++)
            {
                if (i > 0) parameters.Append(",\n");
                parameters.Append($"(@Size{i}, @Year{i}, @Month{i}, @ProdDate{i}, @Batch{i})");

                var row = batchData[i];

                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Size{i}", row[0]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Year{i}", row[1]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Month{i}", row[2]));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@ProdDate{i}", row[3]));

                object batchVal = (row[4] == null || System.String.IsNullOrEmpty(row[4] as string)) ?
                    System.DBNull.Value : row[4];
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Batch{i}", batchVal));
            }

            cmd.CommandText = $@"
                INSERT INTO TLJ500 (Size_mm, Year_folder, Month_folder, Prod_date, Batch_no)
                VALUES {parameters}";

            cmd.Parameters.AddRange(paramList.ToArray());
            cmd.ExecuteNonQuery();
        }

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

        private int GetMonthNumberCached(string monthName)
        {
            if (string.IsNullOrWhiteSpace(monthName)) return 0;

            if (_monthCache.TryGetValue(monthName, out int cachedMonth))
                return cachedMonth;

            try
            {
                int month = System.DateTime.ParseExact(monthName, "MMMM", System.Globalization.CultureInfo.InvariantCulture).Month;
                _monthCache[monthName] = month;
                return month;
            }
            catch
            {
                return 0;
            }
        }

        private int GetMonthNumber(string monthName)
        {
            return GetMonthNumberCached(monthName);
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

        private void UpdateBusbarBatchNumbers(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            try
            {
                using var selectBusbarCmd = connection.CreateCommand();
                selectBusbarCmd.Transaction = transaction;
                selectBusbarCmd.CommandText = @"
                    SELECT Id, Size_mm, Prod_date 
                    FROM Busbar 
                    WHERE (Batch_no IS NULL OR Batch_no = '')
                    ORDER BY Prod_date, Id
                ";

                using var busbarReader = selectBusbarCmd.ExecuteReader();

                var updateBatch = new System.Collections.Generic.List<System.Tuple<int, string>>();

                while (busbarReader.Read())
                {
                    int busbarId = busbarReader.GetInt32(0);
                    string size_mm = busbarReader.GetString(1);
                    string prod_date = busbarReader.GetString(2);

                    string targetTable = DetermineTLJTable(size_mm);
                    string batchNumbers = FindBatchNumbers(connection, transaction, targetTable, size_mm, prod_date);

                    if (!System.String.IsNullOrEmpty(batchNumbers))
                    {
                        updateBatch.Add(System.Tuple.Create(busbarId, batchNumbers));

                        if (updateBatch.Count >= 50)
                        {
                            BatchUpdateBusbarBatches(connection, transaction, updateBatch);
                            updateBatch.Clear();
                        }
                    }
                }

                if (updateBatch.Count > 0)
                {
                    BatchUpdateBusbarBatches(connection, transaction, updateBatch);
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR UpdateBusbarBatchNumbers: {ex.Message}");
                throw;
            }
        }

        private void BatchUpdateBusbarBatches(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            System.Collections.Generic.List<System.Tuple<int, string>> batchData)
        {
            if (batchData.Count == 0) return;

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            var updateCases = new System.Text.StringBuilder();
            var paramList = new System.Collections.Generic.List<Microsoft.Data.Sqlite.SqliteParameter>();

            for (int i = 0; i < batchData.Count; i++)
            {
                var item = batchData[i];
                updateCases.Append($"WHEN @Id{i} THEN @Batch{i} ");
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Id{i}", item.Item1));
                paramList.Add(new Microsoft.Data.Sqlite.SqliteParameter($"@Batch{i}", item.Item2));
            }

            var idParams = System.String.Join(",", System.Linq.Enumerable.Range(0, batchData.Count).Select(i => $"@Id{i}"));

            cmd.CommandText = $@"
                UPDATE Busbar 
                SET Batch_no = CASE Id 
                    {updateCases.ToString()}
                    ELSE Batch_no 
                END
                WHERE Id IN ({idParams})";

            cmd.Parameters.AddRange(paramList.ToArray());
            cmd.ExecuteNonQuery();
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

        private string FindBatchNumbers(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string tableName,
            string size_mm,
            string targetDate)
        {
            try
            {
                System.DateTime targetDateTime;
                if (!System.DateTime.TryParseExact(
                    targetDate,
                    "dd/MM/yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out targetDateTime))
                {
                    if (!System.DateTime.TryParse(targetDate, out targetDateTime))
                    {
                        return string.Empty;
                    }
                }

                using var cmdSameDate = connection.CreateCommand();
                cmdSameDate.Transaction = transaction;
                cmdSameDate.CommandText = $@"
                    SELECT Batch_no, Prod_date 
                    FROM {tableName} 
                    WHERE Size_mm = @Size_mm 
                      AND Prod_date = @TargetDate
                    ORDER BY Prod_date DESC
                ";

                cmdSameDate.Parameters.AddWithValue("@Size_mm", size_mm);
                cmdSameDate.Parameters.AddWithValue("@TargetDate", targetDate);

                using var readerSameDate = cmdSameDate.ExecuteReader();
                if (readerSameDate.HasRows)
                {
                    return ExtractBatchNumbers(readerSameDate);
                }

                using var cmdBeforeDate = connection.CreateCommand();
                cmdBeforeDate.Transaction = transaction;
                cmdBeforeDate.CommandText = $@"
                    SELECT Batch_no, Prod_date 
                    FROM {tableName} 
                    WHERE Size_mm = @Size_mm 
                      AND Prod_date < @TargetDate
                    ORDER BY Prod_date DESC
                    LIMIT 1
                ";

                cmdBeforeDate.Parameters.AddWithValue("@Size_mm", size_mm);
                cmdBeforeDate.Parameters.AddWithValue("@TargetDate", targetDate);

                using var readerBeforeDate = cmdBeforeDate.ExecuteReader();
                if (readerBeforeDate.HasRows)
                {
                    return ExtractBatchNumbers(readerBeforeDate);
                }

                return string.Empty;
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FindBatchNumbers for {size_mm} on {targetDate}: {ex.Message}");
                return string.Empty;
            }
        }

        private string ExtractBatchNumbers(Microsoft.Data.Sqlite.SqliteDataReader reader)
        {
            System.Collections.Generic.List<string> batchList = new System.Collections.Generic.List<string>();

            while (reader.Read())
            {
                if (!reader.IsDBNull(0))
                {
                    string batchData = reader.GetString(0);

                    if (!System.String.IsNullOrEmpty(batchData))
                    {
                        string[] batches = batchData.Split(
                            new[] { '\n', '\r' },
                            System.StringSplitOptions.RemoveEmptyEntries
                        );

                        foreach (string batch in batches)
                        {
                            string trimmedBatch = batch.Trim();
                            if (!System.String.IsNullOrEmpty(trimmedBatch))
                            {
                                batchList.Add(trimmedBatch);
                            }
                        }
                    }
                }
            }

            return System.String.Join("\n", batchList);
        }

        private void ResetCounters()
        {
            _totalFilesFound = 0;
            _totalRowsInserted = 0;
            _debugLog = string.Empty;
            _monthCache.Clear();
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